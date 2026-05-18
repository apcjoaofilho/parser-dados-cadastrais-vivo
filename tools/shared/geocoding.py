#!/usr/bin/env python3
"""
Módulo de geocodificação usando Nominatim (OpenStreetMap) com cache SQLite.
"""

import hashlib
import logging
import sqlite3
import time
from urllib.parse import quote

import requests

logger = logging.getLogger(__name__)

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
USER_AGENT = "VivoParser/1.0 (parser_vivo.py)"
RATE_LIMIT_SECONDS = 1


def _normalizar_endereco(endereco: str) -> str:
    """Limpa endereços mal formatados comuns nos dados Vivo."""
    e = endereco.strip()
    # Remove duplicações de tipo de logradouro
    fixes = [
        ("Rua R ", "Rua "),
        ("Rua AVENIDA ", "Avenida "),
        ("Rua ALAMEDA ", "Alameda "),
        ("Rua TRAVESSA ", "Travessa "),
        ("Rua RUA ", "Rua "),
        ("Rua EST ", "Estrada "),
        ("Rua ESTRADA ", "Estrada "),
    ]
    for bad, good in fixes:
        if e.upper().startswith(bad):
            e = good + e[len(bad):]
            break
    # Remove múltiplos espaços
    while "  " in e:
        e = e.replace("  ", " ")
    return e


def _hash_endereco(endereco_completo: str) -> str:
    """Gera hash MD5 do endereço para uso como chave de cache."""
    return hashlib.md5(endereco_completo.encode("utf-8")).hexdigest()


def criar_tabela_cache(conn: sqlite3.Connection) -> None:
    """
    Cria tabela de cache de geocodificação se não existir.

    AVISO DE SEGURANÇA: Esta tabela armazena `endereco_completo` em texto
    plano (PII). O arquivo `.db` que a contém deve ser protegido como
    qualquer outro arquivo com dados pessoais.
    """
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS geocoding_cache (
            endereco_hash TEXT PRIMARY KEY,
            endereco_completo TEXT NOT NULL,
            latitude REAL,
            longitude REAL,
            google_maps_url TEXT,
            data_consulta TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    conn.commit()


def consultar_cache(conn: sqlite3.Connection, endereco_hash: str) -> dict | None:
    """Consulta cache por hash de endereço. Retorna dict ou None."""
    cursor = conn.execute(
        "SELECT latitude, longitude, google_maps_url FROM geocoding_cache WHERE endereco_hash = ?",
        (endereco_hash,),
    )
    row = cursor.fetchone()
    if row:
        return {"latitude": row[0], "longitude": row[1], "google_maps_url": row[2]}
    return None


def salvar_cache(
    conn: sqlite3.Connection,
    endereco_hash: str,
    endereco_completo: str,
    latitude: float | None,
    longitude: float | None,
    google_maps_url: str | None,
) -> None:
    """Salva resultado de geocodificação no cache."""
    conn.execute(
        """
        INSERT INTO geocoding_cache (endereco_hash, endereco_completo, latitude, longitude, google_maps_url)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(endereco_hash) DO UPDATE SET
            latitude=excluded.latitude,
            longitude=excluded.longitude,
            google_maps_url=excluded.google_maps_url,
            data_consulta=CURRENT_TIMESTAMP
        """,
        (endereco_hash, endereco_completo, latitude, longitude, google_maps_url),
    )
    conn.commit()


def _extrair_logradouro_sem_numero(endereco: str) -> str:
    """Remove número e complementos óbvios do endereço."""
    import re
    # Remove tudo após o primeiro número de 1-5 dígitos (provavelmente número da casa)
    # mas mantém nomes que contenham números como parte do logradouro (ex: "25 de Março")
    m = re.search(r"\s+\d{1,5}(?:\s+|$)", endereco)
    if m:
        return endereco[:m.start()].strip()
    return endereco


def geocodificar_endereco(endereco_completo: str) -> tuple[float | None, float | None]:
    """
    Consulta Nominatim para obter coordenadas do endereço.
    Tenta variações do endereço (completo → sem número → cidade/estado).
    Retorna (latitude, longitude) ou (None, None) em caso de falha.
    """
    headers = {"User-Agent": USER_AGENT}
    variacoes = [endereco_completo]

    # Tenta sem número
    partes = endereco_completo.split(", ")
    if len(partes) >= 5:
        logradouro_limpo = _extrair_logradouro_sem_numero(partes[0])
        if logradouro_limpo and logradouro_limpo != partes[0]:
            variacoes.append(", ".join([logradouro_limpo] + partes[1:]))
        # Tenta logradouro limpo + cidade + estado + Brazil (sem bairro/CEP)
        variacoes.append(f"{logradouro_limpo or partes[0]}, {partes[2]}, {partes[3]}, Brazil")
        # Tenta apenas cidade/estado/país
        variacoes.append(f"{partes[2]}, {partes[3]}, Brazil")

    for tentativa in variacoes:
        params = {"q": tentativa, "format": "json", "limit": 1}
        try:
            response = requests.get(NOMINATIM_URL, params=params, headers=headers, timeout=30)
            response.raise_for_status()
            data = response.json()
            if data:
                lat = float(data[0]["lat"])
                lon = float(data[0]["lon"])
                return lat, lon
        except Exception as e:
            logger.debug("Falha na variação '%s...': %s", tentativa[:50], e)
            continue

    logger.warning("Nenhuma coordenada encontrada para '%s...'", endereco_completo[:50])
    return None, None


def _google_maps_url(lat: float, lon: float) -> str:
    """Gera URL de direções do Google Maps para as coordenadas."""
    return f"https://www.google.com/maps/dir/?api=1&destination={lat},{lon}"


def processar_geocoding(df, db_path: str) -> None:
    """
    Adiciona colunas latitude, longitude e google_maps_url ao DataFrame.
    Usa cache SQLite em db_path para evitar requisições duplicadas.
    Modifica o DataFrame in-place.
    """
    conn = sqlite3.connect(db_path)
    criar_tabela_cache(conn)

    latitudes = []
    longitudes = []
    urls = []

    ultima_req = 0

    for _, row in df.iterrows():
        endereco_limpo = _normalizar_endereco(str(row["endereco"]))
        endereco_completo = f"{endereco_limpo}, {row['bairro']}, {row['municipio']}, {row['estado']}, {row['cep']}, Brazil"
        endereco_hash = _hash_endereco(endereco_completo)

        cached = consultar_cache(conn, endereco_hash)
        if cached:
            latitudes.append(cached["latitude"])
            longitudes.append(cached["longitude"])
            urls.append(cached["google_maps_url"])
            continue

        # Rate limit: apenas após cache miss
        agora = time.time()
        elapsed = agora - ultima_req
        if elapsed < RATE_LIMIT_SECONDS:
            time.sleep(RATE_LIMIT_SECONDS - elapsed)

        lat, lon = geocodificar_endereco(endereco_completo)
        ultima_req = time.time()

        url = _google_maps_url(lat, lon) if lat is not None and lon is not None else None

        salvar_cache(conn, endereco_hash, endereco_completo, lat, lon, url)

        latitudes.append(lat)
        longitudes.append(lon)
        urls.append(url)

    conn.close()

    df["latitude"] = latitudes
    df["longitude"] = longitudes
    df["google_maps_url"] = urls
    logger.info("Geocodificação concluída. %d registros processados.", len(df))
