#!/usr/bin/env python3
"""Coleta e consolida os dados do mercado secundário de debêntures da ANBIMA."""

import datetime as dt
import os
import tempfile
from io import StringIO
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from bizdays import Calendar

BASE_URL = "https://www.anbima.com.br/informacoes/merc-sec-debentures/arqs/db"
DOWNLOADS = Path("downloads/debentures")
BASE = Path("bases/debentures_base.xlsx")
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; coleta-debentures/1.0)"}


def salvar_excel(df, path, sheet_name):
    path.parent.mkdir(parents=True, exist_ok=True)
    if len(df) + 1 > 1_048_576 or len(df.columns) > 16_384:
        raise ValueError(f"{path}: dados excedem os limites do Excel")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", dir=path.parent, delete=False) as temp:
        temp_path = Path(temp.name)
    try:
        with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            for row in writer.sheets[sheet_name].iter_rows():
                for cell in row:
                    if cell.data_type == "f":
                        cell.data_type = "s"
        os.replace(temp_path, path)
    finally:
        temp_path.unlink(missing_ok=True)


def baixar_data(data, path):
    url = f"{BASE_URL}{data:%y%m%d}.txt"
    response = requests.get(url, headers=HEADERS, timeout=60)
    if response.status_code == 404:
        print(f"{data}: arquivo ainda não disponível.")
        return False
    response.raise_for_status()
    content = response.content.decode("latin1")
    first_line = content.splitlines()[0] if content else ""
    if not content.strip() or "não há dados disponíveis" in first_line.lower():
        print(f"{data}: sem dados disponíveis.")
        return False
    if "<" in first_line or "error" in first_line.lower():
        raise ValueError(f"{data}: resposta inesperada da ANBIMA")

    df = pd.read_csv(StringIO(content), sep="@", skiprows=2)
    df = df.dropna(how="all")
    if df.empty or len(df.columns) < 2:
        raise ValueError(f"{data}: arquivo sem registros válidos")
    df.insert(0, "dt_referencia", data.isoformat())
    salvar_excel(df, path, "Debentures")
    print(f"Salvo: {path}")
    return True


def consolidar(downloads, destino):
    arquivos = sorted(downloads.glob("*.xlsx"))
    if not arquivos:
        print("Nenhum arquivo disponível para consolidação.")
        return
    frames = [pd.read_excel(arquivo) for arquivo in arquivos]
    if any("dt_referencia" not in frame.columns for frame in frames):
        raise ValueError("Arquivo individual sem dt_referencia")
    consolidado = pd.concat(frames, ignore_index=True)
    consolidado = consolidado.sort_values("dt_referencia", kind="stable").drop_duplicates()
    salvar_excel(consolidado, destino, "Base_Consolidada")
    print(f"Base consolidada: {destino} ({len(consolidado)} linhas)")


def main():
    hoje = dt.datetime.now(ZoneInfo("America/Sao_Paulo")).date()
    calendario = Calendar.load(name="ANBIMA")
    inicio = pd.Timestamp(calendario.offset(hoje, -6)).date()
    DOWNLOADS.mkdir(parents=True, exist_ok=True)
    baixou = False
    erros = []
    for dias in range((hoje - inicio).days + 1):
        data = inicio + dt.timedelta(days=dias)
        path = DOWNLOADS / f"{data:%y%m%d}.xlsx"
        if not calendario.isbizday(data) or path.exists():
            continue
        try:
            baixou = baixar_data(data, path) or baixou
        except (requests.RequestException, ValueError, OSError, pd.errors.ParserError) as exc:
            print(f"Erro ao coletar {data}: {exc}")
            erros.append(data)

    if baixou or not BASE.exists():
        consolidar(DOWNLOADS, BASE)
    if erros:
        raise SystemExit(f"Falha ao coletar {len(erros)} data(s); tente novamente.")


if __name__ == "__main__":
    main()
