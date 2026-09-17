from __future__ import annotations

import io
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

TZ = ZoneInfo("America/Sao_Paulo")
REQUIRED_COLUMNS = ["Local", "Nome Fantasia", "Vencimento", "Valor Vencido"]


def _norm(value: str) -> str:
    text = str(value or "").strip().lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    return text


def _canonicalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    aliases = {
        "liquidacao": "Liquidação",
        "local": "Local",
        "valor titulo": "Valor Título",
        "cliente": "Cliente",
        "nome fantasia": "Nome Fantasia",
        "grupo": "Grupo",
        "classificacao": "Classificação",
        "observacao": "Observação",
        "representante": "Representante",
        "descricao conta": "Descrição Conta",
        "num titulo boleto": "Num. Título Boleto",
        "nota": "Nota",
        "contrato": "Contrato",
        "fatura": "Fatura",
        "pedido": "Pedido",
        "os": "OS",
        "telefone": "Telefone",
        "celular": "Celular",
        "documento": "Documento",
        "tipo documento": "Tipo Documento",
        "emissao": "Emissão",
        "vencimento": "Vencimento",
        "valor recebido": "Valor Recebido",
        "valor vencido": "Valor Vencido",
        "valor a vencer": "Valor a Vencer",
        "multas devidas": "Multas Devidas",
        "juros devidos": "Juros Devidos",
        "vencido com juros": "Vencido com Juros",
    }
    mapping = {}
    for col in df.columns:
        mapping[col] = aliases.get(_norm(col), str(col).strip())
    return df.rename(columns=mapping)


def _to_number(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)

    def parse(v):
        if pd.isna(v):
            return 0.0
        s = str(v).strip().replace("R$", "").replace(" ", "")
        if not s:
            return 0.0
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except Exception:
            return 0.0

    return series.map(parse).astype(float)


def _has_positive_number(value) -> bool:
    try:
        return pd.notna(value) and float(value) > 0
    except Exception:
        return False


def prepare_receivables(df: pd.DataFrame, today=None) -> pd.DataFrame:
    out = _canonicalize_columns(df.copy())
    missing = [c for c in REQUIRED_COLUMNS if c not in out.columns]
    if missing:
        raise ValueError(
            "A planilha de inadimplência não possui as colunas necessárias. Faltando: "
            + ", ".join(missing)
        )

    optional = [
        "Liquidação", "Valor Título", "Cliente", "Grupo", "Classificação", "Observação",
        "Representante", "Descrição Conta", "Num. Título Boleto", "Nota", "Contrato",
        "Fatura", "Pedido", "OS", "Telefone", "Celular", "Documento", "Tipo Documento",
        "Emissão", "Valor Recebido", "Valor a Vencer", "Multas Devidas", "Juros Devidos",
        "Vencido com Juros",
    ]
    for col in optional:
        if col not in out.columns:
            out[col] = ""

    text_cols = [
        "Local", "Nome Fantasia", "Grupo", "Classificação", "Observação", "Representante",
        "Descrição Conta", "Num. Título Boleto", "Contrato", "Telefone", "Celular",
        "Documento", "Tipo Documento",
    ]
    for col in text_cols:
        out[col] = out[col].fillna("").astype(str).str.strip()

    for col in [
        "Valor Título", "Valor Recebido", "Valor Vencido", "Valor a Vencer",
        "Multas Devidas", "Juros Devidos", "Vencido com Juros", "Fatura", "Pedido", "Nota", "OS",
    ]:
        out[col] = _to_number(out[col])

    out["Vencimento"] = pd.to_datetime(out["Vencimento"], errors="coerce", dayfirst=True).dt.normalize()
    out["Emissão"] = pd.to_datetime(out["Emissão"], errors="coerce", dayfirst=True).dt.normalize()

    if today is None:
        today = datetime.now(TZ).date()
    today_ts = pd.Timestamp(today)
    out["Dias em atraso"] = (today_ts - out["Vencimento"]).dt.days.astype("Int64")

    # O MonitorReceber traz em Valor Vencido o saldo que continua em aberto.
    out["Em aberto"] = out["Valor Vencido"].clip(lower=0)
    out["É inadimplente"] = (
        out["Vencimento"].notna()
        & (out["Dias em atraso"] > 0)
        & (out["Em aberto"] > 0)
    )

    # Origem operacional: Pedido preenchido caracteriza venda; contrato INTERM é intermediação;
    # demais contratos são locação. Sem pedido/contrato fica sem identificação.
    def origem(row) -> str:
        pedido = row.get("Pedido", 0)
        contrato = str(row.get("Contrato", "") or "").strip()
        contrato_norm = _norm(contrato)
        if _has_positive_number(pedido):
            return "Venda"
        if contrato_norm.startswith("interm"):
            return "Intermediação"
        if contrato:
            return "Locação"
        return "Não identificado"

    out["Origem"] = out.apply(origem, axis=1)

    # Faturamento: considera número de Fatura ou marca DATA FAT na observação.
    # Isso deixa visíveis os casos que ainda não têm evidência de faturamento no MonitorReceber.
    obs_norm = out["Observação"].fillna("").astype(str).map(_norm)
    has_fatura = out["Fatura"].map(_has_positive_number)
    has_data_fat = obs_norm.str.contains("data fat", na=False)
    out["Status faturamento"] = (has_fatura | has_data_fat).map({True: "Faturado", False: "Não faturado"})

    def bucket(days):
        if pd.isna(days) or days <= 0:
            return "Não vencido"
        days = int(days)
        if days == 1:
            return "24h"
        if days == 2:
            return "48h"
        if days <= 7:
            return "72h a 7 dias"
        if days <= 15:
            return "8 a 15 dias"
        if days <= 30:
            return "16 a 30 dias"
        if days <= 60:
            return "31 a 60 dias"
        if days <= 90:
            return "61 a 90 dias"
        if days <= 120:
            return "91 a 120 dias"
        return "121 dias +"

    out["Faixa atraso"] = out["Dias em atraso"].map(bucket)
    return out


def read_receivables_bytes(file_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0)
    return prepare_receivables(raw)


def read_receivables_path(path: str | Path) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=0)
    return prepare_receivables(raw)


def build_receivables_export(df: pd.DataFrame) -> bytes:
    cols = [
        "Local", "Nome Fantasia", "Cliente", "Origem", "Status faturamento", "Contrato", "Pedido",
        "Fatura", "Documento", "Tipo Documento", "Emissão", "Vencimento", "Dias em atraso", "Faixa atraso",
        "Valor Título", "Valor Recebido", "Em aberto", "Multas Devidas", "Juros Devidos",
        "Vencido com Juros", "Telefone", "Celular", "Representante", "Observação",
    ]
    export = df[[c for c in cols if c in df.columns]].copy()
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        export.to_excel(writer, index=False, sheet_name="Inadimplencia")
        wb = writer.book
        ws = writer.sheets["Inadimplencia"]
        header = wb.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#2F4EA2", "align": "center"})
        money = wb.add_format({"num_format": 'R$ #,##0.00'})
        date_fmt = wb.add_format({"num_format": "dd/mm/yyyy"})
        for i, col in enumerate(export.columns):
            ws.write(0, i, col, header)
            sample = export[col].astype(str).head(100)
            width = min(max([len(col)] + [len(v) for v in sample]) + 2, 38)
            fmt = None
            if col in ["Valor Título", "Valor Recebido", "Em aberto", "Multas Devidas", "Juros Devidos", "Vencido com Juros"]:
                fmt = money
            elif col in ["Emissão", "Vencimento"]:
                fmt = date_fmt
            ws.set_column(i, i, width, fmt)
        if len(export):
            ws.autofilter(0, 0, len(export), len(export.columns) - 1)
        ws.freeze_panes(1, 0)
    return output.getvalue()
