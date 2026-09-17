from __future__ import annotations

import io
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Iterable
from zoneinfo import ZoneInfo

import pandas as pd

TZ = ZoneInfo("America/Sao_Paulo")
REQUIRED_COLUMNS = [
    "LOCAL",
    "Equipamento",
    "Status",
    "Nome Fantasia",
    "Produto",
    "Grupo",
    "Contrato",
    "DATA FINAL",
]

COLUMN_ALIASES = {
    "local": "LOCAL",
    "unidade": "LOCAL",
    "equipamento": "Equipamento",
    "status": "Status",
    "nome fantasia": "Nome Fantasia",
    "cliente": "Nome Fantasia",
    "proprietario": "Proprietário",
    "cod produto": "Cód. Produto",
    "codigo produto": "Cód. Produto",
    "produto": "Produto",
    "grupo": "Grupo",
    "contrato": "Contrato",
    "os aberta": "OS Aberta",
    "data final": "DATA FINAL",
    "data vencimento": "DATA FINAL",
    "vencimento": "DATA FINAL",
    "obs": "OBS",
    "observacao": "OBS",
}

LOCADO_STATUS_KEYWORDS = (
    "em contrato",
    "com contrato",
    "locado",
    "alugado",
)


def now_sp() -> datetime:
    return datetime.now(TZ)


def today_sp() -> date:
    return now_sp().date()


def _clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _normalize_text(value: str) -> str:
    text = _clean_text(value).lower()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _canonicalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    mapping = {}
    for col in df.columns:
        norm = _normalize_text(str(col)).replace(".", "")
        mapping[col] = COLUMN_ALIASES.get(norm, str(col).strip())
    return df.rename(columns=mapping)


def _parse_excel_date(series: pd.Series) -> pd.Series:
    # Primeiro tenta datas já reconhecidas pelo Excel / pandas.
    parsed = pd.to_datetime(series, errors="coerce", dayfirst=True)

    # Converte somente valores realmente numéricos que representem serial do Excel.
    # Isso evita tratar Timestamp como seu valor interno em nanossegundos.
    numeric_mask = series.map(
        lambda v: isinstance(v, (int, float)) and not isinstance(v, bool) and pd.notna(v)
    )
    if numeric_mask.any():
        numeric = pd.to_numeric(series.where(numeric_mask), errors="coerce")
        valid_serial = numeric.between(1, 100000, inclusive="both")
        excel_dates = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
        excel_dates.loc[valid_serial] = (
            pd.Timestamp("1899-12-30")
            + pd.to_timedelta(numeric.loc[valid_serial], unit="D")
        )
        parsed = parsed.where(parsed.notna(), excel_dates)

    return parsed.dt.normalize()


def _extract_client_name(value: str) -> str:
    text = _clean_text(value)
    if not text:
        return ""
    return re.sub(r"\s*\[[^\]]+\]\s*$", "", text).strip()


def _extract_client_code(value: str) -> str:
    text = _clean_text(value)
    match = re.search(r"\[([^\]]+)\]\s*$", text)
    return match.group(1).strip() if match else ""


def _is_locado_row(row: pd.Series) -> bool:
    status = _normalize_text(row.get("Status", ""))
    contrato = _clean_text(row.get("Contrato", ""))
    status_locado = any(keyword in status for keyword in LOCADO_STATUS_KEYWORDS)
    return bool(status_locado or contrato)


def validate_dataframe(df: pd.DataFrame) -> tuple[bool, list[str]]:
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    return len(missing) == 0, missing


def read_excel_bytes(file_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0)
    raw = _canonicalize_columns(raw)
    ok, missing = validate_dataframe(raw)
    if not ok:
        raise ValueError(
            "A planilha não possui todas as colunas necessárias. Faltando: "
            + ", ".join(missing)
        )
    return prepare_dataframe(raw)


def read_excel_path(path: str | Path) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name=0)
    raw = _canonicalize_columns(raw)
    ok, missing = validate_dataframe(raw)
    if not ok:
        raise ValueError(
            "A planilha não possui todas as colunas necessárias. Faltando: "
            + ", ".join(missing)
        )
    return prepare_dataframe(raw)


def prepare_dataframe(df: pd.DataFrame, today: date | None = None) -> pd.DataFrame:
    today = today or today_sp()
    out = df.copy()

    # Garante colunas opcionais para manter o app resiliente.
    for col in ["Proprietário", "Cód. Produto", "OS Aberta", "OBS"]:
        if col not in out.columns:
            out[col] = ""

    text_cols = [
        "LOCAL",
        "Equipamento",
        "Status",
        "Nome Fantasia",
        "Proprietário",
        "Produto",
        "Grupo",
        "Contrato",
        "OS Aberta",
        "OBS",
    ]
    for col in text_cols:
        out[col] = out[col].map(_clean_text)

    out["DATA FINAL"] = _parse_excel_date(out["DATA FINAL"])
    out["Cliente"] = out["Nome Fantasia"].map(_extract_client_name)
    out["Código Cliente"] = out["Nome Fantasia"].map(_extract_client_code)

    out["É Locado"] = out.apply(_is_locado_row, axis=1)
    out["É Disponível"] = out["Status"].map(_normalize_text).eq("disponivel")
    out["Tem OS"] = out["OS Aberta"].astype(str).str.strip().ne("")

    today_ts = pd.Timestamp(today)
    out["Dias Atraso"] = 0
    mask_date = out["DATA FINAL"].notna()
    delta_days = (today_ts - out.loc[mask_date, "DATA FINAL"]).dt.days
    out.loc[mask_date, "Dias Atraso"] = delta_days
    out["Dias Atraso"] = out["Dias Atraso"].astype(int)

    out["Atrasado"] = out["É Locado"] & out["DATA FINAL"].notna() & (out["Dias Atraso"] > 0)
    out["Vence Hoje"] = out["É Locado"] & out["DATA FINAL"].eq(today_ts)

    dias_ate_vencer = pd.Series(pd.NA, index=out.index, dtype="Int64")
    dias_ate_vencer.loc[mask_date] = (
        out.loc[mask_date, "DATA FINAL"] - today_ts
    ).dt.days.astype("Int64")
    out["Dias até Vencer"] = dias_ate_vencer
    out["Vence em 5 dias"] = (
        out["É Locado"]
        & out["Dias até Vencer"].notna()
        & out["Dias até Vencer"].between(1, 5, inclusive="both")
    )

    def faixa_atraso(row) -> str:
        if not row["Atrasado"]:
            return ""
        dias = int(row["Dias Atraso"])
        if dias == 1:
            return "24h"
        if dias == 2:
            return "48h"
        return "72h+"

    out["Faixa Atraso"] = out.apply(faixa_atraso, axis=1)

    def prioridade(row) -> str:
        if row["Atrasado"]:
            dias = int(row["Dias Atraso"])
            if dias >= 3:
                return "🔴 Crítico 72h+"
            if dias == 2:
                return "🟠 Atenção 48h"
            return "🟡 Alerta 24h"
        if row["Vence Hoje"]:
            return "🔵 Vence hoje"
        if row["Vence em 5 dias"]:
            return "🟣 Próximos 5 dias"
        if row["É Locado"]:
            return "🟢 Em dia"
        if row["É Disponível"]:
            return "⚪ Disponível"
        return "⚫ Outro status"

    out["Prioridade"] = out.apply(prioridade, axis=1)

    # Normaliza vazios de unidade/grupo para não quebrar filtros.
    out["LOCAL"] = out["LOCAL"].replace("", "SEM UNIDADE")
    out["Grupo"] = out["Grupo"].replace("", "SEM GRUPO")
    out["Produto"] = out["Produto"].replace("", "SEM PRODUTO")

    return out


@dataclass(frozen=True)
class Metrics:
    total: int
    locados: int
    disponiveis: int
    outros: int
    atrasados: int
    vence_hoje: int
    vence_5_dias: int
    faixa_24h: int
    faixa_48h: int
    faixa_72h: int
    acumulado_24h: int
    acumulado_48h: int
    acumulado_72h: int
    ocupacao_pct: float
    atraso_locados_pct: float


def calculate_metrics(df: pd.DataFrame) -> Metrics:
    total = int(len(df))
    locados = int(df["É Locado"].sum()) if total else 0
    disponiveis = int(df["É Disponível"].sum()) if total else 0
    atrasados = int(df["Atrasado"].sum()) if total else 0
    vence_hoje = int(df["Vence Hoje"].sum()) if total else 0
    vence_5 = int(df["Vence em 5 dias"].sum()) if total else 0

    atraso_days = df.loc[df["Atrasado"], "Dias Atraso"] if total else pd.Series(dtype=int)
    faixa_24h = int((atraso_days == 1).sum())
    faixa_48h = int((atraso_days == 2).sum())
    faixa_72h = int((atraso_days >= 3).sum())
    acumulado_24h = int((atraso_days >= 1).sum())
    acumulado_48h = int((atraso_days >= 2).sum())
    acumulado_72h = int((atraso_days >= 3).sum())

    outros = max(total - locados - disponiveis, 0)
    ocupacao = (locados / total * 100) if total else 0.0
    atraso_locados = (atrasados / locados * 100) if locados else 0.0

    return Metrics(
        total=total,
        locados=locados,
        disponiveis=disponiveis,
        outros=outros,
        atrasados=atrasados,
        vence_hoje=vence_hoje,
        vence_5_dias=vence_5,
        faixa_24h=faixa_24h,
        faixa_48h=faixa_48h,
        faixa_72h=faixa_72h,
        acumulado_24h=acumulado_24h,
        acumulado_48h=acumulado_48h,
        acumulado_72h=acumulado_72h,
        ocupacao_pct=ocupacao,
        atraso_locados_pct=atraso_locados,
    )


def filter_scope(
    df: pd.DataFrame,
    unidades: Iterable[str] | None = None,
    grupos: Iterable[str] | None = None,
    search: str = "",
) -> pd.DataFrame:
    out = df.copy()
    unidades = list(unidades or [])
    grupos = list(grupos or [])

    if unidades:
        out = out[out["LOCAL"].isin(unidades)]
    if grupos:
        out = out[out["Grupo"].isin(grupos)]

    search = search.strip()
    if search:
        s = search.casefold()
        haystack = (
            out["Cliente"].astype(str)
            + " | "
            + out["Contrato"].astype(str)
            + " | "
            + out["Equipamento"].astype(str)
            + " | "
            + out["Produto"].astype(str)
        ).str.casefold()
        out = out[haystack.str.contains(re.escape(s), na=False)]

    return out


def filter_view(df: pd.DataFrame, visao: str, faixas: Iterable[str] | None = None) -> pd.DataFrame:
    out = df.copy()
    if visao == "Atrasados":
        out = out[out["Atrasado"]]
    elif visao == "Vence hoje":
        out = out[out["Vence Hoje"]]
    elif visao == "Vence em 5 dias":
        out = out[out["Vence em 5 dias"]]
    elif visao == "Locados em dia":
        out = out[out["É Locado"] & ~out["Atrasado"]]
    elif visao == "Disponíveis":
        out = out[out["É Disponível"]]
    elif visao == "Com OS":
        out = out[out["Tem OS"]]

    faixas = list(faixas or [])
    if faixas:
        out = out[out["Faixa Atraso"].isin(faixas)]
    return out


def priority_sort(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    rank = {
        "🔴 Crítico 72h+": 0,
        "🟠 Atenção 48h": 1,
        "🟡 Alerta 24h": 2,
        "🔵 Vence hoje": 3,
        "🟣 Próximos 5 dias": 4,
        "🟢 Em dia": 5,
        "⚪ Disponível": 6,
        "⚫ Outro status": 7,
    }
    out["__rank"] = out["Prioridade"].map(rank).fillna(99)
    out = out.sort_values(
        ["__rank", "Dias Atraso", "DATA FINAL"],
        ascending=[True, False, True],
        na_position="last",
    )
    return out.drop(columns="__rank")


def dashboard_table(df: pd.DataFrame) -> pd.DataFrame:
    cols = [
        "Prioridade",
        "LOCAL",
        "Cliente",
        "Código Cliente",
        "Contrato",
        "Equipamento",
        "Produto",
        "Grupo",
        "Status",
        "DATA FINAL",
        "Dias Atraso",
        "Dias até Vencer",
        "OS Aberta",
        "OBS",
    ]
    existing = [c for c in cols if c in df.columns]
    return priority_sort(df)[existing].copy()


def build_excel_export(df: pd.DataFrame, sheet_name: str = "Pendencias") -> bytes:
    output = io.BytesIO()
    export_df = dashboard_table(df)
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        export_df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        workbook = writer.book
        worksheet = writer.sheets[sheet_name[:31]]
        header_fmt = workbook.add_format(
            {
                "bold": True,
                "font_color": "#FFFFFF",
                "bg_color": "#2F4EA2",
                "border": 0,
                "align": "center",
                "valign": "vcenter",
            }
        )
        date_fmt = workbook.add_format({"num_format": "dd/mm/yyyy"})
        for col_idx, col in enumerate(export_df.columns):
            worksheet.write(0, col_idx, col, header_fmt)
            sample = export_df[col].astype(str).head(150)
            width = min(max([len(col)] + [len(x) for x in sample]) + 2, 42)
            worksheet.set_column(col_idx, col_idx, width)
        if "DATA FINAL" in export_df.columns:
            idx = export_df.columns.get_loc("DATA FINAL")
            worksheet.set_column(idx, idx, 13, date_fmt)
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, len(export_df), max(len(export_df.columns) - 1, 0))
    return output.getvalue()


def snapshot_records(df: pd.DataFrame, snapshot_date: date | None = None) -> pd.DataFrame:
    snapshot_date = snapshot_date or today_sp()
    records = []
    scopes = [("TODAS", df)]
    for unidade in sorted(df["LOCAL"].dropna().astype(str).unique()):
        scopes.append((unidade, df[df["LOCAL"] == unidade]))

    for unidade, scope in scopes:
        m = calculate_metrics(scope)
        records.append(
            {
                "data": snapshot_date.isoformat(),
                "unidade": unidade,
                "total": m.total,
                "locados": m.locados,
                "disponiveis": m.disponiveis,
                "atrasados": m.atrasados,
                "vence_hoje": m.vence_hoje,
                "vence_5_dias": m.vence_5_dias,
                "ocupacao_pct": round(m.ocupacao_pct, 4),
                "atraso_locados_pct": round(m.atraso_locados_pct, 4),
            }
        )
    return pd.DataFrame(records)


def upsert_history(df: pd.DataFrame, history_path: str | Path) -> None:
    history_path = Path(history_path)
    history_path.parent.mkdir(parents=True, exist_ok=True)
    new = snapshot_records(df)

    if history_path.exists():
        try:
            old = pd.read_csv(history_path, dtype={"unidade": str})
        except Exception:
            old = pd.DataFrame()
    else:
        old = pd.DataFrame()

    if not old.empty and {"data", "unidade"}.issubset(old.columns):
        keys = set(zip(new["data"], new["unidade"]))
        keep_mask = [
            (str(row.get("data")), str(row.get("unidade"))) not in keys
            for _, row in old.iterrows()
        ]
        old = old.loc[keep_mask]

    combined = pd.concat([old, new], ignore_index=True)
    combined = combined.sort_values(["data", "unidade"])
    combined.to_csv(history_path, index=False)


def read_history(history_path: str | Path) -> pd.DataFrame:
    path = Path(history_path)
    if not path.exists():
        return pd.DataFrame()
    try:
        hist = pd.read_csv(path)
        hist["data"] = pd.to_datetime(hist["data"], errors="coerce")
        return hist.dropna(subset=["data"])
    except Exception:
        return pd.DataFrame()
