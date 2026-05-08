# -*- coding: utf-8 -*-
import pandas as pd
from io import BytesIO

KI_KEYS = ["hotshot", "cozy fit", "glide-step", "glide step"]

PRIORITY_LABEL_TO_ID = {
    "Descuento Retail": "retail",
    "Tech Sports": "techsports",
    "Compras Retail-Ecomm-BTS": "compras",
    "Listado Prod. 1P": "onep",
    "Key Initiative": "ki",
}

COMPRA_GROUPS = [
    ("COMPRA RETAIL BTS-26", ["M"]),
    ("COMPRA RETAIL 2026-1", ["N", "O", "P"]),
    ("COMPRA ECOMM 2026-1", ["Q"]),
    ("COMPRA ECOMM BTS-26", ["R"]),
    ("COMPRA RETAIL 2025-2", ["S", "T", "U"]),
    ("COMPRA RETAIL 2026-2", ["V", "W", "X"]),
    ("COMPRA ECOMM 2026-2", ["Y"]),
]


def _norm_sc(x) -> str:
    if x is None or pd.isna(x):
        return ""
    return str(x).strip()


def _norm_key(x) -> str:
    return _norm_sc(x).replace("-", "").replace(" ", "").upper()


def _safe_str(v) -> str:
    if v is None or pd.isna(v):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def _cell_is_value(v) -> bool:
    s = _safe_str(v)
    return bool(s and s != "-")


def _to_float_discount(v):
    if v is None or pd.isna(v):
        return None

    s = str(v).strip()
    if not s or s == "-":
        return None

    s = s.replace("%", "").replace(",", ".").strip()

    try:
        f = float(s)
    except Exception:
        return None

    if 0 < f <= 1:
        return round(f * 100, 4)

    return float(f)


def _config_from_editor(df, key_col):
    config = {}

    if df is None or df.empty:
        return config

    for _, row in df.iterrows():
        key = _safe_str(row.get(key_col, ""))
        if not key:
            continue

        enabled = bool(row.get("Permite descuento", False))
        pct = _to_float_discount(row.get("%", 0)) or 0.0

        config[key] = {
            "enabled": enabled,
            "pct": float(pct),
        }

    return config


def _season_config_from_editor(df):
    config = {}

    if df is None or df.empty:
        return config

    for _, row in df.iterrows():
        season = _safe_str(row.get("Season Actual", ""))
        if not season:
            continue

        pct = _to_float_discount(row.get("%", 0)) or 0.0
        config[season] = float(pct)

    return config


def load_style_colors(uploaded_file):
    df = pd.read_excel(uploaded_file, header=0, usecols=[0])
    vals = [_norm_sc(v) for v in df.iloc[:, 0].tolist()]
    return [v for v in vals if v]


def load_1p(path):
    df = pd.read_excel(path, header=0, usecols="A,D,E,F,G,H")
    df.columns = ["Style-Color", "Falabella", "Meli", "Paris", "Ripley", "Ecommerce"]
    df["Style-Color"] = df["Style-Color"].map(_norm_sc)

    for c in ["Falabella", "Meli", "Paris", "Ripley", "Ecommerce"]:
        df[c] = df[c].fillna("-").astype(str).str.strip()

    return df.drop_duplicates(subset=["Style-Color"], keep="first").set_index("Style-Color", drop=True)


def load_techsports(path):
    # A = SKECHERS / Style-Color
    # G = STYLE NAME
    # I = SEASON ACTUAL
    # K = DETALLE PARA DSCTO
    df = pd.read_excel(path, header=0, usecols="A,G,I,K")
    df.columns = ["Style-Color", "Style Name", "SEASON_ACTUAL", "Detalle"]

    df["Style-Color"] = df["Style-Color"].map(_norm_sc)
    df["Style Name"] = df["Style Name"].map(_safe_str)
    df["SEASON_ACTUAL"] = df["SEASON_ACTUAL"].map(_safe_str)
    df["Detalle"] = df["Detalle"].fillna("-").astype(str).str.strip()

    return df.drop_duplicates(subset=["Style-Color"], keep="first").set_index("Style-Color", drop=True)


def load_compra(path):
    df = pd.read_excel(path, header=0, usecols="B,M,N,O,P,Q,R,S,T,U,V,W,X,Y")
    df.columns = ["Style-Color", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y"]
    df["Style-Color"] = df["Style-Color"].map(_norm_sc)

    return df.drop_duplicates(subset=["Style-Color"], keep="first").set_index("Style-Color", drop=True)


def load_retail(path):
    lookup = {}
    xl = pd.ExcelFile(path)

    pares_sheet = "PARES" if "PARES" in xl.sheet_names else xl.sheet_names[0]
    acc_sheet = "ACC" if "ACC" in xl.sheet_names else xl.sheet_names[min(1, len(xl.sheet_names) - 1)]

    df_pares = pd.read_excel(path, sheet_name=pares_sheet, header=0, usecols="C,I")
    df_pares.columns = ["StyleColor", "Descuento"]

    df_acc = pd.read_excel(path, sheet_name=acc_sheet, header=0, usecols="C,G")
    df_acc.columns = ["StyleColor", "Descuento"]

    for df in [df_pares, df_acc]:
        df["StyleColor"] = df["StyleColor"].map(_safe_str)
        df["Descuento"] = df["Descuento"].apply(_to_float_discount)

        for _, row in df.iterrows():
            key = _norm_key(row["StyleColor"])
            pct = row["Descuento"]

            if not key or pct is None:
                continue

            prev = lookup.get(key)
            if prev is None or pct > prev:
                lookup[key] = float(pct)

    return lookup


def load_disponible(path):
    # Disponible.xlsx / hoja ALL
    # A = SKECHERS
    # AH = QTY
    # AI = Q
    raw = pd.read_excel(path, sheet_name="ALL", header=None)

    key_pos = 0
    qty_pos = 33
    q_pos = 34

    header_row = None

    for idx, row in raw.iterrows():
        a = str(row.iloc[key_pos]).strip().upper() if len(row) > key_pos else ""
        ah = str(row.iloc[qty_pos]).strip().upper() if len(row) > qty_pos else ""
        ai = str(row.iloc[q_pos]).strip().upper() if len(row) > q_pos else ""

        if a == "SKECHERS" and ah == "QTY" and ai == "Q":
            header_row = idx
            break

    if header_row is None:
        raise ValueError("No encontré encabezados en Disponible: A=SKECHERS, AH=QTY, AI=Q.")

    out = raw.iloc[header_row + 1:, [key_pos, qty_pos, q_pos]].copy()
    out.columns = ["Style-Color", "QTY", "Q"]
    out["Style-Color"] = out["Style-Color"].map(_norm_sc)
    out = out[out["Style-Color"] != ""]

    return out.drop_duplicates(subset=["Style-Color"], keep="first").set_index("Style-Color", drop=True)


def _get_retail_discount(sc, retail_lookup):
    return retail_lookup.get(_norm_key(sc))


def _onep_matches(sc, onep_idx):
    matches = []

    if sc not in onep_idx.index:
        return matches

    row = onep_idx.loc[sc]

    canales = [
        ("Falabella", "Falabella"),
        ("Meli", "Meli"),
        ("Paris", "Paris"),
        ("Ripley", "Ripley"),
        ("Ecommerce", "Ecommerce"),
    ]

    for col, canal in canales:
        val = _safe_str(row.get(col, "-"))
        if val and val != "-":
            matches.append((val, canal))

    return matches


def _tech_match(sc, tech_idx):
    if sc not in tech_idx.index:
        return None

    row = tech_idx.loc[sc]
    detalle = _safe_str(row.get("Detalle", "-"))
    season = _safe_str(row.get("SEASON_ACTUAL", ""))
    name = _safe_str(row.get("Style Name", ""))

    if not detalle or detalle == "-":
        return None

    return detalle, season, name


def _compra_matches(sc, compra_idx):
    matches = []

    if sc not in compra_idx.index:
        return matches

    row = compra_idx.loc[sc]

    for group_name, cols in COMPRA_GROUPS:
        if any(_cell_is_value(row.get(c)) for c in cols):
            matches.append(group_name)

    return matches


def _eval_ki(name):
    name_l = (name or "").lower()

    for k in KI_KEYS:
        if k in name_l:
            if k in ["glide step", "glide-step"]:
                return "Glide-Step"
            if k == "cozy fit":
                return "Cozy Fit"
            if k == "hotshot":
                return "Hotshot"

    return None


def _eval_configured(rule_name, matches, config, detail_builder=None):
    if not matches:
        return None

    enabled = []
    disabled = []

    for raw in matches:
        key = raw[0] if isinstance(raw, tuple) else raw
        cfg = config.get(key, {"enabled": False, "pct": 0.0})

        if cfg.get("enabled", False):
            enabled.append((raw, float(cfg.get("pct", 0.0))))
        else:
            disabled.append(raw)

    if enabled:
        max_pct = max(p for _, p in enabled)
        selected = [raw for raw, p in enabled if p == max_pct]

        if detail_builder:
            detail = detail_builder(selected)
        else:
            detail = ", ".join([str(x[0] if isinstance(x, tuple) else x) for x in selected])

        return {
            "habilitado": True,
            "pct": max_pct,
            "comentario": f"{rule_name} - {detail} - {max_pct:g}%",
        }

    if disabled:
        if detail_builder:
            detail = detail_builder(disabled)
        else:
            detail = ", ".join([str(x[0] if isinstance(x, tuple) else x) for x in disabled])

        return {
            "habilitado": False,
            "pct": 0.0,
            "comentario": f"{rule_name} - {detail}",
        }

    return None


def _eval_rule(rule_id, sc, name, onep_idx, tech_idx, compra_idx, retail_lookup, cfg_1p, cfg_tech, cfg_tech_season, cfg_compras):
    if rule_id == "retail":
        pct = _get_retail_discount(sc, retail_lookup)
        if pct is not None:
            return {
                "habilitado": True,
                "pct": float(pct),
                "comentario": f"Descuento Retail - {float(pct):g}%",
            }
        return None

    if rule_id == "onep":
        matches = _onep_matches(sc, onep_idx)
        return _eval_configured(
            "Listado Prod. 1P",
            matches,
            cfg_1p,
            detail_builder=lambda items: ", ".join([f"{v} {canal}" for v, canal in items])
        )

    if rule_id == "techsports":
        match = _tech_match(sc, tech_idx)
        if not match:
            return None

        detalle, season_actual, tech_name = match
        cfg = cfg_tech.get(detalle, {"enabled": False, "pct": 0.0})

        if not cfg.get("enabled", False):
            return {
                "habilitado": False,
                "pct": 0.0,
                "comentario": f"Tech Sports - {detalle}",
            }

        if detalle.upper().strip() == "APTO PARA DSCTO" and cfg_tech_season:
            pct = float(cfg_tech_season.get(season_actual, cfg.get("pct", 0.0)) or 0.0)
            return {
                "habilitado": True,
                "pct": pct,
                "comentario": f"Tech Sports - {detalle} {season_actual} - {pct:g}%",
            }

        pct = float(cfg.get("pct", 0.0))
        return {
            "habilitado": True,
            "pct": pct,
            "comentario": f"Tech Sports - {detalle} - {pct:g}%",
        }

    if rule_id == "compras":
        matches = _compra_matches(sc, compra_idx)
        return _eval_configured("Compras", matches, cfg_compras)

    if rule_id == "ki":
        ki = _eval_ki(name)
        if ki:
            return {
                "habilitado": False,
                "pct": 0.0,
                "comentario": f"Key Initiative - {ki}",
            }
        return None

    return None


def calcular_descuentos(
    style_file,
    files,
    prioridades,
    config_1p_df,
    config_tech_df,
    config_tech_season_df,
    config_compras_df,
):
    style_colors = load_style_colors(style_file)

    onep_idx = load_1p(files["Listado Prod. 1P"])
    tech_idx = load_techsports(files["Tech Sports"])
    compra_idx = load_compra(files["Compras Retail-Ecomm-BTS"])
    retail_lookup = load_retail(files["Descuentos Retail"])
    disponible_idx = load_disponible(files["Disponible"])

    cfg_1p = _config_from_editor(config_1p_df, "Valor detectado")
    cfg_tech = _config_from_editor(config_tech_df, "Valor detectado")
    cfg_tech_season = _season_config_from_editor(config_tech_season_df)
    cfg_compras = _config_from_editor(config_compras_df, "Grupo detectado")

    priority_order = [
        PRIORITY_LABEL_TO_ID.get(p, p)
        for p in prioridades
        if PRIORITY_LABEL_TO_ID.get(p, p)
    ]

    rows = []

    for sc in style_colors:
        name = ""

        tech_match = _tech_match(sc, tech_idx)
        if tech_match:
            _, _, name = tech_match

        result = None

        for rule_id in priority_order:
            result = _eval_rule(
                rule_id=rule_id,
                sc=sc,
                name=name,
                onep_idx=onep_idx,
                tech_idx=tech_idx,
                compra_idx=compra_idx,
                retail_lookup=retail_lookup,
                cfg_1p=cfg_1p,
                cfg_tech=cfg_tech,
                cfg_tech_season=cfg_tech_season,
                cfg_compras=cfg_compras,
            )

            if result is not None:
                break

        if result is None:
            result = {
                "habilitado": True,
                "pct": 0.0,
                "comentario": "Habilitado a tener descuento",
            }

        if sc in disponible_idx.index:
            qty = disponible_idx.loc[sc].get("QTY", "")
            q = disponible_idx.loc[sc].get("Q", "")
        else:
            qty = ""
            q = ""

        rows.append({
            "Style-Color": sc,
            "%": float(result["pct"]),
            "Habilitado/No Habilitado": (
                "Habilitado para tener Descuento"
                if result["habilitado"]
                else "No Habilitado para tener Descuento"
            ),
            "Comentario": result["comentario"],
            "QTY": qty,
            "Q": q,
        })

    return pd.DataFrame(
        rows,
        columns=[
            "Style-Color",
            "%",
            "Habilitado/No Habilitado",
            "Comentario",
            "QTY",
            "Q",
        ]
    )


def dataframe_to_excel_bytes(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sheet1")

    output.seek(0)
    return output.getvalue()
