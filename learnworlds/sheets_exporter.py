"""Export LearnWorlds course data to a Google Sheets spreadsheet."""

import os
import json
from datetime import datetime

import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ── Auth ──────────────────────────────────────────────────────────────────────

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

STATUS_PT = {
    "completed":  "Concluído",
    "in_progress": "Em Andamento",
    "not_started": "Não Iniciado",
    "unknown":    "Desconhecido",
}

# ── Color palette ─────────────────────────────────────────────────────────────

def _rgb(r, g, b):
    return {"red": round(r/255, 4), "green": round(g/255, 4), "blue": round(b/255, 4)}

NAVY        = _rgb(30,  58,  95)
NAVY_LIGHT  = _rgb(52,  73, 104)
WHITE       = _rgb(255, 255, 255)
LIGHT_GRAY  = _rgb(248, 249, 250)
MID_GRAY    = _rgb(222, 226, 230)
ORANGE      = _rgb(245, 124,  0)

GREEN_BG    = _rgb(212, 237, 218)
GREEN_TEXT  = _rgb( 21,  87,  36)
AMBER_BG    = _rgb(255, 243, 205)
AMBER_TEXT  = _rgb(133, 100,   4)
RED_BG      = _rgb(248, 215, 218)
RED_TEXT    = _rgb(114,  28,  36)
GRAY_BG     = _rgb(233, 236, 239)
GRAY_TEXT   = _rgb( 73,  80,  87)

STATUS_FMT = {
    "Concluído":     {"bg": GREEN_BG, "fg": GREEN_TEXT},
    "Em Andamento":  {"bg": AMBER_BG, "fg": AMBER_TEXT},
    "Não Iniciado":  {"bg": RED_BG,   "fg": RED_TEXT},
    "Desconhecido":  {"bg": GRAY_BG,  "fg": GRAY_TEXT},
}

# ── Low-level helpers ─────────────────────────────────────────────────────────

def _get_gc():
    sa_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "google-credentials.json")
    if os.path.exists(sa_file):
        creds = Credentials.from_service_account_file(sa_file, scopes=SCOPES)
    else:
        sa_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
        if not sa_json:
            raise FileNotFoundError(
                f"Google credentials not found at '{sa_file}'. "
                "Set GOOGLE_SERVICE_ACCOUNT_FILE or GOOGLE_SERVICE_ACCOUNT_JSON in .env"
            )
        creds = Credentials.from_service_account_info(json.loads(sa_json), scopes=SCOPES)
    return gspread.authorize(creds)


def _col(n: int) -> str:
    """1-indexed column number → letter (e.g. 1→A, 27→AA)."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _a1(row: int, col: int) -> str:
    return f"{_col(col)}{row}"


def _range(r1, c1, r2, c2) -> str:
    return f"{_a1(r1, c1)}:{_a1(r2, c2)}"


def _fmt_req(sheet_id, r1, c1, r2, c2, *, bg=None, bold=None, font_size=None,
             fg=None, h_align=None, v_align=None, wrap=None, border_bottom=False):
    fmt, fields = {}, []
    if bg:
        fmt["backgroundColor"] = bg
        fields.append("backgroundColor")
    tf = {}
    if bold is not None:
        tf["bold"] = bold
        fields.append("textFormat.bold")
    if font_size:
        tf["fontSize"] = font_size
        fields.append("textFormat.fontSize")
    if fg:
        tf["foregroundColor"] = fg
        fields.append("textFormat.foregroundColor")
    if tf:
        fmt["textFormat"] = tf
    if h_align:
        fmt["horizontalAlignment"] = h_align
        fields.append("horizontalAlignment")
    if v_align:
        fmt["verticalAlignment"] = v_align
        fields.append("verticalAlignment")
    if wrap:
        fmt["wrapStrategy"] = wrap
        fields.append("wrapStrategy")
    if border_bottom:
        fmt["borders"] = {"bottom": {"style": "SOLID_MEDIUM", "color": NAVY}}
        fields.append("borders.bottom")

    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r2,
                "startColumnIndex": c1 - 1, "endColumnIndex": c2,
            },
            "cell": {"userEnteredFormat": fmt},
            "fields": ",".join(f"userEnteredFormat.{f}" for f in fields),
        }
    }


def _col_width_req(sheet_id, col_idx_0, px):
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                      "startIndex": col_idx_0, "endIndex": col_idx_0 + 1},
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }


def _row_height_req(sheet_id, row_idx_0, px):
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                      "startIndex": row_idx_0, "endIndex": row_idx_0 + 1},
            "properties": {"pixelSize": px},
            "fields": "pixelSize",
        }
    }


def _merge_req(sheet_id, r1, c1, r2, c2):
    return {
        "mergeCells": {
            "range": {"sheetId": sheet_id, "startRowIndex": r1 - 1, "endRowIndex": r2,
                      "startColumnIndex": c1 - 1, "endColumnIndex": c2},
            "mergeType": "MERGE_ALL",
        }
    }


def _freeze_req(sheet_id, rows=1, cols=0):
    return {
        "updateSheetProperties": {
            "properties": {
                "sheetId": sheet_id,
                "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols},
            },
            "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
        }
    }


def _cond_gradient_req(sheet_id, r1, c1, r2, c2):
    """Red → Amber → Green gradient for 0-100 range."""
    return {
        "addConditionalFormatRule": {
            "rule": {
                "ranges": [{"sheetId": sheet_id, "startRowIndex": r1 - 1,
                            "endRowIndex": r2, "startColumnIndex": c1 - 1, "endColumnIndex": c2}],
                "gradientRule": {
                    "minpoint": {"color": RED_BG, "type": "NUMBER", "value": "0"},
                    "midpoint": {"color": AMBER_BG, "type": "NUMBER", "value": "50"},
                    "maxpoint": {"color": GREEN_BG, "type": "NUMBER", "value": "100"},
                },
            },
            "index": 0,
        }
    }


def _batch(spreadsheet, requests):
    if requests:
        spreadsheet.batch_update({"requests": requests})


# ── Sheet builders ─────────────────────────────────────────────────────────────

def _build_resumo(ws, course: dict, df: pd.DataFrame, reviews: list):
    sid = ws.id
    total = len(df)
    completed   = int((df["completion_percentage"] >= 100).sum())
    in_progress = int(((df["completion_percentage"] > 0) & (df["completion_percentage"] < 100)).sum())
    not_started = int((df["completion_percentage"] == 0).sum())
    avg_comp   = round(df["completion_percentage"].mean(), 1) if total else 0
    avg_hours  = round(df["time_spent_hours"].mean(), 1) if total else 0
    total_hours = round(df["time_spent_hours"].sum(), 1) if total else 0
    certs      = int(df["certificate_issued"].sum())
    avg_quiz   = round(df["quiz_score"].dropna().mean(), 1) if df["quiz_score"].notna().any() else "—"
    comp_rate  = round(completed / total * 100, 1) if total else 0
    avg_rating = round(pd.DataFrame(reviews)["rating"].dropna().mean(), 1) if reviews else "—"
    generated  = datetime.now().strftime("%d/%m/%Y às %H:%M")
    title      = course.get("title", "Curso")

    rows = [
        [title],
        [f"Relatório gerado em: {generated}"],
        [""],
        ["MÉTRICAS GERAIS", ""],
        ["Total Matriculados",          total],
        ["Taxa de Conclusão",           f"{comp_rate}%"],
        ["Concluíram",                  completed],
        ["Em Andamento",               in_progress],
        ["Não Iniciaram",              not_started],
        [""],
        ["PERFORMANCE"],
        ["Progresso Médio",            f"{avg_comp}%"],
        ["Tempo Médio por Aluno",      f"{avg_hours}h"],
        ["Tempo Total (turma)",        f"{total_hours}h"],
        ["Certificados Emitidos",      certs],
        ["Média Quizzes",             avg_quiz if avg_quiz == "—" else f"{avg_quiz}%"],
        ["Satisfação Média",          avg_rating if avg_rating == "—" else f"{avg_rating}/5"],
        [""],
        ["DISTRIBUIÇÃO DE STATUS", "Qtd", "%"],
        ["Concluído",      completed,   f"{comp_rate}%"],
        ["Em Andamento",  in_progress,  f"{round(in_progress/total*100,1) if total else 0}%"],
        ["Não Iniciado",  not_started,  f"{round(not_started/total*100,1) if total else 0}%"],
    ]
    ws.update("A1", rows)

    reqs = [
        _merge_req(sid, 1, 1, 1, 3),
        _fmt_req(sid, 1, 1, 1, 3, bg=NAVY, fg=WHITE, bold=True, font_size=14, h_align="CENTER", v_align="MIDDLE"),
        _row_height_req(sid, 1, 42),
        _fmt_req(sid, 2, 1, 2, 3, fg=GRAY_TEXT, font_size=9),
        _fmt_req(sid, 4, 1, 4, 2, bg=NAVY, fg=WHITE, bold=True),
        _fmt_req(sid, 11, 1, 11, 2, bg=NAVY, fg=WHITE, bold=True),
        _fmt_req(sid, 19, 1, 19, 3, bg=NAVY_LIGHT, fg=WHITE, bold=True),
        # Status coloring rows 20-22
        _fmt_req(sid, 20, 1, 20, 3, bg=GREEN_BG, fg=GREEN_TEXT),
        _fmt_req(sid, 21, 1, 21, 3, bg=AMBER_BG, fg=AMBER_TEXT),
        _fmt_req(sid, 22, 1, 22, 3, bg=RED_BG, fg=RED_TEXT),
        # Column widths
        _col_width_req(sid, 0, 220),
        _col_width_req(sid, 1, 110),
        _col_width_req(sid, 2, 80),
    ]
    _batch(ws.spreadsheet, reqs)
    ws.freeze(rows=1)


def _build_alunos(ws, df: pd.DataFrame):
    sid = ws.id
    headers = [
        "Nome", "Email", "Status", "Progresso (%)",
        "Aulas Concluídas", "Total Aulas", "% Aulas",
        "Tempo (h)", "Quizzes Feitos", "Score Quiz (%)",
        "Certificado", "Dias Matriculado", "Último Acesso", "Dias s/ Acesso",
    ]

    df_s = df.sort_values("completion_percentage", ascending=False).copy()
    df_s["status_pt"] = df_s["status"].map(STATUS_PT).fillna("Desconhecido")
    df_s["pct_aulas"] = df_s.apply(
        lambda r: round(r["completed_lessons"] / r["total_lessons"] * 100, 1)
        if r["total_lessons"] > 0 else 0, axis=1
    )
    df_s["cert_label"] = df_s["certificate_issued"].map({True: "Sim", False: "Não"})
    df_s["last_activity_fmt"] = pd.to_datetime(df_s["last_activity"], errors="coerce").dt.strftime("%d/%m/%Y")

    data = [headers] + [
        [
            r["username"] or "",
            r["email"],
            r["status_pt"],
            r["completion_percentage"],
            int(r["completed_lessons"]),
            int(r["total_lessons"]),
            r["pct_aulas"],
            r["time_spent_hours"],
            int(r.get("quiz_units_completed", 0) or 0),
            r["quiz_score"] if pd.notna(r.get("quiz_score")) else "",
            r["cert_label"],
            int(r["days_since_enrollment"]) if pd.notna(r.get("days_since_enrollment")) else "",
            r["last_activity_fmt"] if pd.notna(r.get("last_activity_fmt")) else "",
            int(r["days_since_last_access"]) if pd.notna(r.get("days_since_last_access")) else "",
        ]
        for _, r in df_s.iterrows()
    ]

    ws.update("A1", data)
    n = len(df_s)
    ncols = len(headers)

    reqs = [
        _fmt_req(sid, 1, 1, 1, ncols, bg=NAVY, fg=WHITE, bold=True, h_align="CENTER"),
        _freeze_req(sid, rows=1, cols=2),
        _cond_gradient_req(sid, 2, 4, n + 1, 4),   # Progresso col
        _cond_gradient_req(sid, 2, 7, n + 1, 7),   # % Aulas col
        # Column widths
        _col_width_req(sid, 0, 160),  # Nome
        _col_width_req(sid, 1, 220),  # Email
        _col_width_req(sid, 2, 120),  # Status
        _col_width_req(sid, 3, 100),  # Progresso
        _col_width_req(sid, 4, 110),
        _col_width_req(sid, 5, 90),
        _col_width_req(sid, 6, 80),
        _col_width_req(sid, 7, 80),
        _col_width_req(sid, 8, 100),
        _col_width_req(sid, 9, 110),
        _col_width_req(sid, 10, 90),
        _col_width_req(sid, 11, 110),
        _col_width_req(sid, 12, 110),
        _col_width_req(sid, 13, 110),
    ]

    # Color status cells individually
    status_col = 3  # column C (1-indexed)
    for i, (_, row) in enumerate(df_s.iterrows(), start=2):
        status_pt = STATUS_PT.get(row["status"], "Desconhecido")
        fmt_colors = STATUS_FMT.get(status_pt, STATUS_FMT["Desconhecido"])
        reqs.append(_fmt_req(sid, i, status_col, i, status_col,
                             bg=fmt_colors["bg"], fg=fmt_colors["fg"], bold=True, h_align="CENTER"))

    # Alternate row shading
    for i in range(2, n + 2):
        if i % 2 == 0:
            reqs.append(_fmt_req(sid, i, 1, i, ncols, bg=LIGHT_GRAY))

    _batch(ws.spreadsheet, reqs)


def _build_presenca(ws, df: pd.DataFrame):
    sid = ws.id
    headers = [
        "Nome", "Email", "Status",
        "Data Matrícula", "Último Acesso",
        "Dias s/ Acesso", "Dias Matriculado",
        "Tempo (h)", "Progresso (%)",
        "Situação de Acesso",
    ]

    df_s = df.sort_values("days_since_last_access", ascending=False, na_position="first").copy()
    df_s["status_pt"] = df_s["status"].map(STATUS_PT).fillna("Desconhecido")
    df_s["enrolled_fmt"] = pd.to_datetime(df_s["enrolled_at"], errors="coerce").dt.strftime("%d/%m/%Y")
    df_s["last_fmt"] = pd.to_datetime(df_s["last_activity"], errors="coerce").dt.strftime("%d/%m/%Y")

    def _acesso(days):
        if pd.isna(days): return "Nunca acessou"
        if days <= 7:     return "Ativo (≤7 dias)"
        if days <= 30:    return "Recente (8-30 dias)"
        return "Inativo (>30 dias)"

    df_s["situacao"] = df_s["days_since_last_access"].apply(_acesso)

    data = [headers] + [
        [
            r["username"] or "",
            r["email"],
            r["status_pt"],
            r["enrolled_fmt"] if pd.notna(r.get("enrolled_fmt")) else "",
            r["last_fmt"] if pd.notna(r.get("last_fmt")) else "",
            int(r["days_since_last_access"]) if pd.notna(r.get("days_since_last_access")) else "—",
            int(r["days_since_enrollment"]) if pd.notna(r.get("days_since_enrollment")) else "—",
            r["time_spent_hours"],
            r["completion_percentage"],
            r["situacao"],
        ]
        for _, r in df_s.iterrows()
    ]

    ws.update("A1", data)
    n = len(df_s)
    ncols = len(headers)

    SITUACAO_FMT = {
        "Ativo (≤7 dias)":       {"bg": GREEN_BG, "fg": GREEN_TEXT},
        "Recente (8-30 dias)":   {"bg": AMBER_BG, "fg": AMBER_TEXT},
        "Inativo (>30 dias)":    {"bg": RED_BG,   "fg": RED_TEXT},
        "Nunca acessou":         {"bg": GRAY_BG,  "fg": GRAY_TEXT},
    }

    reqs = [
        _fmt_req(sid, 1, 1, 1, ncols, bg=NAVY, fg=WHITE, bold=True, h_align="CENTER"),
        _freeze_req(sid, rows=1, cols=2),
        _cond_gradient_req(sid, 2, 9, n + 1, 9),
        _col_width_req(sid, 0, 160),
        _col_width_req(sid, 1, 220),
        _col_width_req(sid, 2, 120),
        _col_width_req(sid, 3, 110),
        _col_width_req(sid, 4, 110),
        _col_width_req(sid, 5, 100),
        _col_width_req(sid, 6, 110),
        _col_width_req(sid, 7, 80),
        _col_width_req(sid, 8, 100),
        _col_width_req(sid, 9, 150),
    ]

    for i, (_, row) in enumerate(df_s.iterrows(), start=2):
        # Status
        status_pt = STATUS_PT.get(row["status"], "Desconhecido")
        sc = STATUS_FMT.get(status_pt, STATUS_FMT["Desconhecido"])
        reqs.append(_fmt_req(sid, i, 3, i, 3, bg=sc["bg"], fg=sc["fg"], bold=True, h_align="CENTER"))
        # Situação de acesso
        sit = row["situacao"]
        sc2 = SITUACAO_FMT.get(sit, {"bg": GRAY_BG, "fg": GRAY_TEXT})
        reqs.append(_fmt_req(sid, i, 10, i, 10, bg=sc2["bg"], fg=sc2["fg"], bold=True, h_align="CENTER"))
        if i % 2 == 0:
            reqs.append(_fmt_req(sid, i, 1, i, ncols, bg=LIGHT_GRAY))

    _batch(ws.spreadsheet, reqs)


def _build_aulas_quiz(ws, df: pd.DataFrame):
    sid = ws.id
    headers = [
        "Nome", "Email",
        "Aulas Concluídas", "Total Aulas", "% Aulas",
        "Vídeos Assistidos",
        "Quizzes Feitos", "Total Quizzes", "Score Quiz (%)",
        "Tempo (h)", "Tempo/Aula (min)", "Status",
    ]

    df_s = df.sort_values("completed_lessons", ascending=False).copy()
    df_s["status_pt"] = df_s["status"].map(STATUS_PT).fillna("Desconhecido")
    df_s["pct_aulas"] = df_s.apply(
        lambda r: round(r["completed_lessons"] / r["total_lessons"] * 100, 1)
        if r["total_lessons"] > 0 else 0, axis=1
    )
    df_s["min_per_lesson"] = df_s.apply(
        lambda r: round(r["time_spent_minutes"] / r["completed_lessons"], 1)
        if r["completed_lessons"] > 0 else 0, axis=1
    )

    data = [headers] + [
        [
            r["username"] or "",
            r["email"],
            int(r["completed_lessons"]),
            int(r["total_lessons"]),
            r["pct_aulas"],
            int(r.get("video_units_completed", 0) or 0),
            int(r.get("quiz_units_completed", 0) or 0),
            int(r.get("quiz_units_total", 0) or 0),
            r["quiz_score"] if pd.notna(r.get("quiz_score")) else "",
            r["time_spent_hours"],
            r["min_per_lesson"],
            r["status_pt"],
        ]
        for _, r in df_s.iterrows()
    ]

    ws.update("A1", data)
    n = len(df_s)
    ncols = len(headers)

    reqs = [
        _fmt_req(sid, 1, 1, 1, ncols, bg=NAVY, fg=WHITE, bold=True, h_align="CENTER"),
        _freeze_req(sid, rows=1, cols=2),
        _cond_gradient_req(sid, 2, 5, n + 1, 5),   # % Aulas
        _cond_gradient_req(sid, 2, 9, n + 1, 9),   # Score Quiz
        _col_width_req(sid, 0, 160),
        _col_width_req(sid, 1, 220),
        *[_col_width_req(sid, c, 100) for c in range(2, ncols)],
    ]

    for i, (_, row) in enumerate(df_s.iterrows(), start=2):
        if i % 2 == 0:
            reqs.append(_fmt_req(sid, i, 1, i, ncols, bg=LIGHT_GRAY))
        status_pt = STATUS_PT.get(row["status"], "Desconhecido")
        sc = STATUS_FMT.get(status_pt, STATUS_FMT["Desconhecido"])
        reqs.append(_fmt_req(sid, i, 12, i, 12, bg=sc["bg"], fg=sc["fg"], bold=True, h_align="CENTER"))

    _batch(ws.spreadsheet, reqs)


def _build_modulos(ws, section_df: pd.DataFrame):
    """Pivot table: student vs. section completion %."""
    sid = ws.id
    if section_df.empty:
        ws.update("A1", [["Dados de módulos não disponíveis via API."]])
        return

    pivot = section_df.pivot_table(
        index="email", columns="section_title",
        values="section_completion_pct", fill_value=0,
    ).reset_index()

    section_cols = [c for c in pivot.columns if c != "email"]
    headers = ["Email"] + section_cols + ["Média Geral (%)"]

    pivot["media"] = pivot[section_cols].mean(axis=1).round(1)

    data = [headers] + [
        [r["email"]] + [r[c] for c in section_cols] + [r["media"]]
        for _, r in pivot.iterrows()
    ]
    ws.update("A1", data)

    n = len(pivot)
    ncols = len(headers)
    reqs = [
        _fmt_req(sid, 1, 1, 1, ncols, bg=NAVY, fg=WHITE, bold=True, h_align="CENTER"),
        _freeze_req(sid, rows=1, cols=1),
        _col_width_req(sid, 0, 220),
        *[_col_width_req(sid, c, 130) for c in range(1, ncols)],
    ]
    # Gradient for each section column
    for col_i in range(2, ncols + 1):
        reqs.append(_cond_gradient_req(sid, 2, col_i, n + 1, col_i))

    _batch(ws.spreadsheet, reqs)


def _build_em_risco(ws, df: pd.DataFrame):
    sid = ws.id
    at_risk = df[
        (df["completion_percentage"] < 25) |
        (df["days_since_last_access"].fillna(999) > 14)
    ].sort_values("completion_percentage").copy()

    headers = [
        "Nome", "Email", "Status", "Progresso (%)",
        "Dias s/ Acesso", "Último Acesso", "Tempo (h)", "Motivo de Atenção",
    ]

    def _motivo(r):
        motivos = []
        if r["completion_percentage"] < 25:
            motivos.append(f"Progresso baixo ({r['completion_percentage']:.0f}%)")
        days = r.get("days_since_last_access")
        if pd.notna(days) and days > 14:
            motivos.append(f"Sem acesso há {int(days)} dias")
        if pd.isna(days) and r["completion_percentage"] == 0:
            motivos.append("Nunca acessou")
        return " | ".join(motivos) if motivos else "—"

    if at_risk.empty:
        ws.update("A1", [["✅ Todos os alunos estão engajados! Nenhum aluno em risco no momento."]])
        return

    at_risk["status_pt"] = at_risk["status"].map(STATUS_PT).fillna("Desconhecido")
    at_risk["last_fmt"] = pd.to_datetime(at_risk["last_activity"], errors="coerce").dt.strftime("%d/%m/%Y")

    data = [headers] + [
        [
            r["username"] or "",
            r["email"],
            r["status_pt"],
            r["completion_percentage"],
            int(r["days_since_last_access"]) if pd.notna(r.get("days_since_last_access")) else "—",
            r["last_fmt"] if pd.notna(r.get("last_fmt")) else "—",
            r["time_spent_hours"],
            _motivo(r),
        ]
        for _, r in at_risk.iterrows()
    ]

    ws.update("A1", data)
    n = len(at_risk)
    ncols = len(headers)
    reqs = [
        _fmt_req(sid, 1, 1, 1, ncols, bg=_rgb(180, 0, 0), fg=WHITE, bold=True, h_align="CENTER"),
        _freeze_req(sid, rows=1, cols=2),
        _cond_gradient_req(sid, 2, 4, n + 1, 4),
        _col_width_req(sid, 0, 160),
        _col_width_req(sid, 1, 220),
        _col_width_req(sid, 2, 120),
        _col_width_req(sid, 3, 100),
        _col_width_req(sid, 4, 100),
        _col_width_req(sid, 5, 110),
        _col_width_req(sid, 6, 80),
        _col_width_req(sid, 7, 280),
    ]
    for i, (_, row) in enumerate(at_risk.iterrows(), start=2):
        status_pt = STATUS_PT.get(row["status"], "Desconhecido")
        sc = STATUS_FMT.get(status_pt, STATUS_FMT["Desconhecido"])
        reqs.append(_fmt_req(sid, i, 3, i, 3, bg=sc["bg"], fg=sc["fg"], bold=True, h_align="CENTER"))
        reqs.append(_fmt_req(sid, i, 1, i, ncols, bg=RED_BG))

    _batch(ws.spreadsheet, reqs)


def _build_satisfacao(ws, reviews: list):
    sid = ws.id
    if not reviews:
        ws.update("A1", [["Nenhuma avaliação registrada para este curso."]])
        return

    df_r = pd.DataFrame(reviews)
    if "rating" not in df_r.columns or df_r["rating"].dropna().empty:
        ws.update("A1", [["Dados de satisfação não disponíveis via API."]])
        return

    avg = round(df_r["rating"].dropna().mean(), 2)
    dist = df_r["rating"].value_counts().sort_index()

    summary = [
        ["Avaliação Média", avg],
        ["Total de Avaliações", len(df_r)],
        [""],
        ["Nota", "Qtd", "% do Total"],
    ] + [
        [int(k), int(v), round(v / len(df_r) * 100, 1)]
        for k, v in dist.items()
    ] + [
        [""],
        ["AVALIAÇÕES INDIVIDUAIS"],
        ["Email", "Nota", "Comentário", "Data"],
    ] + [
        [
            r.get("user_id", ""),
            r.get("rating", ""),
            r.get("comment", ""),
            r.get("created_at", ""),
        ]
        for r in reviews
    ]

    ws.update("A1", summary)
    reqs = [
        _fmt_req(sid, 1, 1, 1, 2, bg=NAVY, fg=WHITE, bold=True),
        _col_width_req(sid, 0, 200),
        _col_width_req(sid, 1, 80),
        _col_width_req(sid, 2, 350),
        _col_width_req(sid, 3, 120),
    ]
    _batch(ws.spreadsheet, reqs)


# ── Main export ────────────────────────────────────────────────────────────────

def export_to_sheets(course_data: dict, spreadsheet_title: str = None,
                     share_email: str = None) -> str:
    """
    Export all course data to a new Google Sheets spreadsheet.
    Returns the spreadsheet URL.
    """
    course = course_data.get("course", {})
    records = course_data.get("progress_records", [])
    section_records = course_data.get("section_records", [])
    reviews = course_data.get("reviews", [])

    if not records:
        raise ValueError("No student data found. Cannot create spreadsheet.")

    df = pd.DataFrame(records)
    df["completion_percentage"] = pd.to_numeric(df["completion_percentage"], errors="coerce").fillna(0)
    df["time_spent_hours"] = pd.to_numeric(df["time_spent_hours"], errors="coerce").fillna(0)
    df["time_spent_minutes"] = pd.to_numeric(df["time_spent_minutes"], errors="coerce").fillna(0)
    df["completed_lessons"] = pd.to_numeric(df["completed_lessons"], errors="coerce").fillna(0)
    df["total_lessons"] = pd.to_numeric(df["total_lessons"], errors="coerce").fillna(0)

    section_df = pd.DataFrame(section_records) if section_records else pd.DataFrame()

    title = spreadsheet_title or f"Dashboard — {course.get('title', 'Curso')} — {datetime.now().strftime('%d/%m/%Y')}"

    gc = _get_gc()
    print(f"[sheets] Criando planilha: {title}")
    ss = gc.create(title)

    # Share with user if email provided
    if share_email:
        ss.share(share_email, perm_type="user", role="writer")
        print(f"[sheets] Planilha compartilhada com: {share_email}")

    sheets_config = [
        ("📊 Resumo Geral",    lambda ws: _build_resumo(ws, course, df, reviews)),
        ("👥 Roster de Alunos", lambda ws: _build_alunos(ws, df)),
        ("📅 Presença & Acesso", lambda ws: _build_presenca(ws, df)),
        ("📝 Aulas & Quizzes",  lambda ws: _build_aulas_quiz(ws, df)),
        ("🔥 Por Módulo",       lambda ws: _build_modulos(ws, section_df)),
        ("⚠️ Em Risco",         lambda ws: _build_em_risco(ws, df)),
        ("⭐ Satisfação",       lambda ws: _build_satisfacao(ws, reviews)),
    ]

    # Rename first sheet and create the rest
    first_ws = ss.sheet1
    first_ws.update_title(sheets_config[0][0])
    worksheets = [first_ws]

    for name, _ in sheets_config[1:]:
        worksheets.append(ss.add_worksheet(title=name, rows=500, cols=30))

    # Populate each sheet
    for ws, (name, builder) in zip(worksheets, sheets_config):
        print(f"[sheets] Populando aba: {name}")
        builder(ws)

    url = ss.url
    print(f"\n[sheets] ✅ Planilha criada com sucesso!")
    print(f"[sheets] 🔗 {url}\n")
    return url
