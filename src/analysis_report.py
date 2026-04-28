from __future__ import annotations

from pathlib import Path
import re

import pandas as pd
import plotly.graph_objects as go
from plotly.offline import plot


INPUT_FILE = "opps_corrigido.xlsx"
OUTPUT_FILE = "analise.html"


def parse_numeric_series(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)

    cleaned = (
        series.astype(str)
        .str.replace("\u00A0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(r"[^\d,.\-()]", "", regex=True)
    )

    cleaned = cleaned.apply(normalize_numeric_text)
    numeric = pd.to_numeric(cleaned, errors="coerce")
    return numeric.fillna(0.0)


def normalize_numeric_text(value: str) -> str:
    s = value.strip()
    if not s:
        return ""

    negative = s.startswith("(") and s.endswith(")")
    if negative:
        s = s[1:-1]

    s = re.sub(r"[^0-9,.\-]", "", s)
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) != 3:
            s = parts[0] + "." + parts[1]
        else:
            s = "".join(parts)
    elif "." in s and s.count(".") > 1:
        parts = s.split(".")
        s = "".join(parts[:-1]) + "." + parts[-1]

    if negative and not s.startswith("-"):
        s = "-" + s
    return s


def parse_date_series(series: pd.Series) -> pd.Series:
    parsed = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")

    numeric_values = pd.to_numeric(series, errors="coerce")
    serial_mask = numeric_values.between(20000, 80000, inclusive="both")
    if serial_mask.any():
        parsed.loc[serial_mask] = pd.to_datetime(
            numeric_values.loc[serial_mask],
            unit="D",
            origin="1899-12-30",
            errors="coerce",
        )

    unresolved = parsed.isna()
    if unresolved.any():
        text = series.loc[unresolved].astype(str).str.strip()
        first_pass = pd.to_datetime(text, errors="coerce", dayfirst=False)
        second_mask = first_pass.isna()
        if second_mask.any():
            first_pass.loc[second_mask] = pd.to_datetime(
                text.loc[second_mask], errors="coerce", dayfirst=True
            )
        parsed.loc[unresolved] = first_pass

    return parsed


def choose_performance_dimension(df: pd.DataFrame) -> str:
    candidates = [col for col in ("Lead_Source", "Lead_Office") if col in df.columns]
    if not candidates:
        return "Lead_Source"

    best_col = candidates[0]
    best_score = -1
    for col in candidates:
        non_unknown = (df[col].astype(str).str.strip().str.lower() != "unknown").sum()
        unique_count = df[col].nunique(dropna=True)
        score = int(non_unknown) * 1000 + int(unique_count)
        if score > best_score:
            best_score = score
            best_col = col
    return best_col


def load_and_prepare_data(base_dir: Path) -> pd.DataFrame:
    input_path = base_dir / INPUT_FILE
    if not input_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {input_path}")

    df = pd.read_excel(input_path)

    for col in ("Stage", "Amount", "Close_Date", "Lead_Source", "Lead_Office"):
        if col not in df.columns:
            df[col] = pd.NA

    df["Stage"] = df["Stage"].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    df["Amount"] = parse_numeric_series(df["Amount"])
    df["Close_Date"] = parse_date_series(df["Close_Date"])
    df["Lead_Source"] = (
        df["Lead_Source"].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    )
    df["Lead_Office"] = (
        df["Lead_Office"].fillna("Unknown").astype(str).str.strip().replace("", "Unknown")
    )

    return df


def build_stage_chart(stage_summary: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=stage_summary["Stage"],
            y=stage_summary["Valor_Total"],
            name="Valor Total",
            marker_color="#0b66ff",
            yaxis="y1",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=stage_summary["Stage"],
            y=stage_summary["Volume"],
            name="Volume",
            mode="lines+markers",
            marker_color="#f97316",
            yaxis="y2",
        )
    )
    fig.update_layout(
        title="Pipeline por Stage",
        xaxis_title="Stage",
        yaxis=dict(title="Valor Total"),
        yaxis2=dict(title="Volume", overlaying="y", side="right"),
        legend=dict(orientation="h", y=1.12),
        template="plotly_white",
        margin=dict(l=40, r=40, t=70, b=40),
    )
    return fig


def build_monthly_chart(monthly_summary: pd.DataFrame) -> go.Figure:
    fig = go.Figure(
        go.Scatter(
            x=monthly_summary["Mes"],
            y=monthly_summary["Valor_Total"],
            mode="lines+markers",
            line=dict(color="#0d9488", width=3),
            marker=dict(size=7),
            name="Valor Total",
        )
    )
    fig.update_layout(
        title="Evolução Mensal do Pipeline (Close Date)",
        xaxis_title="Mês",
        yaxis_title="Valor Total",
        template="plotly_white",
        margin=dict(l=40, r=30, t=70, b=40),
    )
    return fig


def build_performance_chart(perf_summary: pd.DataFrame, dimension: str) -> go.Figure:
    fig = go.Figure(
        go.Bar(
            x=perf_summary["Valor_Total"],
            y=perf_summary[dimension],
            orientation="h",
            marker_color="#2563eb",
            text=perf_summary["Volume"],
            textposition="outside",
            name="Valor Total",
        )
    )
    fig.update_layout(
        title=f"Performance por {dimension}",
        xaxis_title="Valor Total",
        yaxis_title=dimension,
        template="plotly_white",
        margin=dict(l=40, r=30, t=70, b=40),
    )
    return fig


def generate_insights(
    stage_summary: pd.DataFrame, monthly_summary: pd.DataFrame, perf_summary: pd.DataFrame, dimension: str
) -> list[str]:
    insights: list[str] = []

    total_value = float(stage_summary["Valor_Total"].sum()) if not stage_summary.empty else 0.0
    if not stage_summary.empty and total_value > 0:
        top_stage = stage_summary.iloc[0]
        share = (float(top_stage["Valor_Total"]) / total_value) * 100
        insights.append(
            f"O estágio '{top_stage['Stage']}' concentra {share:.1f}% do valor do pipeline; "
            "é o principal ponto de alavancagem (ou risco) para previsão."
        )

    if not stage_summary.empty:
        top3_value = float(stage_summary.head(3)["Valor_Total"].sum())
        share_top3 = (top3_value / total_value * 100) if total_value > 0 else 0.0
        insights.append(
            f"Os 3 maiores estágios somam {share_top3:.1f}% do valor total; priorizar hygiene e governança nesses estágios melhora a confiabilidade do forecast."
        )

    if len(monthly_summary) >= 2:
        last = float(monthly_summary.iloc[-1]["Valor_Total"])
        prev = float(monthly_summary.iloc[-2]["Valor_Total"])
        delta = last - prev
        pct = ((delta / prev) * 100) if prev else 0.0
        trend_word = "aceleração" if delta >= 0 else "desaceleração"
        insights.append(
            f"Entre os dois últimos meses houve {trend_word} de {pct:.1f}% no valor do pipeline; ajuste metas e capacidade comercial conforme essa direção."
        )
    else:
        insights.append(
            "A série mensal ainda é curta para tendência robusta; manter acompanhamento mensal para estabilizar leitura de sazonalidade."
        )

    if not perf_summary.empty:
        top_dim = perf_summary.iloc[0]
        avg_ticket = float(top_dim["Ticket_Medio"])
        insights.append(
            f"O melhor desempenho por {dimension} é '{top_dim[dimension]}', com ticket médio de {avg_ticket:,.2f}; use como referência de perfil para prospecção."
        )

    if len(insights) < 3:
        insights.append(
            "Padronize cadastro de Stage, origem e valores para reduzir ruído analítico e melhorar decisões de alocação do time."
        )

    return insights[:5]


def build_html(
    stage_chart_div: str,
    monthly_chart_div: str,
    performance_chart_div: str,
    insights: list[str],
    performance_dimension: str,
    total_rows: int,
) -> str:
    insights_html = "\n".join(f"<li>{text}</li>" for text in insights)

    return f"""<!doctype html>
<html lang="pt-BR">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Análise Comercial - Oportunidades</title>
  <style>
    :root {{
      --bg: #f3f6fb;
      --surface: #ffffff;
      --border: #d6dfed;
      --text: #13243a;
      --muted: #5b6b80;
      --shadow: 0 10px 24px rgba(9, 25, 58, 0.08);
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      background: linear-gradient(180deg, #f8fbff 0%, var(--bg) 62%);
      color: var(--text);
      font-family: "Segoe UI", Tahoma, Arial, sans-serif;
    }}
    .container {{
      max-width: 1200px;
      margin: 0 auto;
      padding: 22px 14px 34px;
    }}
    .hero, .section {{
      background: var(--surface);
      border: 1px solid var(--border);
      border-radius: 14px;
      box-shadow: var(--shadow);
    }}
    .hero {{
      padding: 20px;
      margin-bottom: 14px;
    }}
    h1 {{
      margin: 0 0 8px;
      font-size: 1.7rem;
    }}
    h2 {{
      margin: 0 0 10px;
      font-size: 1.2rem;
    }}
    .muted {{
      margin: 0;
      color: var(--muted);
    }}
    .section {{
      padding: 16px;
      margin-top: 14px;
    }}
    .insights {{
      margin: 0;
      padding-left: 18px;
    }}
    .insights li + li {{
      margin-top: 8px;
    }}
    .badge {{
      display: inline-block;
      margin-top: 10px;
      padding: 4px 10px;
      border-radius: 999px;
      border: 1px solid #c7d9ff;
      background: #edf4ff;
      color: #1042a4;
      font-size: 0.82rem;
    }}
    @media (max-width: 700px) {{
      h1 {{ font-size: 1.35rem; }}
      .container {{ padding: 12px 8px 24px; }}
    }}
  </style>
</head>
<body>
  <div class="container">
    <section class="hero">
      <h1>Análise de Pipeline Comercial</h1>
      <p class="muted">Base única utilizada: <strong>{INPUT_FILE}</strong>. Registros analisados: <strong>{total_rows}</strong>.</p>
      <p class="muted">Dimensão de performance selecionada automaticamente: <strong>{performance_dimension}</strong>.</p>
      <div class="badge">Fase 4 · Simplicidade > Complexidade</div>
    </section>

    <section class="section">
      <h2>1) Pipeline por Stage (volume e valor)</h2>
      {stage_chart_div}
    </section>

    <section class="section">
      <h2>2) Evolução por Close Date (mês/ano)</h2>
      {monthly_chart_div}
    </section>

    <section class="section">
      <h2>3) Performance por {performance_dimension}</h2>
      {performance_chart_div}
    </section>

    <section class="section">
      <h2>4) Insights acionáveis</h2>
      <ul class="insights">
        {insights_html}
      </ul>
    </section>
  </div>
</body>
</html>
"""


def main() -> None:
    base_dir = Path(__file__).resolve().parent.parent
    output_path = base_dir / OUTPUT_FILE

    df = load_and_prepare_data(base_dir)
    performance_dimension = choose_performance_dimension(df)

    stage_summary = (
        df.groupby("Stage", dropna=False)
        .agg(Volume=("Stage", "size"), Valor_Total=("Amount", "sum"))
        .reset_index()
        .sort_values("Valor_Total", ascending=False)
    )

    monthly_df = df.dropna(subset=["Close_Date"]).copy()
    monthly_df["Mes"] = monthly_df["Close_Date"].dt.to_period("M").dt.to_timestamp()
    monthly_summary = (
        monthly_df.groupby("Mes", dropna=False)["Amount"]
        .sum()
        .rename("Valor_Total")
        .reset_index()
        .sort_values("Mes")
    )

    perf_summary = (
        df.groupby(performance_dimension, dropna=False)
        .agg(
            Volume=(performance_dimension, "size"),
            Valor_Total=("Amount", "sum"),
            Ticket_Medio=("Amount", "mean"),
        )
        .reset_index()
        .sort_values("Valor_Total", ascending=False)
        .head(12)
    )

    stage_fig = build_stage_chart(stage_summary)
    monthly_fig = build_monthly_chart(monthly_summary)
    performance_fig = build_performance_chart(perf_summary, performance_dimension)

    stage_chart_div = plot(
        stage_fig,
        output_type="div",
        include_plotlyjs=True,
        config={"displayModeBar": False, "responsive": True},
    )
    monthly_chart_div = plot(
        monthly_fig,
        output_type="div",
        include_plotlyjs=False,
        config={"displayModeBar": False, "responsive": True},
    )
    performance_chart_div = plot(
        performance_fig,
        output_type="div",
        include_plotlyjs=False,
        config={"displayModeBar": False, "responsive": True},
    )

    insights = generate_insights(stage_summary, monthly_summary, perf_summary, performance_dimension)
    html = build_html(
        stage_chart_div=stage_chart_div,
        monthly_chart_div=monthly_chart_div,
        performance_chart_div=performance_chart_div,
        insights=insights,
        performance_dimension=performance_dimension,
        total_rows=len(df),
    )

    output_path.write_text(html, encoding="utf-8")
    print(f"Análise gerada em: {output_path}")


if __name__ == "__main__":
    main()
