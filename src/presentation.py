from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.utils import simpleSplit
    from reportlab.pdfgen import canvas
except ImportError as exc:
    raise SystemExit(
        "Dependência ausente: reportlab. Instale com: "
        "python -m pip install reportlab pandas openpyxl"
    ) from exc


INPUT_FILE = "opps_corrigido.xlsx"
OUTPUT_FILE = "apresentacao.pdf"

BASE_ROWS = 413
FINAL_ROWS = 261
DUPLICATES_REMOVED = 152
UNKNOWN = "Unknown"


@dataclass
class Slide:
    title: str
    subtitle: str
    bullets: list[str]


class ExecutiveDeck:
    def __init__(self, output_path: Path) -> None:
        self.output_path = output_path
        self.pdf = canvas.Canvas(str(output_path), pagesize=A4)
        self.page_width, self.page_height = A4
        self.margin_left = 48
        self.margin_right = 48
        self.margin_top = 58
        self.margin_bottom = 56
        self.content_width = self.page_width - self.margin_left - self.margin_right

    def draw_slide(self, slide: Slide, page_number: int, total_pages: int) -> None:
        y = self.page_height - self.margin_top

        self.pdf.setFont("Helvetica-Bold", 20)
        self.pdf.drawString(self.margin_left, y, slide.title)
        y -= 30

        self.pdf.setFont("Helvetica", 11)
        y = self._draw_wrapped_text(
            slide.subtitle,
            x=self.margin_left,
            y=y,
            font_name="Helvetica",
            font_size=11,
            line_height=15,
            spacing_after=12,
        )

        y -= 4
        y = self._draw_bullets(slide.bullets, start_y=y, bullet_indent=14, line_height=14)

        self._draw_footer(page_number=page_number, total_pages=total_pages)
        self.pdf.showPage()

    def _draw_wrapped_text(
        self,
        text: str,
        x: float,
        y: float,
        font_name: str,
        font_size: int,
        line_height: int,
        spacing_after: int = 0,
    ) -> float:
        self.pdf.setFont(font_name, font_size)
        lines = simpleSplit(text, font_name, font_size, self.content_width)
        for line in lines:
            if y < self.margin_bottom:
                break
            self.pdf.drawString(x, y, line)
            y -= line_height
        return y - spacing_after

    def _draw_bullets(self, bullets: Iterable[str], start_y: float, bullet_indent: int, line_height: int) -> float:
        y = start_y
        bullet_symbol = u"\u2022"
        text_x = self.margin_left + bullet_indent
        available_width = self.content_width - bullet_indent

        self.pdf.setFont("Helvetica", 11)
        for bullet in bullets:
            wrapped = simpleSplit(bullet, "Helvetica", 11, available_width)
            if not wrapped:
                continue

            if y < self.margin_bottom + 2 * line_height:
                break

            self.pdf.drawString(self.margin_left, y, bullet_symbol)
            self.pdf.drawString(text_x, y, wrapped[0])
            y -= line_height

            for continuation in wrapped[1:]:
                if y < self.margin_bottom:
                    break
                self.pdf.drawString(text_x, y, continuation)
                y -= line_height

            y -= 6

        return y

    def _draw_footer(self, page_number: int, total_pages: int) -> None:
        self.pdf.setFont("Helvetica", 9)
        footer_text = f"Case Monks - Apresentacao Executiva | Pagina {page_number}/{total_pages}"
        self.pdf.drawString(self.margin_left, self.margin_bottom - 16, footer_text)

    def save(self) -> None:
        self.pdf.save()


def build_business_insights(base_dir: Path) -> list[str]:
    input_path = base_dir / INPUT_FILE
    if not input_path.exists():
        return [
            "A base limpa reduz inflacao de pipeline e melhora confiabilidade do forecast comercial.",
            "Oportunidades em estagios iniciais devem ser monitoradas com criterios minimos de preenchimento.",
            "Padronizacao de origem e moeda melhora comparabilidade entre times e regioes.",
            "Rotina semanal de qualidade de dados reduz retrabalho comercial e ruido analitico.",
        ]

    df = pd.read_excel(input_path)

    if "Amount" not in df.columns:
        df["Amount"] = 0.0
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0.0)

    if "Stage" not in df.columns:
        df["Stage"] = UNKNOWN
    df["Stage"] = df["Stage"].fillna(UNKNOWN).astype(str).str.strip().replace("", UNKNOWN)

    if "Lead_Source" not in df.columns:
        df["Lead_Source"] = UNKNOWN
    df["Lead_Source"] = (
        df["Lead_Source"].fillna(UNKNOWN).astype(str).str.strip().replace("", UNKNOWN)
    )

    total_amount = float(df["Amount"].sum())
    stage_summary = (
        df.groupby("Stage", dropna=False)["Amount"]
        .sum()
        .reset_index()
        .sort_values("Amount", ascending=False)
    )
    source_summary = (
        df.groupby("Lead_Source", dropna=False)["Amount"]
        .sum()
        .reset_index()
        .sort_values("Amount", ascending=False)
    )

    insights: list[str] = []
    if not stage_summary.empty and total_amount > 0:
        top_stage = stage_summary.iloc[0]
        top_stage_share = float(top_stage["Amount"]) / total_amount * 100
        insights.append(
            f"O stage '{top_stage['Stage']}' concentra {top_stage_share:.1f}% do valor total do pipeline e deve ter governanca reforcada de avancos."
        )

        top3_share = float(stage_summary.head(3)["Amount"].sum()) / total_amount * 100
        insights.append(
            f"Os 3 maiores estagios somam {top3_share:.1f}% do valor total; foco nesses pontos aumenta previsibilidade de receita."
        )

    if not source_summary.empty and total_amount > 0:
        top_source = source_summary.iloc[0]
        source_share = float(top_source["Amount"]) / total_amount * 100
        insights.append(
            f"A origem '{top_source['Lead_Source']}' responde por {source_share:.1f}% do valor, sugerindo priorizacao comercial nesse canal."
        )

    unknown_ratio = (
        (df["Lead_Source"].str.lower() == UNKNOWN.lower()).mean() * 100
        if len(df) > 0
        else 0.0
    )
    insights.append(
        f"{unknown_ratio:.1f}% dos registros estao com Lead_Source='Unknown'; reduzir esse indice melhora atribuicao de performance e ROI de acquisicao."
    )

    return insights[:5]


def build_slides(insights: list[str]) -> list[Slide]:
    page_1 = Slide(
        title="Fase 5 e 6 - Encerramento Executivo",
        subtitle=(
            "Contexto do case: consolidacao e saneamento da base de oportunidades comerciais "
            "para suportar decisao de negocio e previsao de receita."
        ),
        bullets=[
            f"Base original recebida: {BASE_ROWS} linhas.",
            "Objetivo: remover inconsistencias e criar uma camada analitica confiavel para operacao comercial.",
            "Escopo tecnico: limpeza, relatorio de qualidade, analise comercial e apresentacao executiva.",
            "Foco de negocio: confiabilidade de forecast, priorizacao comercial e governanca de CRM.",
        ],
    )

    page_2 = Slide(
        title="Problema dos Dados",
        subtitle=(
            "A qualidade da base impactava diretamente os indicadores comerciais e a leitura do pipeline."
        ),
        bullets=[
            f"Evolucao do volume: {BASE_ROWS} linhas brutas -> {FINAL_ROWS} linhas limpas.",
            f"Duplicatas removidas: {DUPLICATES_REMOVED} registros (aprox. {DUPLICATES_REMOVED / BASE_ROWS * 100:.1f}% da base original).",
            "Impacto direto sem tratamento: inflacao de pipeline, forecast superestimado e esforco duplicado do time de vendas.",
            "Regra adotada para duplicatas: maior completude de campos e, em empate, Created_Date mais antiga.",
            "Resultado: base unica por oportunidade, com menor ruido analitico e maior rastreabilidade.",
        ],
    )

    page_3 = Slide(
        title="Principais Achados de Negocio",
        subtitle="Insights acionaveis extraidos da base limpa para orientar alocacao e execucao comercial.",
        bullets=insights,
    )

    page_4 = Slide(
        title="Recomendacoes de Processo (Blindagem CRM/Salesforce)",
        subtitle="Ajustes de processo para evitar regressao de qualidade e manter previsibilidade do funil.",
        bullets=[
            "Criar validation rules obrigando preenchimento minimo por stage (Amount, Close_Date, Owner e fonte).",
            "Bloquear novos duplicados com chave canonica (Opportunity_ID + Account_ID) e rotina automatica de merge.",
            "Padronizar campos de moeda com lista controlada ISO-3 e sincronizacao entre Amount_Currency e Total_Product_Amount_Currency.",
            "Instituir monitor semanal de qualidade (duplicatas, nulos criticos, valores invalidos, datas inconsistentes) com SLA por squad.",
            "Definir data owner comercial e ritual mensal de melhoria continua da qualidade no CRM.",
        ],
    )

    page_5 = Slide(
        title="Conclusao e Aprendizados",
        subtitle="Fechamento executivo da iniciativa e direcao para sustentacao.",
        bullets=[
            "A limpeza elevou a confiabilidade da base para uso em acompanhamento de pipeline e tomada de decisao.",
            "A principal alavanca de valor foi a deduplicacao criteriosa, eliminando vies operacional e analitico.",
            "Tratamento explicito de nulos e padronizacao cambial sem FX reduziram ambiguidades sem criar distorcoes financeiras.",
            "Proximo ciclo recomendado: acoplar controles no CRM para prevenir erro na origem, nao apenas corrigir na saida.",
        ],
    )

    return [page_1, page_2, page_3, page_4, page_5]


def main() -> None:
    base_dir = Path(__file__).resolve().parent.parent
    output_path = base_dir / OUTPUT_FILE

    insights = build_business_insights(base_dir)
    slides = build_slides(insights)

    deck = ExecutiveDeck(output_path=output_path)
    total_pages = len(slides)
    for idx, slide in enumerate(slides, start=1):
        deck.draw_slide(slide=slide, page_number=idx, total_pages=total_pages)
    deck.save()

    print(f"Apresentacao gerada em: {output_path}")


if __name__ == "__main__":
    main()

