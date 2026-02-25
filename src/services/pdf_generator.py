"""
Servicio de generación de PDFs para informes con formato profesional.
"""

from typing import Dict, Optional
from datetime import datetime
from pathlib import Path
import sys
import os
import tempfile
import shutil

from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.units import cm, inch
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from PIL import Image
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from urllib.parse import quote

from src.core.report_outline import get_content_outline
from src.core.result_schemes import RESULT_SCHEMES


DEFAULT_TECHNICAL_TEAM = [
    {
        "name": "Licda. Yara Lizeth Pérez A.",
        "details": [
            "Fonoaudióloga,",
            "Registro 180.",
        ],
    },
    {
        "name": "Licda. Stephanie María Thorne.",
        "details": [
            "Terapeuta Respiratoria,",
            "Registro 124.",
        ],
    },
]


class PDFGenerator:
    """Generador de PDFs para informes clínicos."""

    def __init__(self, template_path: Optional[str] = None):
        self.template_path = template_path
        self.page_width, self.page_height = landscape(letter)
        self.left_margin = 2.5 * cm
        self.right_margin = 2.5 * cm
        self.top_margin = 3.0 * cm
        self.bottom_margin = 3.0 * cm

    def generate(self, report_data: Dict, output_path: str, logo_path: Optional[str] = None) -> bool:
        """Genera el PDF del informe aplicando el logo como marca de agua."""
        temp_pdf_path: Optional[str] = None
        try:
            output_dir = os.path.dirname(output_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)

            temp_pdf = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            temp_pdf_path = temp_pdf.name
            temp_pdf.close()

            pdf_canvas = canvas.Canvas(temp_pdf_path, pagesize=landscape(letter))

            watermark_logo = logo_path if logo_path and os.path.exists(logo_path) else None
            if watermark_logo:
                self._add_watermark_logo(pdf_canvas, watermark_logo)

            self._draw_header(pdf_canvas, report_data)
            self._draw_general_info(pdf_canvas, report_data)
            self._draw_footer(pdf_canvas, page_number=1)
            last_page = self._draw_table_of_contents(
                pdf_canvas,
                report_data,
                watermark_logo=watermark_logo,
                start_page=2,
            )
            company_last_page = self._draw_company_profile_page(
                pdf_canvas,
                report_data,
                watermark_logo=watermark_logo,
                page_number=last_page + 1,
            )

            evaluated_people = report_data.get("evaluated") or []
            grouped_entries = self._group_evaluated_entries(evaluated_people)
            attachment_sections = []
            current_page = company_last_page
            for dataset_key in self._determine_result_dataset_keys(report_data.get("type", "")):
                entries = grouped_entries.get(dataset_key, [])
                if not entries:
                    continue
                current_page = self._draw_results_table_page(
                    pdf_canvas,
                    report_data,
                    dataset_key,
                    entries,
                    page_number=current_page + 1,
                    watermark_logo=watermark_logo,
                )
                current_page = self._draw_results_statistics_page(
                    pdf_canvas,
                    report_data,
                    dataset_key,
                    entries,
                    page_number=current_page + 1,
                    watermark_logo=watermark_logo,
                )

            conclusion_text = (report_data.get("conclusion") or "").strip()
            if conclusion_text:
                current_page = self._draw_conclusion_page(
                    pdf_canvas,
                    report_data,
                    conclusion_text,
                    page_number=current_page + 1,
                    watermark_logo=watermark_logo,
                )

            recommendations_text = (report_data.get("recommendations") or "").strip()
            if recommendations_text:
                current_page = self._draw_recommendations_page(
                    pdf_canvas,
                    report_data,
                    recommendations_text,
                    page_number=current_page + 1,
                    watermark_logo=watermark_logo,
                )

            technical_team = self._resolve_technical_team(report_data)
            if technical_team:
                current_page = self._draw_technical_team_page(
                    pdf_canvas,
                    report_data,
                    technical_team,
                    page_number=current_page + 1,
                    watermark_logo=watermark_logo,
                )

            calibration_certificates = self._resolve_calibration_certificates(report_data)
            if calibration_certificates:
                current_page = self._draw_calibration_certificates_page(
                    pdf_canvas,
                    report_data,
                    calibration_certificates,
                    page_number=current_page + 1,
                    watermark_logo=watermark_logo,
                )

            calibration_files = self._resolve_calibration_files(report_data)
            if calibration_files:
                cover_start = current_page + 1
                current_page = self._draw_attachment_cover_page(
                    pdf_canvas,
                    report_data,
                    "CERTIFICADOS DE CALIBRACIÓN ADJUNTOS:",
                    "Los documentos oficiales emitidos por los laboratorios se incluyen a continuación.",
                    calibration_files,
                    page_number=current_page + 1,
                    watermark_logo=watermark_logo,
                    folder_path=self._resolve_attachment_folder(report_data, "calibration", calibration_files),
                )
                attachment_sections.append({
                    "files": calibration_files,
                    "cover_end": current_page,
                    "cover_start": cover_start,
                })

            audiogram_files = self._resolve_file_list(report_data.get("audiogram_files"))
            if audiogram_files:
                cover_start = current_page + 1
                current_page = self._draw_attachment_cover_page(
                    pdf_canvas,
                    report_data,
                    "AUDIOGRAMAS ADJUNTOS:",
                    "Se anexan los audiogramas exportados desde el equipo de medición.",
                    audiogram_files,
                    page_number=current_page + 1,
                    watermark_logo=watermark_logo,
                    folder_path=self._resolve_attachment_folder(report_data, "audiometria", audiogram_files),
                )
                attachment_sections.append({
                    "files": audiogram_files,
                    "cover_end": current_page,
                    "cover_start": cover_start,
                })

            spirometry_files = self._resolve_file_list(report_data.get("spirometry_files"))
            if spirometry_files:
                cover_start = current_page + 1
                current_page = self._draw_attachment_cover_page(
                    pdf_canvas,
                    report_data,
                    "REPORTES DE ESPIROMETRÍA ADJUNTOS:",
                    "Se incluyen los reportes completos generados por el espirómetro.",
                    spirometry_files,
                    page_number=current_page + 1,
                    watermark_logo=watermark_logo,
                    folder_path=self._resolve_attachment_folder(report_data, "espirometria", spirometry_files),
                )
                attachment_sections.append({
                    "files": spirometry_files,
                    "cover_end": current_page,
                    "cover_start": cover_start,
                })

            attendance_files = self._resolve_file_list(report_data.get("attendance_files"))
            if attendance_files:
                cover_start = current_page + 1
                current_page = self._draw_attachment_cover_page(
                    pdf_canvas,
                    report_data,
                    "LISTADO DE ASISTENCIA:",
                    "Los listados firmados se anexan inmediatamente después de esta página.",
                    attendance_files,
                    page_number=current_page + 1,
                    watermark_logo=watermark_logo,
                    folder_path=self._resolve_attachment_folder(report_data, "attendance", attendance_files),
                )
                attachment_sections.append({
                    "files": attendance_files,
                    "cover_end": current_page,
                    "cover_start": cover_start,
                })

            pdf_canvas.save()

            protocol_pdf = None
            report_type = report_data.get("type", "")
            normalized_type = (report_type or "").lower()
            has_audio = "audiometr" in normalized_type
            has_espiro = "espiro" in normalized_type

            if has_audio and has_espiro:
                protocol_pdf = [
                    self._render_audiometry_protocol_pdf(report_data, watermark_logo),
                    self._render_spirometry_protocol_pdf(report_data, watermark_logo),
                ]
            elif self._should_include_audiometry_protocol(report_type):
                protocol_pdf = self._render_audiometry_protocol_pdf(report_data, watermark_logo)
            elif self._should_include_spirometry_protocol(report_type):
                protocol_pdf = self._render_spirometry_protocol_pdf(report_data, watermark_logo)

            if attachment_sections or protocol_pdf:
                try:
                    if isinstance(protocol_pdf, list):
                        trailing = [path for path in protocol_pdf if path]
                    else:
                        trailing = [protocol_pdf] if protocol_pdf else []
                    self._append_supporting_pdfs(temp_pdf_path, output_path, attachment_sections, trailing)
                    temp_pdf_path = None
                except Exception as merge_exc:  # pragma: no cover - merge fallback
                    print(f"Advertencia al adjuntar anexos: {merge_exc}")
                    shutil.move(temp_pdf_path, output_path)
                    temp_pdf_path = None
            else:
                shutil.move(temp_pdf_path, output_path)
                temp_pdf_path = None

            return True
        except Exception as exc:  # pragma: no cover - logging simple error
            print(f"Error al generar PDF: {exc}")
            return False
        finally:
            if temp_pdf_path and os.path.exists(temp_pdf_path):
                os.remove(temp_pdf_path)

    def _add_watermark_logo(self, canvas_obj: canvas.Canvas, logo_path: str) -> None:
        """Coloca el logo de fondo con opacidad muy baja."""

        try:
            img = Image.open(logo_path)

            max_width = self.page_width * 0.60
            ratio = max_width / img.width
            new_height = int(img.height * ratio)
            img = img.resize((int(max_width), new_height), Image.Resampling.LANCZOS)

            img = img.convert("RGBA")
            alpha = img.split()[3]
            alpha = alpha.point(lambda p: int(p * 0.06))
            img.putalpha(alpha)

            temp_img_path = "temp_watermark.png"
            img.save(temp_img_path, format="PNG")

            x = (self.page_width - max_width) / 2
            y = (self.page_height - new_height) / 2

            canvas_obj.drawImage(
                temp_img_path,
                x,
                y,
                width=max_width,
                height=new_height,
                mask="auto",
            )

            if os.path.exists(temp_img_path):
                os.remove(temp_img_path)
        except Exception as exc:  # pragma: no cover
            print(f"Error al agregar watermark: {exc}")

    def _draw_header_branding(self, pdf_canvas: canvas.Canvas, report_data: Dict) -> float:
        """Dibuja el encabezado superior con logo y datos de contacto."""

        y = self.page_height - self.top_margin

        pdf_canvas.setFont("Helvetica-Bold", 11)
        pdf_canvas.drawCentredString(
            self.page_width / 2,
            y,
            "CENTRO DE ATENCIÓN INTEGRAL TERAPÉUTICO PANAMÁ (CAIT PANAMÁ)",
        )

        y -= 0.35 * inch
        pdf_canvas.setFont("Helvetica", 8)
        pdf_canvas.drawCentredString(self.page_width / 2, y, "Ruc.: 8-749-2471 B.V. 17")

        y -= 0.25 * inch
        pdf_canvas.drawCentredString(
            self.page_width / 2,
            y,
            "Teléfono : 6022-9400 / 6671-4015 correo electrónico: caitpanama@gmail.com",
        )

        y -= 0.25 * inch
        pdf_canvas.setFont("Helvetica", 7)
        pdf_canvas.drawCentredString(
            self.page_width / 2,
            y,
            "Panamá,Panamá Oeste",
        )

        return y

    def _draw_header(self, pdf_canvas: canvas.Canvas, report_data: Dict) -> None:
        """Dibuja el encabezado institucional con espaciado más holgado."""

        y = self._draw_header_branding(pdf_canvas, report_data)

        y -= 0.5 * inch
        pdf_canvas.setFont("Helvetica-Bold", 14)
        pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
        pdf_canvas.drawCentredString(self.page_width / 2, y, "INFORME:")

        y -= 0.35 * inch
        pdf_canvas.setFont("Helvetica-Bold", 13)
        report_type_text = f"EVALUACIÓN DE {report_data.get('type', 'evaluación').upper()}S"
        pdf_canvas.drawCentredString(self.page_width / 2, y, report_type_text)

        y -= 0.45 * inch
        pdf_canvas.setFont("Helvetica-Bold", 11)
        pdf_canvas.drawCentredString(self.page_width / 2, y, "EMPRESA:")

        y -= 0.3 * inch
        pdf_canvas.setFont("Helvetica-Bold", 11)
        pdf_canvas.setFillColor(colors.black)
        pdf_canvas.drawCentredString(self.page_width / 2, y, report_data.get("company", "N/A").upper())

        y -= 0.4 * inch
        pdf_canvas.setFont("Helvetica-Bold", 10)
        pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
        pdf_canvas.drawCentredString(
            self.page_width / 2,
            y,
            f"ESTUDIO OCUPACIONAL: \"{report_data.get('location', 'N/A').upper()}\"",
        )

        y -= 0.45 * inch
        pdf_canvas.drawCentredString(self.page_width / 2, y, "PREPARADO POR:")

        technical_team = self._resolve_technical_team(report_data)
        is_combined_report = (
            "audiometr" in (report_data.get("type", "").lower())
            and "espiro" in (report_data.get("type", "").lower())
        )

        if is_combined_report and len(technical_team) >= 2:
            left_member = technical_team[0]
            right_member = technical_team[1]

            def _draw_member_block(member: Dict, center_x: float, top_y: float) -> None:
                name = (member.get("name") or "N/A").upper()
                details = [
                    str(detail).upper()
                    for detail in (member.get("details") or [])
                    if str(detail).strip()
                ]

                line_y = top_y
                pdf_canvas.setFont("Helvetica-Bold", 11)
                pdf_canvas.setFillColor(colors.HexColor("#007A3D"))
                pdf_canvas.drawCentredString(center_x, line_y, name)

                pdf_canvas.setFont("Helvetica-Bold", 10)
                for detail in details[:2]:
                    line_y -= 0.24 * inch
                    pdf_canvas.drawCentredString(center_x, line_y, detail)

            y -= 0.3 * inch
            left_x = self.page_width * 0.23
            right_x = self.page_width * 0.77
            _draw_member_block(left_member, left_x, y)
            _draw_member_block(right_member, right_x, y)

            y -= 0.72 * inch
        else:
            evaluator_profile = self._resolve_evaluator_profile(report_data)
            evaluator_name = (evaluator_profile.get("name") or report_data.get("evaluator") or "N/A").upper()
            header_label = (evaluator_profile.get("header_label") or evaluator_profile.get("title_label") or "LICENCIADA").upper()

            y -= 0.3 * inch
            pdf_canvas.setFont("Helvetica", 9)
            pdf_canvas.setFillColor(colors.black)
            pdf_canvas.drawCentredString(
                self.page_width / 2,
                y,
                f"{header_label}: {evaluator_name}",
            )

            profession_line = evaluator_profile.get("profession") or "TERAPEUTA RESPIRATORIA"
            y -= 0.25 * inch
            pdf_canvas.drawCentredString(self.page_width / 2, y, profession_line)

            registry_line = evaluator_profile.get("registry") or "REGISTRO IDÓNEO 124"
            y -= 0.2 * inch
            pdf_canvas.drawCentredString(self.page_width / 2, y, registry_line)

        counterpart_name = (report_data.get("company_counterpart") or "").strip()
        counterpart_role = (report_data.get("counterpart_role") or "").strip()
        if counterpart_name:
            y -= 0.45 * inch
            pdf_canvas.setFont("Helvetica-Bold", 10)
            pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
            pdf_canvas.drawCentredString(self.page_width / 2, y, "CONTRAPARTE TÉCNICA:")

            y -= 0.3 * inch
            pdf_canvas.setFont("Helvetica", 9)
            pdf_canvas.setFillColor(colors.black)
            pdf_canvas.drawCentredString(self.page_width / 2, y, counterpart_name.upper())

            if counterpart_role:
                y -= 0.25 * inch
                pdf_canvas.drawCentredString(self.page_width / 2, y, counterpart_role.upper())

        y -= 0.45 * inch
        pdf_canvas.setFont("Helvetica-Bold", 10)
        pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
        pdf_canvas.drawCentredString(self.page_width / 2, y, "FECHA DE REALIZACIÓN:")

        y -= 0.25 * inch
        pdf_canvas.setFont("Helvetica-Bold", 11)
        pdf_canvas.setFillColor(colors.black)
        pdf_canvas.drawCentredString(self.page_width / 2, y, report_data.get("date", "N/A").upper())

        y -= 0.5 * inch

    def _draw_general_info(self, pdf_canvas: canvas.Canvas, report_data: Dict) -> None:
        """Información general (ya incluida en el encabezado)."""

    def _draw_table_of_contents(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        watermark_logo: Optional[str] = None,
        start_page: int = 2,
    ) -> int:
        """Dibuja la página estática de contenido en la segunda página."""

        pdf_canvas.showPage()
        current_page = start_page

        if watermark_logo:
            self._add_watermark_logo(pdf_canvas, watermark_logo)

        y = self._draw_header_branding(pdf_canvas, report_data)
        y -= 0.35 * inch

        pdf_canvas.setFont("Helvetica-Bold", 14)
        pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
        pdf_canvas.drawString(self.left_margin, y, "CONTENIDO:")

        y -= 0.15 * inch
        pdf_canvas.setLineWidth(2)
        pdf_canvas.setStrokeColor(colors.HexColor("#2E7D32"))
        pdf_canvas.line(self.left_margin, y, self.page_width - self.right_margin, y)
        y -= 0.3 * inch

        pdf_canvas.setFont("Helvetica", 11)
        pdf_canvas.setFillColor(colors.black)
        for title in get_content_outline(report_data.get("type", "")):
            if y <= self.bottom_margin + 0.5 * inch:
                self._draw_footer(pdf_canvas, page_number=current_page)
                pdf_canvas.showPage()
                current_page += 1
                if watermark_logo:
                    self._add_watermark_logo(pdf_canvas, watermark_logo)
                y = self._draw_header_branding(pdf_canvas, report_data)
                y -= 0.35 * inch
                pdf_canvas.setFont("Helvetica-Bold", 14)
                pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
                pdf_canvas.drawString(self.left_margin, y, "CONTENIDO (cont.):")
                y -= 0.15 * inch
                pdf_canvas.setLineWidth(2)
                pdf_canvas.setStrokeColor(colors.HexColor("#2E7D32"))
                pdf_canvas.line(self.left_margin, y, self.page_width - self.right_margin, y)
                y -= 0.3 * inch
                pdf_canvas.setFont("Helvetica", 11)
                pdf_canvas.setFillColor(colors.black)

            pdf_canvas.drawString(self.left_margin, y, f"- {title}")
            y -= 0.35 * inch

        self._draw_footer(pdf_canvas, page_number=current_page)
        return current_page

    def _draw_attachment_cover_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        title: str,
        description: str,
        file_paths: list,
        page_number: int,
        watermark_logo: Optional[str] = None,
        folder_path: Optional[str] = None,
    ) -> int:
        """Agrega una página que anuncia los anexos que se incluirán enseguida."""

        pdf_canvas.showPage()
        current_page = page_number

        if watermark_logo:
            self._add_watermark_logo(pdf_canvas, watermark_logo)

        y = self._draw_header_branding(pdf_canvas, report_data)
        y -= 0.35 * inch

        pdf_canvas.setFont("Helvetica-Bold", 14)
        pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
        pdf_canvas.drawString(self.left_margin, y, title)
        y -= 0.25 * inch

        pdf_canvas.setFont("Helvetica", 11)
        pdf_canvas.setFillColor(colors.black)
        pdf_canvas.drawString(self.left_margin, y, description)
        y -= 0.3 * inch

        if file_paths:
            pdf_canvas.setFont("Helvetica", 10)
            for file_path in file_paths:
                if y <= self.bottom_margin + 0.5 * inch:
                    self._draw_footer(pdf_canvas, page_number=current_page)
                    pdf_canvas.showPage()
                    current_page += 1
                    if watermark_logo:
                        self._add_watermark_logo(pdf_canvas, watermark_logo)
                    y = self._draw_header_branding(pdf_canvas, report_data)
                    y -= 0.35 * inch
                    pdf_canvas.setFont("Helvetica-Bold", 14)
                    pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
                    pdf_canvas.drawString(self.left_margin, y, f"{title} (cont.)")
                    y -= 0.3 * inch
                    pdf_canvas.setFont("Helvetica", 10)
                    pdf_canvas.setFillColor(colors.black)

                label = f"• {Path(file_path).name}"
                pdf_canvas.drawString(
                    self.left_margin + 12,
                    y,
                    label,
                )
                if self._links_are_relative(report_data):
                    self._add_file_link(
                        pdf_canvas,
                        str(Path(folder_path or "") / Path(file_path).name) if folder_path else Path(file_path).name,
                        self.left_margin + 12,
                        y - 2,
                        420,
                        12,
                        relative=True,
                    )
                y -= 0.28 * inch

        self._draw_footer(pdf_canvas, page_number=current_page)
        return current_page

    def _draw_company_profile_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        watermark_logo: Optional[str],
        page_number: int,
    ) -> int:
        """Genera la página de datos de la empresa (página 3)."""

        pdf_canvas.showPage()

        if watermark_logo:
            self._add_watermark_logo(pdf_canvas, watermark_logo)

        y = self._draw_header_branding(pdf_canvas, report_data)
        y -= 0.35 * inch

        pdf_canvas.setFont("Helvetica-Bold", 14)
        pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
        pdf_canvas.drawString(self.left_margin, y, "DATOS DE LA EMPRESA:")

        y -= 0.18 * inch
        pdf_canvas.setStrokeColor(colors.HexColor("#2E7D32"))
        pdf_canvas.setLineWidth(3)
        pdf_canvas.line(self.left_margin, y, self.page_width - self.right_margin, y)
        y -= 0.35 * inch

        counterpart_name = (report_data.get("company_counterpart") or "").strip()
        counterpart_role = (report_data.get("counterpart_role") or "").strip()

        rows = [
            ("Nombre de la empresa:", report_data.get("company", "N/A")),
            (
                "Planta evaluada:",
                report_data.get("plant") or report_data.get("location", "N/A"),
            ),
            ("Actividad principal:", report_data.get("activity", "N/A")),
            (
                "País en que se realizó:",
                report_data.get("country") or report_data.get("location", "N/A"),
            ),
            (
                "Fechas del estudio:",
                report_data.get("study_dates", report_data.get("date", "N/A")),
            ),
        ]

        if counterpart_name:
            rows.insert(
                3,
                ("Contraparte técnica por la empresa:", counterpart_name),
            )
            if counterpart_role:
                rows.insert(4, ("Cargo del personal encargado:", counterpart_role))

        bullet_symbol = "❖"
        bullet_color = colors.HexColor("#2E7D32")
        bullet_x = self.left_margin
        label_x = bullet_x + 0.2 * inch
        value_x = label_x + 2.7 * inch
        label_width = value_x - label_x - 0.15 * inch
        value_width = self.page_width - self.right_margin - value_x
        line_height = 0.22 * inch

        for label, value in rows:
            pdf_canvas.setFillColor(bullet_color)
            pdf_canvas.setFont("Helvetica", 12)
            pdf_canvas.drawString(bullet_x, y, bullet_symbol)

            label_lines = self._wrap_text(label, "Helvetica-Bold", 11, label_width, pdf_canvas)
            pdf_canvas.setFont("Helvetica-Bold", 11)
            label_line_y = y
            for line in label_lines:
                pdf_canvas.drawString(label_x, label_line_y, line)
                label_line_y -= line_height

            value_lines = self._wrap_text(str(value or "N/A"), "Helvetica", 11, value_width, pdf_canvas)
            pdf_canvas.setFont("Helvetica", 11)
            pdf_canvas.setFillColor(colors.black)
            value_line_y = y
            for line in value_lines:
                pdf_canvas.drawString(value_x, value_line_y, line)
                value_line_y -= line_height

            row_height = max(len(label_lines), len(value_lines)) * line_height
            y -= row_height + 0.15 * inch

            if y <= self.bottom_margin + 0.75 * inch:
                # Si la página se llena inesperadamente, continuar en la siguiente
                self._draw_footer(pdf_canvas, page_number=page_number)
                page_number += 1
                pdf_canvas.showPage()
                if watermark_logo:
                    self._add_watermark_logo(pdf_canvas, watermark_logo)
                y = self._draw_header_branding(pdf_canvas, report_data)
                y -= 0.35 * inch
                pdf_canvas.setFont("Helvetica-Bold", 14)
                pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
                pdf_canvas.drawString(self.left_margin, y, "DATOS DE LA EMPRESA (cont.):")
                y -= 0.18 * inch
                pdf_canvas.setStrokeColor(colors.HexColor("#2E7D32"))
                pdf_canvas.setLineWidth(3)
                pdf_canvas.line(self.left_margin, y, self.page_width - self.right_margin, y)
                y -= 0.35 * inch

        self._draw_footer(pdf_canvas, page_number=page_number)
        return page_number

    def _resolve_attachment_folder(self, report_data: Dict, key: str, file_paths: list) -> Optional[str]:
        """Determina la carpeta a la que debe apuntar el enlace de anexos."""

        links = report_data.get("attachment_folder_links") or {}
        linked_path = links.get(key) if isinstance(links, dict) else None
        if linked_path:
            return str(linked_path)

        for file_path in file_paths or []:
            try:
                return str(Path(file_path).resolve().parent)
            except OSError:
                return str(Path(file_path).parent)
        return None

    def _draw_results_table_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        dataset_key: str,
        entries: list,
        page_number: int,
        watermark_logo: Optional[str] = None,
    ) -> int:
        """Genera la página (o páginas) con el cuadro detallado de resultados."""

        if not entries:
            return page_number - 1

        scheme = RESULT_SCHEMES.get(dataset_key, RESULT_SCHEMES["audiometria"])

        table_width = self.page_width - self.left_margin - self.right_margin
        columns = [
            {"key": "index", "title": "N°", "ratio": 0.06, "align": "center"},
            {"key": "name", "title": "NOMBRE", "ratio": 0.27, "align": "left"},
            {"key": "identification", "title": "CÉDULA", "ratio": 0.16, "align": "center"},
            {"key": "age", "title": "EDAD", "ratio": 0.09, "align": "center"},
            {"key": "position", "title": "CARGO", "ratio": 0.20, "align": "left"},
            {"key": "result", "title": "RESULTADO", "ratio": 0.22, "align": "center"},
        ]
        column_widths = [table_width * col["ratio"] for col in columns]
        header_height = 0.45 * inch
        line_spacing = 0.18 * inch
        min_row_height = 0.4 * inch
        current_page = page_number - 1
        current_y = None

        for idx, entry in enumerate(entries, start=1):
            row_height = self._estimate_results_row_height(
                pdf_canvas,
                entry,
                columns,
                column_widths,
                line_spacing,
                min_row_height,
                dataset_key,
                row_index=idx,
            )
            if current_y is None or current_y - row_height <= self.bottom_margin:
                if current_y is not None:
                    self._draw_footer(pdf_canvas, page_number=current_page)
                current_page += 1
                current_y = self._prepare_results_table_page(
                    pdf_canvas,
                    report_data,
                    scheme,
                    columns,
                    column_widths,
                    header_height,
                    watermark_logo,
                )

            current_y = self._draw_results_table_row(
                pdf_canvas,
                entry,
                row_height,
                current_y,
                column_widths,
                columns,
                dataset_key=dataset_key,
                row_index=idx,
                line_spacing=line_spacing,
            )

        if current_y is not None:
            self._draw_footer(pdf_canvas, page_number=current_page)

        return current_page

    def _group_evaluated_entries(self, entries: list) -> Dict[str, list]:
        """Agrupa los registros según el tipo de prueba declarado."""

        grouped = {key: [] for key in RESULT_SCHEMES}
        for entry in entries:
            dataset = self._infer_dataset_from_entry(entry)
            grouped.setdefault(dataset, []).append(entry)
        return grouped

    def _determine_result_dataset_keys(self, report_type: str) -> list:
        """Mantiene el orden esperado de secciones según la selección del usuario."""

        normalized = (report_type or "").lower()
        has_audio = "audiometr" in normalized
        has_espiro = "espiro" in normalized

        if has_audio and has_espiro:
            return ["audiometria", "espirometria"]
        if has_espiro:
            return ["espirometria"]
        return ["audiometria"]

    def _infer_dataset_from_entry(self, entry: Dict) -> str:
        """Determina el dataset más probable para mantener compatibilidad retro."""

        explicit = (entry.get("test_type") or "").lower()
        if "espiro" in explicit:
            return "espirometria"
        if "audio" in explicit:
            return "audiometria"

        code = (entry.get("result_code") or "").lower()
        espiro_keys = {opt["key"] for opt in RESULT_SCHEMES["espirometria"]["options"]}
        if code in espiro_keys:
            return "espirometria"

        label = (entry.get("result_label") or entry.get("result") or "").lower()
        if any(term in label for term in ("obstrucci", "restricci")):
            return "espirometria"
        return "audiometria"

    def _prepare_results_table_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        scheme: Dict,
        columns: list,
        column_widths: list,
        header_height: float,
        watermark_logo: Optional[str] = None,
    ) -> float:
        """Configura una nueva página para la tabla sin fondo de marca de agua."""

        pdf_canvas.showPage()
        if watermark_logo:
            self._add_watermark_logo(pdf_canvas, watermark_logo)
        y = self._draw_header_branding(pdf_canvas, report_data)
        y -= 0.35 * inch

        pdf_canvas.setFont("Helvetica-Bold", 15)
        pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
        pdf_canvas.drawCentredString(
            self.page_width / 2,
            y,
            scheme.get("title", self._resolve_results_title(report_data.get("type", ""))),
        )

        y -= 0.25 * inch
        plant_label = report_data.get("plant") or report_data.get("location", "")
        pdf_canvas.setFont("Helvetica-Bold", 12)
        pdf_canvas.drawCentredString(
            self.page_width / 2,
            y,
            f"PLANTA: {plant_label.upper() if plant_label else 'N/D'}",
        )

        y -= 0.2 * inch
        pdf_canvas.setStrokeColor(colors.HexColor("#2E7D32"))
        pdf_canvas.setLineWidth(2)
        pdf_canvas.line(self.left_margin, y, self.page_width - self.right_margin, y)
        y -= 0.35 * inch

        self._draw_results_table_header(pdf_canvas, columns, column_widths, header_height, y)
        return y - header_height

    def _draw_results_table_header(
        self,
        pdf_canvas: canvas.Canvas,
        columns: list,
        column_widths: list,
        header_height: float,
        top_y: float,
    ) -> None:
        """Renderiza la fila de encabezados del cuadro de resultados."""

        total_width = sum(column_widths)
        pdf_canvas.setFillColor(colors.HexColor("#B7D58A"))
        pdf_canvas.rect(self.left_margin, top_y - header_height, total_width, header_height, stroke=0, fill=1)
        pdf_canvas.setStrokeColor(colors.HexColor("#2E7D32"))
        pdf_canvas.setLineWidth(1)

        x = self.left_margin
        pdf_canvas.setFont("Helvetica-Bold", 11)
        pdf_canvas.setFillColor(colors.black)
        for idx, column in enumerate(columns):
            width = column_widths[idx]
            pdf_canvas.rect(x, top_y - header_height, width, header_height, stroke=1, fill=0)
            pdf_canvas.drawCentredString(
                x + width / 2,
                top_y - (header_height / 2) + 2,
                column["title"],
            )
            x += width

    def _estimate_results_row_height(
        self,
        pdf_canvas: canvas.Canvas,
        entry: Dict,
        columns: list,
        column_widths: list,
        line_spacing: float,
        min_row_height: float,
        dataset_key: str,
        row_index: int,
    ) -> float:
        """Calcula altura necesaria para la fila en base a texto envuelto."""

        max_lines = 1
        for idx, column in enumerate(columns):
            width = column_widths[idx] - 8
            value = self._resolve_entry_value(entry, column["key"], row_index)
            if column["key"] == "result":
                style = self._result_style_for_entry(entry, dataset_key)
                value = style["label"]
            lines = self._wrap_text(str(value or "N/A"), "Helvetica", 10, width, pdf_canvas)
            max_lines = max(max_lines, len(lines))

        return max(min_row_height, max_lines * line_spacing + 0.18 * inch)

    def _draw_results_table_row(
        self,
        pdf_canvas: canvas.Canvas,
        entry: Dict,
        row_height: float,
        current_y: float,
        column_widths: list,
        columns: list,
        dataset_key: str,
        row_index: int,
        line_spacing: float,
    ) -> float:
        """Dibuja una fila de la tabla con resaltado en la columna de resultado."""

        table_width = sum(column_widths)
        row_bottom = current_y - row_height
        fill_color = colors.HexColor("#F8FBF7") if row_index % 2 == 0 else colors.white
        pdf_canvas.setFillColor(fill_color)
        pdf_canvas.rect(self.left_margin, row_bottom, table_width, row_height, stroke=0, fill=1)

        x = self.left_margin
        pdf_canvas.setFont("Helvetica", 10)
        pdf_canvas.setLineWidth(1)

        for idx, column in enumerate(columns):
            width = column_widths[idx]
            value = self._resolve_entry_value(entry, column["key"], row_index)
            lines = self._wrap_text(str(value or "N/A"), "Helvetica", 10, width - 8, pdf_canvas)

            if column["key"] == "result":
                style = self._result_style_for_entry(entry, dataset_key)
                pdf_canvas.setFillColor(style["background"])
                pdf_canvas.rect(x, row_bottom, width, row_height, stroke=0, fill=1)
                pdf_canvas.setFillColor(style["text"])
                value = style["label"]
                lines = self._wrap_text(str(value or "N/A"), "Helvetica", 10, width - 8, pdf_canvas)
            else:
                pdf_canvas.setFillColor(colors.black)

            pdf_canvas.rect(x, row_bottom, width, row_height, stroke=1, fill=0)

            text_y = row_bottom + row_height - 0.18 * inch
            for line in lines:
                if column["align"] == "center":
                    pdf_canvas.drawCentredString(
                        x + width / 2,
                        text_y,
                        line,
                    )
                else:
                    pdf_canvas.drawString(
                        x + 4,
                        text_y,
                        line,
                    )
                text_y -= line_spacing

            x += width

        return row_bottom

    def _result_style_for_entry(self, entry: Dict, dataset_key: str) -> Dict:
        """Asigna colores al resultado según su clasificación y esquema."""

        scheme = RESULT_SCHEMES.get(dataset_key, RESULT_SCHEMES["audiometria"])
        palette = {opt["key"]: opt for opt in scheme.get("options", [])}
        key = self._resolve_result_code(entry, dataset_key)
        option = palette.get(key)

        label = (
            entry.get("result_label")
            or entry.get("result")
            or (entry.get("results") or {}).get("status")
        )
        if not label:
            label = option.get("label") if option else "N/A"

        background = colors.HexColor(option.get("bg", "#E0E0E0")) if option else colors.HexColor("#E0E0E0")
        text_color = colors.HexColor(option.get("fg", "#1B5E20")) if option else colors.HexColor("#1B5E20")

        return {
            "label": label.upper(),
            "background": background,
            "text": text_color,
        }

    def _draw_results_statistics_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        dataset_key: str,
        entries: list,
        page_number: int,
        watermark_logo: Optional[str] = None,
    ) -> int:
        """Crea la página de análisis estadístico inmediatamente después de la tabla."""

        scheme = RESULT_SCHEMES.get(dataset_key, RESULT_SCHEMES["audiometria"])
        stats = self._compute_results_stats(dataset_key, entries)
        if stats.get("total", 0) == 0:
            return page_number - 1

        pdf_canvas.showPage()
        if watermark_logo:
            self._add_watermark_logo(pdf_canvas, watermark_logo)
        y = self._draw_header_branding(pdf_canvas, report_data)
        y -= 0.35 * inch

        pdf_canvas.setFont("Helvetica-Bold", 14)
        pdf_canvas.setFillColor(colors.HexColor("#1B5E20"))
        pdf_canvas.drawString(self.left_margin, y, "ESTADÍSTICA DE LOS RESULTADOS:")

        y -= 0.15 * inch
        pdf_canvas.setStrokeColor(colors.HexColor("#2E7D32"))
        pdf_canvas.setLineWidth(2)
        pdf_canvas.line(self.left_margin, y, self.page_width - self.right_margin, y)

        table_top = y - 0.35 * inch
        table_width = (self.page_width - self.left_margin - self.right_margin) * 0.38
        column_widths = [table_width * 0.7, table_width * 0.3]
        header_height = 0.4 * inch
        row_height = 0.32 * inch

        # Encabezado de la tabla
        pdf_canvas.setFillColor(colors.HexColor("#DAEBC8"))
        pdf_canvas.rect(self.left_margin, table_top - header_height, table_width, header_height, stroke=1, fill=1)
        pdf_canvas.setFillColor(colors.black)
        header_title = f"RESULTADO DE LAS {scheme.get('chart_label', self._get_results_label(report_data.get('type', '')))}"
        header_lines = self._wrap_text(header_title, "Helvetica-Bold", 9, column_widths[0] - 8, pdf_canvas)
        line_y = table_top - 0.12 * inch
        for line in header_lines:
            pdf_canvas.setFont("Helvetica-Bold", 9)
            pdf_canvas.drawCentredString(
                self.left_margin + column_widths[0] / 2,
                line_y,
                line,
            )
            line_y -= 0.14 * inch
        pdf_canvas.drawCentredString(
            self.left_margin + column_widths[0] + column_widths[1] / 2,
            table_top - (header_height / 2) + 2,
            "CANTIDAD",
        )

        rows = [
            (option.get("table_label", option["label"]).upper(), stats.get(option["key"], 0))
            for option in scheme.get("options", [])
        ]
        rows.append(("TOTAL", stats.get("total", 0)))

        current_y = table_top - header_height
        for label, value in rows:
            current_y -= row_height
            is_total = label == "TOTAL"
            fill_color = colors.HexColor("#F7FDF1") if not is_total else colors.HexColor("#FFF9E7")
            pdf_canvas.setFillColor(fill_color)
            pdf_canvas.rect(self.left_margin, current_y, table_width, row_height, stroke=1, fill=1)

            font_style = ("Helvetica-Bold", 10) if is_total else ("Helvetica", 10)
            pdf_canvas.setFont(*font_style)
            pdf_canvas.setFillColor(colors.black)
            pdf_canvas.drawString(self.left_margin + 6, current_y + row_height / 2 - 4, label)
            pdf_canvas.drawCentredString(
                self.left_margin + column_widths[0] + column_widths[1] / 2,
                current_y + row_height / 2 - 4,
                str(value),
            )

        chart_bottom = current_y
        chart_path = self._create_results_chart_image(dataset_key, stats, scheme)
        if chart_path:
            try:
                chart_gap = 0.5 * inch
                chart_x = self.left_margin + table_width + chart_gap
                available_width = self.page_width - self.right_margin - chart_x
                if available_width > 1.0 * inch:
                    chart_width = min(available_width, 5.2 * inch)
                    chart_height = chart_width * 0.72
                    chart_top = table_top - 0.1 * inch
                    chart_bottom = chart_top - chart_height
                    pdf_canvas.drawImage(
                        chart_path,
                        chart_x,
                        chart_bottom,
                        width=chart_width,
                        height=chart_height,
                        mask="auto",
                    )
            finally:
                if os.path.exists(chart_path):
                    os.remove(chart_path)

        analysis_text = self._build_results_analysis_text(dataset_key, stats)
        content_bottom = min(current_y, chart_bottom)
        analysis_y = content_bottom - 0.35 * inch
        pdf_canvas.setFont("Helvetica-Bold", 10)
        pdf_canvas.drawString(self.left_margin, analysis_y, "Análisis de los resultados:")
        analysis_y -= 0.18 * inch

        pdf_canvas.setFont("Helvetica", 10)
        text_width = self.page_width - self.left_margin - self.right_margin
        for line in self._wrap_text(analysis_text, "Helvetica", 10, text_width, pdf_canvas):
            pdf_canvas.drawString(self.left_margin, analysis_y, line)
            analysis_y -= 0.2 * inch

        self._draw_footer(pdf_canvas, page_number=page_number)
        return page_number

    def _draw_conclusion_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        text: str,
        page_number: int,
        watermark_logo: Optional[str] = None,
    ) -> int:
        """Genera la página de conclusiones respetando el formato institucional."""

        paragraphs = self._split_into_paragraphs(text)
        return self._draw_textual_section_page(
            pdf_canvas,
            report_data,
            title="CONCLUSIONES:",
            entries=paragraphs,
            page_number=page_number,
            watermark_logo=watermark_logo,
            use_bullets=False,
        )

    def _draw_recommendations_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        text: str,
        page_number: int,
        watermark_logo: Optional[str] = None,
    ) -> int:
        """Genera la página de recomendaciones usando viñetas en verde."""

        bullet_items = self._split_into_bullet_items(text)
        return self._draw_textual_section_page(
            pdf_canvas,
            report_data,
            title="RECOMENDACIONES:",
            entries=bullet_items,
            page_number=page_number,
            watermark_logo=watermark_logo,
            use_bullets=True,
        )

    def _draw_protocol_header(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        protocol_title: str,
        section_title: str,
        watermark_logo: Optional[str],
        start_new_page: bool,
    ) -> float:
        """Dibuja el encabezado del protocolo y el titulo de seccion inicial."""

        if start_new_page:
            pdf_canvas.showPage()
        if watermark_logo:
            self._add_watermark_logo(pdf_canvas, watermark_logo)

        y = self._draw_header_branding(pdf_canvas, report_data)
        y -= 0.35 * inch

        pdf_canvas.setFont("Helvetica-Bold", 13)
        pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
        pdf_canvas.drawString(self.left_margin, y, protocol_title)
        y -= 0.18 * inch
        pdf_canvas.setStrokeColor(colors.HexColor("#2E7D32"))
        pdf_canvas.setLineWidth(2)
        pdf_canvas.line(self.left_margin, y, self.page_width - self.right_margin, y)
        y -= 0.35 * inch

        pdf_canvas.setFont("Helvetica-Bold", 12)
        pdf_canvas.setFillColor(colors.black)
        pdf_canvas.drawString(self.left_margin, y, section_title)
        y -= 0.16 * inch
        pdf_canvas.setStrokeColor(colors.black)
        pdf_canvas.setLineWidth(1)
        pdf_canvas.line(self.left_margin, y, self.page_width - self.right_margin, y)
        return y - 0.28 * inch

    def _draw_audiometry_protocol_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        page_number: int,
        watermark_logo: Optional[str] = None,
        start_new_page: bool = True,
    ) -> int:
        """Página con el protocolo específico de audiometría."""

        intro_text = (
            "Es una prueba que evalúa el funcionamiento del sistema auditivo, que permite "
            "determinar la capacidad de una persona para escuchar los sonidos y así conocer "
            "si el proceso de audición está alterado."
        )

        steps = [
            "Verificar si el conducto auditivo esta obstruido por cerumen por medio de la "
            "otoscopia, de ser así debe realizársele un lavado de oído y 2 a 3 días después "
            "se puede proceder a realizar la prueba.",
            "Se utiliza el audiómetro con el que se va a realizar el estudio, el mismo debe de "
            "tener su certificado de calibración vigente, para garantizar la fiabilidad de los resultados.",
            "Se le explica al colaborador el método de respuestas a dar al evaluador, un fonoaudiólogo "
            "idóneo, el mismo debe de indicar cuando escucha los sonidos y esto se grafica en un audiograma, "
            "que es lo que se le entrega a la empresa como resultado posteriormente, con su debido cuestionario "
            "de preguntas ocupacionales y personales de cada colaborador.",
            "Se le indica al colaborador, el resultado de su estudio de audiometría una vez finalizada la prueba, "
            "para que pueda crear conciencia de como limpiar sus oídos y la importancia de la utilización de los "
            "protectores auditivos, como parte de la conservación auditiva de la empresa.",
        ]

        protocol_images = self._resolve_protocol_image_paths()
        intro_image = protocol_images[0] if protocol_images else None
        step_images = protocol_images[1:5]

        current_page = page_number
        y = self._draw_protocol_header(
            pdf_canvas,
            report_data,
            "PROTOCOLO DE LA AUDIOMETRÍA:",
            "LA AUDIOMETRÍA:",
            watermark_logo,
            start_new_page,
        )

        text_width = self.page_width - self.left_margin - self.right_margin
        line_spacing = 0.23 * inch

        for line in self._wrap_text(intro_text, "Helvetica", 11, text_width, pdf_canvas):
            if y - line_spacing <= self.bottom_margin:
                self._draw_footer(pdf_canvas, page_number=current_page)
                current_page += 1
                y = self._prepare_protocol_continuation_page(
                    pdf_canvas,
                    report_data,
                    watermark_logo,
                )
            pdf_canvas.setFont("Helvetica", 11)
            pdf_canvas.setFillColor(colors.black)
            pdf_canvas.drawString(self.left_margin, y, line)
            y -= line_spacing

        if intro_image:
            current_page, y = self._draw_protocol_image(
                pdf_canvas,
                report_data,
                intro_image,
                y,
                current_page,
                watermark_logo,
            )

        self._draw_footer(pdf_canvas, page_number=current_page)
        current_page += 1
        y = self._prepare_text_section_header(
            pdf_canvas,
            report_data,
            "PASOS PARA REALIZAR LA AUDIOMETRIA:",
            watermark_logo,
        )

        for idx, step in enumerate(steps, start=1):
            if idx == 3:
                self._draw_footer(pdf_canvas, page_number=current_page)
                current_page += 1
                y = self._prepare_protocol_continuation_page(
                    pdf_canvas,
                    report_data,
                    watermark_logo,
                )

            available_width = text_width - 0.35 * inch
            lines = self._wrap_text(step, "Helvetica", 11, available_width, pdf_canvas)
            required_height = len(lines) * line_spacing + 0.12 * inch

            if y - required_height <= self.bottom_margin:
                self._draw_footer(pdf_canvas, page_number=current_page)
                current_page += 1
                y = self._prepare_protocol_continuation_page(
                    pdf_canvas,
                    report_data,
                    watermark_logo,
                )

            pdf_canvas.setFont("Helvetica-Bold", 11)
            pdf_canvas.setFillColor(colors.black)
            pdf_canvas.drawString(self.left_margin, y, f"{idx}.")

            text_x = self.left_margin + 0.35 * inch
            line_y = y
            pdf_canvas.setFont("Helvetica", 11)
            for line in lines:
                pdf_canvas.drawString(text_x, line_y, line)
                line_y -= line_spacing

            y = line_y - 0.12 * inch

            image_path = step_images[idx - 1] if idx - 1 < len(step_images) else None
            if image_path:
                current_page, y = self._draw_protocol_image(
                    pdf_canvas,
                    report_data,
                    image_path,
                    y,
                    current_page,
                    watermark_logo,
                    max_height=1.35 * inch,
                )

        self._draw_footer(pdf_canvas, page_number=current_page)
        return current_page

    def _draw_spirometry_protocol_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        page_number: int,
        watermark_logo: Optional[str] = None,
        start_new_page: bool = True,
    ) -> int:
        """Pagina con el protocolo especifico de espirometria."""

        intro_text = (
            "Es la prueba que evalua la capacidad pulmonar, de la persona y asi determinar "
            "si presenta patologias respiratorias en su sistema."
        )

        steps = [
            "Verificar antes de proceder a realizar la prueba, si el paciente tiene historial de "
            "enfermedad pulmonar obstructiva cronica (EPOC), asma y otras enfermedades que "
            "afectan la respiracion, para tenerlo en cuenta.",
            "El equipo que se utiliza para realizar la prueba es el espirometro, el mismo tiene "
            "su certificado de calibracion vigente, para garantizar la fiabilidad de los resultados.",
            "Se le explica al colaborador la tecnica adecuada que debe de realizar, el evaluador "
            "es un Terapeuta Respiratorio idoneo, el cual va a estimular constantemente al paciente "
            "para que pueda cumplir con todos los pasos de la prueba, ya que la misma es de mucho "
            "esfuerzo recordando que mide en todo momento su capacidad al espirar.",
            "Se le indica al colaborador, el resultado de su estudio de espirometria una vez "
            "finalizada la prueba, para que pueda crear conciencia de como proteger sus pulmones "
            "de toda clase de particulas o polvo, que pueden llegar a afectar su sistema respiratorio.",
        ]

        protocol_images = self._resolve_spirometry_protocol_image_paths()
        intro_image = protocol_images[0] if protocol_images else None
        step_images = protocol_images[1:5]

        current_page = page_number
        y = self._draw_protocol_header(
            pdf_canvas,
            report_data,
            "PROTOCOLO DE LA ESPIROMETRÍA:",
            "LA ESPIROMETRÍA:",
            watermark_logo,
            start_new_page,
        )

        text_width = self.page_width - self.left_margin - self.right_margin
        line_spacing = 0.23 * inch

        for line in self._wrap_text(intro_text, "Helvetica", 11, text_width, pdf_canvas):
            if y - line_spacing <= self.bottom_margin:
                self._draw_footer(pdf_canvas, page_number=current_page)
                current_page += 1
                y = self._prepare_protocol_continuation_page(
                    pdf_canvas,
                    report_data,
                    watermark_logo,
                )
            pdf_canvas.setFont("Helvetica", 11)
            pdf_canvas.setFillColor(colors.black)
            pdf_canvas.drawString(self.left_margin, y, line)
            y -= line_spacing

        if intro_image:
            current_page, y = self._draw_protocol_image(
                pdf_canvas,
                report_data,
                intro_image,
                y,
                current_page,
                watermark_logo,
            )

        self._draw_footer(pdf_canvas, page_number=current_page)
        current_page += 1
        y = self._prepare_text_section_header(
            pdf_canvas,
            report_data,
            "PASOS PARA REALIZAR LA ESPIROMETRÍA:",
            watermark_logo,
        )

        for idx, step in enumerate(steps, start=1):
            if idx == 3:
                self._draw_footer(pdf_canvas, page_number=current_page)
                current_page += 1
                y = self._prepare_protocol_continuation_page(
                    pdf_canvas,
                    report_data,
                    watermark_logo,
                )

            available_width = text_width - 0.35 * inch
            lines = self._wrap_text(step, "Helvetica", 11, available_width, pdf_canvas)
            required_height = len(lines) * line_spacing + 0.12 * inch

            if y - required_height <= self.bottom_margin:
                self._draw_footer(pdf_canvas, page_number=current_page)
                current_page += 1
                y = self._prepare_protocol_continuation_page(
                    pdf_canvas,
                    report_data,
                    watermark_logo,
                )

            pdf_canvas.setFont("Helvetica-Bold", 11)
            pdf_canvas.setFillColor(colors.black)
            pdf_canvas.drawString(self.left_margin, y, f"{idx}.")

            text_x = self.left_margin + 0.35 * inch
            line_y = y
            pdf_canvas.setFont("Helvetica", 11)
            for line in lines:
                pdf_canvas.drawString(text_x, line_y, line)
                line_y -= line_spacing

            y = line_y - 0.12 * inch

            image_path = step_images[idx - 1] if idx - 1 < len(step_images) else None
            if image_path:
                current_page, y = self._draw_protocol_image(
                    pdf_canvas,
                    report_data,
                    image_path,
                    y,
                    current_page,
                    watermark_logo,
                    max_height=1.35 * inch,
                )

        self._draw_footer(pdf_canvas, page_number=current_page)
        return current_page

    def _resolve_protocol_image_paths(self) -> list:
        """Devuelve las imagenes del protocolo en el orden esperado."""

        base_dir = self._get_resource_base() / "imagenes de protocolo"
        ordered_names = [
            "Imagen1.jpg",
            "Imagen2.jpg",
            "Imagen3.jpg",
            "Imagen4.jpg",
            "Imagen5.jpg",
        ]
        resolved = []
        for name in ordered_names:
            candidate = base_dir / name
            if candidate.exists():
                resolved.append(str(candidate))
        return resolved

    def _resolve_spirometry_protocol_image_paths(self) -> list:
        """Devuelve las imagenes del protocolo de espirometria en orden."""

        base_dir = self._get_resource_base() / "imagenes de protocolo"
        ordered_names = [
            "espirometria 1.jpg",
            "espirometria 2.jpg",
            "espirometria 3.jpg",
            "espirometria 4.jpg",
            "espirometria 5.jpg",
        ]
        resolved = []
        for name in ordered_names:
            candidate = base_dir / name
            if candidate.exists():
                resolved.append(str(candidate))
        return resolved

    def _draw_protocol_image(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        image_path: str,
        y: float,
        current_page: int,
        watermark_logo: Optional[str],
        max_height: float = 2.6 * inch,
    ) -> tuple[int, float]:
        """Dibuja una imagen del protocolo y maneja el salto de pagina si es necesario."""

        if not image_path or not os.path.exists(image_path):
            return current_page, y

        try:
            with Image.open(image_path) as img:
                width_px, height_px = img.size
        except Exception:  # pragma: no cover
            return current_page, y

        max_width = self.page_width - self.left_margin - self.right_margin
        scale = min(max_width / width_px, max_height / height_px, 1.0)
        draw_width = width_px * scale
        draw_height = height_px * scale

        if y - draw_height - 0.2 * inch <= self.bottom_margin:
            self._draw_footer(pdf_canvas, page_number=current_page)
            current_page += 1
            y = self._prepare_protocol_continuation_page(
                pdf_canvas,
                report_data,
                watermark_logo,
            )

        x = (self.page_width - draw_width) / 2
        y -= draw_height
        pdf_canvas.drawImage(
            ImageReader(image_path),
            x,
            y,
            width=draw_width,
            height=draw_height,
            mask="auto",
        )
        y -= 0.2 * inch
        return current_page, y

    def _render_audiometry_protocol_pdf(
        self,
        report_data: Dict,
        watermark_logo: Optional[str],
    ) -> Optional[str]:
        """Genera un PDF temporal con el protocolo de audiometría."""

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        temp_path = temp_file.name
        temp_file.close()

        pdf_canvas = canvas.Canvas(temp_path, pagesize=landscape(letter))
        self._draw_audiometry_protocol_page(
            pdf_canvas,
            report_data,
            page_number=1,
            watermark_logo=watermark_logo,
            start_new_page=False,
        )
        pdf_canvas.save()
        return temp_path

    def _render_spirometry_protocol_pdf(
        self,
        report_data: Dict,
        watermark_logo: Optional[str],
    ) -> Optional[str]:
        """Genera un PDF temporal con el protocolo de espirometria."""

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        temp_path = temp_file.name
        temp_file.close()

        pdf_canvas = canvas.Canvas(temp_path, pagesize=landscape(letter))
        self._draw_spirometry_protocol_page(
            pdf_canvas,
            report_data,
            page_number=1,
            watermark_logo=watermark_logo,
            start_new_page=False,
        )
        pdf_canvas.save()
        return temp_path

    def _prepare_protocol_continuation_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        watermark_logo: Optional[str],
    ) -> float:
        """Abre una pagina nueva del protocolo sin titulo de seccion."""

        pdf_canvas.showPage()
        if watermark_logo:
            self._add_watermark_logo(pdf_canvas, watermark_logo)

        y = self._draw_header_branding(pdf_canvas, report_data)
        return y - 0.35 * inch

    def _get_resource_base(self) -> Path:
        """Obtiene la raiz de recursos tanto en modo normal como en ejecutable."""

        base_path = getattr(sys, "_MEIPASS", None)
        if base_path:
            return Path(base_path)
        return Path(__file__).resolve().parents[2]

    def _draw_technical_team_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        entries: list,
        page_number: int,
        watermark_logo: Optional[str] = None,
    ) -> int:
        """Dibuja la sección del equipo técnico usando el formato institucional."""

        normalized_entries = [entry for entry in entries if entry.get("name")]
        if not normalized_entries:
            return page_number - 1

        current_page = page_number
        y = self._prepare_text_section_header(
            pdf_canvas,
            report_data,
            "EQUIPO TÉCNICO:",
            watermark_logo,
        )

        text_width = self.page_width - self.left_margin - self.right_margin
        bullet_symbol = "❖"
        bullet_color = colors.HexColor("#2E7D32")
        bullet_indent = 0.28 * inch
        line_spacing = 0.24 * inch

        for member in normalized_entries:
            available_width = text_width - bullet_indent
            name_lines = self._wrap_text(member.get("name", ""), "Helvetica-Bold", 12, available_width, pdf_canvas)
            detail_lines = []
            for detail in member.get("details", []):
                detail = (detail or "").strip()
                if not detail:
                    continue
                detail_lines.extend(self._wrap_text(detail, "Helvetica", 11, available_width, pdf_canvas))

            total_lines = len(name_lines) + len(detail_lines)
            required_height = total_lines * line_spacing + 0.3 * inch

            if y - required_height <= self.bottom_margin:
                self._draw_footer(pdf_canvas, page_number=current_page)
                current_page += 1
                y = self._prepare_text_section_header(
                    pdf_canvas,
                    report_data,
                    "EQUIPO TÉCNICO (cont.):",
                    watermark_logo,
                )

            pdf_canvas.setFont("Helvetica-Bold", 13)
            pdf_canvas.setFillColor(bullet_color)
            pdf_canvas.drawString(self.left_margin, y, bullet_symbol)

            text_x = self.left_margin + bullet_indent
            line_y = y
            pdf_canvas.setFillColor(colors.black)
            pdf_canvas.setFont("Helvetica-Bold", 12)
            for line in name_lines:
                pdf_canvas.drawString(text_x, line_y, line)
                line_y -= line_spacing

            if detail_lines:
                pdf_canvas.setFont("Helvetica", 11)
                for line in detail_lines:
                    pdf_canvas.drawString(text_x, line_y, line)
                    line_y -= line_spacing

            y = line_y - 0.22 * inch

        credential_links = []
        if self._links_are_relative(report_data):
            credential_links = report_data.get("evaluator_credentials_links") or []
        if credential_links:
            link_y = y
            if link_y - 0.6 * inch <= self.bottom_margin:
                self._draw_footer(pdf_canvas, page_number=current_page)
                current_page += 1
                y = self._prepare_text_section_header(
                    pdf_canvas,
                    report_data,
                    "EQUIPO TÉCNICO (cont.):",
                    watermark_logo,
                )
                link_y = y

            pdf_canvas.setFont("Helvetica-Bold", 11)
            pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
            pdf_canvas.drawString(self.left_margin, link_y, "Idoneidad adjunta:")
            link_y -= 0.2 * inch

            pdf_canvas.setFont("Helvetica", 10)
            pdf_canvas.setFillColor(colors.HexColor("#1B5E20"))
            for item in credential_links:
                if not isinstance(item, dict):
                    continue
                file_path = item.get("file") or ""
                name = item.get("name") or ""
                if not file_path:
                    continue
                label = f"• {name}" if name else f"• {Path(file_path).name}"
                pdf_canvas.drawString(self.left_margin + 10, link_y, label)
                self._add_file_link(
                    pdf_canvas,
                    file_path,
                    self.left_margin + 10,
                    link_y - 2,
                    320,
                    12,
                    relative=True,
                )
                link_y -= 0.22 * inch

        self._draw_footer(pdf_canvas, page_number=current_page)
        return current_page

    def _draw_calibration_certificates_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        certificates: list,
        page_number: int,
        watermark_logo: Optional[str] = None,
    ) -> int:
        """Renderiza la tabla oficial de certificados de calibración."""

        normalized = [entry for entry in certificates if entry.get("equipment") or entry.get("certificate")]
        if not normalized:
            return page_number - 1

        columns = [
            {"key": "equipment", "title": "Equipo", "ratio": 0.20},
            {"key": "model", "title": "Modelo", "ratio": 0.12},
            {"key": "serial", "title": "Serie", "ratio": 0.12},
            {"key": "certificate", "title": "Certificado", "ratio": 0.16},
            {"key": "calibration_date", "title": "Calibración", "ratio": 0.12},
            {"key": "valid_until", "title": "Vigencia", "ratio": 0.12},
            {"key": "provider", "title": "Laboratorio", "ratio": 0.16},
        ]
        table_width = self.page_width - self.left_margin - self.right_margin
        column_widths = [table_width * column["ratio"] for column in columns]
        header_height = 0.38 * inch
        line_spacing = 0.18 * inch

        current_page = page_number
        y = self._prepare_text_section_header(
            pdf_canvas,
            report_data,
            "CERTIFICADOS DE CALIBRACIÓN:",
            watermark_logo,
        )
        current_y = self._draw_calibration_table_header(
            pdf_canvas,
            columns,
            column_widths,
            header_height,
            y,
        )

        for idx, entry in enumerate(normalized, start=1):
            row_height = self._estimate_calibration_row_height(
                pdf_canvas,
                entry,
                columns,
                column_widths,
                line_spacing,
            )

            if current_y - row_height <= self.bottom_margin:
                self._draw_footer(pdf_canvas, page_number=current_page)
                current_page += 1
                y = self._prepare_text_section_header(
                    pdf_canvas,
                    report_data,
                    "CERTIFICADOS DE CALIBRACIÓN (cont.):",
                    watermark_logo,
                )
                current_y = self._draw_calibration_table_header(
                    pdf_canvas,
                    columns,
                    column_widths,
                    header_height,
                    y,
                )

            current_y = self._draw_calibration_table_row(
                pdf_canvas,
                entry,
                columns,
                column_widths,
                current_y,
                line_spacing,
                row_index=idx,
            )

        has_attachments = any(entry.get("attachment") for entry in normalized)
        if has_attachments:
            note_height = 0.28 * inch
            if current_y - note_height <= self.bottom_margin:
                self._draw_footer(pdf_canvas, page_number=current_page)
                current_page += 1
                current_y = self._prepare_text_section_header(
                    pdf_canvas,
                    report_data,
                    "CERTIFICADOS DE CALIBRACIÓN (cont.):",
                    watermark_logo,
                )
            pdf_canvas.setFont("Helvetica-Oblique", 9)
            pdf_canvas.setFillColor(colors.HexColor("#424242"))
            pdf_canvas.drawString(
                self.left_margin,
                current_y - 0.2 * inch,
                "Los certificados completos se presentan en las páginas de anexos inmediatamente posteriores.",
            )
            current_y -= 0.35 * inch

        self._draw_footer(pdf_canvas, page_number=current_page)
        return current_page

    def _draw_calibration_table_header(
        self,
        pdf_canvas: canvas.Canvas,
        columns: list,
        column_widths: list,
        header_height: float,
        top_y: float,
    ) -> float:
        """Dibuja el encabezado de la tabla de certificados."""

        total_width = sum(column_widths)
        pdf_canvas.setFillColor(colors.HexColor("#DAEBC8"))
        pdf_canvas.rect(self.left_margin, top_y - header_height, total_width, header_height, stroke=0, fill=1)
        pdf_canvas.setStrokeColor(colors.HexColor("#2E7D32"))
        pdf_canvas.setLineWidth(1)
        x = self.left_margin
        pdf_canvas.rect(self.left_margin, top_y - header_height, total_width, header_height, stroke=1, fill=0)
        for idx, column in enumerate(columns):
            width = column_widths[idx]
            pdf_canvas.drawCentredString(
                x + width / 2,
                top_y - (header_height / 2) + 2,
                column["title"].upper(),
            )
            if idx < len(columns) - 1:
                x += width
                pdf_canvas.line(x, top_y, x, top_y - header_height)
            else:
                x += width
        return top_y - header_height

    def _estimate_calibration_row_height(
        self,
        pdf_canvas: canvas.Canvas,
        entry: Dict,
        columns: list,
        column_widths: list,
        line_spacing: float,
    ) -> float:
        """Calcula la altura requerida por el registro para predecir saltos de página."""

        max_lines = 1
        for idx, column in enumerate(columns):
            text = entry.get(column["key"], "") or "N/A"
            max_width = column_widths[idx] - 8
            lines = self._wrap_text(text, "Helvetica", 10, max_width, pdf_canvas)
            max_lines = max(max_lines, len(lines))
        return max_lines * line_spacing + 0.18 * inch

    def _draw_calibration_table_row(
        self,
        pdf_canvas: canvas.Canvas,
        entry: Dict,
        columns: list,
        column_widths: list,
        current_y: float,
        line_spacing: float,
        row_index: int,
    ) -> float:
        """Dibuja una fila completa de la tabla de certificados."""

        total_width = sum(column_widths)
        x = self.left_margin
        lines_per_column = []
        max_lines = 1
        for idx, column in enumerate(columns):
            text = entry.get(column["key"], "") or "N/A"
            width = column_widths[idx] - 8
            lines = self._wrap_text(text, "Helvetica", 10, width, pdf_canvas)
            lines_per_column.append(lines)
            max_lines = max(max_lines, len(lines))

        row_height = max_lines * line_spacing + 0.18 * inch
        row_bottom = current_y - row_height
        fill_color = colors.white if row_index % 2 else colors.HexColor("#F8FBF3")
        pdf_canvas.setFillColor(fill_color)
        pdf_canvas.rect(self.left_margin, row_bottom, total_width, row_height, stroke=0, fill=1)
        pdf_canvas.setStrokeColor(colors.HexColor("#2E7D32"))
        pdf_canvas.rect(self.left_margin, row_bottom, total_width, row_height, stroke=1, fill=0)

        x = self.left_margin
        for idx, lines in enumerate(lines_per_column):
            width = column_widths[idx]
            text_x = x + 4
            text_y = current_y - 0.18 * inch
            pdf_canvas.setFont("Helvetica", 10)
            pdf_canvas.setFillColor(colors.black)
            for line in lines:
                pdf_canvas.drawString(text_x, text_y, line)
                text_y -= line_spacing
            x += width
            if idx < len(lines_per_column) - 1:
                pdf_canvas.line(x, row_bottom, x, current_y)

        return row_bottom

    def _draw_textual_section_page(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        title: str,
        entries: list,
        page_number: int,
        watermark_logo: Optional[str] = None,
        use_bullets: bool = False,
    ) -> int:
        """Dibuja una sección de texto libre (con o sin viñetas) manteniendo estilos."""

        filtered_entries = [entry.strip() for entry in entries if entry and entry.strip()]
        if not filtered_entries:
            return page_number - 1

        current_page = page_number
        text_width = self.page_width - self.left_margin - self.right_margin
        bullet_indent = 0.28 * inch
        line_spacing = 0.23 * inch

        y = self._prepare_text_section_header(
            pdf_canvas,
            report_data,
            title,
            watermark_logo,
        )

        for entry in filtered_entries:
            available_width = text_width - (bullet_indent if use_bullets else 0)
            lines = self._wrap_text(entry, "Helvetica", 11, available_width, pdf_canvas)
            required_height = len(lines) * line_spacing + 0.2 * inch

            if y - required_height <= self.bottom_margin:
                self._draw_footer(pdf_canvas, page_number=current_page)
                current_page += 1
                y = self._prepare_text_section_header(
                    pdf_canvas,
                    report_data,
                    f"{title} (cont.)",
                    watermark_logo,
                )

            line_y = y
            if use_bullets:
                pdf_canvas.setFont("Helvetica-Bold", 12)
                pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
                pdf_canvas.drawString(self.left_margin, line_y, "❖")
                text_x = self.left_margin + bullet_indent
            else:
                text_x = self.left_margin

            pdf_canvas.setFont("Helvetica", 11)
            pdf_canvas.setFillColor(colors.black)
            for line in lines:
                pdf_canvas.drawString(text_x, line_y, line)
                line_y -= line_spacing

            y = line_y - (0.08 * inch if use_bullets else 0.14 * inch)

        self._draw_footer(pdf_canvas, page_number=current_page)
        return current_page

    def _prepare_text_section_header(
        self,
        pdf_canvas: canvas.Canvas,
        report_data: Dict,
        title: str,
        watermark_logo: Optional[str],
    ) -> float:
        """Abre una página nueva con el encabezado institucional estándar."""

        pdf_canvas.showPage()
        if watermark_logo:
            self._add_watermark_logo(pdf_canvas, watermark_logo)

        y = self._draw_header_branding(pdf_canvas, report_data)
        y -= 0.35 * inch

        pdf_canvas.setFont("Helvetica-Bold", 14)
        pdf_canvas.setFillColor(colors.HexColor("#2E7D32"))
        pdf_canvas.drawString(self.left_margin, y, title)

        y -= 0.18 * inch
        pdf_canvas.setStrokeColor(colors.HexColor("#2E7D32"))
        pdf_canvas.setLineWidth(2)
        pdf_canvas.line(self.left_margin, y, self.page_width - self.right_margin, y)

        return y - 0.35 * inch

    def _should_include_audiometry_protocol(self, report_type: str) -> bool:
        """Retorna True cuando el informe es solo de audiometría."""

        normalized = (report_type or "").lower()
        has_audio = "audiometr" in normalized
        has_espiro = "espiro" in normalized
        return has_audio and not has_espiro

    def _should_include_spirometry_protocol(self, report_type: str) -> bool:
        """Retorna True cuando el informe es solo de espirometría."""

        normalized = (report_type or "").lower()
        has_audio = "audiometr" in normalized
        has_espiro = "espiro" in normalized
        return has_espiro and not has_audio

    def _split_into_paragraphs(self, text: str) -> list:
        """Divide un texto en párrafos preservando saltos en blanco."""

        if not text:
            return []

        paragraphs = []
        buffer = []
        for raw_line in text.replace("\r", "").splitlines():
            stripped = raw_line.strip()
            if stripped:
                buffer.append(stripped)
            elif buffer:
                paragraphs.append(" ".join(buffer))
                buffer = []

        if buffer:
            paragraphs.append(" ".join(buffer))

        return paragraphs

    def _split_into_bullet_items(self, text: str) -> list:
        """Convierte un bloque de texto en viñetas individuales limpiando símbolos previos."""

        if not text:
            return []

        items = []
        symbols = "•*-❖\uf0b6·"
        for raw_line in text.replace("\r", "").splitlines():
            stripped = raw_line.strip()
            if not stripped:
                continue
            while stripped and stripped[0] in symbols:
                stripped = stripped[1:].strip()
            items.append(stripped or raw_line.strip())

        return items

    def _resolve_evaluator_profile(self, report_data: Dict) -> Dict:
        """Devuelve siempre un perfil con los datos mostrados en el encabezado."""

        profile = report_data.get("evaluator_profile")
        if isinstance(profile, dict):
            return profile

        fallback_name = (report_data.get("evaluator") or "N/A").strip() or "N/A"
        return {
            "name": fallback_name,
            "header_label": "LICENCIADA",
            "profession": "TERAPEUTA RESPIRATORIA",
            "registry": "REGISTRO IDÓNEO 124",
            "technical_details": [
                "Terapeuta Respiratoria,",
                "Registro Idóneo 124.",
            ],
        }

    def _resolve_technical_team(self, report_data: Dict) -> list:
        """Obtiene la lista de integrantes del equipo técnico con datos normalizados."""

        raw_team = report_data.get("technical_team")
        resolved = []

        if isinstance(raw_team, list):
            for member in raw_team:
                if isinstance(member, dict):
                    name = (member.get("name") or "").strip()
                    if not name:
                        continue
                    details = []
                    details_field = member.get("details")
                    if isinstance(details_field, list):
                        details.extend([
                            str(detail).strip()
                            for detail in details_field
                            if str(detail).strip()
                        ])
                    elif isinstance(details_field, str) and details_field.strip():
                        details.append(details_field.strip())

                    for key in (
                        "profession",
                        "role",
                        "position",
                        "title",
                        "credential",
                        "license",
                        "registration",
                        "registry",
                    ):
                        value = member.get(key)
                        normalized = str(value).strip() if value else ""
                        if normalized and normalized not in details:
                            details.append(normalized)

                    resolved.append({"name": name, "details": details})

                elif isinstance(member, str):
                    name = member.strip()
                    if name:
                        resolved.append({"name": name, "details": []})

        if resolved:
            return resolved

        profile_entry = self._technical_entry_from_profile(report_data.get("evaluator_profile"))
        if profile_entry:
            return [profile_entry]

        if raw_team is None:
            return [
                {
                    "name": item["name"],
                    "details": list(item.get("details", [])),
                }
                for item in DEFAULT_TECHNICAL_TEAM
            ]

        return []

    def _resolve_calibration_certificates(self, report_data: Dict) -> list:
        """Normaliza la estructura de certificados recibida desde la aplicación."""

        raw_entries = report_data.get("calibration_certificates")
        resolved = []
        if isinstance(raw_entries, list):
            for entry in raw_entries:
                if not isinstance(entry, dict):
                    continue
                equipment = (entry.get("equipment") or entry.get("device") or "").strip()
                certificate = (entry.get("certificate") or entry.get("certificate_number") or "").strip()
                provider = (entry.get("provider") or entry.get("laboratory") or "").strip()
                if not equipment and not certificate:
                    continue
                resolved.append(
                    {
                        "equipment": equipment or certificate or "Equipo",
                        "model": (entry.get("model") or entry.get("equipment_model") or "").strip(),
                        "serial": (entry.get("serial") or entry.get("serial_number") or "").strip(),
                        "certificate": certificate or "N/A",
                        "calibration_date": (entry.get("calibration_date") or "").strip(),
                        "valid_until": (entry.get("valid_until") or entry.get("expiration_date") or "").strip(),
                        "provider": provider,
                        "attachment": (
                            entry.get("attachment")
                            or entry.get("file")
                            or entry.get("file_path")
                            or ""
                        ).strip(),
                    }
                )
        return resolved

    def _resolve_calibration_files(self, report_data: Dict) -> list:
        """Obtiene rutas válidas de certificados para anexarlos al PDF final."""

        files = []
        candidates = []
        raw_files = report_data.get("calibration_files")
        if raw_files:
            candidates.extend(raw_files if isinstance(raw_files, list) else [raw_files])

        raw_entries = report_data.get("calibration_certificates")
        if raw_entries:
            candidates.extend(raw_entries if isinstance(raw_entries, list) else [raw_entries])

        for entry in candidates:
            file_path = None
            if isinstance(entry, str):
                file_path = entry
            elif isinstance(entry, dict):
                file_path = entry.get("file") or entry.get("file_path") or entry.get("attachment")
            if not file_path:
                continue
            normalized = file_path.strip()
            if normalized and os.path.exists(normalized) and normalized not in files:
                files.append(normalized)
        return files

    def _resolve_file_list(self, raw_value) -> list:
        """Normaliza un conjunto arbitrario de rutas de archivo provenientes del informe."""

        resolved = []
        if not raw_value:
            return resolved

        candidates = raw_value if isinstance(raw_value, list) else [raw_value]
        for item in candidates:
            if not item:
                continue
            candidate = str(item).strip()
            if candidate and os.path.exists(candidate) and candidate not in resolved:
                resolved.append(candidate)
        return resolved

    def _append_supporting_pdfs(
        self,
        base_pdf_path: str,
        output_path: str,
        attachment_sections: list,
        trailing_pdfs: Optional[list] = None,
    ) -> None:
        """Concatena anexos PDF inmediatamente después de su portada y agrega PDFs al final."""

        reader = PdfReader(base_pdf_path)
        writer = PdfWriter()

        ordered_sections = list(attachment_sections)
        section_index = 0
        trailing_pdfs = [path for path in (trailing_pdfs or []) if path]

        for page_idx, page in enumerate(reader.pages, start=1):
            writer.add_page(page)

            while section_index < len(ordered_sections):
                section = ordered_sections[section_index]
                cover_end = section.get("cover_end")
                if cover_end != page_idx:
                    break

                for file_path in section.get("files", []):
                    if not os.path.exists(file_path):
                        continue
                    try:
                        attachment_reader = PdfReader(file_path)
                        for attachment_page in attachment_reader.pages:
                            if self._is_blank_pdf_page(attachment_page):
                                continue
                            writer.add_page(attachment_page)
                    except Exception as exc:  # pragma: no cover
                        print(f"Advertencia al adjuntar {file_path}: {exc}")

                section_index += 1

        for trailing_path in trailing_pdfs:
            if not os.path.exists(trailing_path):
                continue
            try:
                trailing_reader = PdfReader(trailing_path)
                for trailing_page in trailing_reader.pages:
                    if self._is_blank_pdf_page(trailing_page):
                        continue
                    writer.add_page(trailing_page)
            except Exception as exc:  # pragma: no cover
                print(f"Advertencia al adjuntar {trailing_path}: {exc}")

        with open(output_path, "wb") as destination:
            writer.write(destination)

        if os.path.exists(base_pdf_path):
            os.remove(base_pdf_path)

        for trailing_path in trailing_pdfs:
            if trailing_path and os.path.exists(trailing_path):
                os.remove(trailing_path)

    def _is_blank_pdf_page(self, page) -> bool:
        """Determina si una página PDF está vacía (solo espacios en blanco)."""

        try:
            text = page.extract_text() or ""
            if text.strip():
                return False

            contents = page.get_contents()
            if not contents:
                return True

            if isinstance(contents, list):
                data = b"".join(item.get_data() for item in contents if item)
            else:
                data = contents.get_data()

            stripped = data.translate(None, b" \t\r\n")
            return len(stripped) == 0
        except Exception:
            return False

    def _ensure_landscape_pdf(self, file_path: str) -> tuple[str, Optional[str]]:
        """Garantiza orientación horizontal para los anexos PDF."""

        try:
            reader = PdfReader(file_path)
        except Exception:
            return file_path, None

        processed_pages = []
        needs_rotation = False
        for page in reader.pages:
            box = page.mediabox
            width = float(box.right) - float(box.left)
            height = float(box.top) - float(box.bottom)
            rotated_page = page
            if height > width:
                rotated_page = page.rotate(90)
                needs_rotation = True
            processed_pages.append(rotated_page)

        if not needs_rotation:
            return file_path, None

        writer = PdfWriter()
        for page in processed_pages:
            writer.add_page(page)

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        temp_file.close()
        with open(temp_file.name, "wb") as temp_destination:
            writer.write(temp_destination)

        return temp_file.name, temp_file.name

    def _technical_entry_from_profile(self, profile: Optional[Dict]) -> Optional[Dict]:
        """Convierte el perfil del evaluador en un bloque para la sección técnica."""

        if not isinstance(profile, dict):
            return None

        name = (profile.get("name") or "").strip()
        if not name:
            return None

        details = []
        raw_details = profile.get("technical_details")
        if isinstance(raw_details, list):
            details = [
                str(detail).strip()
                for detail in raw_details
                if isinstance(detail, (str, int, float)) and str(detail).strip()
            ]

        if not details:
            profession = (profile.get("profession") or "").strip()
            registry = (profile.get("registry") or "").strip()
            if profession:
                suffix = "" if profession.endswith((",", ".")) else ","
                details.append(f"{profession}{suffix}")
            if registry:
                suffix = "" if registry.endswith((",", ".")) else "."
                details.append(f"{registry}{suffix}")

        if not details:
            return None

        return {"name": name, "details": details}

    def _compute_results_stats(self, dataset_key: str, entries: list) -> Dict[str, int]:
        """Cuenta ocurrencias por tipo de resultado para el esquema indicado."""

        scheme = RESULT_SCHEMES.get(dataset_key, RESULT_SCHEMES["audiometria"])
        stats = {option["key"]: 0 for option in scheme.get("options", [])}
        for entry in entries:
            key = self._resolve_result_code(entry, dataset_key)
            if key not in stats:
                stats[key] = 0
            stats[key] += 1
        stats["total"] = sum(stats.values())
        return stats

    def _build_results_analysis_text(self, dataset_key: str, stats: Dict[str, int]) -> str:
        """Genera un texto descriptivo con porcentajes según el esquema."""

        total = stats.get("total", 0)
        if total == 0:
            return "No se registraron personas evaluadas para este estudio."

        def pct_for(keys):
            if not isinstance(keys, (list, tuple)):
                keys = [keys]
            return (sum(stats.get(key, 0) for key in keys) / total) * 100 if total else 0.0

        if dataset_key == "espirometria":
            normal_pct = pct_for("espiro_normal")
            leves_pct = pct_for(["restriccion_leve", "obstruccion_leve", "obstruccion_restriccion_leve"])
            moderadas_pct = pct_for(["restriccion_moderada", "obstruccion_moderada"])
            grave_pct = pct_for("restriccion_grave")
            return (
                f"Del total de espirometrías aplicadas, el {normal_pct:.0f}% se mantiene dentro de parámetros normales, "
                f"el {leves_pct:.0f}% presenta restricciones u obstrucciones leves, "
                f"el {moderadas_pct:.0f}% evidencia compromiso moderado y el {grave_pct:.0f}% corresponde a casos graves."
            )

        normal_pct = pct_for("normal")
        unilateral_pct = pct_for("caida_unilateral")
        bilateral_pct = pct_for("caida_bilateral")
        return (
            f"En base a los colaboradores evaluados, el {normal_pct:.0f}% presenta respuestas dentro de los límites normales, "
            f"el {unilateral_pct:.0f}% requiere vigilancia auditiva por caída unilateral y el {bilateral_pct:.0f}% presenta caída bilateral."
        )

    def _create_results_chart_image(
        self,
        dataset_key: str,
        stats: Dict[str, int],
        scheme: Dict,
    ) -> Optional[str]:
        """Genera una gráfica de pastel con Seaborn para el conjunto indicado."""


        try:
            import matplotlib.pyplot as plt
            import seaborn as sns
            from matplotlib import patheffects
        except Exception as exc:  # pragma: no cover
            print(f"No se pudo cargar seaborn/matplotlib: {exc}")

            return None

        segments = []
        for option in scheme.get("options", []):
            key = option.get("key")
            label = option.get("chart_label") or option.get("label") or key
            segments.append((label, stats.get(key, 0), key, option))

        total_count = sum(val for _label, val, _code, _opt in segments)

        if total_count == 0:
            return None


        sns.set_theme(style="white")
        fig, ax = plt.subplots(figsize=(4.6, 3.6), dpi=200)

        fig.patch.set_alpha(0)

        def hex_to_rgb_fraction(hex_color: str) -> tuple:

            hex_color = hex_color.lstrip("#")
            return tuple(int(hex_color[i : i + 2], 16) / 255.0 for i in (0, 2, 4))


        vivid_palette_hex = [
            "#0072CE",  # azul vivo
            "#FF6F00",  # naranja vivo
            "#00A86B",  # verde esmeralda
            "#D81B60",  # magenta
            "#6A1B9A",  # violeta
            "#E53935",  # rojo vivo
            "#00897B",  # turquesa
            "#F9A825",  # amarillo ámbar
        ]
        fallback_color = hex_to_rgb_fraction(vivid_palette_hex[0])

        data_values = []
        data_labels = []
        palette = []
        color_idx = 0
        for label, value, _key, option in segments:
            if value > 0:
                data_values.append(value)
                data_labels.append(str(label).upper())
            palette_hex = vivid_palette_hex[color_idx % len(vivid_palette_hex)]
            palette.append(hex_to_rgb_fraction(palette_hex))
            color_idx += 1


        if not data_values:
            data_values = [1]
            data_labels = ["SIN DATOS"]
            palette = [fallback_color]

        def label_fmt(pct: float) -> str:
            absolute = int(round(pct * total_count / 100.0))
            return f"{pct:.0f}%\n({absolute})"

        def adjust_color(color: tuple, factor: float) -> tuple:

            return tuple(min(1, max(0, c * factor)) for c in color)

        base_colors = [adjust_color(col, 0.8) for col in palette]

        ax.pie(

            data_values,
            colors=base_colors,
            startangle=90,
            radius=1.02,
            wedgeprops={"linewidth": 0, "edgecolor": "none"},
            center=(0, -0.08),
        )

        explode = [0.06] * len(data_values)
        wedges, texts, autotexts = ax.pie(
            data_values,
            colors=palette,
            startangle=90,

            explode=explode,
            shadow=False,
            autopct=label_fmt,
            pctdistance=0.74,

            textprops={"fontsize": 8, "color": "#212121"},
            wedgeprops={"linewidth": 1.4, "edgecolor": "#FFFFFF"},
        )

        for wedge in wedges:
            wedge.set_linewidth(1.4)
            wedge.set_edgecolor("#FFFFFF")
        for autotext in autotexts:
            autotext.set_fontweight("bold")

        total_text = ax.text(
            0,
            -0.02,
            f"TOTAL\n{total_count}",
            ha="center",

            va="center",
            fontsize=10,
            fontweight="bold",

            color="#FFFFFF",
        )
        total_text.set_path_effects([
            patheffects.withStroke(linewidth=2.2, foreground="#1B5E20")
        ])

        legend_labels = []
        for label, val in zip(data_labels, data_values):
            percent = (val / total_count) * 100 if total_count else 0
            legend_labels.append(f"{label}: {val} ({percent:.0f}%)")

        ax.legend(
            wedges,
            legend_labels,
            loc="center left",
            bbox_to_anchor=(1.1, 0.5),
            frameon=False,
            fontsize=8,
            labelcolor=["#2b2b2b"] * len(legend_labels),
        )
        chart_label = scheme.get("chart_label", self._get_results_label(dataset_key))
        ax.set_title(
            f"Distribución porcentual - {chart_label}",
            fontsize=10,
            pad=12,
        )
        ax.axis("equal")

        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
        temp_file_path = temp_file.name
        temp_file.close()

        fig.savefig(temp_file_path, format="png", bbox_inches="tight", transparent=True)
        plt.close(fig)
        return temp_file_path

    def _resolve_result_code(self, entry: Dict, dataset_key: str = "audiometria") -> str:
        """Normaliza un registro para ubicarlo dentro del esquema solicitado."""

        scheme = RESULT_SCHEMES.get(dataset_key, RESULT_SCHEMES["audiometria"])
        valid_keys = {opt["key"] for opt in scheme.get("options", [])}

        raw_code = entry.get("result_code")
        if raw_code in valid_keys:
            return raw_code

        label_source = entry.get("result_label") or entry.get("result") or (entry.get("results") or {}).get("status", "")
        normalized = self._normalize_text(label_source)

        if dataset_key == "espirometria":
            mapping = [
                ("espiro_normal", ["normal"]),
                ("restriccion_leve", ["restriccion leve", "leve"]),
                ("obstruccion_leve", ["obstruccion leve"]),
                ("obstruccion_restriccion_leve", ["obstruccion a restriccion", "obstruccion a restriccion leve"]),
                ("restriccion_moderada", ["restriccion moderada"]),
                ("obstruccion_moderada", ["obstruccion moderada"]),
                ("restriccion_grave", ["restriccion grave", "grave"]),
            ]
            for key, tokens in mapping:
                if any(token in normalized for token in tokens):
                    return key
            return "espiro_normal"

        if "caida" in normalized and "bilateral" in normalized:
            return "caida_bilateral"
        if "caida" in normalized and "uni" in normalized:
            return "caida_unilateral"
        return "normal"

    @staticmethod
    def _normalize_text(value: Optional[str]) -> str:
        """Elimina tildes y convierte a minúsculas para comparaciones simples."""

        if not value:
            return ""
        normalized = value.lower()
        for accented, plain in (("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"), ("ñ", "n")):
            normalized = normalized.replace(accented, plain)
        return normalized

    def _resolve_results_title(self, report_type: str) -> str:
        """Genera el título según el tipo de estudio seleccionado."""

        study = self._get_results_label(report_type)
        return f"CUADRO DE RESULTADOS DE LAS PRUEBAS DE {study}."

    def _get_results_label(self, report_type: str) -> str:
        """Devuelve la etiqueta base para AUDIOMETRÍAS/ESPIROMETRÍAS."""

        normalized = (report_type or "").lower()
        if "audiometr" in normalized and "espirom" in normalized:
            return "AUDIOMETRÍAS Y ESPIROMETRÍAS"
        if "audiometr" in normalized:
            return "AUDIOMETRÍAS"
        return "ESPIROMETRÍAS"

    def _resolve_entry_value(self, entry: Dict, key: str, index: int) -> str:
        """Obtiene el valor visible para cada columna manejando respaldos."""

        if key == "index":
            return str(index)
        if key == "name":
            return entry.get("name") or entry.get("full_name") or "N/A"
        if key == "identification":
            return entry.get("identification") or entry.get("cedula") or "N/A"
        if key == "age":
            value = entry.get("age") or entry.get("edad")
            return str(value) if value else ""
        if key == "position":
            return entry.get("position") or entry.get("cargo") or ""
        if key == "result":
            return (
                entry.get("result_label")
                or entry.get("result")
                or (entry.get("results") or {}).get("status")
                or "N/A"
            )
        return ""

    def _wrap_text(
        self,
        text: str,
        font_name: str,
        font_size: int,
        max_width: float,
        pdf_canvas: canvas.Canvas,
    ) -> list:
        """Divide texto largo en múltiples líneas para evitar desbordes."""

        if not text:
            return ["N/A"]

        words = text.split()
        lines = []
        current_line = ""

        for word in words:
            test_line = f"{current_line} {word}".strip()
            if pdf_canvas.stringWidth(test_line, font_name, font_size) <= max_width:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                current_line = word

        if current_line:
            lines.append(current_line)

        return lines or ["N/A"]

    def _add_file_link(
        self,
        pdf_canvas: canvas.Canvas,
        target_path: str,
        x: float,
        y: float,
        width: float,
        height: float,
        relative: bool = False,
    ) -> None:
        """Crea un enlace local (file://) o relativo al PDF."""

        if not target_path:
            return

        file_url = self._build_file_url(target_path, relative=relative)
        if not file_url:
            return

        pdf_canvas.linkURL(file_url, (x, y, x + width, y + height), relative=1 if relative else 0)

    def _build_file_url(self, path_value: str, relative: bool = False) -> str:
        """Construye un enlace file:// compatible con rutas locales o relativas."""

        if relative:
            relative_path = path_value.replace("\\", "/")
            return quote(relative_path)

        try:
            resolved = Path(path_value).resolve()
        except OSError:
            resolved = Path(path_value)

        posix_path = resolved.as_posix()
        if ":" in posix_path:
            posix_path = "/" + posix_path
        return "file://" + quote(posix_path)

    def _links_are_relative(self, report_data: Dict) -> bool:
        """Indica si los enlaces deben ser relativos (solo ZIP)."""

        return report_data.get("link_mode") == "relative"

    def _draw_footer(self, pdf_canvas: canvas.Canvas, page_number: int = 1) -> None:
        """Pie de página con información de generación y numeración."""

        y = self.bottom_margin / 2
        motto_y = y + 0.22 * inch
        pdf_canvas.setFont("Helvetica-Oblique", 9)
        pdf_canvas.setFillColor(colors.HexColor("#666666"))
        pdf_canvas.drawCentredString(self.page_width / 2, motto_y, '"EL PILAR DE TUS SENTIDOS".')

        pdf_canvas.setFont("Helvetica", 8)
        pdf_canvas.setFillColor(colors.black)
        pdf_canvas.drawString(self.left_margin, y, f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        pdf_canvas.drawString(
            self.page_width - self.right_margin - 0.5 * inch,
            y,
            f"Página {page_number}",
        )
