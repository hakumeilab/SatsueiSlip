from __future__ import annotations

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.pdfmetrics import registerFont
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from satsuei_slip.models import DeliveryInfo, VideoItem

JP_FONT_NAME = "YuGothic"
JP_FONT_BOLD_NAME = "YuGothic-Bold"
FALLBACK_FONT_NAME = "HeiseiKakuGo-W5"
YU_GOTHIC_REGULAR_PATH = Path("C:/Windows/Fonts/YuGothR.ttc")
YU_GOTHIC_BOLD_PATH = Path("C:/Windows/Fonts/YuGothB.ttc")
CUT_FPS = 24
ROWS_PER_COLUMN = 25
ITEMS_PER_PAGE = 50
PAGE_WIDTH = 190 * mm
NUMBER_GUTTER_WIDTH = 10 * mm
BOX_WIDTH = 180 * mm
HALF_BOX_WIDTH = 90 * mm


class DeliverySlipPdfExporter:
    def __init__(self) -> None:
        self.font_name, self.bold_font_name = self._register_fonts()
        styles = getSampleStyleSheet()
        self.normal_style = ParagraphStyle(
            "SlipBody",
            parent=styles["BodyText"],
            fontName=self.bold_font_name,
            fontSize=10,
            leading=12,
            alignment=TA_LEFT,
            textColor=colors.HexColor("#111111"),
        )
        self.right_style = ParagraphStyle(
            "SlipBodyRight",
            parent=self.normal_style,
            alignment=TA_RIGHT,
        )
        self.center_style = ParagraphStyle(
            "SlipBodyCenter",
            parent=self.normal_style,
            alignment=TA_CENTER,
        )
        self.gray_center_style = ParagraphStyle(
            "SlipGrayCenter",
            parent=self.meta_label_style if hasattr(self, "meta_label_style") else self.normal_style,
            fontName=self.font_name,
            fontSize=9,
            leading=11,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#8a8a8a"),
        )
        self.title_style = ParagraphStyle(
            "SlipTitle",
            parent=self.normal_style,
            fontName=self.bold_font_name,
            fontSize=16,
            leading=18,
        )
        self.small_style = ParagraphStyle(
            "SlipSmall",
            parent=self.normal_style,
            fontName=self.bold_font_name,
            fontSize=9,
            leading=11,
        )
        self.meta_label_style = ParagraphStyle(
            "SlipMetaLabel",
            parent=self.normal_style,
            fontName=self.font_name,
            fontSize=8,
            leading=9,
            textColor=colors.HexColor("#8a8a8a"),
        )
        self.bold_label_style = ParagraphStyle(
            "SlipBoldLabel",
            parent=self.normal_style,
            fontName=self.bold_font_name,
            fontSize=10,
            leading=12,
        )

    def export(self, output_path: Path, delivery_info: DeliveryInfo, items: list[VideoItem]) -> None:
        doc = SimpleDocTemplate(
            str(output_path),
            pagesize=A4,
            leftMargin=10 * mm,
            rightMargin=10 * mm,
            topMargin=10 * mm,
            bottomMargin=10 * mm,
            title="納品伝票",
        )

        story = []
        total_pages = max(1, (len(items) + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE)
        for page_index in range(total_pages):
            page_items = items[page_index * ITEMS_PER_PAGE : (page_index + 1) * ITEMS_PER_PAGE]
            story.append(self._build_header(delivery_info, items))
            story.append(Spacer(1, 16))
            story.append(
                self._build_detail_table(
                    page_items,
                    delivery_info.head_trim_frames,
                )
            )
            story.append(Spacer(1, 18))
            story.append(self._build_footer_table(delivery_info, page_items))
            if page_index < total_pages - 1:
                story.append(PageBreak())

        doc.build(story)

    def _build_header(self, delivery_info: DeliveryInfo, items: list[VideoItem]) -> Table:
        right_info = Table(
            [
                ["date", delivery_info.delivery_date.strftime("%Y-%m-%d")],
                ["size", self._summary_resolution(items)],
                ["bold", str(max(0, delivery_info.head_trim_frames))],
                ["fps", self._summary_fps(items)],
            ],
            colWidths=[24 * mm, 66 * mm],
            rowHeights=[6 * mm, 6 * mm, 6 * mm, 6 * mm],
        )
        right_info.setStyle(self._plain_box_style(label_cols=(0,)))

        right_second = Table(
            [["episode", self._p(delivery_info.episode_name), "roll", self._p(delivery_info.folder_name)]],
            colWidths=[18 * mm, 24 * mm, 22 * mm, 26 * mm],
            rowHeights=[12 * mm],
        )
        right_second.setStyle(self._plain_box_style(label_cols=(0, 2)))

        rows = [
            [
                "",
                self._header_label_value_box(
                    "company",
                    self._company_display_name(delivery_info.company_name),
                    24 * mm,
                    max_font_size=16,
                    min_font_size=9,
                    text_width=HALF_BOX_WIDTH - 8 * mm,
                ),
                right_info,
            ],
            [
                "",
                self._header_label_value_box(
                    "title",
                    delivery_info.project_name,
                    12 * mm,
                    max_font_size=16,
                    min_font_size=9,
                    text_width=HALF_BOX_WIDTH - 8 * mm,
                ),
                right_second,
            ],
        ]
        table = Table(
            rows,
            colWidths=[NUMBER_GUTTER_WIDTH, HALF_BOX_WIDTH, HALF_BOX_WIDTH],
            rowHeights=[24 * mm, 12 * mm],
        )
        table.setStyle(
            TableStyle(
                [
                    ("BOX", (1, 0), (2, -1), 1.0, colors.black),
                    ("INNERGRID", (1, 0), (2, -1), 0.8, colors.black),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (0, -1), 0),
                    ("RIGHTPADDING", (0, 0), (0, -1), 0),
                    ("TOPPADDING", (0, 0), (0, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (0, -1), 0),
                    ("LEFTPADDING", (1, 0), (1, -1), 6),
                    ("RIGHTPADDING", (1, 0), (1, -1), 6),
                    ("TOPPADDING", (1, 0), (1, -1), 4),
                    ("BOTTOMPADDING", (1, 0), (1, -1), 4),
                    ("LEFTPADDING", (2, 0), (2, -1), 0),
                    ("RIGHTPADDING", (2, 0), (2, -1), 0),
                    ("TOPPADDING", (2, 0), (2, -1), 0),
                    ("BOTTOMPADDING", (2, 0), (2, -1), 0),
                ]
            )
        )
        return table

    def _build_detail_table(
        self,
        items: list[VideoItem],
        head_trim_frames: int,
    ) -> Table:
        rows = []
        for row_index in range(ROWS_PER_COLUMN):
            left_item = items[row_index] if row_index < len(items) else None
            right_pos = row_index + ROWS_PER_COLUMN
            right_item = items[right_pos] if right_pos < len(items) else None
            rows.append(
                [
                    str(self._left_display_number(row_index)),
                    self._p(left_item.file_name if left_item else "", self.small_style),
                    self._duration_cell(left_item, head_trim_frames),
                    str(self._right_display_number(row_index)),
                    self._p(right_item.file_name if right_item else "", self.small_style),
                    self._duration_cell(right_item, head_trim_frames),
                ]
            )

        table = Table(
            rows,
            colWidths=[10 * mm, 63 * mm, 22 * mm, 10 * mm, 63 * mm, 22 * mm],
            rowHeights=[6.8 * mm] * ROWS_PER_COLUMN,
        )
        table.setStyle(
            TableStyle(
                [
                    ("FONTNAME", (0, 0), (-1, -1), self.bold_font_name),
                    ("FONTNAME", (0, 0), (0, -1), self.font_name),
                    ("FONTNAME", (3, 0), (3, -1), self.font_name),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                    ("BOX", (1, 0), (2, -1), 1.0, colors.black),
                    ("BOX", (4, 0), (5, -1), 1.0, colors.black),
                    ("LINEBEFORE", (1, 0), (1, -1), 1.0, colors.black),
                    ("LINEAFTER", (2, 0), (2, -1), 1.0, colors.black),
                    ("LINEBEFORE", (4, 0), (4, -1), 1.0, colors.black),
                    ("LINEAFTER", (5, 0), (5, -1), 1.0, colors.black),
                    ("INNERGRID", (1, 0), (2, -1), 0.5, colors.HexColor("#777777")),
                    ("INNERGRID", (4, 0), (5, -1), 0.5, colors.HexColor("#777777")),
                    ("ALIGN", (0, 0), (0, -1), "RIGHT"),
                    ("ALIGN", (2, 0), (2, -1), "CENTER"),
                    ("ALIGN", (3, 0), (3, -1), "RIGHT"),
                    ("ALIGN", (5, 0), (5, -1), "CENTER"),
                    ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#8a8a8a")),
                    ("TEXTCOLOR", (3, 0), (3, -1), colors.HexColor("#8a8a8a")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 2),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ]
                + [
                    ("BACKGROUND", (0, row), (-1, row), colors.HexColor("#f5f5f5"))
                    for row in range(1, ROWS_PER_COLUMN, 2)
                ]
            )
        )
        return table

    def _build_footer_table(self, delivery_info: DeliveryInfo, items: list[VideoItem]) -> Table:
        total_frames = sum(item.trimmed_frame_count(delivery_info.head_trim_frames) for item in items)
        total_seconds, total_remain = divmod(total_frames, CUT_FPS)
        memo_text = delivery_info.note
        estimated_count = sum(1 for item in items if item.frame_count_estimated)
        if estimated_count > 0:
            suffix = f"※ *付き相当の{estimated_count}件は推定フレームを含みます。"
            memo_text = f"{memo_text}\n{suffix}" if memo_text else suffix

        rows = [
            [
                "",
                self._inline_meta_value("count", f"{len(items)} / 50", align_right=False),
                self._inline_meta_value("total", f"{total_seconds} + {total_remain:02d}", align_right=True),
            ],
            ["", "", ""],
            ["", self._memo_box(memo_text), ""],
            ["", self._p(delivery_info.sender_footer, self.right_style), ""],
        ]
        table = Table(
            rows,
            colWidths=[NUMBER_GUTTER_WIDTH, HALF_BOX_WIDTH, HALF_BOX_WIDTH],
            rowHeights=[12 * mm, 6 * mm, 18 * mm, 16 * mm],
        )
        table.setStyle(
            TableStyle(
                [
                    ("BOX", (1, 0), (2, 0), 1.0, colors.black),
                    ("BOX", (1, 2), (2, 3), 1.0, colors.black),
                    ("LINEAFTER", (1, 0), (1, 0), 0.8, colors.black),
                    ("LINEBELOW", (1, 2), (2, 2), 0.8, colors.black),
                    ("LINEBEFORE", (2, 0), (2, 0), 0.8, colors.black),
                    ("SPAN", (1, 1), (2, 1)),
                    ("SPAN", (1, 2), (2, 2)),
                    ("SPAN", (1, 3), (2, 3)),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("FONTNAME", (0, 0), (-1, -1), self.font_name),
                    ("FONTSIZE", (0, 0), (-1, -1), 10),
                    ("LEFTPADDING", (0, 0), (0, -1), 0),
                    ("RIGHTPADDING", (0, 0), (0, -1), 0),
                    ("TOPPADDING", (0, 0), (0, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (0, -1), 0),
                    ("LEFTPADDING", (1, 0), (2, -1), 6),
                    ("RIGHTPADDING", (1, 0), (2, -1), 6),
                    ("TOPPADDING", (1, 0), (2, -1), 4),
                    ("BOTTOMPADDING", (1, 0), (2, -1), 4),
                    ("TOPPADDING", (1, 1), (2, 1), 0),
                    ("BOTTOMPADDING", (1, 1), (2, 1), 0),
                ]
            )
        )
        return table

    def _duration_cell(self, item: VideoItem | None, head_trim_frames: int) -> str | Table:
        if item is None:
            return ""
        return self._p(item.cut_duration_text(head_trim_frames, CUT_FPS), self.center_style)

    def _summary_resolution(self, items: list[VideoItem]) -> str:
        values = {item.resolution_text for item in items if item.resolution_text != "-"}
        if not values:
            return ""
        if len(values) == 1:
            return values.pop()
        return "混在"

    def _summary_fps(self, items: list[VideoItem]) -> str:
        values = {item.fps_text for item in items if item.fps_text != "-"}
        if not values:
            return ""
        if len(values) == 1:
            return values.pop()
        return "混在"

    def _plain_box_style(self, label_cols: tuple[int, ...] = (0,)) -> TableStyle:
        commands = [
            ("BOX", (0, 0), (-1, -1), 0.8, colors.black),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("FONTNAME", (0, 0), (-1, -1), self.bold_font_name),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]
        for col in label_cols:
            commands.append(("BACKGROUND", (col, 0), (col, -1), colors.HexColor("#f5f5f5")))
            commands.append(("TEXTCOLOR", (col, 0), (col, -1), colors.HexColor("#8a8a8a")))
            commands.append(("FONTNAME", (col, 0), (col, -1), self.font_name))
        return TableStyle(
            commands
        )

    def _p(self, text: str, style: ParagraphStyle | None = None) -> Paragraph:
        escaped = (
            (text or "")
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\n", "<br/>")
        )
        return Paragraph(escaped if escaped else "&nbsp;", style or self.normal_style)

    def _meta_value_block(
        self,
        label: str,
        value: str,
        value_style: ParagraphStyle,
    ) -> Table:
        table = Table(
            [[Paragraph(label, self.meta_label_style)], [self._p(value, value_style)]],
            colWidths=[None],
            rowHeights=[4 * mm, None],
        )
        table.setStyle(
            TableStyle(
                [
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ]
            )
        )
        return table

    def _inline_meta_value(self, label: str, value: str, align_right: bool) -> Table:
        table = Table(
            [[Paragraph(label, self.meta_label_style), self._p(value, self.center_style)]],
            colWidths=[14 * mm, 70 * mm],
        )
        table.setStyle(
            TableStyle(
                [
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ]
            )
        )
        return table

    def _memo_box(self, memo_text: str) -> Table:
        table = Table(
            [[Paragraph("memo", self.meta_label_style)], [self._p(memo_text, self.normal_style)]],
            colWidths=[BOX_WIDTH],
            rowHeights=[4 * mm, 14 * mm],
        )
        table.setStyle(
            TableStyle(
                [
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ]
            )
        )
        return table

    def _header_label_value_box(
        self,
        label: str,
        value: str,
        total_height,
        max_font_size: int,
        min_font_size: int,
        text_width,
    ) -> Table:
        label_height = 4 * mm
        value_style = self._fit_text_style(
            value,
            self.title_style,
            max_font_size=max_font_size,
            min_font_size=min_font_size,
            text_width=text_width,
        )
        table = Table(
            [[Paragraph(label, self.meta_label_style)], [self._p(value, value_style)]],
            colWidths=[HALF_BOX_WIDTH],
            rowHeights=[label_height, total_height - label_height],
        )
        table.setStyle(
            TableStyle(
                [
                    ("LINEBELOW", (0, 0), (0, 0), 0.5, colors.HexColor("#b0b0b0")),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 1),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                    ("VALIGN", (0, 0), (0, 0), "TOP"),
                    ("VALIGN", (0, 1), (0, 1), "MIDDLE"),
                ]
            )
        )
        return table

    def _fit_text_style(
        self,
        text: str,
        base_style: ParagraphStyle,
        max_font_size: int,
        min_font_size: int,
        text_width,
    ) -> ParagraphStyle:
        value = text or ""
        font_size = max_font_size
        while font_size > min_font_size:
            if stringWidth(value, base_style.fontName, font_size) <= text_width:
                break
            font_size -= 1

        style = ParagraphStyle(
            f"{base_style.name}_{font_size}",
            parent=base_style,
            fontName=base_style.fontName,
            fontSize=font_size,
            leading=font_size + 2,
            alignment=base_style.alignment,
            textColor=base_style.textColor,
        )
        return style

    def _register_fonts(self) -> tuple[str, str]:
        try:
            if YU_GOTHIC_REGULAR_PATH.is_file() and YU_GOTHIC_BOLD_PATH.is_file():
                registerFont(TTFont(JP_FONT_NAME, str(YU_GOTHIC_REGULAR_PATH)))
                registerFont(TTFont(JP_FONT_BOLD_NAME, str(YU_GOTHIC_BOLD_PATH)))
                return JP_FONT_NAME, JP_FONT_BOLD_NAME
        except Exception:
            pass

        registerFont(UnicodeCIDFont(FALLBACK_FONT_NAME))
        return FALLBACK_FONT_NAME, FALLBACK_FONT_NAME

    def _company_display_name(self, company_name: str) -> str:
        value = company_name.strip()
        if not value:
            return ""
        if value.endswith("様"):
            return value
        return f"{value} 様"

    def _left_display_number(self, row_index: int) -> int:
        return row_index + 1

    def _right_display_number(self, row_index: int) -> int:
        return row_index + ROWS_PER_COLUMN + 1
