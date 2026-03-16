import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re
import traceback


class PDFProcessor:
    def __init__(self, translator, translation_level="normal"):
        self.translator = translator
        self.translation_level = translation_level

    def _extract_line_text_with_spacing(self, line):
        """Extract text from a line's spans with proper spacing.

        PDF spans often lack spaces between them. Use position gaps to detect
        where spaces should be inserted.
        """
        spans = line.get("spans", [])
        if not spans:
            return "", 0, False

        parts = []
        prev_end_x = None
        max_size = 0
        is_bold = False

        for span in spans:
            text = span["text"]
            if not text:
                continue

            if prev_end_x is not None:
                gap = span["origin"][0] - prev_end_x
                char_width = span["size"] * 0.5
                if gap > char_width * 0.3:
                    parts.append(" ")

            parts.append(text)
            prev_end_x = span.get("bbox", [0, 0, 0, 0])[2] if "bbox" in span else (
                span["origin"][0] + len(text) * span["size"] * 0.5
            )

            if span["size"] > max_size:
                max_size = span["size"]
            if "bold" in span.get("font", "").lower():
                is_bold = True

        return "".join(parts), max_size, is_bold

    def _extract_block_text(self, block):
        """Extract full block text by joining lines with spaces.
        Returns (text, max_font_size, is_bold, bbox).
        """
        block_lines = []
        max_size = 0
        is_bold = False

        for line in block.get("lines", []):
            line_text, line_size, line_bold = self._extract_line_text_with_spacing(line)
            line_text = line_text.strip()
            if not line_text:
                continue
            block_lines.append(line_text)
            if line_size > max_size:
                max_size = line_size
            if line_bold:
                is_bold = True

        full_text = " ".join(block_lines)
        return full_text, max_size, is_bold, block["bbox"]

    def _is_footer_or_header(self, text, y_top, y_bottom, page_height):
        """Detect footer/header/copyright text that should be skipped entirely."""
        stripped = text.strip()

        # Page number prefix pattern: "3INTERNAL", "12INTERNAL", "3 INTERNAL"
        if re.match(r'^\d+\s*INTERNAL', stripped, re.IGNORECASE):
            return True

        # "INTERNAL – SAP" patterns
        if re.search(r'INTERNAL\s*[–\-]\s*SAP', stripped, re.IGNORECASE):
            return True

        # Standalone "INTERNAL" at bottom of page
        if y_bottom > page_height * 0.90 and re.search(r'INTERNAL', stripped, re.IGNORECASE):
            return True

        # Copyright notices
        if '© SAP' in stripped or 'All rights reserved' in stripped:
            return True

        # Bottom 10% of page with very short text (page numbers, marks)
        if y_bottom > page_height * 0.90 and len(stripped) <= 30:
            return True

        return False

    def _should_skip_text(self, text):
        """Check if a specific text string should skip translation."""
        stripped = text.strip()
        if not stripped:
            return True

        # Already Korean
        if any('\uac00' <= c <= '\ud7a3' for c in stripped):
            return True

        # Pure numbers, dates, or version strings
        if re.match(r'^[\d\.\-/\s:,]+$', stripped):
            return True

        # All-caps short acronyms (2-6 chars, allowing /)
        if re.match(r'^[A-Z/]{2,6}$', stripped):
            return True

        # Fiori app IDs
        if re.match(r'^F\d{4,5}$', stripped):
            return True

        # Single characters or just punctuation
        if len(stripped) <= 2 and not stripped.isalpha():
            return True

        # Check against do-not-translate glossary entries (exact match)
        if hasattr(self.translator, '_do_not_translate'):
            if stripped.lower() in self.translator._do_not_translate:
                return True

        return False

    def _should_skip_as_label(self, text, font_size):
        """For 'normal' level: skip very short label-like text with large font (section titles)."""
        if self.translation_level == "thorough":
            return False

        stripped = text.strip()
        words = stripped.split()

        # Very short text with very large font (≥ 22pt) = likely a major section title
        if font_size >= 22 and len(words) <= 5:
            return True

        return False

    def process_pdf(self, input_stream, output_stream, progress_callback=None):
        """Extract text from PDF, translate, and generate DOCX output."""
        try:
            from concurrent.futures import ThreadPoolExecutor, as_completed

            doc = fitz.open(stream=input_stream.read(), filetype="pdf")
            total_pages = len(doc)

            # Step 1: Extract all text blocks from all pages
            all_pages_data = []  # list of (page_num, items_list)

            for page_num in range(total_pages):
                page = doc[page_num]
                page_height = page.rect.height

                text_dict = page.get_text("dict")
                page_items = []

                for block in text_dict.get("blocks", []):
                    if block["type"] != 0:  # Skip image blocks
                        continue

                    text, font_size, is_bold, bbox = self._extract_block_text(block)
                    text = text.strip()
                    if not text:
                        continue

                    y_top = bbox[1]
                    y_bottom = bbox[3]

                    # Skip footer/header/copyright
                    if self._is_footer_or_header(text, y_top, y_bottom, page_height):
                        continue

                    # Skip non-translatable text
                    if self._should_skip_text(text):
                        page_items.append({
                            "text": text,
                            "size": font_size,
                            "bold": is_bold,
                            "skip": True,
                            "y": y_top
                        })
                        continue

                    # Check label skip
                    if self._should_skip_as_label(text, font_size):
                        page_items.append({
                            "text": text,
                            "size": font_size,
                            "bold": is_bold,
                            "skip": True,
                            "y": y_top
                        })
                        continue

                    # Translate this text
                    page_items.append({
                        "text": text,
                        "size": font_size,
                        "bold": is_bold,
                        "skip": False,
                        "y": y_top
                    })

                # Sort items by vertical position
                page_items.sort(key=lambda x: x["y"])
                all_pages_data.append((page_num, page_items))

            # Step 2: Collect unique texts to translate
            unique_texts = set()
            for page_num, items in all_pages_data:
                for item in items:
                    if not item["skip"]:
                        unique_texts.add(item["text"])

            # Step 3: Translate unique texts in parallel
            total_unique = len(unique_texts)
            translation_map = {}
            translation_errors = []
            processed_count = 0

            if unique_texts:
                max_workers = 15
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    future_to_text = {
                        executor.submit(self.translator.translate, text): text
                        for text in unique_texts
                    }

                    for future in as_completed(future_to_text):
                        text = future_to_text[future]
                        try:
                            result = future.result()
                            translation_map[text] = result if result else text
                        except Exception as e:
                            err_msg = f"'{text[:30]}...' 번역 실패: {str(e)}"
                            translation_errors.append(err_msg)
                            translation_map[text] = text

                        processed_count += 1
                        if progress_callback:
                            progress_callback(min((processed_count / total_unique) * 0.8, 0.8))

            # Step 4: Generate DOCX output
            docx_doc = Document()

            # Set default font
            style = docx_doc.styles['Normal']
            font = style.font
            font.name = 'Malgun Gothic'
            font.size = Pt(10)

            # Reduce paragraph spacing
            style.paragraph_format.space_before = Pt(1)
            style.paragraph_format.space_after = Pt(1)

            pages_written = 0
            for page_idx, (page_num, items) in enumerate(all_pages_data):
                # Skip pages with no content
                if not items:
                    continue

                # Add page separator (except for first page with content)
                if pages_written > 0:
                    docx_doc.add_page_break()

                # Add page header
                header_para = docx_doc.add_paragraph()
                header_run = header_para.add_run(f"── 페이지 {page_num + 1} / {total_pages} ──")
                header_run.font.size = Pt(8)
                header_run.font.color.rgb = RGBColor(150, 150, 150)
                header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

                for item in items:
                    original_text = item["text"]
                    para = docx_doc.add_paragraph()

                    if item["skip"]:
                        # Skipped text: keep original English, gray color
                        run = para.add_run(original_text)
                        run.font.color.rgb = RGBColor(100, 100, 100)
                        if item.get("bold"):
                            run.bold = True
                        if item.get("size", 0) >= 16:
                            run.font.size = Pt(min(int(item["size"] * 0.75), 20))
                    else:
                        # Translated text
                        translated = translation_map.get(original_text, original_text)
                        run = para.add_run(translated)
                        run.font.name = 'Malgun Gothic'
                        if item.get("bold"):
                            run.bold = True
                        if item.get("size", 0) >= 14:
                            run.font.size = Pt(min(int(item["size"] * 0.7), 18))

                pages_written += 1
                if progress_callback:
                    progress_callback(0.8 + (page_idx / max(len(all_pages_data), 1)) * 0.2)

            if progress_callback:
                progress_callback(1.0)

            docx_doc.save(output_stream)
            doc.close()
            return output_stream, translation_errors

        except Exception as e:
            print(f"Critical error in process_pdf: {e}")
            traceback.print_exc()
            raise e
