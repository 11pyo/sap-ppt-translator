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

        # Fiori app IDs (e.g., "F1234", "F0842")
        if re.match(r'^F\d{4,5}$', stripped):
            return True

        # Check against do-not-translate glossary entries
        if hasattr(self.translator, '_do_not_translate'):
            if stripped.lower() in self.translator._do_not_translate:
                return True

        return False

    def _is_heading(self, block, page_height):
        """Determine if a text block is likely a heading/title."""
        if self.translation_level == "thorough":
            return False

        text = block["text"].strip()
        if not text:
            return False

        # Short text (≤ 40 chars) at the top 15% of the page
        y_position = block["bbox"][1]  # top y coordinate
        if y_position < page_height * 0.15 and len(text) <= 60:
            return True

        # Large font size (likely a title/heading)
        if block.get("size", 0) >= 18 and len(text) <= 80:
            return True

        # Short label-like text (≤ 4 words, no sentence punctuation)
        words = text.split()
        if (len(words) <= 4 and len(text) <= 40
                and not re.search(r'[.!?。,]', text)):
            return True

        return False

    def _is_footer(self, block, page_height):
        """Detect footer/watermark text."""
        y_position = block["bbox"][3]  # bottom y coordinate
        text = block["text"].strip()

        if y_position > page_height * 0.9 and len(text) <= 30:
            return True
        return False

    def process_pdf(self, input_stream, output_stream, progress_callback=None):
        """Extract text from PDF, translate, and generate DOCX output."""
        try:
            from concurrent.futures import ThreadPoolExecutor

            doc = fitz.open(stream=input_stream.read(), filetype="pdf")
            total_pages = len(doc)

            # Step 1: Extract all text blocks from all pages
            all_blocks = []  # list of (page_num, block_info_list)
            for page_num in range(total_pages):
                page = doc[page_num]
                page_height = page.rect.height

                # Extract text as dict for detailed info (position, font size, etc.)
                text_dict = page.get_text("dict")
                page_blocks = []

                for block in text_dict.get("blocks", []):
                    if block["type"] != 0:  # Skip image blocks
                        continue

                    # Combine all spans in all lines of this block
                    block_text = ""
                    max_size = 0
                    is_bold = False
                    for line in block.get("lines", []):
                        line_text = ""
                        for span in line.get("spans", []):
                            line_text += span["text"]
                            if span["size"] > max_size:
                                max_size = span["size"]
                            if "bold" in span.get("font", "").lower():
                                is_bold = True
                        if line_text.strip():
                            if block_text:
                                block_text += " "
                            block_text += line_text

                    if not block_text.strip():
                        continue

                    block_info = {
                        "text": block_text.strip(),
                        "bbox": block["bbox"],
                        "size": max_size,
                        "bold": is_bold,
                        "page": page_num
                    }

                    # Apply filtering
                    if self._is_footer(block_info, page_height):
                        block_info["skip"] = True
                    elif self._is_heading(block_info, page_height):
                        block_info["skip"] = True
                    elif self._should_skip_text(block_info["text"]):
                        block_info["skip"] = True
                    else:
                        block_info["skip"] = False

                    page_blocks.append(block_info)

                all_blocks.append((page_num, page_blocks))

            # Step 2: Collect unique texts to translate
            unique_texts = set()
            for page_num, blocks in all_blocks:
                for b in blocks:
                    if not b["skip"]:
                        unique_texts.add(b["text"])

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

                    for future in future_to_text:
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
                            progress_callback((processed_count / total_unique) * 0.8)

            # Step 4: Generate DOCX output
            docx_doc = Document()

            # Set default font
            style = docx_doc.styles['Normal']
            font = style.font
            font.name = 'Malgun Gothic'
            font.size = Pt(10)

            for page_idx, (page_num, blocks) in enumerate(all_blocks):
                # Add page separator (except for first page)
                if page_idx > 0:
                    docx_doc.add_page_break()

                # Add page header
                header_para = docx_doc.add_paragraph()
                header_run = header_para.add_run(f"── 페이지 {page_num + 1} ──")
                header_run.font.size = Pt(8)
                header_run.font.color.rgb = RGBColor(150, 150, 150)
                header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

                for b in blocks:
                    original_text = b["text"]

                    if b["skip"]:
                        # Skipped text: keep original English
                        para = docx_doc.add_paragraph()
                        run = para.add_run(original_text)
                        run.font.color.rgb = RGBColor(100, 100, 100)
                        if b.get("bold"):
                            run.bold = True
                        if b.get("size", 0) >= 16:
                            run.font.size = Pt(min(int(b["size"] * 0.8), 24))
                    else:
                        # Translated text
                        translated = translation_map.get(original_text, original_text)
                        para = docx_doc.add_paragraph()
                        run = para.add_run(translated)
                        run.font.name = 'Malgun Gothic'
                        if b.get("bold"):
                            run.bold = True

                if progress_callback:
                    progress_callback(0.8 + (page_idx / len(all_blocks)) * 0.2)

            docx_doc.save(output_stream)
            doc.close()
            return output_stream, translation_errors

        except Exception as e:
            print(f"Critical error in process_pdf: {e}")
            traceback.print_exc()
            raise e
