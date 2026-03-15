from pptx import Presentation
from pptx.util import Pt
import os
import re
import traceback

class PPTXProcessor:
    def __init__(self, translator, translation_level="normal"):
        self.translator = translator
        self.korean_font = "Malgun Gothic"
        self.translation_level = translation_level

    def _should_skip_shape(self, shape):
        """Determine if a shape's text should be excluded from translation."""
        level = self.translation_level

        # 1. Title/Subtitle placeholders - skip for normal and minimal
        if level in ("normal", "minimal") and shape.is_placeholder:
            try:
                idx = shape.placeholder_format.idx
                # 0=title, 1=center title, 13=subtitle
                if idx in (0, 1, 13):
                    return True
            except Exception:
                pass
            # Also check placeholder type enum if available
            try:
                from pptx.enum.shapes import PP_PLACEHOLDER
                ptype = shape.placeholder_format.type
                if ptype in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.SUBTITLE,
                             PP_PLACEHOLDER.CENTER_TITLE):
                    return True
            except Exception:
                pass

        # 2. Large font heuristic (diagram/chart display labels)
        if shape.has_text_frame:
            if level == "minimal":
                font_threshold = Pt(20)
            elif level == "normal":
                font_threshold = Pt(24)
            else:  # thorough
                font_threshold = Pt(32)

            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.size and run.font.size >= font_threshold:
                        return True

        # 3. Short label detection (very short non-sentence text)
        if shape.has_text_frame:
            text = shape.text_frame.text.strip()
            max_len = 8 if level == "minimal" else 5
            if 1 < len(text) <= max_len and ' ' not in text:
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

        # Check against do-not-translate glossary entries
        if hasattr(self.translator, '_do_not_translate'):
            if stripped.lower() in self.translator._do_not_translate:
                return True

        return False

    def process_presentation(self, input_path, output_path, progress_callback=None):
        try:
            from concurrent.futures import ThreadPoolExecutor
            prs = Presentation(input_path)

            # Step 1: Collect all text targets
            text_frames = []
            def collect_frames(shapes, is_master=False, nesting_depth=0):
                if not shapes: return
                for shape in shapes:
                    # For Masters/Layouts, only translate if it's NOT a placeholder
                    # and has actual content (placeholders often overlap or cause shifts)
                    if is_master and (shape.is_placeholder or not shape.has_text_frame):
                        continue

                    # Smart shape-level filtering (not applied to master slides)
                    if not is_master and self._should_skip_shape(shape):
                        continue

                    if shape.has_text_frame:
                        # For deeply nested group shapes (depth >= 2),
                        # skip short text as it's likely a diagram node label
                        if nesting_depth >= 2:
                            text = shape.text_frame.text.strip()
                            if len(text) < 30:
                                continue
                        text_frames.append(shape.text_frame)
                    if shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                text_frames.append(cell.text_frame)
                    if shape.shape_type == 6:  # Group shape
                        collect_frames(shape.shapes, is_master, nesting_depth + 1)

            # 1.1: Collect from slides (Higher priority)
            for slide in prs.slides:
                collect_frames(slide.shapes, is_master=False)
                if slide.has_notes_slide:
                    collect_frames(slide.notes_slide.shapes, is_master=False)

            # 1.2: Collect from layouts/masters (Only if user wanted 'thorough' but stay safe)
            for master in prs.slide_masters:
                collect_frames(master.shapes, is_master=True)
                for layout in master.slide_layouts:
                    collect_frames(layout.shapes, is_master=True)

            # Step 2: Extract all unique paragraphs (with text-level filtering)
            unique_texts = set()
            paragraphs_to_translate = []
            for tf in text_frames:
                for p in tf.paragraphs:
                    if p.text.strip() and len(p.text.strip()) > 1: # Skip single chars/bullets
                        if not self._should_skip_text(p.text):
                            unique_texts.add(p.text)
                            paragraphs_to_translate.append(p)

            # Step 3: Translate unique texts in parallel
            total_unique = len(unique_texts)
            translation_map = {}
            translation_errors = []
            processed_count = 0

            # Only translate if there's actually something to translate
            if unique_texts:
                max_workers = 15
                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    future_to_text = {executor.submit(self.translator.translate, text): text for text in unique_texts}

                    for i, future in enumerate(future_to_text):
                        text = future_to_text[future]
                        try:
                            result = future.result()
                            if result:
                                translation_map[text] = result
                            else:
                                translation_map[text] = text
                        except Exception as e:
                            err_msg = f"'{text[:30]}...' 번역 실패 ({type(e).__name__}): {str(e)}"
                            translation_errors.append(err_msg)
                            translation_map[text] = text

                        processed_count += 1
                        if progress_callback:
                            progress_callback((processed_count / total_unique) * 0.8)

            # Step 4: Apply translations
            for i, p in enumerate(paragraphs_to_translate):
                if p.text in translation_map:
                    translated_text = translation_map[p.text]
                    if translated_text and translated_text != p.text:
                        self._update_paragraph_text(p, translated_text)

                if progress_callback:
                    progress_callback(0.8 + (i / len(paragraphs_to_translate)) * 0.2)

            prs.save(output_path)
            return output_path, translation_errors
        except Exception as e:
            print(f"Critical error in process_presentation: {e}")
            traceback.print_exc()
            raise e

    def _update_paragraph_text(self, paragraph, translated_text):
        """Update paragraph text while preserving ALL original formatting.

        Key insight: Instead of paragraph.text = ... (which destroys all runs/XML),
        we modify individual run.text in-place so all XML attributes are preserved.
        """
        if translated_text is None:
            return

        try:
            runs = paragraph.runs

            if not runs:
                # No runs exist — fallback to direct assignment (rare edge case)
                paragraph.text = translated_text
                return

            if len(runs) == 1:
                # Single run: simply replace its text (keeps ALL formatting)
                runs[0].text = translated_text
            else:
                # Multiple runs: put all translated text in the first run,
                # clear the remaining runs (preserves first run's formatting)
                runs[0].text = translated_text
                for run in runs[1:]:
                    run.text = ""

            # Only set Korean font name — do NOT touch size, color, bold, etc.
            # They are already preserved from the original run XML
            for run in runs:
                if run.text:  # Only set font on runs that have text
                    run.font.name = self.korean_font

        except Exception as e:
            print(f"Error updating paragraph text: {e}")
            traceback.print_exc()
