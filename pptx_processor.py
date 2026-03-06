from pptx import Presentation
from pptx.util import Pt
import os
import traceback

class PPTXProcessor:
    def __init__(self, translator):
        self.translator = translator
        self.korean_font = "Malgun Gothic"

    def process_presentation(self, input_path, output_path, progress_callback=None):
        try:
            from concurrent.futures import ThreadPoolExecutor
            prs = Presentation(input_path)
            
            # Step 1: Collect all text targets
            text_frames = []
            def collect_frames(shapes, is_master=False):
                if not shapes: return
                for shape in shapes:
                    # For Masters/Layouts, only translate if it's NOT a placeholder
                    # and has actual content (placeholders often overlap or cause shifts)
                    if is_master and (shape.is_placeholder or not shape.has_text_frame):
                        continue
                        
                    if shape.has_text_frame:
                        text_frames.append(shape.text_frame)
                    if shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                text_frames.append(cell.text_frame)
                    if shape.shape_type == 6:  # Group shape
                        collect_frames(shape.shapes, is_master)
            
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
            
            # Step 2: Extract all unique paragraphs
            unique_texts = set()
            paragraphs_to_translate = []
            for tf in text_frames:
                for p in tf.paragraphs:
                    if p.text.strip() and len(p.text.strip()) > 1: # Skip single chars/bullets
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
