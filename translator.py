import deepl
import openai
from deep_translator import GoogleTranslator
import json
import os
import re
from dotenv import load_dotenv

load_dotenv()

class TranslationService:
    def __init__(self, service_type="DeepL", api_key=None):
        self.service_type = service_type
        self.api_key = api_key
        self.glossary = self._load_glossary()
        self.cache = {} # In-memory cache for the current session
        # Build do-not-translate set (entries where key == value, excluding _comment keys)
        self._do_not_translate = {
            k.lower(): k for k, v in self.glossary.items()
            if k == v and not k.startswith("_")
        }

    def _load_glossary(self):
        try:
            with open("glossary.json", "r", encoding="utf-8") as f:
                return json.load(f)
        except FileNotFoundError:
            return {}

    def _is_do_not_translate(self, text):
        """Check if text matches a do-not-translate glossary entry."""
        stripped = text.strip()
        if not stripped:
            return False
        # Exact match (case-insensitive)
        if stripped.lower() in self._do_not_translate:
            return True
        return False

    def _protect_glossary_terms(self, text):
        """Replace do-not-translate terms with placeholders before translation.

        This ensures terms like 'End-to-End' are preserved even when they
        appear inside longer sentences (e.g., 'Provides end-to-end visibility').
        """
        protected = {}
        result = text
        # Sort by length descending to match longer terms first
        # (e.g., "SAP Business Network" before "Business Network")
        for key in sorted(self._do_not_translate.values(), key=len, reverse=True):
            pattern = re.compile(re.escape(key), re.IGNORECASE)
            if pattern.search(result):
                placeholder = f"__DNT{len(protected)}__"
                result = pattern.sub(placeholder, result)
                protected[placeholder] = key
        return result, protected

    def _restore_glossary_terms(self, text, protected):
        """Restore placeholders back to original terms after translation."""
        result = text
        for placeholder, original in protected.items():
            result = result.replace(placeholder, original)
        return result

    def translate(self, text):
        if not text or not text.strip():
            return text

        # Check cache
        if text in self.cache:
            return self.cache[text]

        # Simple check if text is already Korean (contains Hangul)
        if any('\uac00' <= char <= '\ud7a3' for char in text):
            self.cache[text] = text
            return text

        # Check do-not-translate list before calling any API
        if self._is_do_not_translate(text):
            self.cache[text] = text
            return text

        # Skip pure numbers, dates, version strings
        if re.match(r'^[\d\.\-/\s:,]+$', text.strip()):
            self.cache[text] = text
            return text

        # Skip all-caps short acronyms (2-6 chars)
        if re.match(r'^[A-Z/]{2,6}$', text.strip()):
            self.cache[text] = text
            return text

        # Protect glossary terms with placeholders before translation
        protected_text, protected_map = self._protect_glossary_terms(text)

        result = protected_text
        if self.service_type == "DeepL":
            result = self._translate_deepl(protected_text)
        elif self.service_type == "OpenAI":
            result = self._translate_openai(protected_text)
        elif self.service_type == "Free (Google)":
            result = self._translate_free(protected_text)
        elif self.service_type == "Smart (OpenAI -> Free)":
            # Primary: OpenAI
            result = self._translate_openai(protected_text)
            # If OpenAI fails (returns original text or None/Error), try Free
            if result is None or result == protected_text:
                result = self._translate_free(protected_text)

        # FINAL GUARD: Ensure we NEVER return None
        if result is None:
            result = text

        # Restore glossary terms from placeholders
        if protected_map:
            result = self._restore_glossary_terms(result, protected_map)

        self.cache[text] = result
        return result

    def _translate_free(self, text):
        import time
        max_retries = 3
        for i in range(max_retries):
            try:
                # GoogleTranslator from deep-translator often works without a key
                result = GoogleTranslator(source='en', target='ko').translate(text)
                if result:
                    return result
                print(f"Free Translator Attempt {i+1} returned empty result.")
            except Exception as e:
                print(f"Free Translator Attempt {i+1} failed: {e}")

            if i < max_retries - 1:
                time.sleep(1)
        return text

    def _translate_deepl(self, text):
        if not self.api_key:
            return text # Just return original if no key
        try:
            translator = deepl.Translator(self.api_key)
            result = translator.translate_text(text, target_lang="KO")
            return result.text
        except Exception as e:
            print(f"DeepL Error: {e}")
            return text # Fallback to original text on error

    def _translate_openai(self, text):
        if not self.api_key:
            return text # Just return original if no key
        try:
            client = openai.OpenAI(api_key=self.api_key)

            # Build glossary strings
            translate_entries = []
            keep_entries = []
            for k, v in self.glossary.items():
                if k.startswith("_"):
                    continue
                if k == v:
                    keep_entries.append(k)
                else:
                    translate_entries.append(f"{k} -> {v}")

            glossary_str = ", ".join(translate_entries)
            keep_str = ", ".join(keep_entries[:30])  # Limit to avoid token overflow

            system_prompt = f"""You are a professional SAP consultant translating English to Korean for SAP business presentations.

RULES:
1. Translate descriptive body text naturally into professional Korean.
2. DO NOT translate the following - return them exactly as-is in English:
   - SAP product/solution names (SAP Ariba, S/4HANA, Business Network, SAP Fiori, etc.)
   - SAP module abbreviations (MM, SD, PP, FI, CO, etc.)
   - Common IT/business terms widely understood in Korean business context: {keep_str}
   - Proper nouns, brand names, and company names
   - Technical labels that are 3 words or fewer (e.g., "Purchase Requisition", "Goods Receipt")
   - Process flow labels (e.g., "Procure-to-Pay", "Order-to-Cash")
   - SAP Fiori app names (e.g., "Manage Purchase Orders", "Create Purchase Requisition", "My Inbox")
3. IMPORTANT: When these terms appear INSIDE a sentence, keep the term in English within the Korean translation.
   Example: "Provides end-to-end visibility" → "End-to-End 가시성을 제공합니다"
4. Use these glossary term mappings for translation: {glossary_str}
5. Keep the tone professional, suitable for executive-level SAP presentations.
6. IMPORTANT: Output ONLY the translated/preserved text, nothing else."""

            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": text}
                ]
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"OpenAI Error: {e}")
            return text # Fallback to original text on error
