"""
translator.py
=============
Telugu -> English TRANSLITERATION & FAST BATCH TRANSLATION.

Features:
- Phonetic transliteration for names/addresses (e.g., "కొంపల్లి" -> "Kompally").
- High-speed batch translation for Excel files using deep-translator.
- Automatic clean-up of temporary files and direct output to User Downloads.
"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Optional

from utils import logger

try:
    from indic_transliteration import sanscript
    from indic_transliteration.sanscript import transliterate

    _TRANSLIT_AVAILABLE = True
except ImportError as exc:
    _TRANSLIT_AVAILABLE = False
    _TRANSLIT_IMPORT_ERROR = exc

try:
    import pandas as pd
    from deep_translator import GoogleTranslator

    _DEEP_TRANSLATION_AVAILABLE = True
except ImportError:
    pd = None
    GoogleTranslator = None
    _DEEP_TRANSLATION_AVAILABLE = False

# Unicode range for Telugu script
_TELUGU_RE = re.compile(r"[\u0C00-\u0C7F]")

# Diacritic-stripping map
_DIACRITIC_MAP = {
    "ā": "a",
    "ī": "i",
    "ū": "u",
    "ṛ": "ri",
    "ṝ": "ri",
    "ē": "e",
    "ō": "o",
    "ṃ": "m",
    "ḥ": "h",
    "ṅ": "n",
    "ñ": "n",
    "ṭ": "t",
    "ḍ": "d",
    "ṇ": "n",
    "ś": "sh",
    "ṣ": "sh",
    "ḷ": "ll",
    "ḻ": "l",
}


def contains_telugu(text: str) -> bool:
    return bool(_TELUGU_RE.search(text or ""))


def _strip_diacritics(text: str) -> str:
    for src, dst in _DIACRITIC_MAP.items():
        text = text.replace(src, dst).replace(src.upper(), dst.capitalize())
    return text


def transliterate_telugu(text: str) -> str:
    """Transliterate Telugu-script string to Roman-script approximation."""
    if not text or not contains_telugu(text):
        return text

    if not _TRANSLIT_AVAILABLE:
        logger.warning(
            "indic_transliteration not installed; returning original text. Error: %s",
            _TRANSLIT_IMPORT_ERROR,
        )
        return text

    try:
        romanized = transliterate(text, sanscript.TELUGU, sanscript.IAST)
        romanized = _strip_diacritics(romanized)
        words = romanized.split()
        romanized = " ".join(
            w.capitalize() if w.isalpha() else w for w in words
        )
        return romanized
    except Exception as exc:
        logger.error("Transliteration failed for %r: %s", text, exc)
        return text


def transliterate_address(address: str) -> str:
    """Transliterate address string word-by-word."""
    if not address:
        return address
    tokens = address.split()
    out_tokens = [
        transliterate_telugu(tok) if contains_telugu(tok) else tok
        for tok in tokens
    ]
    return " ".join(out_tokens)


def translate_excel_file(
    input_path: str, output_path: Optional[str] = None, batch_size: int = 50
) -> str:
    """Reads Excel, translates Telugu columns in high-speed batches,

    saves directly to Downloads folder, and cleans up the input file.
    """
    if not _DEEP_TRANSLATION_AVAILABLE or pd is None or GoogleTranslator is None:
        raise RuntimeError(
            "pandas and deep-translator are required for translation. "
            "Install via pip install pandas deep-translator openpyxl"
        )

    # 1. Prepare target save path (Default to Downloads folder if output_path is not explicitly provided)
    input_p = Path(input_path)

    if not output_path:
        downloads_folder = Path.home() / "Downloads"
        final_output = downloads_folder / f"{input_p.stem}_English.xlsx"
    else:
        final_output = Path(output_path)

    df = pd.read_excel(input_path, engine="openpyxl")
    translator = GoogleTranslator(source="te", target="en")

    # 2. Translate Telugu content in batches
    for col in df.select_dtypes(include=[object]).columns:
        mask = df[col].notna() & df[col].astype(str).str.contains(_TELUGU_RE)

        if mask.any():
            indices = df[mask].index.tolist()
            texts = df.loc[indices, col].astype(str).tolist()

            translated_list = []
            for i in range(0, len(texts), batch_size):
                batch = texts[i : i + batch_size]
                try:
                    translated_batch = translator.translate_batch(batch)
                    translated_list.extend(translated_batch)
                except Exception as exc:
                    logger.error(f"Batch translation error on col {col}: {exc}")
                    translated_list.extend(batch)  # Keep original if batch fails

            # Update translated values in DataFrame
            for idx, val in zip(indices, translated_list):
                df.at[idx, col] = val

    # 3. Save to final destination
    df.to_excel(final_output, index=False, engine="openpyxl")
    logger.info(f"Translated file successfully saved to: {final_output}")

    # 4. Clean up original Telugu file if different from final path
    if input_p.exists() and input_p.resolve() != final_output.resolve():
        try:
            os.remove(input_p)
            logger.info(f"Cleaned up temporary Telugu file: {input_p}")
        except Exception as e:
            logger.warning(
                f"Could not delete temporary file {input_p}: {e}"
            )

    return str(final_output)