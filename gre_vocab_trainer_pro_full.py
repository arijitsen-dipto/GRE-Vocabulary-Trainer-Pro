"""
GRE Vocabulary Trainer Pro
==========================

Designed for: GRE_VOCAB_DATABASE_FINAL_CLEAN.xlsx + GRE_USER_DATA.xlsx

Main features
-------------
1. Stable word IDs based on the normalized word, so progress is not broken if
   workbook sheet order changes.
2. Five clear study modes:
   - Flashcard
   - Reverse Flashcard
   - Typing (random Word <-> Meaning direction)
   - Multiple Choice (7 options)
   - Mixed Mode
3. Multiple Choice and Typing use only word/meaning directions. They do
   NOT use synonym, antonym, or example prompts.
4. Right-side panel shows mnemonic, favorite/difficult state, source/group info,
   progress, and user notes. User mnemonic is hidden from normal notes and is
   visible only through the mnemonic buttons: Default, Mine, Both, Hide.
5. Plan Study with Daily / Weekly / Monthly plans: select source sheet, group
   range, mode, order, target words/reviews, include due words, and optional
   group lock.
6. Clean logs and progress dashboard in GRE_USER_DATA.xlsx:
   - user_notes
   - word_flags
   - progress
   - app_settings
   - study_plan
   - session_log
   - review_log
   - daily_stats

Place this file in the same folder as:
- GRE_VOCAB_DATABASE_FINAL_CLEAN.xlsx
- GRE_USER_DATA.xlsx

Run:
    python gre_vocab_trainer_pro_full.py
"""

from __future__ import annotations

import datetime as _dt
import hashlib
import os
import random
import re
import time
import uuid
from dataclasses import dataclass, field
from typing import Any, Optional

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog


# ---------------------------------------------------------------------------
# File names
# ---------------------------------------------------------------------------

VOCAB_FILE = "GRE_VOCAB_DATABASE_FINAL_CLEAN.xlsx"
USER_FILE = "GRE_USER_DATA.xlsx"


# ---------------------------------------------------------------------------
# Utility functions
# ---------------------------------------------------------------------------

def clean_text(value: Any) -> str:
    """Return a safe clean string; avoid Excel NaN becoming the text 'nan'."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat"}:
        return ""
    return text


def normalize_word(word: str) -> str:
    """Normalize a word for duplicate merging and stable IDs."""
    word = clean_text(word).lower()
    word = re.sub(r"\s+", " ", word)
    return word.strip()


def stable_word_id(normalized_word: str) -> str:
    """Stable ID that will not change if sheet order changes."""
    digest = hashlib.sha1(normalized_word.encode("utf-8")).hexdigest()[:12]
    return f"W_{digest}"


def now_iso() -> str:
    return _dt.datetime.now().replace(microsecond=0).isoformat(sep=" ")


def today_iso() -> str:
    return _dt.date.today().isoformat()


def parse_group_number(group: Any) -> int:
    text = clean_text(group)
    m = re.search(r"\d+", text)
    if m:
        return int(m.group())
    return 10**9


def natural_sort_key(value: Any) -> tuple:
    text = clean_text(value)
    if text.lower() == "all":
        return (-1, "")
    n = parse_group_number(text)
    if n != 10**9:
        return (0, n)
    return (1, text.lower())


def tokenize(text: str) -> set[str]:
    """Tokenize text for rough answer checking."""
    text = clean_text(text).lower()
    return set(re.findall(r"[a-z]{3,}", text))


def compact_meaning(meaning: str, max_len: int = 180) -> str:
    meaning = clean_text(meaning)
    if len(meaning) <= max_len:
        return meaning
    cut = meaning[:max_len].rsplit(" ", 1)[0]
    return cut + "..."


def ensure_unique_list(items: list[str]) -> list[str]:
    seen = set()
    out: list[str] = []
    for item in items:
        key = clean_text(item).lower()
        if key and key not in seen:
            out.append(clean_text(item))
            seen.add(key)
    return out


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class WordEntry:
    word_id: str
    display_word: str
    normalized_word: str
    primary: dict[str, str] = field(default_factory=dict)
    sources: list[tuple[str, str]] = field(default_factory=list)

    def __post_init__(self) -> None:
        defaults = {
            "meaning": "",
            "part_of_speech": "",
            "group": "",
            "example": "",
            "example2": "",
            "synonym": "",
            "antonym": "",
            "root": "",
            "mnemonic": "",
            "tone": "",
            "frequency": "",
            "source_note": "",
        }
        defaults.update(self.primary or {})
        self.primary = defaults

    def update_from_row(self, row: dict[str, Any], sheet_name: str) -> None:
        group = clean_text(row.get("group", ""))
        self.sources.append((sheet_name, group))
        for field_name in self.primary:
            value = clean_text(row.get(field_name, ""))
            if value and not self.primary.get(field_name):
                self.primary[field_name] = value

    def source_names(self) -> str:
        return ", ".join(ensure_unique_list([s for s, _ in self.sources]))

    def groups_for_source(self, source: str) -> list[str]:
        return ensure_unique_list([g for s, g in self.sources if s == source and g])


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class DataModel:
    expected_vocab_columns = [
        "word", "meaning", "part_of_speech", "group", "example", "example2",
        "synonym", "antonym", "root", "mnemonic", "tone", "frequency",
        "source_note",
    ]

    user_sheets = {
        "user_notes": [
            "word_id", "normalized_word", "display_word", "user_note",
            "user_meaning", "user_mnemonic", "user_translation", "user_tag",
            "updated_at",
        ],
        "word_flags": [
            "word_id", "favorite", "difficult", "updated_at",
        ],
        "progress": [
            "word_id", "normalized_word", "display_word", "status",
            "review_count", "correct_count", "wrong_count", "ease_factor",
            "interval_days", "last_review", "next_review", "streak",
            "leech_score", "last_rating",
        ],
        "app_settings": ["setting", "value", "updated_at"],
        "study_plan": [
            "plan_name", "plan_period", "enabled", "source", "start_group", "end_group",
            "new_words_per_day", "weekly_new_words", "monthly_new_words",
            "max_reviews_per_day", "weekly_review_target", "monthly_review_target",
            "mode", "order", "include_due", "group_lock", "start_date", "end_date",
            "created_at", "updated_at",
        ],
        "session_log": [
            "session_id", "start_time", "end_time", "source", "group",
            "mode", "order", "deck_size", "answered", "correct", "wrong",
            "accuracy", "new_words", "due_words", "duration_sec",
        ],
        "review_log": [
            "timestamp", "session_id", "word_id", "normalized_word",
            "display_word", "source", "group", "mode", "prompt_type",
            "prompt", "correct_answer", "user_answer", "is_correct", "rating",
            "response_time_sec", "old_status", "new_status", "old_interval",
            "new_interval", "next_review",
        ],
        "daily_stats": [
            "date", "studied", "correct", "wrong", "accuracy", "new_words",
            "reviews", "mastered_today", "time_spent_sec",
        ],
    }

    default_settings = {
        "mnemonic_view": "Both",
        "auto_save": "True",
        "mcq_options": "7",
    }

    def __init__(self, vocab_file: str = VOCAB_FILE, user_file: str = USER_FILE) -> None:
        self.vocab_file = vocab_file
        self.user_file = user_file
        self.sources: dict[str, pd.DataFrame] = {}
        self.words: dict[str, WordEntry] = {}
        self.word_id_map: dict[str, WordEntry] = {}
        self.user_notes: dict[str, dict[str, str]] = {}
        self.word_flags: dict[str, dict[str, bool]] = {}
        self.progress: dict[str, dict[str, Any]] = {}
        self.app_settings: dict[str, str] = dict(self.default_settings)
        self.study_plan: dict[str, Any] = {}
        self.session_log_rows: list[dict[str, Any]] = []
        self.review_log_rows: list[dict[str, Any]] = []
        self.daily_stats_rows: list[dict[str, Any]] = []

    # ------------------------------------------------------------------
    # Vocabulary loading
    # ------------------------------------------------------------------
    def load_vocabulary(self) -> None:
        if not os.path.exists(self.vocab_file):
            raise FileNotFoundError(
                f"Vocabulary file '{self.vocab_file}' not found. Place it in the same folder as this script."
            )
        xls = pd.ExcelFile(self.vocab_file)
        self.sources.clear()
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, dtype=str).fillna("")
            df.columns = [clean_text(c).strip() for c in df.columns]
            for col in self.expected_vocab_columns:
                if col not in df.columns:
                    df[col] = ""
            df = df[self.expected_vocab_columns].copy()
            for col in df.columns:
                df[col] = df[col].map(clean_text)
            df = df[df["word"].str.strip() != ""].reset_index(drop=True)
            self.sources[sheet] = df

    def build_global_registry(self) -> None:
        self.words.clear()
        self.word_id_map.clear()
        for sheet_name, df in self.sources.items():
            for _, row in df.iterrows():
                word = clean_text(row.get("word", ""))
                normalized = normalize_word(word)
                if not normalized:
                    continue
                if normalized not in self.words:
                    wid = stable_word_id(normalized)
                    entry = WordEntry(wid, word, normalized)
                    self.words[normalized] = entry
                    self.word_id_map[wid] = entry
                self.words[normalized].update_from_row(row.to_dict(), sheet_name)

    # ------------------------------------------------------------------
    # User data file loading/saving/upgrading
    # ------------------------------------------------------------------
    def create_empty_user_data(self) -> None:
        with pd.ExcelWriter(self.user_file, engine="openpyxl") as writer:
            for sheet_name, columns in self.user_sheets.items():
                rows: list[dict[str, Any]] = []
                if sheet_name == "app_settings":
                    rows = [
                        {"setting": k, "value": v, "updated_at": now_iso()}
                        for k, v in self.default_settings.items()
                    ]
                elif sheet_name == "study_plan":
                    rows = [{
                        "plan_name": "Default",
                        "plan_period": "Daily",
                        "enabled": "False",
                        "source": "Ultimate Total List",
                        "start_group": "1",
                        "end_group": "1",
                        "new_words_per_day": "30",
                        "weekly_new_words": "210",
                        "monthly_new_words": "900",
                        "max_reviews_per_day": "200",
                        "weekly_review_target": "1000",
                        "monthly_review_target": "4000",
                        "mode": "Mixed Mode",
                        "order": "Due First",
                        "include_due": "True",
                        "group_lock": "True",
                        "start_date": today_iso(),
                        "end_date": "",
                        "created_at": now_iso(),
                        "updated_at": now_iso(),
                    }]
                pd.DataFrame(rows, columns=columns).to_excel(writer, sheet_name=sheet_name, index=False)

    def upgrade_user_data_file(self) -> None:
        if not os.path.exists(self.user_file):
            self.create_empty_user_data()
            return
        try:
            existing = pd.read_excel(self.user_file, sheet_name=None, dtype=str)
        except Exception:
            self.create_empty_user_data()
            return
        with pd.ExcelWriter(self.user_file, engine="openpyxl") as writer:
            for sheet_name, columns in self.user_sheets.items():
                df = existing.get(sheet_name, pd.DataFrame())
                if df.empty:
                    df = pd.DataFrame(columns=columns)
                df = df.fillna("")
                for col in columns:
                    if col not in df.columns:
                        df[col] = ""
                df = df[columns]
                if sheet_name == "app_settings" and df.empty:
                    df = pd.DataFrame([
                        {"setting": k, "value": v, "updated_at": now_iso()}
                        for k, v in self.default_settings.items()
                    ], columns=columns)
                if sheet_name == "study_plan" and df.empty:
                    df = pd.DataFrame([{c: "" for c in columns}], columns=columns)
                    df.loc[0, "plan_name"] = "Default"
                    df.loc[0, "plan_period"] = "Daily"
                    df.loc[0, "enabled"] = "False"
                    df.loc[0, "source"] = "Ultimate Total List"
                    df.loc[0, "start_group"] = "1"
                    df.loc[0, "end_group"] = "1"
                    df.loc[0, "new_words_per_day"] = "30"
                    df.loc[0, "weekly_new_words"] = "210"
                    df.loc[0, "monthly_new_words"] = "900"
                    df.loc[0, "max_reviews_per_day"] = "200"
                    df.loc[0, "weekly_review_target"] = "1000"
                    df.loc[0, "monthly_review_target"] = "4000"
                    df.loc[0, "mode"] = "Mixed Mode"
                    df.loc[0, "order"] = "Due First"
                    df.loc[0, "include_due"] = "True"
                    df.loc[0, "group_lock"] = "True"
                    df.loc[0, "start_date"] = today_iso()
                    df.loc[0, "end_date"] = ""
                    df.loc[0, "created_at"] = now_iso()
                    df.loc[0, "updated_at"] = now_iso()
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    def load_user_data(self) -> None:
        self.upgrade_user_data_file()
        xls = pd.read_excel(self.user_file, sheet_name=None, dtype=str)
        self.user_notes.clear()
        self.word_flags.clear()
        self.progress.clear()
        self.app_settings = dict(self.default_settings)
        self.study_plan = {}
        self.session_log_rows = []
        self.review_log_rows = []
        self.daily_stats_rows = []

        # User notes
        notes_df = xls.get("user_notes", pd.DataFrame()).fillna("")
        for _, row in notes_df.iterrows():
            wid = clean_text(row.get("word_id", ""))
            if not wid:
                continue
            self.user_notes[wid] = {
                "user_note": clean_text(row.get("user_note", "")),
                "user_meaning": clean_text(row.get("user_meaning", "")),
                "user_mnemonic": clean_text(row.get("user_mnemonic", "")),
                "user_translation": clean_text(row.get("user_translation", "")),
                "user_tag": clean_text(row.get("user_tag", "")),
                "updated_at": clean_text(row.get("updated_at", "")),
            }

        # Flags
        flags_df = xls.get("word_flags", pd.DataFrame()).fillna("")
        for _, row in flags_df.iterrows():
            wid = clean_text(row.get("word_id", ""))
            if not wid:
                continue
            self.word_flags[wid] = {
                "favorite": clean_text(row.get("favorite", "False")).lower() == "true",
                "difficult": clean_text(row.get("difficult", "False")).lower() == "true",
            }

        # Progress
        prog_df = xls.get("progress", pd.DataFrame()).fillna("")
        for _, row in prog_df.iterrows():
            wid = clean_text(row.get("word_id", ""))
            if not wid:
                continue
            self.progress[wid] = {
                "status": clean_text(row.get("status", "unseen")) or "unseen",
                "review_count": self._to_int(row.get("review_count", "0")),
                "correct_count": self._to_int(row.get("correct_count", "0")),
                "wrong_count": self._to_int(row.get("wrong_count", "0")),
                "ease_factor": self._to_float(row.get("ease_factor", "2.5"), 2.5),
                "interval_days": self._to_int(row.get("interval_days", "0")),
                "last_review": clean_text(row.get("last_review", "")),
                "next_review": clean_text(row.get("next_review", "")),
                "streak": self._to_int(row.get("streak", "0")),
                "leech_score": self._to_int(row.get("leech_score", "0")),
                "last_rating": clean_text(row.get("last_rating", "")),
            }

        # Settings
        settings_df = xls.get("app_settings", pd.DataFrame()).fillna("")
        for _, row in settings_df.iterrows():
            key = clean_text(row.get("setting", ""))
            value = clean_text(row.get("value", ""))
            if key:
                self.app_settings[key] = value

        # Study plan: first row is active/default
        plan_df = xls.get("study_plan", pd.DataFrame()).fillna("")
        if not plan_df.empty:
            row = plan_df.iloc[0]
            self.study_plan = {col: clean_text(row.get(col, "")) for col in plan_df.columns}

        # Logs loaded for append/save
        for sheet_name, attr in [
            ("session_log", "session_log_rows"),
            ("review_log", "review_log_rows"),
            ("daily_stats", "daily_stats_rows"),
        ]:
            df = xls.get(sheet_name, pd.DataFrame()).fillna("")
            setattr(self, attr, df.to_dict(orient="records"))

    @staticmethod
    def _to_int(value: Any, default: int = 0) -> int:
        try:
            return int(float(clean_text(value)))
        except Exception:
            return default

    @staticmethod
    def _to_float(value: Any, default: float = 0.0) -> float:
        try:
            return float(clean_text(value))
        except Exception:
            return default

    def save_user_data(self) -> None:
        """Safely save user data to Excel.

        This version writes to a temporary workbook first, verifies that it
        contains visible sheets, then replaces the real GRE_USER_DATA.xlsx.
        This prevents the common openpyxl error: "At least one sheet must be
        visible" and reduces the chance of corrupting the user file if a save
        fails halfway.
        """
        # Prepare notes
        notes_rows = []
        for wid, data in self.user_notes.items():
            entry = self.word_id_map.get(wid)
            notes_rows.append({
                "word_id": wid,
                "normalized_word": entry.normalized_word if entry else "",
                "display_word": entry.display_word if entry else "",
                "user_note": data.get("user_note", ""),
                "user_meaning": data.get("user_meaning", ""),
                "user_mnemonic": data.get("user_mnemonic", ""),
                "user_translation": data.get("user_translation", ""),
                "user_tag": data.get("user_tag", ""),
                "updated_at": data.get("updated_at", ""),
            })

        flags_rows = []
        for wid, flags in self.word_flags.items():
            flags_rows.append({
                "word_id": wid,
                "favorite": str(bool(flags.get("favorite", False))),
                "difficult": str(bool(flags.get("difficult", False))),
                "updated_at": now_iso(),
            })

        progress_rows = []
        for wid, prog in self.progress.items():
            entry = self.word_id_map.get(wid)
            progress_rows.append({
                "word_id": wid,
                "normalized_word": entry.normalized_word if entry else "",
                "display_word": entry.display_word if entry else "",
                "status": prog.get("status", "unseen"),
                "review_count": prog.get("review_count", 0),
                "correct_count": prog.get("correct_count", 0),
                "wrong_count": prog.get("wrong_count", 0),
                "ease_factor": prog.get("ease_factor", 2.5),
                "interval_days": prog.get("interval_days", 0),
                "last_review": prog.get("last_review", ""),
                "next_review": prog.get("next_review", ""),
                "streak": prog.get("streak", 0),
                "leech_score": prog.get("leech_score", 0),
                "last_rating": prog.get("last_rating", ""),
            })

        settings_rows = [
            {"setting": k, "value": v, "updated_at": now_iso()}
            for k, v in sorted(self.app_settings.items())
        ]
        # Ensure there is always at least one non-empty sheet.
        if not settings_rows:
            settings_rows = [
                {"setting": k, "value": v, "updated_at": now_iso()}
                for k, v in self.default_settings.items()
            ]

        if not self.study_plan:
            self.study_plan = {
                "plan_name": "Default", "plan_period": "Daily", "enabled": "False", "source": "Ultimate Total List",
                "start_group": "1", "end_group": "1", "new_words_per_day": "30", "weekly_new_words": "210",
                "monthly_new_words": "900", "max_reviews_per_day": "200", "weekly_review_target": "1000",
                "monthly_review_target": "4000", "mode": "Mixed Mode", "order": "Due First",
                "include_due": "True", "group_lock": "True", "start_date": today_iso(), "end_date": "",
                "created_at": now_iso(), "updated_at": now_iso(),
            }
        else:
            self.study_plan["updated_at"] = self.study_plan.get("updated_at", now_iso()) or now_iso()
        plan_rows = [self.study_plan]

        sheet_data = {
            "user_notes": (notes_rows, self.user_sheets["user_notes"]),
            "word_flags": (flags_rows, self.user_sheets["word_flags"]),
            "progress": (progress_rows, self.user_sheets["progress"]),
            "app_settings": (settings_rows, self.user_sheets["app_settings"]),
            "study_plan": (plan_rows, self.user_sheets["study_plan"]),
            "session_log": (self.session_log_rows, self.user_sheets["session_log"]),
            "review_log": (self.review_log_rows, self.user_sheets["review_log"]),
            "daily_stats": (self.daily_stats_rows, self.user_sheets["daily_stats"]),
        }

        tmp_file = f"{self.user_file}.tmp.xlsx"
        backup_file = f"{self.user_file}.bak.xlsx"

        # Remove a leftover temp file from a previous failed save.
        try:
            if os.path.exists(tmp_file):
                os.remove(tmp_file)
        except Exception:
            pass

        try:
            with pd.ExcelWriter(tmp_file, engine="openpyxl") as writer:
                for sheet_name, (rows, columns) in sheet_data.items():
                    df = pd.DataFrame(rows or [], columns=columns)
                    for col in columns:
                        if col not in df.columns:
                            df[col] = ""
                    df = df[columns].fillna("")
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Explicitly make sure at least one worksheet is visible.
                visible_sheets = [ws for ws in writer.book.worksheets if ws.sheet_state == "visible"]
                if not visible_sheets and writer.book.worksheets:
                    writer.book.worksheets[0].sheet_state = "visible"

            # Keep one backup of the last working user file.
            try:
                if os.path.exists(self.user_file):
                    import shutil
                    shutil.copy2(self.user_file, backup_file)
            except Exception:
                pass

            # Atomic replace when possible. If the file is open in Excel, this
            # will fail with a clear permission error and the old file remains.
            os.replace(tmp_file, self.user_file)
        finally:
            try:
                if os.path.exists(tmp_file):
                    os.remove(tmp_file)
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Access helpers
    # ------------------------------------------------------------------
    def get_progress(self, wid: str) -> dict[str, Any]:
        if wid not in self.progress:
            self.progress[wid] = {
                "status": "unseen",
                "review_count": 0,
                "correct_count": 0,
                "wrong_count": 0,
                "ease_factor": 2.5,
                "interval_days": 0,
                "last_review": "",
                "next_review": "",
                "streak": 0,
                "leech_score": 0,
                "last_rating": "",
            }
        return self.progress[wid]

    def get_flags(self, wid: str) -> dict[str, bool]:
        if wid not in self.word_flags:
            self.word_flags[wid] = {
                "favorite": False,
                "difficult": False,
            }
        return self.word_flags[wid]

    def get_user_note(self, wid: str) -> dict[str, str]:
        if wid not in self.user_notes:
            self.user_notes[wid] = {
                "user_note": "",
                "user_meaning": "",
                "user_mnemonic": "",
                "user_translation": "",
                "user_tag": "",
                "updated_at": "",
            }
        return self.user_notes[wid]

    def due_words(self, entries: list[WordEntry], limit: Optional[int] = None) -> list[WordEntry]:
        today = today_iso()
        due = []
        for e in entries:
            prog = self.get_progress(e.word_id)
            status = prog.get("status", "unseen")
            next_review = clean_text(prog.get("next_review", ""))
            if status != "unseen" and next_review and next_review <= today:
                due.append(e)
        if limit:
            return due[:limit]
        return due

    def unseen_words(self, entries: list[WordEntry], limit: Optional[int] = None) -> list[WordEntry]:
        new = []
        for e in entries:
            prog = self.get_progress(e.word_id)
            if prog.get("status", "unseen") == "unseen":
                new.append(e)
        if limit:
            return new[:limit]
        return new


# ---------------------------------------------------------------------------
# Trainer Application
# ---------------------------------------------------------------------------

class GRETrainerApp(tk.Tk):
    study_modes = [
        "Flashcard",
        "Reverse Flashcard",
        "Typing",
        "Multiple Choice (7 options)",
        "Mixed Mode",
    ]
    orders = ["Shuffle", "Listwise", "Due First", "Weak First", "New First"]
    filters = [
        "All", "Unseen", "Learning", "Reviewing", "Mastered",
        "Due", "Difficult", "Favorite",
    ]

    def __init__(self, model: DataModel) -> None:
        super().__init__()
        self.model = model
        self.title("GRE Vocabulary Trainer Pro")
        self.geometry("1360x800")
        self.minsize(1200, 700)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.setup_menu()

        self.current_deck: list[WordEntry] = []
        self.current_index: int = 0
        self.current_mode: str = "Flashcard"
        self.current_prompt_type: str = ""
        self.current_options: list[str] = []
        self.current_correct_answer: str = ""
        self.current_prompt_text: str = ""
        self.showing_answer: bool = False
        self.card_start_time: float = time.time()
        self.session_id: str = ""
        self.session_start_time: float = 0.0
        self.session_context: dict[str, Any] = {}
        self.session_answered = 0
        self.session_correct = 0
        self.session_wrong = 0
        self.session_new = 0
        self.session_due = 0
        self.mnemonic_view = self.model.app_settings.get("mnemonic_view", "Both") or "Both"

        self.setup_layout()
        self.bind_shortcuts()
        self.refresh_group_options()

    # ------------------------------------------------------------------
    # UI setup
    # ------------------------------------------------------------------
    def setup_menu(self) -> None:
        """Top menu for important actions that should always be visible."""
        menubar = tk.Menu(self)
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="Find Word", command=self.find_word_dialog)
        tools_menu.add_command(label="Save Now", command=self.save_now)
        tools_menu.add_separator()
        tools_menu.add_command(label="Reset Current Word", command=self.reset_current_word)
        tools_menu.add_command(label="Reset All Progress", command=self.reset_all_progress)
        tools_menu.add_separator()
        tools_menu.add_command(label="Progress Dashboard", command=self.show_progress_dashboard)
        tools_menu.add_command(label="Show Statistics", command=self.show_statistics)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        self.config(menu=menubar)

    def setup_layout(self) -> None:
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        self.left_frame = tk.Frame(self, bd=1, relief="solid", padx=10, pady=10)
        self.left_frame.grid(row=0, column=0, sticky="ns")

        self.main_frame = tk.Frame(self, padx=10, pady=10)
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.rowconfigure(2, weight=1)
        self.main_frame.columnconfigure(0, weight=3)
        self.main_frame.columnconfigure(1, weight=1)

        self.setup_left_panel()
        self.setup_card_area()
        self.setup_mnemonic_panel()
        self.setup_controls()

    def setup_left_panel(self) -> None:
        tk.Label(self.left_frame, text="GRE Trainer Pro", font=("Arial", 16, "bold")).pack(pady=(0, 10))

        tk.Label(self.left_frame, text="Dataset / Source", anchor="w").pack(fill="x")
        self.source_var = tk.StringVar(value="Ultimate Total List")
        source_values = ["Ultimate Total List"] + list(self.model.sources.keys())
        self.source_combo = ttk.Combobox(self.left_frame, textvariable=self.source_var, values=source_values, state="readonly", width=30)
        self.source_combo.pack(fill="x", pady=3)
        self.source_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_group_options())

        tk.Label(self.left_frame, text="Word Finder", anchor="w").pack(fill="x", pady=(6, 0))
        finder_row = tk.Frame(self.left_frame)
        finder_row.pack(fill="x", pady=3)
        self.find_word_var = tk.StringVar()
        self.find_word_entry = tk.Entry(finder_row, textvariable=self.find_word_var, width=22)
        self.find_word_entry.pack(side="left", fill="x", expand=True)
        self.find_word_entry.bind("<Return>", lambda e: self.find_word_dialog())
        tk.Button(finder_row, text="Find", command=self.find_word_dialog, width=6).pack(side="left", padx=(4, 0))

        tk.Label(self.left_frame, text="Group", anchor="w").pack(fill="x")
        self.group_var = tk.StringVar(value="All")
        self.group_combo = ttk.Combobox(self.left_frame, textvariable=self.group_var, values=["All"], state="readonly", width=30)
        self.group_combo.pack(fill="x", pady=3)

        tk.Label(self.left_frame, text="Mode", anchor="w").pack(fill="x")
        self.mode_var = tk.StringVar(value="Flashcard")
        self.mode_combo = ttk.Combobox(self.left_frame, textvariable=self.mode_var, values=self.study_modes, state="readonly", width=30)
        self.mode_combo.pack(fill="x", pady=3)

        tk.Label(self.left_frame, text="Filter", anchor="w").pack(fill="x")
        self.filter_var = tk.StringVar(value="All")
        self.filter_combo = ttk.Combobox(self.left_frame, textvariable=self.filter_var, values=self.filters, state="readonly", width=30)
        self.filter_combo.pack(fill="x", pady=3)

        tk.Label(self.left_frame, text="Root", anchor="w").pack(fill="x")
        self.root_var = tk.StringVar(value="All")
        self.root_combo = ttk.Combobox(self.left_frame, textvariable=self.root_var, values=["All"], state="readonly", width=30)
        self.root_combo.pack(fill="x", pady=3)

        tk.Label(self.left_frame, text="Tone", anchor="w").pack(fill="x")
        self.tone_var = tk.StringVar(value="All")
        self.tone_combo = ttk.Combobox(self.left_frame, textvariable=self.tone_var, values=["All"], state="readonly", width=30)
        self.tone_combo.pack(fill="x", pady=3)

        tk.Label(self.left_frame, text="Frequency", anchor="w").pack(fill="x")
        self.frequency_var = tk.StringVar(value="All")
        self.frequency_combo = ttk.Combobox(self.left_frame, textvariable=self.frequency_var, values=["All"], state="readonly", width=30)
        self.frequency_combo.pack(fill="x", pady=3)

        tk.Label(self.left_frame, text="Number of Cards", anchor="w").pack(fill="x")
        self.count_var = tk.StringVar(value="30")
        self.count_combo = ttk.Combobox(self.left_frame, textvariable=self.count_var, values=["All", "10", "20", "30", "50", "100", "200"], state="readonly", width=30)
        self.count_combo.pack(fill="x", pady=3)

        tk.Label(self.left_frame, text="Order", anchor="w").pack(fill="x")
        self.order_var = tk.StringVar(value="Shuffle")
        self.order_combo = ttk.Combobox(self.left_frame, textvariable=self.order_var, values=self.orders, state="readonly", width=30)
        self.order_combo.pack(fill="x", pady=3)

        tk.Label(self.left_frame, text="MCQ Options", anchor="w").pack(fill="x")
        self.mcq_options_var = tk.StringVar(value="7")
        self.mcq_options_combo = ttk.Combobox(self.left_frame, textvariable=self.mcq_options_var, values=["7"], state="readonly", width=30)
        self.mcq_options_combo.pack(fill="x", pady=3)

        tk.Button(self.left_frame, text="Start Practice", command=self.start_practice, font=("Arial", 11, "bold")).pack(fill="x", pady=(12, 4))
        tk.Button(self.left_frame, text="Start Plan Session", command=self.start_today_plan).pack(fill="x", pady=3)
        tk.Button(self.left_frame, text="Set Plan", command=self.open_plan_window).pack(fill="x", pady=3)
        tk.Button(self.left_frame, text="Progress Dashboard", command=self.show_progress_dashboard).pack(fill="x", pady=3)
        tk.Button(self.left_frame, text="Show Statistics", command=self.show_statistics).pack(fill="x", pady=3)
        tk.Button(self.left_frame, text="Save Now", command=self.save_now).pack(fill="x", pady=3)
        tk.Button(self.left_frame, text="Reset Current Word", command=self.reset_current_word).pack(fill="x", pady=3)
        tk.Button(self.left_frame, text="Reset All Progress", command=self.reset_all_progress).pack(fill="x", pady=3)

        tk.Label(
            self.left_frame,
            text="Shortcuts:\nSpace = reveal/check\nEnter = submit\n←/→ = previous/next\n1 Again | 2 Hard | 3 Good | 4 Easy",
            font=("Arial", 8), justify="left"
        ).pack(anchor="w", pady=(10, 0))

    def setup_card_area(self) -> None:
        self.deck_label = tk.Label(self.main_frame, text="", font=("Arial", 16, "bold"))
        self.deck_label.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))

        self.counter_label = tk.Label(self.main_frame, text="", font=("Arial", 10, "italic"))
        self.counter_label.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        self.card_frame = tk.Frame(self.main_frame, bd=2, relief="solid", padx=18, pady=18)
        self.card_frame.grid(row=2, column=0, sticky="nsew", padx=(0, 8))
        self.card_frame.columnconfigure(0, weight=1)
        self.card_frame.rowconfigure(7, weight=1)

        self.word_label = tk.Label(self.card_frame, text="", font=("Arial", 30, "bold"), wraplength=730, justify="center")
        self.word_label.grid(row=0, column=0, sticky="ew", pady=(0, 6))

        self.pos_label = tk.Label(self.card_frame, text="", font=("Arial", 13, "italic"), fg="darkgreen")
        self.pos_label.grid(row=1, column=0, sticky="ew", pady=2)

        self.prompt_label = tk.Label(self.card_frame, text="", font=("Arial", 18), wraplength=760, justify="center")
        self.prompt_label.grid(row=2, column=0, sticky="ew", pady=8)

        self.answer_label = tk.Label(self.card_frame, text="", font=("Arial", 16), wraplength=760, justify="center")
        self.answer_label.grid(row=3, column=0, sticky="ew", pady=4)

        self.example_label = tk.Label(self.card_frame, text="", font=("Arial", 11), fg="dimgray", wraplength=760, justify="center")
        self.example_label.grid(row=4, column=0, sticky="ew", pady=2)

        self.extra_label = tk.Label(self.card_frame, text="", font=("Arial", 11), fg="darkblue", wraplength=760, justify="center")
        self.extra_label.grid(row=5, column=0, sticky="ew", pady=2)

        self.status_label = tk.Label(self.card_frame, text="", font=("Arial", 11, "bold"))
        self.status_label.grid(row=6, column=0, sticky="ew", pady=5)

        self.input_frame = tk.Frame(self.card_frame)
        self.input_frame.grid(row=7, column=0, sticky="nsew", pady=8)
        self.input_frame.columnconfigure(0, weight=1)

        self.typing_entry = tk.Entry(self.input_frame, font=("Arial", 14))
        self.submit_button = tk.Button(self.input_frame, text="Submit", command=self.submit_text_answer, font=("Arial", 11))

        self.mc_var = tk.StringVar(value="")
        self.mc_buttons: list[tk.Radiobutton] = []
        for i in range(7):
            rb = tk.Radiobutton(
                self.input_frame, text="", variable=self.mc_var, value="",
                anchor="w", justify="left", wraplength=720, font=("Arial", 11)
            )
            self.mc_buttons.append(rb)

    def setup_mnemonic_panel(self) -> None:
        # Right-side panel: mnemonic + flags + user notes + current word source/group info.
        self.mnemonic_frame = tk.Frame(self.main_frame, bd=2, relief="solid", padx=10, pady=10)
        self.mnemonic_frame.grid(row=2, column=1, sticky="nsew")
        self.mnemonic_frame.columnconfigure(0, weight=1)
        self.mnemonic_frame.rowconfigure(2, weight=1)
        self.mnemonic_frame.rowconfigure(7, weight=1)

        tk.Label(self.mnemonic_frame, text="Right Panel", font=("Arial", 14, "bold")).grid(row=0, column=0, sticky="ew")

        button_frame = tk.Frame(self.mnemonic_frame)
        button_frame.grid(row=1, column=0, sticky="ew", pady=6)
        for idx, name in enumerate(["Default", "Mine", "Both", "Hide"]):
            btn = tk.Button(button_frame, text=name, command=lambda n=name: self.set_mnemonic_view(n), width=7)
            btn.grid(row=0, column=idx, padx=1)

        self.mnemonic_text = tk.Text(self.mnemonic_frame, wrap="word", height=8, font=("Arial", 10))
        self.mnemonic_text.grid(row=2, column=0, sticky="nsew")
        self.mnemonic_text.configure(state="disabled")

        self.right_flag_label = tk.Label(
            self.mnemonic_frame, text="Favorite: No | Difficult: No | Not selected",
            font=("Arial", 10, "bold"), wraplength=300, justify="left"
        )
        self.right_flag_label.grid(row=3, column=0, sticky="ew", pady=(8, 2))

        self.right_source_label = tk.Label(
            self.mnemonic_frame, text="Source/group info appears here.",
            font=("Arial", 9), fg="dimgray", wraplength=300, justify="left"
        )
        self.right_source_label.grid(row=4, column=0, sticky="ew", pady=2)

        tk.Label(self.mnemonic_frame, text="My Notes", font=("Arial", 12, "bold")).grid(row=5, column=0, sticky="w", pady=(10, 2))
        self.right_notes_text = tk.Text(self.mnemonic_frame, wrap="word", height=10, font=("Arial", 10))
        self.right_notes_text.grid(row=6, column=0, sticky="nsew")
        self.right_notes_text.configure(state="disabled")

        self.right_progress_label = tk.Label(
            self.mnemonic_frame, text="", font=("Arial", 9), fg="darkblue", wraplength=300, justify="left"
        )
        self.right_progress_label.grid(row=7, column=0, sticky="ew", pady=(8, 0))

    def setup_controls(self) -> None:
        self.control_frame = tk.Frame(self.main_frame, pady=8)
        self.control_frame.grid(row=3, column=0, columnspan=2, sticky="ew")
        self.control_frame.columnconfigure(tuple(range(10)), weight=1)

        self.reveal_button = tk.Button(self.control_frame, text="Reveal", command=self.reveal_answer, width=11)
        self.reveal_button.grid(row=0, column=0, padx=3)
        self.again_button = tk.Button(self.control_frame, text="Again (1)", command=lambda: self.rate_current_card("again"), width=11)
        self.again_button.grid(row=0, column=1, padx=3)
        self.hard_button = tk.Button(self.control_frame, text="Hard (2)", command=lambda: self.rate_current_card("hard"), width=11)
        self.hard_button.grid(row=0, column=2, padx=3)
        self.good_button = tk.Button(self.control_frame, text="Good (3)", command=lambda: self.rate_current_card("good"), width=11)
        self.good_button.grid(row=0, column=3, padx=3)
        self.easy_button = tk.Button(self.control_frame, text="Easy (4)", command=lambda: self.rate_current_card("easy"), width=11)
        self.easy_button.grid(row=0, column=4, padx=3)
        self.prev_button = tk.Button(self.control_frame, text="← Prev", command=self.prev_card, width=10)
        self.prev_button.grid(row=0, column=5, padx=3)
        self.next_button = tk.Button(self.control_frame, text="Next →", command=self.next_card, width=10)
        self.next_button.grid(row=0, column=6, padx=3)
        self.note_button = tk.Button(self.control_frame, text="Edit Note", command=self.open_note_editor, width=10)
        self.note_button.grid(row=0, column=7, padx=3)
        self.favorite_button = tk.Button(self.control_frame, text="Favorite", command=lambda: self.toggle_flag("favorite"), width=10)
        self.favorite_button.grid(row=0, column=8, padx=3)
        self.difficult_button = tk.Button(self.control_frame, text="Difficult", command=lambda: self.toggle_flag("difficult"), width=10)
        self.difficult_button.grid(row=0, column=9, padx=3)

    def bind_shortcuts(self) -> None:
        self.bind("<space>", lambda e: self.space_action())
        self.bind("<Return>", lambda e: self.enter_action())
        self.bind("<Right>", lambda e: self.next_card())
        self.bind("<Left>", lambda e: self.prev_card())
        self.bind("1", lambda e: self.rate_current_card("again"))
        self.bind("2", lambda e: self.rate_current_card("hard"))
        self.bind("3", lambda e: self.rate_current_card("good"))
        self.bind("4", lambda e: self.rate_current_card("easy"))

    # ------------------------------------------------------------------
    # UI data refresh
    # ------------------------------------------------------------------
    def refresh_group_options(self) -> None:
        source = self.source_var.get()
        groups = ["All"]
        roots = ["All"]
        tones = ["All"]
        freqs = ["All"]

        if source == "Ultimate Total List":
            entries = list(self.model.words.values())
            groups_set = set()
            for entry in entries:
                for _, group in entry.sources:
                    if group:
                        groups_set.add(group)
                if entry.primary.get("root"):
                    roots.append(entry.primary["root"])
                if entry.primary.get("tone"):
                    tones.append(entry.primary["tone"])
                if entry.primary.get("frequency"):
                    freqs.append(entry.primary["frequency"])
            groups += sorted(groups_set, key=natural_sort_key)
        else:
            df = self.model.sources.get(source, pd.DataFrame())
            if not df.empty:
                groups += sorted(ensure_unique_list(df["group"].tolist()), key=natural_sort_key)
                roots += sorted(ensure_unique_list(df["root"].tolist()))
                tones += sorted(ensure_unique_list(df["tone"].tolist()))
                freqs += sorted(ensure_unique_list(df["frequency"].tolist()))

        self.group_combo["values"] = groups
        self.root_combo["values"] = ["All"] + sorted(set(r for r in roots if r and r != "All"))
        self.tone_combo["values"] = ["All"] + sorted(set(t for t in tones if t and t != "All"))
        self.frequency_combo["values"] = ["All"] + sorted(set(f for f in freqs if f and f != "All"))
        self.group_var.set("All")
        self.root_var.set("All")
        self.tone_var.set("All")
        self.frequency_var.set("All")

    def clear_card_display(self) -> None:
        self.word_label.config(text="")
        self.pos_label.config(text="")
        self.prompt_label.config(text="")
        self.answer_label.config(text="")
        self.example_label.config(text="")
        self.extra_label.config(text="")
        self.status_label.config(text="")
        self.hide_inputs()
        self.update_right_panel(None, revealed=False)

    def hide_inputs(self) -> None:
        self.typing_entry.grid_forget()
        self.submit_button.grid_forget()
        for rb in self.mc_buttons:
            rb.grid_forget()

    # ------------------------------------------------------------------
    # Deck building
    # ------------------------------------------------------------------
    def unique_entries(self, entries: list[WordEntry]) -> list[WordEntry]:
        """Remove duplicate WordEntry objects without requiring WordEntry to be hashable."""
        seen: set[str] = set()
        unique: list[WordEntry] = []
        for entry in entries:
            key = clean_text(getattr(entry, "word_id", "")) or normalize_word(getattr(entry, "display_word", ""))
            if key and key not in seen:
                unique.append(entry)
                seen.add(key)
        return unique

    def same_group(self, a: str, b: str) -> bool:
        """Compare group values safely, so Excel values like 1 and 1.0 both match group '1'."""
        a = clean_text(a)
        b = clean_text(b)
        if a == b:
            return True
        try:
            return str(int(float(a))) == str(int(float(b)))
        except Exception:
            return False

    def build_entries_from_selection(self, source: str, group: str) -> list[WordEntry]:
        if source == "Ultimate Total List":
            entries = list(self.model.words.values())
            if group != "All":
                entries = [e for e in entries if any(self.same_group(group, g) for _, g in e.sources)]
            return self.unique_entries(entries)

        df = self.model.sources.get(source, pd.DataFrame())
        entries: list[WordEntry] = []
        for _, row in df.iterrows():
            if group != "All" and not self.same_group(clean_text(row.get("group", "")), group):
                continue
            normalized = normalize_word(row.get("word", ""))
            entry = self.model.words.get(normalized)
            if entry:
                entries.append(entry)
        return self.unique_entries(entries)

    def apply_filters(self, entries: list[WordEntry]) -> list[WordEntry]:
        status_filter = self.filter_var.get()
        root_filter = self.root_var.get()
        tone_filter = self.tone_var.get()
        freq_filter = self.frequency_var.get()
        today = today_iso()

        def ok(entry: WordEntry) -> bool:
            flags = self.model.get_flags(entry.word_id)
            prog = self.model.get_progress(entry.word_id)
            status = prog.get("status", "unseen")
            if root_filter != "All" and entry.primary.get("root", "") != root_filter:
                return False
            if tone_filter != "All" and entry.primary.get("tone", "") != tone_filter:
                return False
            if freq_filter != "All" and entry.primary.get("frequency", "") != freq_filter:
                return False
            if status_filter == "All":
                return True
            if status_filter == "Due":
                return status != "unseen" and clean_text(prog.get("next_review", "")) <= today
            if status_filter in ["Unseen", "Learning", "Reviewing", "Mastered"]:
                return status == status_filter.lower()
            if status_filter == "Difficult":
                return flags.get("difficult", False)
            if status_filter == "Favorite":
                return flags.get("favorite", False)
            return True

        return [e for e in entries if ok(e)]

    def sort_entries(self, entries: list[WordEntry], order: str) -> list[WordEntry]:
        entries = list(entries)
        if order == "Shuffle":
            random.shuffle(entries)
        elif order == "Listwise":
            pass
        elif order == "Due First":
            entries.sort(key=lambda e: (
                0 if self.is_due(e) else 1,
                self.model.get_progress(e.word_id).get("next_review", "9999-99-99"),
                e.normalized_word,
            ))
        elif order == "Weak First":
            entries.sort(key=lambda e: (
                -int(self.model.get_progress(e.word_id).get("wrong_count", 0)),
                e.normalized_word,
            ))
        elif order == "New First":
            entries.sort(key=lambda e: (
                0 if self.model.get_progress(e.word_id).get("status", "unseen") == "unseen" else 1,
                e.normalized_word,
            ))
        return entries

    def is_due(self, entry: WordEntry) -> bool:
        prog = self.model.get_progress(entry.word_id)
        status = prog.get("status", "unseen")
        next_review = clean_text(prog.get("next_review", ""))
        return status != "unseen" and bool(next_review) and next_review <= today_iso()

    def limit_entries(self, entries: list[WordEntry], count_value: str) -> list[WordEntry]:
        if count_value == "All":
            return entries
        try:
            count = int(count_value)
            return entries[:count]
        except Exception:
            return entries

    def start_practice(self) -> None:
        source = self.source_var.get()
        group = self.group_var.get()
        mode = self.mode_var.get()
        order = self.order_var.get()
        count = self.count_var.get()

        entries = self.build_entries_from_selection(source, group)
        entries = self.apply_filters(entries)
        entries = self.sort_entries(entries, order)
        entries = self.limit_entries(entries, count)

        if not entries:
            messagebox.showinfo("No words", "No matching words found for the selected criteria.")
            return
        title = self.make_session_title(source, group, mode, entries)
        self.begin_session(entries, source, group, mode, order, title=title)

    def make_session_title(self, source: str, group: str, mode: str, entries: list[WordEntry]) -> str:
        """Build an informative title, especially for Ultimate Total List."""
        if source == "Ultimate Total List":
            dataset_names = set()
            group_nums = set()
            for entry in entries:
                for src, grp in entry.sources:
                    if src:
                        dataset_names.add(src)
                    if grp:
                        n = parse_group_number(grp)
                        if n != 10**9:
                            group_nums.add(n)
            dataset_count = len(dataset_names)
            if group == "All":
                group_text = "All groups"
            else:
                group_text = f"Group {group}"
            return f"Ultimate Total List | {group_text} | Datasets: {dataset_count} | Words: {len(entries)} | {mode}"
        return f"{source} | Group {group} | Words: {len(entries)} | {mode}"

    def word_dataset_count(self, entry: WordEntry) -> int:
        return len(set(src for src, _ in entry.sources if src))

    def word_group_text(self, entry: WordEntry) -> str:
        groups = []
        for _, grp in entry.sources:
            if grp:
                n = parse_group_number(grp)
                groups.append(str(n) if n != 10**9 else clean_text(grp))
        groups = ensure_unique_list(groups)
        return ", ".join(groups[:8]) + ("..." if len(groups) > 8 else "") if groups else "-"

    def begin_session(self, entries: list[WordEntry], source: str, group: str, mode: str, order: str, title: str) -> None:
        self.current_deck = entries
        self.current_index = 0
        self.current_mode = mode
        self.session_id = uuid.uuid4().hex[:12]
        self.session_start_time = time.time()
        self.session_context = {"source": source, "group": group, "mode": mode, "order": order, "title": title}
        self.session_answered = 0
        self.session_correct = 0
        self.session_wrong = 0
        self.session_new = sum(1 for e in entries if self.model.get_progress(e.word_id).get("status", "unseen") == "unseen")
        self.session_due = sum(1 for e in entries if self.is_due(e))
        self.deck_label.config(text=title)
        self.show_card()

    # ------------------------------------------------------------------
    # Plan Study: Daily / Weekly / Monthly
    # ------------------------------------------------------------------
    def open_plan_window(self) -> None:
        """Set a clear Daily / Weekly / Monthly study plan.

        The app still starts one study session at a time, but weekly/monthly plans
        are shown in user-friendly targets and automatically converted into a
        reasonable daily session target when the user clicks Start Plan.
        """
        win = tk.Toplevel(self)
        win.title("Set Study Plan")
        win.geometry("500x680")
        win.grab_set()

        plan = self.model.study_plan or {}
        values_sources = ["Ultimate Total List"] + list(self.model.sources.keys())

        fields: dict[str, tk.StringVar] = {}
        row = 0
        tk.Label(win, text="Study Plan", font=("Arial", 16, "bold")).grid(row=row, column=0, columnspan=2, pady=10)
        row += 1

        tk.Label(
            win,
            text="Choose Daily, Weekly, or Monthly. The app will calculate today’s session from that plan.",
            font=("Arial", 9), fg="dimgray", justify="left", wraplength=455,
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=10, pady=(0, 8))
        row += 1

        def add_combo(label: str, key: str, values: list[str], default: str) -> ttk.Combobox:
            nonlocal row
            tk.Label(win, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=4)
            var = tk.StringVar(value=clean_text(plan.get(key, default)) or default)
            fields[key] = var
            combo = ttk.Combobox(win, textvariable=var, values=values, state="readonly", width=30)
            combo.grid(row=row, column=1, sticky="ew", padx=10, pady=4)
            row += 1
            return combo

        def add_entry(label: str, key: str, default: str) -> tk.Entry:
            nonlocal row
            tk.Label(win, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=4)
            var = tk.StringVar(value=clean_text(plan.get(key, default)) or default)
            fields[key] = var
            ent = tk.Entry(win, textvariable=var, width=32)
            ent.grid(row=row, column=1, sticky="ew", padx=10, pady=4)
            row += 1
            return ent

        add_combo("Plan type", "plan_period", ["Daily", "Weekly", "Monthly"], "Daily")
        add_combo("Dataset / source", "source", values_sources, "Ultimate Total List")
        add_entry("Start group", "start_group", "1")
        add_entry("End group", "end_group", "1")

        tk.Label(win, text="Targets", font=("Arial", 11, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 2))
        row += 1
        add_entry("Daily new words", "new_words_per_day", "30")
        add_entry("Weekly new words", "weekly_new_words", "210")
        add_entry("Monthly new words", "monthly_new_words", "900")
        add_entry("Daily review limit", "max_reviews_per_day", "200")
        add_entry("Weekly review target", "weekly_review_target", "1000")
        add_entry("Monthly review target", "monthly_review_target", "4000")

        tk.Label(win, text="Session settings", font=("Arial", 11, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", padx=10, pady=(10, 2))
        row += 1
        add_combo("Study mode", "mode", self.study_modes, "Mixed Mode")
        add_combo("Order", "order", self.orders, "Due First")
        add_combo("Include due reviews", "include_due", ["True", "False"], "True")
        add_combo("Group lock", "group_lock", ["True", "False"], "True")
        add_combo("Enabled", "enabled", ["True", "False"], "True")
        add_entry("Start date", "start_date", today_iso())
        add_entry("End date", "end_date", "")

        tk.Label(
            win,
            text=(
                "Group lock ON: the app uses the earliest unfinished group in the selected range.\n"
                "Group lock OFF: the app can pull words from all selected groups.\n"
                "Weekly/monthly targets are converted to today’s practical session size."
            ),
            font=("Arial", 9), justify="left", fg="dimgray", wraplength=455,
        ).grid(row=row, column=0, columnspan=2, sticky="w", padx=10, pady=8)
        row += 1

        def save_plan() -> None:
            self.model.study_plan = {
                "plan_name": "Default",
                "plan_period": fields["plan_period"].get(),
                "enabled": fields["enabled"].get(),
                "source": fields["source"].get(),
                "start_group": fields["start_group"].get(),
                "end_group": fields["end_group"].get(),
                "new_words_per_day": fields["new_words_per_day"].get(),
                "weekly_new_words": fields["weekly_new_words"].get(),
                "monthly_new_words": fields["monthly_new_words"].get(),
                "max_reviews_per_day": fields["max_reviews_per_day"].get(),
                "weekly_review_target": fields["weekly_review_target"].get(),
                "monthly_review_target": fields["monthly_review_target"].get(),
                "mode": fields["mode"].get(),
                "order": fields["order"].get(),
                "include_due": fields["include_due"].get(),
                "group_lock": fields["group_lock"].get(),
                "start_date": fields["start_date"].get(),
                "end_date": fields["end_date"].get(),
                "created_at": clean_text(plan.get("created_at", now_iso())) or now_iso(),
                "updated_at": now_iso(),
            }
            self.model.save_user_data()
            messagebox.showinfo("Saved", f"{fields['plan_period'].get()} plan saved.")
            win.destroy()

        tk.Button(win, text="Save Study Plan", command=save_plan, font=("Arial", 11, "bold")).grid(row=row, column=0, columnspan=2, pady=12)
        win.columnconfigure(1, weight=1)

    def start_today_plan(self) -> None:
        plan = self.model.study_plan or {}
        if clean_text(plan.get("enabled", "False")).lower() != "true":
            if not messagebox.askyesno("Plan disabled", "This study plan is disabled. Start it anyway?"):
                return
        source = clean_text(plan.get("source", "Ultimate Total List")) or "Ultimate Total List"
        mode = clean_text(plan.get("mode", "Mixed Mode")) or "Mixed Mode"
        order = clean_text(plan.get("order", "Due First")) or "Due First"
        include_due = clean_text(plan.get("include_due", "True")).lower() == "true"
        group_lock = clean_text(plan.get("group_lock", "True")).lower() == "true"
        plan_period = clean_text(plan.get("plan_period", "Daily")) or "Daily"
        start_group = parse_group_number(plan.get("start_group", "1"))
        end_group = parse_group_number(plan.get("end_group", "1"))

        # Convert Daily / Weekly / Monthly plan into today's practical session target.
        if plan_period == "Weekly":
            weekly_new = self.model._to_int(plan.get("weekly_new_words", "210"), 210)
            weekly_reviews = self.model._to_int(plan.get("weekly_review_target", "1000"), 1000)
            new_limit = max(1, (weekly_new + 6) // 7)
            review_limit = max(1, (weekly_reviews + 6) // 7)
        elif plan_period == "Monthly":
            monthly_new = self.model._to_int(plan.get("monthly_new_words", "900"), 900)
            monthly_reviews = self.model._to_int(plan.get("monthly_review_target", "4000"), 4000)
            new_limit = max(1, (monthly_new + 29) // 30)
            review_limit = max(1, (monthly_reviews + 29) // 30)
        else:
            new_limit = self.model._to_int(plan.get("new_words_per_day", "30"), 30)
            review_limit = self.model._to_int(plan.get("max_reviews_per_day", "200"), 200)

        selected_group = "All"
        if source == "Ultimate Total List":
            entries_all = list(self.model.words.values())
        else:
            entries_all = self.build_entries_from_selection(source, "All")

        # Filter to group range for sheet-based sources. For Ultimate, use any source group within range.
        def in_group_range(entry: WordEntry) -> bool:
            group_nums: list[int] = []
            if source == "Ultimate Total List":
                group_nums = [parse_group_number(g) for _, g in entry.sources if g]
            else:
                group_nums = [parse_group_number(g) for s, g in entry.sources if s == source and g]
            return any(start_group <= n <= end_group for n in group_nums)

        pool = [e for e in entries_all if in_group_range(e)]
        if group_lock:
            # Earliest group in range with unmastered/unseen words.
            for group_no in range(start_group, end_group + 1):
                group_entries = []
                for e in pool:
                    if source == "Ultimate Total List":
                        nums = [parse_group_number(g) for _, g in e.sources if g]
                    else:
                        nums = [parse_group_number(g) for s, g in e.sources if s == source and g]
                    if group_no in nums:
                        prog = self.model.get_progress(e.word_id)
                        if prog.get("status", "unseen") != "mastered":
                            group_entries.append(e)
                if group_entries:
                    pool = group_entries
                    selected_group = str(group_no)
                    break

        due = self.model.due_words(pool, review_limit) if include_due else []
        new = self.model.unseen_words(pool, new_limit)
        deck = self.unique_entries(due + new)
        deck = self.sort_entries(deck, order)
        if not deck:
            messagebox.showinfo("Study Plan", "No due or new words found for the current plan.")
            return
        title = f"{plan_period} Plan | " + self.make_session_title(source, selected_group, mode, deck)
        self.begin_session(deck, source, selected_group, mode, order, title)

    # ------------------------------------------------------------------
    # Card display and modes
    # ------------------------------------------------------------------
    def actual_mode_for_card(self) -> str:
        if self.current_mode != "Mixed Mode":
            return self.current_mode
        # Randomly choose among active learning modes, excluding mixed itself.
        return random.choice([
            "Flashcard",
            "Reverse Flashcard",
            "Typing",
            "Multiple Choice (7 options)",
        ])

    def show_card(self) -> None:
        if not self.current_deck or self.current_index >= len(self.current_deck):
            self.finish_session()
            return
        self.hide_inputs()
        self.showing_answer = False
        self.current_options = []
        self.current_correct_answer = ""
        self.current_prompt_text = ""
        self.current_prompt_type = ""
        self.card_start_time = time.time()

        entry = self.current_deck[self.current_index]
        mode = self.actual_mode_for_card()
        self.active_card_mode = mode

        self.update_status_label(entry)
        self.word_label.config(text="")
        self.pos_label.config(text=entry.primary.get("part_of_speech", ""))
        self.prompt_label.config(text="")
        self.answer_label.config(text="")
        self.example_label.config(text="")
        self.extra_label.config(text="")
        self.update_right_panel(entry, revealed=False)

        if mode == "Flashcard":
            self.word_label.config(text=entry.display_word)
            self.prompt_label.config(text="Think of the meaning, then reveal.")
            self.current_prompt_type = "word_to_meaning"
            self.current_prompt_text = entry.display_word
            self.current_correct_answer = entry.primary.get("meaning", "")
        elif mode == "Reverse Flashcard":
            self.word_label.config(text="Meaning → Word")
            self.prompt_label.config(text=compact_meaning(entry.primary.get("meaning", ""), 300))
            self.current_prompt_type = "meaning_to_word"
            self.current_prompt_text = entry.primary.get("meaning", "")
            self.current_correct_answer = entry.display_word
        elif mode == "Typing: Meaning → Word":
            self.word_label.config(text="Type the word")
            self.prompt_label.config(text=compact_meaning(entry.primary.get("meaning", ""), 300))
            self.current_prompt_type = "meaning_to_word"
            self.current_prompt_text = entry.primary.get("meaning", "")
            self.current_correct_answer = entry.display_word
            self.show_text_input()
        elif mode == "Typing: Word → Meaning":
            self.word_label.config(text=entry.display_word)
            self.prompt_label.config(text="Type a short meaning.")
            self.current_prompt_type = "word_to_meaning"
            self.current_prompt_text = entry.display_word
            self.current_correct_answer = entry.primary.get("meaning", "")
            self.show_text_input()
        elif mode == "Typing":
            self.setup_typing(entry)
        elif mode == "Multiple Choice (7 options)":
            self.setup_mcq(entry)
        elif mode == "Fill Answer":
            self.setup_fill_answer(entry)
        else:
            self.word_label.config(text=entry.display_word)

        source_bits = f"Datasets: {self.word_dataset_count(entry)} | Groups: {self.word_group_text(entry)}"
        self.counter_label.config(
            text=f"Card {self.current_index + 1} of {len(self.current_deck)} | Mode: {mode} | {source_bits} | Session: {self.session_id}"
        )
        self.update_button_states()

    def update_status_label(self, entry: WordEntry) -> None:
        prog = self.model.get_progress(entry.word_id)
        flags = self.model.get_flags(entry.word_id)
        status = prog.get("status", "unseen")
        streak = prog.get("streak", 0)
        interval = prog.get("interval_days", 0)
        next_review = prog.get("next_review", "")
        flag_text = ", ".join([name for name, val in flags.items() if val])
        text = f"Status: {status} | Streak: {streak}/5 | Interval: {interval}d | Next: {next_review or '-'}"
        if flag_text:
            text += f" | Flags: {flag_text}"
        self.status_label.config(text=text)
        self.update_flag_note_summary(entry)

    def setup_mcq(self, entry: WordEntry) -> None:
        direction = random.choice(["word_to_meaning", "meaning_to_word"])
        self.current_prompt_type = direction
        if direction == "word_to_meaning":
            self.word_label.config(text=entry.display_word)
            self.prompt_label.config(text="Choose the correct meaning.")
            self.current_prompt_text = entry.display_word
            self.current_correct_answer = entry.primary.get("meaning", "")
            options = self.build_mcq_options(entry, option_field="meaning", total=7)
        else:
            self.word_label.config(text="Meaning → Word")
            meaning = compact_meaning(entry.primary.get("meaning", ""), 300)
            self.prompt_label.config(text=meaning)
            self.current_prompt_text = entry.primary.get("meaning", "")
            self.current_correct_answer = entry.display_word
            options = self.build_mcq_options(entry, option_field="word", total=7)
        self.current_options = options
        self.submit_button.config(text="Submit MCQ", command=self.submit_mcq_answer)
        self.mc_var.set("")
        for i, rb in enumerate(self.mc_buttons):
            text = options[i] if i < len(options) else ""
            rb.config(text=f"{i + 1}. {text}", value=text)
            rb.grid(row=i, column=0, sticky="ew", padx=5, pady=2)
        self.submit_button.grid(row=8, column=0, pady=8)

    def build_mcq_options(self, entry: WordEntry, option_field: str, total: int = 7) -> list[str]:
        pos = clean_text(entry.primary.get("part_of_speech", "")).lower()
        candidates = list(self.model.words.values())
        random.shuffle(candidates)

        def candidate_value(e: WordEntry) -> str:
            return e.display_word if option_field == "word" else e.primary.get("meaning", "")

        correct = candidate_value(entry)
        options = [correct]

        # Prefer same part of speech distractors, but never use synonym/antonym/example prompts.
        same_pos = [
            e for e in candidates
            if e.word_id != entry.word_id
            and candidate_value(e)
            and clean_text(e.primary.get("part_of_speech", "")).lower() == pos
        ]
        other = [
            e for e in candidates
            if e.word_id != entry.word_id
            and candidate_value(e)
            and e not in same_pos
        ]
        for e in same_pos + other:
            value = candidate_value(e)
            if value and value not in options:
                options.append(value)
            if len(options) >= total:
                break
        while len(options) < total:
            options.append(f"None of these {len(options)}")
        random.shuffle(options)
        return options[:total]

    def setup_fill_answer(self, entry: WordEntry) -> None:
        # Uses only word/meaning direction. No synonym/antonym/example prompt.
        direction = random.choice(["word_to_meaning", "meaning_to_word"])
        self.current_prompt_type = direction
        if direction == "meaning_to_word":
            self.word_label.config(text="Fill the word")
            self.prompt_label.config(text=compact_meaning(entry.primary.get("meaning", ""), 300))
            self.current_prompt_text = entry.primary.get("meaning", "")
            self.current_correct_answer = entry.display_word
        else:
            self.word_label.config(text=entry.display_word)
            self.prompt_label.config(text="Fill in the meaning.")
            self.current_prompt_text = entry.display_word
            self.current_correct_answer = entry.primary.get("meaning", "")
        self.show_text_input()

    def setup_typing(self, entry: WordEntry) -> None:
        """Typing mode: randomly asks meaning -> word or word -> meaning.

        This keeps the UI simple while still testing both recognition and recall.
        It uses only word/meaning logic, never synonym, antonym, or example prompts.
        """
        direction = random.choice(["word_to_meaning", "meaning_to_word"])
        self.current_prompt_type = direction
        if direction == "meaning_to_word":
            self.word_label.config(text="Type the word")
            self.prompt_label.config(text=compact_meaning(entry.primary.get("meaning", ""), 300))
            self.current_prompt_text = entry.primary.get("meaning", "")
            self.current_correct_answer = entry.display_word
        else:
            self.word_label.config(text=entry.display_word)
            self.prompt_label.config(text="Type a short meaning.")
            self.current_prompt_text = entry.display_word
            self.current_correct_answer = entry.primary.get("meaning", "")
        self.show_text_input()

    def show_text_input(self) -> None:
        self.submit_button.config(text="Submit", command=self.submit_text_answer)
        self.typing_entry.delete(0, tk.END)
        self.typing_entry.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.submit_button.grid(row=1, column=0, pady=5)
        self.after(100, self.typing_entry.focus_set)

    def update_button_states(self) -> None:
        mode = getattr(self, "active_card_mode", self.current_mode)
        if mode in ["Flashcard", "Reverse Flashcard"]:
            self.reveal_button.config(state="normal")
        else:
            self.reveal_button.config(state="disabled")
        # Rating buttons are always enabled; for MCQ/typing they become meaningful after submit.
        for b in [self.again_button, self.hard_button, self.good_button, self.easy_button]:
            b.config(state="normal")

    def reveal_answer(self) -> None:
        if not self.current_deck:
            return
        entry = self.current_deck[self.current_index]
        mode = getattr(self, "active_card_mode", self.current_mode)
        if mode == "Flashcard":
            self.answer_label.config(text=f"Meaning: {entry.primary.get('meaning', '')}")
            examples = []
            if entry.primary.get("example"):
                examples.append(f"Example: {entry.primary.get('example')}")
            if entry.primary.get("example2"):
                examples.append(f"Example 2: {entry.primary.get('example2')}")
            self.example_label.config(text="\n".join(examples))
            extra_parts = []
            if entry.primary.get("root"):
                extra_parts.append(f"Root: {entry.primary.get('root')}")
            if entry.primary.get("tone"):
                extra_parts.append(f"Tone: {entry.primary.get('tone')}")
            if entry.primary.get("frequency"):
                extra_parts.append(f"Frequency: {entry.primary.get('frequency')}")
            # Synonym/antonym can remain on revealed flashcard; they are not used in MCQ/Fill prompts.
            if entry.primary.get("synonym"):
                extra_parts.append(f"Syn: {entry.primary.get('synonym')}")
            if entry.primary.get("antonym"):
                extra_parts.append(f"Ant: {entry.primary.get('antonym')}")
            self.extra_label.config(text=" | ".join(extra_parts))
        elif mode == "Reverse Flashcard":
            self.answer_label.config(text=f"Word: {entry.display_word}")
            self.example_label.config(text=f"Meaning: {entry.primary.get('meaning', '')}")
        self.showing_answer = True
        self.update_right_panel(entry, revealed=True)

    def update_right_panel(self, entry: Optional[WordEntry], revealed: bool) -> None:
        self.update_mnemonic_panel(entry, revealed)
        self.update_flag_note_summary(entry)

    def update_mnemonic_panel(self, entry: Optional[WordEntry], revealed: bool) -> None:
        self.mnemonic_text.configure(state="normal")
        self.mnemonic_text.delete("1.0", tk.END)
        if not entry:
            self.mnemonic_text.insert("1.0", "Mnemonic appears here.")
        elif not revealed or self.mnemonic_view == "Hide":
            self.mnemonic_text.insert("1.0", "Mnemonic appears after reveal or after answer submission.")
        else:
            default_m = clean_text(entry.primary.get("mnemonic", ""))
            user_m = clean_text(self.model.get_user_note(entry.word_id).get("user_mnemonic", ""))
            parts = []
            if self.mnemonic_view in ["Default", "Both"]:
                parts.append("Default mnemonic:\n" + (default_m or "No default mnemonic in database."))
            if self.mnemonic_view in ["Mine", "Both"]:
                parts.append("My mnemonic:\n" + (user_m or "No personal mnemonic yet. Use Edit Note to add one."))
            self.mnemonic_text.insert("1.0", "\n\n".join(parts) if parts else "Mnemonic hidden.")
        self.mnemonic_text.configure(state="disabled")

    def update_flag_note_summary(self, entry: Optional[WordEntry]) -> None:
        if not hasattr(self, "right_notes_text"):
            return
        if not entry:
            self.right_flag_label.config(text="Favorite: No | Difficult: No | Not selected")
            self.right_source_label.config(text="Source/group info appears here.")
            self.right_progress_label.config(text="")
            self.right_notes_text.configure(state="normal")
            self.right_notes_text.delete("1.0", tk.END)
            self.right_notes_text.insert("1.0", "User notes appear here.")
            self.right_notes_text.configure(state="disabled")
            return

        flags = self.model.get_flags(entry.word_id)
        fav = "Yes" if flags.get("favorite") else "No"
        diff = "Yes" if flags.get("difficult") else "No"
        selected = []
        if flags.get("favorite"):
            selected.append("Favorite")
        if flags.get("difficult"):
            selected.append("Difficult")
        selected_text = ", ".join(selected) if selected else "Not selected"
        self.right_flag_label.config(text=f"Favorite: {fav} | Difficult: {diff} | {selected_text}")

        sources = ensure_unique_list([src for src, _ in entry.sources if src])
        source_line = f"Datasets: {len(sources)}"
        if sources:
            source_line += "\n" + ", ".join(sources[:4]) + ("..." if len(sources) > 4 else "")
        source_line += f"\nGroups: {self.word_group_text(entry)}"
        self.right_source_label.config(text=source_line)

        note = self.model.get_user_note(entry.word_id)
        note_parts = []
        # Personal mnemonic is intentionally hidden here.
        # It is shown only in the mnemonic panel after reveal when the user selects Mine or Both.
        if clean_text(note.get("user_meaning", "")):
            note_parts.append("Personal meaning:\n" + clean_text(note.get("user_meaning", "")))
        if clean_text(note.get("user_note", "")):
            note_parts.append("General note:\n" + clean_text(note.get("user_note", "")))
        if clean_text(note.get("user_translation", "")):
            note_parts.append("Translation:\n" + clean_text(note.get("user_translation", "")))
        if clean_text(note.get("user_tag", "")):
            note_parts.append("Tags:\n" + clean_text(note.get("user_tag", "")))
        if clean_text(note.get("updated_at", "")):
            note_parts.append("Updated:\n" + clean_text(note.get("updated_at", "")))

        self.right_notes_text.configure(state="normal")
        self.right_notes_text.delete("1.0", tk.END)
        self.right_notes_text.insert("1.0", "\n\n".join(note_parts) if note_parts else "No user notes yet. Use Edit Note to add your own mnemonic, meaning, translation, tags, or general note.")
        self.right_notes_text.configure(state="disabled")

        prog = self.model.get_progress(entry.word_id)
        progress_line = (
            f"Reviews: {prog.get('review_count', 0)} | Correct: {prog.get('correct_count', 0)} | "
            f"Wrong: {prog.get('wrong_count', 0)}\n"
            f"Last rating: {prog.get('last_rating', '-') or '-'} | Next review: {prog.get('next_review', '-') or '-'}"
        )
        self.right_progress_label.config(text=progress_line)

    def set_mnemonic_view(self, view: str) -> None:
        self.mnemonic_view = view
        self.model.app_settings["mnemonic_view"] = view
        if self.current_deck and self.current_index < len(self.current_deck):
            self.update_right_panel(self.current_deck[self.current_index], revealed=self.showing_answer)

    # ------------------------------------------------------------------
    # Answer submission and scoring
    # ------------------------------------------------------------------
    def space_action(self) -> None:
        mode = getattr(self, "active_card_mode", self.current_mode)
        if mode in ["Flashcard", "Reverse Flashcard"]:
            self.reveal_answer()
        elif mode in ["Typing", "Typing: Meaning → Word", "Typing: Word → Meaning", "Fill Answer"]:
            self.submit_text_answer()
        elif mode == "Multiple Choice (7 options)":
            self.submit_mcq_answer()

    def enter_action(self) -> None:
        self.space_action()

    def submit_text_answer(self) -> None:
        if not self.current_deck:
            return
        user_answer = self.typing_entry.get().strip()
        if not user_answer:
            messagebox.showinfo("Input needed", "Please type an answer.")
            return
        correct = self.evaluate_text_answer(user_answer)
        self.show_answer_result(user_answer, correct)
        self.rate_current_card("good" if correct else "again", user_answer=user_answer, is_correct=correct, auto_advance=True)

    def submit_mcq_answer(self) -> None:
        if not self.current_deck:
            return
        selected = self.mc_var.get()
        if not selected:
            messagebox.showinfo("Select an option", "Please select one of the 7 options.")
            return
        correct = clean_text(selected) == clean_text(self.current_correct_answer)
        self.show_answer_result(selected, correct)
        self.rate_current_card("good" if correct else "again", user_answer=selected, is_correct=correct, auto_advance=True)

    def evaluate_text_answer(self, user_answer: str) -> bool:
        correct = clean_text(self.current_correct_answer)
        if self.current_prompt_type == "meaning_to_word":
            return normalize_word(user_answer) == normalize_word(correct)
        # word_to_meaning: accept approximate meaning via keyword overlap.
        user_tokens = tokenize(user_answer)
        correct_tokens = tokenize(correct)
        if not user_tokens or not correct_tokens:
            return False
        overlap = len(user_tokens & correct_tokens)
        # Require at least one strong overlap and about 25% of the smaller token set.
        return overlap >= 1 and overlap / max(1, min(len(user_tokens), len(correct_tokens))) >= 0.25

    def show_answer_result(self, user_answer: str, correct: bool) -> None:
        entry = self.current_deck[self.current_index]
        verdict = "Correct" if correct else "Incorrect"
        self.answer_label.config(
            text=f"{verdict}\nYour answer: {user_answer}\nCorrect answer: {self.current_correct_answer}"
        )
        self.showing_answer = True
        self.update_right_panel(entry, revealed=True)

    def rate_current_card(self, rating: str, user_answer: str = "", is_correct: Optional[bool] = None, auto_advance: bool = False) -> None:
        if not self.current_deck:
            return
        entry = self.current_deck[self.current_index]
        prog = self.model.get_progress(entry.word_id)
        old_status = prog.get("status", "unseen")
        old_interval = int(prog.get("interval_days", 0))
        old_was_unseen = old_status == "unseen"

        if is_correct is None:
            is_correct = rating in ["hard", "good", "easy"]

        elapsed = max(0.0, time.time() - self.card_start_time)
        self.apply_srs(entry, rating, is_correct)
        new_prog = self.model.get_progress(entry.word_id)

        self.log_review(entry, rating, user_answer, bool(is_correct), elapsed, old_status, new_prog.get("status", "unseen"), old_interval, int(new_prog.get("interval_days", 0)))
        self.session_answered += 1
        if is_correct:
            self.session_correct += 1
        else:
            self.session_wrong += 1
        if old_was_unseen:
            self.session_new += 1

        if self.model.app_settings.get("auto_save", "True").lower() == "true":
            try:
                self.model.save_user_data()
            except Exception as exc:
                messagebox.showerror("Save error", f"Could not save user data: {exc}")

        if rating == "again":
            insert_pos = min(self.current_index + 3, len(self.current_deck))
            self.current_deck.insert(insert_pos, entry)
        elif rating == "hard":
            insert_pos = min(self.current_index + 5, len(self.current_deck))
            self.current_deck.insert(insert_pos, entry)

        self.current_index += 1
        if auto_advance:
            self.after(500, self.show_card)
        else:
            self.show_card()

    def apply_srs(self, entry: WordEntry, rating: str, is_correct: bool) -> None:
        prog = self.model.get_progress(entry.word_id)
        prog["review_count"] = int(prog.get("review_count", 0)) + 1
        prog["last_review"] = today_iso()
        prog["last_rating"] = rating

        ease = float(prog.get("ease_factor", 2.5))
        interval = int(prog.get("interval_days", 0))
        streak = int(prog.get("streak", 0))
        leech_score = int(prog.get("leech_score", 0))

        if not is_correct or rating == "again":
            prog["wrong_count"] = int(prog.get("wrong_count", 0)) + 1
            prog["streak"] = 0
            prog["ease_factor"] = max(1.3, ease - 0.2)
            prog["interval_days"] = 1
            prog["next_review"] = (_dt.date.today() + _dt.timedelta(days=1)).isoformat()
            prog["status"] = "learning"
            prog["leech_score"] = leech_score + 1
            return

        prog["correct_count"] = int(prog.get("correct_count", 0)) + 1
        streak += 1
        prog["streak"] = streak
        if rating == "hard":
            prog["ease_factor"] = max(1.3, ease)
            new_interval = 1 if interval == 0 else max(1, int(interval * 1.2))
        elif rating == "good":
            prog["ease_factor"] = max(1.3, ease + 0.1)
            if interval == 0:
                new_interval = 1
            elif interval == 1:
                new_interval = 3
            else:
                new_interval = max(1, int(interval * prog["ease_factor"]))
        else:  # easy
            prog["ease_factor"] = max(1.3, ease + 0.2)
            if interval == 0:
                new_interval = 3
            elif interval == 1:
                new_interval = 5
            else:
                new_interval = max(1, int(interval * prog["ease_factor"] * 1.3))
        prog["interval_days"] = new_interval
        prog["next_review"] = (_dt.date.today() + _dt.timedelta(days=new_interval)).isoformat()
        prog["status"] = "mastered" if streak >= 5 else "reviewing"

    # ------------------------------------------------------------------
    # Logging
    # ------------------------------------------------------------------
    def log_review(self, entry: WordEntry, rating: str, user_answer: str, is_correct: bool, elapsed: float,
                   old_status: str, new_status: str, old_interval: int, new_interval: int) -> None:
        source = self.session_context.get("source", "")
        group = self.session_context.get("group", "")
        self.model.review_log_rows.append({
            "timestamp": now_iso(),
            "session_id": self.session_id,
            "word_id": entry.word_id,
            "normalized_word": entry.normalized_word,
            "display_word": entry.display_word,
            "source": source,
            "group": group,
            "mode": getattr(self, "active_card_mode", self.current_mode),
            "prompt_type": self.current_prompt_type,
            "prompt": self.current_prompt_text,
            "correct_answer": self.current_correct_answer,
            "user_answer": user_answer,
            "is_correct": str(is_correct),
            "rating": rating,
            "response_time_sec": round(elapsed, 2),
            "old_status": old_status,
            "new_status": new_status,
            "old_interval": old_interval,
            "new_interval": new_interval,
            "next_review": self.model.get_progress(entry.word_id).get("next_review", ""),
        })
        self.update_daily_stats(is_correct, old_status, new_status, elapsed)

    def update_daily_stats(self, is_correct: bool, old_status: str, new_status: str, elapsed: float) -> None:
        today = today_iso()
        row = None
        for r in self.model.daily_stats_rows:
            if clean_text(r.get("date", "")) == today:
                row = r
                break
        if row is None:
            row = {
                "date": today,
                "studied": 0,
                "correct": 0,
                "wrong": 0,
                "accuracy": 0,
                "new_words": 0,
                "reviews": 0,
                "mastered_today": 0,
                "time_spent_sec": 0,
            }
            self.model.daily_stats_rows.append(row)
        row["studied"] = self.model._to_int(row.get("studied", 0)) + 1
        if is_correct:
            row["correct"] = self.model._to_int(row.get("correct", 0)) + 1
        else:
            row["wrong"] = self.model._to_int(row.get("wrong", 0)) + 1
        if old_status == "unseen":
            row["new_words"] = self.model._to_int(row.get("new_words", 0)) + 1
        else:
            row["reviews"] = self.model._to_int(row.get("reviews", 0)) + 1
        if old_status != "mastered" and new_status == "mastered":
            row["mastered_today"] = self.model._to_int(row.get("mastered_today", 0)) + 1
        row["time_spent_sec"] = self.model._to_float(row.get("time_spent_sec", 0.0)) + round(elapsed, 2)
        studied = self.model._to_int(row.get("studied", 0))
        correct = self.model._to_int(row.get("correct", 0))
        row["accuracy"] = round(correct / studied * 100, 2) if studied else 0

    def finish_session(self) -> None:
        if not self.current_deck:
            return
        duration = int(time.time() - self.session_start_time) if self.session_start_time else 0
        accuracy = round(self.session_correct / self.session_answered * 100, 2) if self.session_answered else 0
        self.model.session_log_rows.append({
            "session_id": self.session_id,
            "start_time": _dt.datetime.fromtimestamp(self.session_start_time).replace(microsecond=0).isoformat(sep=" ") if self.session_start_time else "",
            "end_time": now_iso(),
            "source": self.session_context.get("source", ""),
            "group": self.session_context.get("group", ""),
            "mode": self.session_context.get("mode", ""),
            "order": self.session_context.get("order", ""),
            "deck_size": len(self.current_deck),
            "answered": self.session_answered,
            "correct": self.session_correct,
            "wrong": self.session_wrong,
            "accuracy": accuracy,
            "new_words": self.session_new,
            "due_words": self.session_due,
            "duration_sec": duration,
        })
        try:
            self.model.save_user_data()
        except Exception as exc:
            messagebox.showerror("Save error", f"Could not save user data: {exc}")
        messagebox.showinfo("Session complete", f"Session complete. Accuracy: {accuracy}%")
        self.current_deck = []
        self.current_index = 0
        self.clear_card_display()
        self.deck_label.config(text="")
        self.counter_label.config(text="")

    # ------------------------------------------------------------------
    # Word finder
    # ------------------------------------------------------------------
    def find_word_dialog(self) -> None:
        query = ""
        if hasattr(self, "find_word_var"):
            query = clean_text(self.find_word_var.get())
        if not query:
            query = simpledialog.askstring("Find Word", "Enter a word or part of a word:", parent=self) or ""
            query = clean_text(query)
        if not query:
            return
        q_norm = normalize_word(query)
        if hasattr(self, "find_word_var"):
            self.find_word_var.set(query)

        exact_matches: list[WordEntry] = []
        partial_matches: list[WordEntry] = []
        for entry in self.model.words.values():
            if q_norm == entry.normalized_word:
                exact_matches.append(entry)
            elif q_norm in entry.normalized_word:
                partial_matches.append(entry)

        if exact_matches:
            self.open_found_word(exact_matches[0])
            return

        matches = sorted(partial_matches, key=lambda e: e.normalized_word)
        if not matches:
            messagebox.showinfo("Find Word", f"No matching word found for: {query}")
            return
        if len(matches) == 1:
            self.open_found_word(matches[0])
            return
        self.show_word_search_results(matches[:200], query)

    def show_word_search_results(self, matches: list[WordEntry], query: str) -> None:
        win = tk.Toplevel(self)
        win.title(f"Find Word: {query}")
        win.geometry("520x420")
        win.grab_set()
        tk.Label(win, text=f"Matches for '{query}'", font=("Arial", 13, "bold")).pack(pady=8)
        tk.Label(win, text="Select a word and click Open.", fg="dimgray").pack()
        listbox = tk.Listbox(win, height=15, font=("Arial", 11))
        listbox.pack(fill="both", expand=True, padx=10, pady=8)
        for entry in matches:
            listbox.insert(tk.END, f"{entry.display_word}  |  Datasets: {self.word_dataset_count(entry)}  |  Groups: {self.word_group_text(entry)}")

        def open_selected() -> None:
            sel = listbox.curselection()
            if not sel:
                messagebox.showinfo("Find Word", "Select a word first.")
                return
            entry = matches[sel[0]]
            win.destroy()
            self.open_found_word(entry)

        listbox.bind("<Double-Button-1>", lambda e: open_selected())
        listbox.bind("<Return>", lambda e: open_selected())
        win.bind("<Escape>", lambda e: win.destroy())
        if matches:
            listbox.selection_set(0)
            listbox.activate(0)
            listbox.focus_set()
        tk.Button(win, text="Open Selected Word", command=open_selected, font=("Arial", 11, "bold")).pack(pady=8)

    def open_found_word(self, entry: WordEntry) -> None:
        """Open a found word as a normal usable flashcard.

        This fixes the common Word Finder problem where a word is found but the
        card area does not become an active session. The found word becomes a
        one-card Flashcard session unless it is already in the current deck.
        """
        self.mode_var.set("Flashcard")
        if hasattr(self, "find_word_var"):
            self.find_word_var.set(entry.display_word)

        for idx, deck_entry in enumerate(self.current_deck):
            if deck_entry.word_id == entry.word_id:
                self.current_index = idx
                self.current_mode = "Flashcard"
                self.deck_label.config(text=f"Word Finder | {entry.display_word} | Datasets: {self.word_dataset_count(entry)}")
                self.show_card()
                return

        self.begin_session(
            [entry],
            "Word Finder",
            "Single",
            "Flashcard",
            "Listwise",
            title=f"Word Finder | {entry.display_word} | Datasets: {self.word_dataset_count(entry)} | Groups: {self.word_group_text(entry)}",
        )
        # Ensure the card is immediately usable after Find.
        self.reveal_button.config(state="normal")
        self.after(50, lambda: self.focus_force())

    # ------------------------------------------------------------------
    # Notes, flags, reset, stats
    # ------------------------------------------------------------------
    def open_note_editor(self) -> None:
        if not self.current_deck:
            return
        entry = self.current_deck[self.current_index]
        note = self.model.get_user_note(entry.word_id)
        win = tk.Toplevel(self)
        win.title(f"Edit notes: {entry.display_word}")
        win.geometry("560x650")
        win.grab_set()

        tk.Label(win, text=entry.display_word, font=("Arial", 18, "bold")).pack(pady=8)
        tk.Label(win, text=f"Meaning: {entry.primary.get('meaning', '')}", wraplength=520, justify="left").pack(fill="x", padx=10)
        tk.Label(win, text="Default mnemonic:", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
        default_box = tk.Text(win, height=4, wrap="word")
        default_box.pack(fill="x", padx=10)
        default_box.insert("1.0", entry.primary.get("mnemonic", "") or "No default mnemonic in database.")
        default_box.configure(state="disabled")

        fields: dict[str, tk.Text] = {}

        def add_text(label: str, key: str, height: int) -> None:
            tk.Label(win, text=label, font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=(8, 0))
            box = tk.Text(win, height=height, wrap="word")
            box.pack(fill="x", padx=10)
            box.insert("1.0", note.get(key, ""))
            fields[key] = box

        add_text("My mnemonic", "user_mnemonic", 3)
        add_text("Personal meaning", "user_meaning", 2)
        add_text("General note", "user_note", 3)
        add_text("Translation", "user_translation", 2)
        add_text("Tags", "user_tag", 1)

        def save_note() -> None:
            self.model.user_notes[entry.word_id] = {
                "user_note": fields["user_note"].get("1.0", "end").strip(),
                "user_meaning": fields["user_meaning"].get("1.0", "end").strip(),
                "user_mnemonic": fields["user_mnemonic"].get("1.0", "end").strip(),
                "user_translation": fields["user_translation"].get("1.0", "end").strip(),
                "user_tag": fields["user_tag"].get("1.0", "end").strip(),
                "updated_at": now_iso(),
            }
            self.model.save_user_data()
            self.update_right_panel(entry, revealed=self.showing_answer)
            win.destroy()

        tk.Button(win, text="Save", command=save_note, font=("Arial", 11, "bold")).pack(pady=12)

    def toggle_flag(self, flag: str) -> None:
        if not self.current_deck:
            return
        entry = self.current_deck[self.current_index]
        flags = self.model.get_flags(entry.word_id)
        flags[flag] = not flags.get(flag, False)
        self.model.save_user_data()
        self.update_status_label(entry)
        self.update_right_panel(entry, revealed=self.showing_answer)
        messagebox.showinfo("Flag updated", f"{flag.capitalize()} = {flags[flag]} for {entry.display_word}")

    def reset_current_word(self) -> None:
        if not self.current_deck:
            return
        entry = self.current_deck[self.current_index]
        if not messagebox.askyesno("Reset", f"Reset study progress and Favorite/Difficult flags for {entry.display_word}? Notes and mnemonics will remain."):
            return
        self.model.progress.pop(entry.word_id, None)
        self.model.word_flags.pop(entry.word_id, None)
        self.model.save_user_data()
        self.show_card()

    def reset_all_progress(self) -> None:
        if not messagebox.askyesno(
            "Reset All Progress",
            "Reset ALL progress and flags for every word?\n\nNotes, personal mnemonics, translations, tags, and logs will remain. This cannot be undone.",
        ):
            return
        self.model.progress.clear()
        self.model.word_flags.clear()
        self.model.save_user_data()
        if self.current_deck:
            self.show_card()
        else:
            self.clear_card()
        messagebox.showinfo("Reset Complete", "Progress and Favorite/Difficult flags have been reset. Notes, personal mnemonics, translations, tags, and logs were kept.")

    def show_progress_dashboard(self) -> None:
        """User-friendly progress dashboard for GRE study."""
        win = tk.Toplevel(self)
        win.title("GRE Progress Dashboard")
        win.geometry("720x620")
        nb = ttk.Notebook(win)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        def make_tab(title: str) -> tk.Text:
            frame = tk.Frame(nb)
            nb.add(frame, text=title)
            txt = tk.Text(frame, wrap="word", font=("Arial", 10))
            txt.pack(fill="both", expand=True)
            return txt

        total = len(self.model.words)
        counts = {"unseen": 0, "learning": 0, "reviewing": 0, "mastered": 0}
        due = 0
        favorite = 0
        difficult = 0
        for entry in self.model.words.values():
            prog = self.model.get_progress(entry.word_id)
            status = prog.get("status", "unseen")
            counts[status] = counts.get(status, 0) + 1
            if self.is_due(entry):
                due += 1
            flags = self.model.get_flags(entry.word_id)
            if flags.get("favorite"):
                favorite += 1
            if flags.get("difficult"):
                difficult += 1

        overview = make_tab("Overview")
        mastered_pct = round(counts.get("mastered", 0) / total * 100, 2) if total else 0
        overview.insert("1.0", (
            f"Total words: {total}\n"
            f"Mastered: {counts.get('mastered', 0)} ({mastered_pct}%)\n"
            f"Reviewing: {counts.get('reviewing', 0)}\n"
            f"Learning: {counts.get('learning', 0)}\n"
            f"Unseen: {counts.get('unseen', 0)}\n"
            f"Due today: {due}\n"
            f"Favorite: {favorite}\n"
            f"Difficult: {difficult}\n\n"
            "Recommended GRE routine:\n"
            "1. Finish due reviews first.\n"
            "2. Add planned new words.\n"
            "3. End with Difficult words.\n"
        ))
        overview.configure(state="disabled")

        # Daily / weekly / monthly summary from daily_stats
        stats = make_tab("Daily / Weekly / Monthly")
        today = _dt.date.today()
        week_start = today - _dt.timedelta(days=today.weekday())
        month_start = today.replace(day=1)

        def sum_rows(start_date: _dt.date, end_date: _dt.date) -> dict[str, float]:
            out = {"studied": 0, "correct": 0, "wrong": 0, "new_words": 0, "reviews": 0, "time_spent_sec": 0.0}
            for r in self.model.daily_stats_rows:
                try:
                    d = _dt.date.fromisoformat(clean_text(r.get("date", "")))
                except Exception:
                    continue
                if start_date <= d <= end_date:
                    for k in out:
                        if k == "time_spent_sec":
                            out[k] += self.model._to_float(r.get(k, 0), 0.0)
                        else:
                            out[k] += self.model._to_int(r.get(k, 0), 0)
            return out

        today_sum = sum_rows(today, today)
        week_sum = sum_rows(week_start, today)
        month_sum = sum_rows(month_start, today)

        def format_block(name: str, data: dict[str, float]) -> str:
            studied = int(data["studied"])
            correct = int(data["correct"])
            wrong = int(data["wrong"])
            acc = round(correct / studied * 100, 2) if studied else 0
            minutes = round(float(data["time_spent_sec"]) / 60, 1)
            return (
                f"{name}\n"
                f"  Studied: {studied}\n"
                f"  Correct: {correct}\n"
                f"  Wrong: {wrong}\n"
                f"  Accuracy: {acc}%\n"
                f"  New words: {int(data['new_words'])}\n"
                f"  Reviews: {int(data['reviews'])}\n"
                f"  Time: {minutes} minutes\n\n"
            )

        stats.insert("1.0", format_block("Today", today_sum) + format_block("This week", week_sum) + format_block("This month", month_sum))
        stats.configure(state="disabled")

        # Dataset accuracy from review log
        dataset_tab = make_tab("Datasets")
        ds: dict[str, dict[str, int]] = {}
        for r in self.model.review_log_rows:
            source = clean_text(r.get("source", "Unknown")) or "Unknown"
            d = ds.setdefault(source, {"answered": 0, "correct": 0})
            d["answered"] += 1
            if clean_text(r.get("is_correct", "False")).lower() == "true":
                d["correct"] += 1
        lines = []
        for source, d in sorted(ds.items(), key=lambda kv: (-kv[1]["answered"], kv[0])):
            acc = round(d["correct"] / d["answered"] * 100, 2) if d["answered"] else 0
            lines.append(f"{source}: {d['correct']}/{d['answered']} correct ({acc}%)")
        dataset_tab.insert("1.0", "\n".join(lines) if lines else "No dataset log yet.")
        dataset_tab.configure(state="disabled")

        # Weak words by wrong count
        weak_tab = make_tab("Weak Words")
        weak = []
        for entry in self.model.words.values():
            prog = self.model.get_progress(entry.word_id)
            wrong = self.model._to_int(prog.get("wrong_count", 0), 0)
            reviews = self.model._to_int(prog.get("review_count", 0), 0)
            if wrong > 0:
                weak.append((wrong, reviews, entry.display_word, entry.primary.get("meaning", "")))
        weak.sort(key=lambda x: (-x[0], -x[1], x[2].lower()))
        weak_lines = [f"{w} wrong / {r} reviews — {word}: {compact_meaning(meaning, 120)}" for w, r, word, meaning in weak[:50]]
        weak_tab.insert("1.0", "\n".join(weak_lines) if weak_lines else "No weak-word history yet.")
        weak_tab.configure(state="disabled")

        tk.Button(win, text="Close", command=win.destroy).pack(pady=(0, 8))

    def show_statistics(self) -> None:
        total = len(self.model.words)
        counts = {"unseen": 0, "learning": 0, "reviewing": 0, "mastered": 0}
        due = 0
        for entry in self.model.words.values():
            prog = self.model.get_progress(entry.word_id)
            status = prog.get("status", "unseen")
            counts[status] = counts.get(status, 0) + 1
            if self.is_due(entry):
                due += 1
        flag_counts = {"favorite": 0, "difficult": 0}
        for flags in self.model.word_flags.values():
            for k in flag_counts:
                if flags.get(k):
                    flag_counts[k] += 1
        text = (
            f"Total unique words: {total}\n"
            f"Unseen: {counts.get('unseen', 0)}\n"
            f"Learning: {counts.get('learning', 0)}\n"
            f"Reviewing: {counts.get('reviewing', 0)}\n"
            f"Mastered: {counts.get('mastered', 0)}\n"
            f"Due today: {due}\n\n"
            f"Favorite: {flag_counts['favorite']}\n"
            f"Difficult: {flag_counts['difficult']}\n\n"
            f"Review log rows: {len(self.model.review_log_rows)}\n"
            f"Session log rows: {len(self.model.session_log_rows)}"
        )
        messagebox.showinfo("Study Statistics", text)

    def save_now(self) -> None:
        try:
            self.model.save_user_data()
            messagebox.showinfo("Saved", f"Saved to {self.model.user_file}")
        except Exception as exc:
            messagebox.showerror("Save error", str(exc))

    # ------------------------------------------------------------------
    # Navigation and close
    # ------------------------------------------------------------------
    def next_card(self) -> None:
        if not self.current_deck:
            return
        if self.current_index < len(self.current_deck) - 1:
            self.current_index += 1
            self.show_card()

    def prev_card(self) -> None:
        if not self.current_deck:
            return
        if self.current_index > 0:
            self.current_index -= 1
            self.show_card()

    def on_close(self) -> None:
        try:
            if self.current_deck and self.session_answered > 0:
                # Save a partial session summary without blocking exit.
                duration = int(time.time() - self.session_start_time) if self.session_start_time else 0
                accuracy = round(self.session_correct / self.session_answered * 100, 2) if self.session_answered else 0
                self.model.session_log_rows.append({
                    "session_id": self.session_id,
                    "start_time": _dt.datetime.fromtimestamp(self.session_start_time).replace(microsecond=0).isoformat(sep=" ") if self.session_start_time else "",
                    "end_time": now_iso(),
                    "source": self.session_context.get("source", ""),
                    "group": self.session_context.get("group", ""),
                    "mode": self.session_context.get("mode", ""),
                    "order": self.session_context.get("order", ""),
                    "deck_size": len(self.current_deck),
                    "answered": self.session_answered,
                    "correct": self.session_correct,
                    "wrong": self.session_wrong,
                    "accuracy": accuracy,
                    "new_words": self.session_new,
                    "due_words": self.session_due,
                    "duration_sec": duration,
                })
            self.model.save_user_data()
        except Exception as exc:
            messagebox.showerror("Save error", f"Could not save before exit: {exc}")
        self.destroy()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    model = DataModel(VOCAB_FILE, USER_FILE)
    root = tk.Tk()
    root.withdraw()
    try:
        model.load_vocabulary()
        model.build_global_registry()
        model.load_user_data()
    except Exception as exc:
        messagebox.showerror("Startup error", str(exc))
        root.destroy()
        return
    root.destroy()
    app = GRETrainerApp(model)
    app.mainloop()


if __name__ == "__main__":
    main()
