import json
from pathlib import Path
from typing import Any, Dict

from flask import current_app

BASE_DIR = Path(__file__).resolve().parents[2]
FEEDBACK_FILE_PATH = BASE_DIR / 'data' / 'feedback_entries.jsonl'


def save_feedback_entry(entry: Dict[str, Any]) -> bool:
    """Добавляет отзыв пользователя в локальный JSONL-файл."""
    try:
        FEEDBACK_FILE_PATH.parent.mkdir(parents=True, exist_ok=True)
        with FEEDBACK_FILE_PATH.open('a', encoding='utf-8') as feedback_file:
            json.dump(entry, feedback_file, ensure_ascii=False)
            feedback_file.write('\n')
        return True
    except Exception as err:  # pragma: no cover - defensive logging
        logger = getattr(current_app, 'logger', None)
        if logger:
            logger.error('Error saving feedback entry: %s', err)
        return False
