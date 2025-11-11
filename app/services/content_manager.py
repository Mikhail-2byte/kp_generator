from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

from app.services import datasets


@dataclass
class ContentEntry:
    identifier: Optional[str]
    title: str
    summary: str
    files: List[Dict[str, Any]]
    updated_at: Optional[str] = None

    @classmethod
    def from_dict(cls, payload: Dict[str, Any]) -> 'ContentEntry':
        raw_id = payload.get('id') or payload.get('identifier')
        identifier = str(raw_id) if raw_id else None
        return cls(
            identifier=identifier,
            title=str(payload.get('title') or payload.get('name') or ''),
            summary=str(payload.get('summary') or ''),
            files=deepcopy(payload.get('files', [])),
            updated_at=payload.get('updated_at') or payload.get('date'),
        )

    def to_dict(self) -> Dict[str, Any]:
        return {
            'id': self.identifier,
            'title': self.title,
            'summary': self.summary,
            'files': deepcopy(self.files),
            'updated_at': self.updated_at,
        }


class ContentManager:
    def __init__(self, *, actor: Optional[str] = None):
        self.actor = actor
        self._version_keys = {
            'orders': 'orders_documents',
            'orders_documents': 'orders_documents',
            'task_templates': 'task_templates',
            'templates': 'task_templates',
            'instructions': 'instructions_tasks',
            'instructions_tasks': 'instructions_tasks',
        }

    def list_orders(self) -> List[ContentEntry]:
        return [ContentEntry.from_dict(entry) for entry in datasets.get_orders_documents()]

    def list_templates(self) -> List[ContentEntry]:
        return [ContentEntry.from_dict(entry) for entry in datasets.get_task_templates()]

    def list_instructions(self) -> List[ContentEntry]:
        return [ContentEntry.from_dict(entry) for entry in datasets.get_task_instructions()]

    def save_order_entries(self, entries: List[ContentEntry]) -> None:
        datasets.save_orders_documents([entry.to_dict() for entry in entries], actor=self.actor)

    def save_template_entries(self, entries: List[ContentEntry]) -> None:
        datasets.save_task_templates([entry.to_dict() for entry in entries], actor=self.actor)

    def save_instruction_entries(self, entries: List[ContentEntry]) -> None:
        datasets.save_task_instructions([entry.to_dict() for entry in entries], actor=self.actor)

    def add_entry(self, collection: str, payload: Dict[str, Any]) -> ContentEntry:
        entry = ContentEntry.from_dict(payload)
        entry.identifier = entry.identifier or datasets.make_content_id(collection)
        entry.updated_at = payload.get('updated_at') or self._current_date()

        entries = self._get_collection(collection)
        entries.append(entry)
        self._persist(collection, entries)
        return entry

    def update_entry(self, collection: str, identifier: str, payload: Dict[str, Any]) -> Optional[ContentEntry]:
        entries = self._get_collection(collection)
        for index, entry in enumerate(entries):
            if entry.identifier == identifier:
                entries[index] = ContentEntry(
                    identifier=identifier,
                    title=payload.get('title', entry.title),
                    summary=payload.get('summary', entry.summary),
                    files=deepcopy(payload.get('files', entry.files)),
                    updated_at=payload.get('updated_at') or self._current_date(),
                )
                self._persist(collection, entries)
                return entries[index]
        return None

    def delete_entry(self, collection: str, identifier: str) -> bool:
        entries = self._get_collection(collection)
        filtered = [entry for entry in entries if entry.identifier != identifier]
        if len(filtered) == len(entries):
            return False
        self._persist(collection, filtered)
        return True

    def list_versions(self, collection: str) -> List[Dict[str, Any]]:
        key = self._version_keys.get(collection, collection)
        versions = datasets.list_content_versions(key)
        if not versions and key != collection:
            return datasets.list_content_versions(collection)
        return versions

    def load_version(self, collection: str, filename: str) -> Optional[Dict[str, Any]]:
        key = self._version_keys.get(collection, collection)
        version = datasets.load_content_version(key, filename)
        if version is None and key != collection:
            version = datasets.load_content_version(collection, filename)
        return version

    def restore_version(self, collection: str, filename: str) -> bool:
        version = self.load_version(collection, filename)
        if version is None:
            return False

        data = version.get('data')
        if not isinstance(data, dict):
            return False

        entries = self._convert_raw_collection(collection, data)
        self._persist(collection, entries)
        return True

    def _get_collection(self, collection: str) -> List[ContentEntry]:
        if collection == 'orders':
            return self.list_orders()
        if collection == 'task_templates':
            return self.list_templates()
        if collection == 'instructions':
            return self.list_instructions()
        raise ValueError(f'Unsupported collection: {collection}')

    def _persist(self, collection: str, entries: List[ContentEntry]) -> None:
        if collection == 'orders':
            self.save_order_entries(entries)
        elif collection == 'task_templates':
            self.save_template_entries(entries)
        elif collection == 'instructions':
            self.save_instruction_entries(entries)
        else:
            raise ValueError(f'Unsupported collection: {collection}')

    def _convert_raw_collection(self, collection: str, payload: Dict[str, Any]) -> List[ContentEntry]:
        if collection == 'orders':
            entries = payload.get('orders', [])
        elif collection == 'task_templates':
            entries = payload.get('templates', [])
        elif collection == 'instructions':
            entries = payload.get('instructions', [])
        else:
            entries = []

        return [ContentEntry.from_dict(entry) for entry in entries]

    @staticmethod
    def _current_date() -> str:
        return datetime.now(timezone.utc).strftime('%d.%m.%Y')


def build_manager(actor: Optional[str] = None) -> ContentManager:
    manager = ContentManager(actor=actor)
    return manager
