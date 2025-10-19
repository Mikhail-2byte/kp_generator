from __future__ import annotations

import json
from typing import List

import pytest

from app.services import datasets
from app.services.content_manager import ContentEntry, ContentManager


class StubContentManager(ContentManager):
    def __init__(self):
        super().__init__(actor='tester')
        self._orders: List[ContentEntry] = []
        self._templates: List[ContentEntry] = []
        self._instructions: List[ContentEntry] = []

    def list_orders(self) -> List[ContentEntry]:
        return list(self._orders)

    def list_templates(self) -> List[ContentEntry]:
        return list(self._templates)

    def list_instructions(self) -> List[ContentEntry]:
        return list(self._instructions)

    def save_order_entries(self, entries: List[ContentEntry]) -> None:
        self._orders = list(entries)

    def save_template_entries(self, entries: List[ContentEntry]) -> None:
        self._templates = list(entries)

    def save_instruction_entries(self, entries: List[ContentEntry]) -> None:
        self._instructions = list(entries)


def make_payload(identifier: str | None = None) -> dict:
    return {
        'id': identifier,
        'title': 'Test title',
        'summary': 'Description',
        'files': [{'label': 'DOCX', 'filename': 'file.docx'}],
        'updated_at': '19.10.2025'
    }


@pytest.fixture
def manager():
    return StubContentManager()


def test_add_entry_generates_id(manager):
    entry = manager.add_entry('orders', make_payload(None))
    assert entry.identifier.startswith('orders-')
    assert len(manager.list_orders()) == 1


def test_update_entry_overrides_fields(manager):
    added = manager.add_entry('orders', make_payload('existing'))
    manager.update_entry('orders', 'existing', {
        'title': 'Updated',
        'summary': 'New summary',
        'files': [{'label': 'PDF', 'filename': 'new.pdf'}],
        'updated_at': '20.10.2025'
    })

    stored = manager.list_orders()[0]
    assert stored.title == 'Updated'
    assert stored.summary == 'New summary'
    assert stored.files[0]['filename'] == 'new.pdf'
    assert stored.updated_at == '20.10.2025'


def test_delete_entry(manager):
    manager.add_entry('orders', make_payload('for-delete'))
    assert manager.delete_entry('orders', 'for-delete')
    assert not manager.list_orders()


def test_restore_version(manager, tmp_path, monkeypatch):
    versions_path = tmp_path / 'versions'
    collection_dir = versions_path / 'orders'
    collection_dir.mkdir(parents=True)

    payload = {
        'saved_at': '20251019T120000',
        'actor': 'tester',
        'data': {
            'orders': [
                {
                    'id': 'restored-1',
                    'title': 'Restored',
                    'summary': 'From snapshot',
                    'files': [{'label': 'DOCX', 'filename': 'doc.docx'}],
                    'updated_at': '19.10.2025'
                }
            ]
        }
    }

    version_file = collection_dir / '20251019T120000.json'
    version_file.write_text(json.dumps(payload, ensure_ascii=False), encoding='utf-8')

    monkeypatch.setattr(datasets, 'VERSIONS_DIR', versions_path)
    assert manager.restore_version('orders', '20251019T120000.json')
    restored = manager.list_orders()
    assert len(restored) == 1
    assert restored[0].identifier == 'restored-1'


def test_restore_version_failure(manager):
    assert not manager.restore_version('orders', 'missing.json')
