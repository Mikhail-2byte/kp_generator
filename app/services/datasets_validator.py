from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any, Dict, List

from . import datasets


@dataclass
class DatasetCheckResult:
    name: str
    label: str
    status: str  # ok | warning | error
    records: int
    issues: List[str]

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def _status_for(records: int, issues: List[str]) -> str:
    if issues:
        # Если есть проблемы — как минимум warning
        has_errors = any(issue.lower().startswith("error:") for issue in issues)
        return "error" if has_errors else "warning"
    # Нет проблем — но 0 записей для ключевых справочников считаем предупреждением.
    return "ok" if records > 0 else "warning"


def validate_gb_materials() -> DatasetCheckResult:
    items = datasets.load_gb_materials()
    issues: List[str] = []
    if not items:
        issues.append("warning: gb_materials is empty or file missing.")
    return DatasetCheckResult(
        name="gb_materials",
        label="Аналоги материалов (GB)",
        status=_status_for(len(items), issues),
        records=len(items),
        issues=issues,
    )


def validate_duty_catalog_combined() -> DatasetCheckResult:
    """
    Проверяет объединённый каталог ставок пошлин (ручные записи + ТН ВЭД CSV).
    Это единый справочник, который используется в приложении через get_duty_catalog().
    """
    issues: List[str] = []
    
    # Проверяем ручные записи (duty_rates.json)
    manual_path = datasets.CONFIG_DIR / "duty_rates.json"
    manual_items = []
    if not manual_path.exists():
        issues.append("warning: Файл duty_rates.json не найден (ручные записи недоступны).")
    else:
        manual_items = datasets.load_duty_rates()
        for idx, item in enumerate(manual_items):
            if "duty_percent" not in item:
                issues.append(f"error: Ручная запись {idx} не содержит duty_percent.")
    
    # Проверяем каталог ТН ВЭД (CSV)
    tnved_path: Path | None = datasets._tnved_catalog_path()  # type: ignore[attr-defined]
    tnved_items = []
    if tnved_path is None:
        issues.append("warning: Файл каталога ТН ВЭД (CSV) не найден.")
    else:
        tnved_items = datasets.load_tnved_catalog()
        if not tnved_items:
            issues.append("warning: Каталог ТН ВЭД пуст или нечитаем.")
    
    # Объединённый каталог (как в get_duty_catalog)
    combined_catalog = datasets.get_duty_catalog()
    total_records = len(combined_catalog)
    
    # Определяем статус:
    # - error: если есть критические ошибки (невалидные записи)
    # - warning: если оба источника пусты или один недоступен
    # - ok: если есть хотя бы один источник с данными и нет ошибок
    has_errors = any("error:" in issue.lower() for issue in issues)
    if has_errors:
        final_status = "error"
    elif total_records == 0:
        final_status = "warning"
    elif not manual_items and not tnved_items:
        final_status = "warning"
    else:
        final_status = "ok"
    
    # Добавляем информационные сообщения о составе
    if manual_items and tnved_items:
        issues.append(f"info: Ручных записей: {len(manual_items)}, записей из ТН ВЭД: {len(tnved_items)}.")
    elif manual_items:
        issues.append(f"info: Используются только ручные записи ({len(manual_items)} шт.).")
    elif tnved_items:
        issues.append(f"info: Используется только каталог ТН ВЭД ({len(tnved_items)} шт.).")
    
    return DatasetCheckResult(
        name="duty_catalog_combined",
        label="Ставки пошлин (объединённый каталог)",
        status=final_status,
        records=total_records,
        issues=issues,
    )


def validate_logistics() -> DatasetCheckResult:
    items = datasets.load_all_logistics_cities()
    issues: List[str] = []
    for idx, city in enumerate(items):
        if not city.get("name"):
            issues.append(f"error: city {idx} has empty name.")
        if not city.get("region"):
            issues.append(f"warning: city {idx} has empty region.")
    return DatasetCheckResult(
        name="logistics_cities_combined",
        label="Логистика (города)",
        status=_status_for(len(items), issues),
        records=len(items),
        issues=issues,
    )


def _validate_files_exist(
    entries: List[Dict[str, Any]],
    static_subdir: str,
    dataset_name: str,
) -> DatasetCheckResult:
    static_dir = datasets.BASE_DIR / "static" / static_subdir
    issues: List[str] = []
    for entry_idx, entry in enumerate(entries):
        for file_idx, file_item in enumerate(entry.get("files", [])):
            filename = str(file_item.get("filename", "")).strip()
            if not filename:
                continue
            path = static_dir / filename
            if not path.exists():
                issues.append(
                    f"warning: {dataset_name} entry {entry_idx} file {file_idx} "
                    f"missing in static/{static_subdir}: {filename}"
                )
    return DatasetCheckResult(
        name=dataset_name,
        label={
            "orders_documents": "Распоряжения",
            "task_templates": "Шаблоны задач",
            "instructions_tasks": "Инструкции по задачам",
        }.get(dataset_name, dataset_name),
        status=_status_for(len(entries), issues),
        records=len(entries),
        issues=issues,
    )


def validate_orders_documents() -> DatasetCheckResult:
    entries = datasets.load_orders_documents()
    return _validate_files_exist(entries, "orders", "orders_documents")


def validate_task_templates() -> DatasetCheckResult:
    entries = datasets.load_task_templates()
    return _validate_files_exist(entries, "task_templates", "task_templates")


def validate_task_instructions() -> DatasetCheckResult:
    entries = datasets.load_task_instructions()
    return _validate_files_exist(entries, "instructions", "instructions_tasks")


def run_all_validations() -> List[DatasetCheckResult]:
    return [
        validate_gb_materials(),
        validate_duty_catalog_combined(),  # Объединённый каталог (ручные + ТН ВЭД)
        validate_logistics(),
        validate_orders_documents(),
        validate_task_templates(),
        validate_task_instructions(),
    ]


def summarize_status(results: List[DatasetCheckResult]) -> str:
    has_error = any(r.status == "error" for r in results)
    has_warning = any(r.status == "warning" for r in results)
    if has_error:
        return "error"
    if has_warning:
        return "warning"
    return "ok"


