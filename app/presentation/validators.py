from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

from .helpers import extract_positions_from_form


@dataclass
class FormValidationResult:
    """Результат валидации формы расчёта КП."""

    errors: List[str] = field(default_factory=list)
    invalid_fields: List[str] = field(default_factory=list)
    cleaned_data: Dict[str, Any] = field(default_factory=dict)
    positions: List[Dict[str, Any]] = field(default_factory=list)


def _normalize_decimal(value: Any) -> Optional[float]:
    """Преобразует строку с числом в float, поддерживая запятые и пробелы."""

    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    normalized = text.replace(' ', '').replace(',', '.')
    if normalized in {'-', '+', '.', '-.', '+.'}:
        raise ValueError('Invalid numeric format')

    return float(normalized)


def _normalize_integer(value: Any) -> Optional[int]:
    """Преобразует строку в целое число и проверяет целочисленность значения."""

    decimal_value = _normalize_decimal(value)
    if decimal_value is None:
        return None

    if not float(decimal_value).is_integer():
        raise ValueError('Value must be an integer')

    return int(round(decimal_value))


def _format_number_for_form(value: Optional[float], *, as_int: bool = False) -> str:
    """Возвращает строковое представление числа для повторного заполнения формы."""

    if value is None:
        return ''

    if as_int:
        return str(int(value))

    text = f"{float(value):.12f}".rstrip('0').rstrip('.')
    return text if text else '0'


def validate_form_data(form_data: Dict[str, Any]) -> FormValidationResult:
    """Проверяет корректность данных формы и возвращает очищенные значения."""

    cleaned_data: Dict[str, Any] = {
        key: value.strip() if isinstance(value, str) else value
        for key, value in (form_data or {}).items()
    }

    errors: List[str] = []
    invalid_fields: set[str] = set()

    field_labels = {
        'company': 'Название компании',
        'logistics': 'Стоимость логистики',
        'margin_percent': 'Маржа, %',
        'delivery_time': 'Срок поставки (дней)',
        'payment_terms': 'Условия оплаты',
        'proposal_validity': 'Срок действия предложения',
        'warranty_period': 'Гарантийный срок эксплуатации изделия',
        'comment': 'Комментарий',
        'product': 'Наименование товара',
        'drawing_number': 'Номер чертежа',
        'material': 'Материал',
        'cost_price': 'Сумма закупа',
        'cost_price_per_kg': 'Цена за кг',
        'quantity': 'Количество',
        'weight': 'Вес за шт.',
        'duty_percent': 'Пошлина, %'
    }

    positions, position_field_keys = extract_positions_from_form(cleaned_data, include_field_keys=True)

    if not positions:
        errors.append('Необходимо указать хотя бы одну позицию.')
        invalid_fields.update({'product', 'quantity', 'cost_price', 'weight'})
        return FormValidationResult(errors=errors, invalid_fields=sorted(invalid_fields), cleaned_data=cleaned_data)

    company = cleaned_data.get('company', '') or ''
    if not company:
        errors.append('Укажите название компании.')
        invalid_fields.add('company')
    elif len(company) < 2:
        errors.append('Название компании должно содержать минимум 2 символа.')
        invalid_fields.add('company')
    elif len(company) > 200:
        errors.append('Название компании не должно превышать 200 символов.')
        invalid_fields.add('company')
    else:
        cleaned_data['company'] = company

    delivery_address = cleaned_data.get('delivery_address')
    if isinstance(delivery_address, str):
        cleaned_data['delivery_address'] = delivery_address.strip()

    payment_terms = cleaned_data.get('payment_terms', '') or ''
    if payment_terms and len(payment_terms) > 500:
        errors.append('Условия оплаты не должны превышать 500 символов.')
        invalid_fields.add('payment_terms')
    cleaned_data['payment_terms'] = payment_terms

    proposal_validity = cleaned_data.get('proposal_validity', '') or ''
    if proposal_validity and len(proposal_validity) > 200:
        errors.append('Срок действия предложения не должен превышать 200 символов.')
        invalid_fields.add('proposal_validity')
    cleaned_data['proposal_validity'] = proposal_validity

    warranty_period = cleaned_data.get('warranty_period', '') or ''
    if warranty_period and len(warranty_period) > 200:
        errors.append('Гарантийный срок эксплуатации изделия не должен превышать 200 символов.')
        invalid_fields.add('warranty_period')
    cleaned_data['warranty_period'] = warranty_period

    comment = cleaned_data.get('comment', '') or ''
    if comment and len(comment) > 2000:
        errors.append('Комментарий не должен превышать 2000 символов.')
        invalid_fields.add('comment')
    cleaned_data['comment'] = comment

    general_numeric_fields = {
        'logistics': {
            'min': 0,
            'message_min': 'Стоимость логистики должна быть неотрицательной.',
        },
        'margin_percent': {
            'min': 0,
            'max': 100,
            'message_min': 'Маржа должна быть неотрицательной.',
            'message_max': 'Маржа не может превышать 100%.',
        },
        'delivery_time': {
            'min': 1,
            'as_int': True,
            'message_min': 'Срок поставки не может быть меньше 1 дня.',
        },
    }

    for field, rules in general_numeric_fields.items():
        raw_value = cleaned_data.get(field, '')
        label = field_labels.get(field, field)

        if not raw_value:
            errors.append(f'Поле "{label}" является обязательным.')
            invalid_fields.add(field)
            continue

        try:
            if rules.get('as_int'):
                value = _normalize_integer(raw_value)
            else:
                value = _normalize_decimal(raw_value)
        except ValueError:
            errors.append(f'Поле "{label}" должно быть числом.')
            invalid_fields.add(field)
            continue

        if value is None:
            errors.append(f'Поле "{label}" является обязательным.')
            invalid_fields.add(field)
            continue

        if 'min' in rules and value < rules['min']:
            errors.append(rules.get('message_min') or f'Поле "{label}" должно быть не меньше {rules["min"]}.')
            invalid_fields.add(field)

        if 'max' in rules and value > rules['max']:
            errors.append(rules.get('message_max') or f'Поле "{label}" не может превышать {rules["max"]}.')
            invalid_fields.add(field)

        cleaned_data[field] = _format_number_for_form(value, as_int=bool(rules.get('as_int')))

    normalized_positions: List[Dict[str, Any]] = []

    required_position_fields = ['product', 'quantity', 'cost_price', 'weight']
    numeric_position_rules = {
        'cost_price': {'min': 0},
        'cost_price_per_kg': {'min': 0, 'optional': True},
        'weight': {'min': 0},
        'duty_percent': {'min': 0, 'max': 100, 'optional': True},
        'quantity': {'min': 1, 'as_int': True},
    }

    for index, (position, key_map) in enumerate(zip(positions, position_field_keys), start=1):
        normalized_position: Dict[str, Any] = {}

        for field in ['product', 'drawing_number', 'material']:
            value = (position.get(field) or '').strip() if isinstance(position.get(field), str) else position.get(field)
            if field == 'product':
                if not value:
                    errors.append(f'Позиция {index}: укажите наименование товара.')
                    invalid_fields.add(key_map.get(field, field))
                elif len(value) > 200:
                    errors.append(f'Позиция {index}: наименование товара не должно превышать 200 символов.')
                    invalid_fields.add(key_map.get(field, field))
            if value:
                normalized_position[field] = value
                cleaned_data[key_map.get(field, field)] = value

        for field in required_position_fields:
            if field in ['quantity', 'cost_price', 'weight', 'product']:
                key_name = key_map.get(field, field)
                raw_value = cleaned_data.get(key_name, position.get(field))
                if (raw_value is None or str(raw_value).strip() == '') and field != 'product':
                    label = field_labels.get(field, field)
                    errors.append(f'Позиция {index}: поле "{label}" является обязательным.')
                    invalid_fields.add(key_name)

        for field, rules in numeric_position_rules.items():
            key_name = key_map.get(field, field)
            raw_value = cleaned_data.get(key_name, position.get(field))
            label = field_labels.get(field, field)

            if (raw_value is None or str(raw_value).strip() == ''):
                if not rules.get('optional'):
                    errors.append(f'Позиция {index}: поле "{label}" является обязательным.')
                    invalid_fields.add(key_name)
                continue

            try:
                if rules.get('as_int'):
                    value = _normalize_integer(raw_value)
                else:
                    value = _normalize_decimal(raw_value)
            except ValueError:
                errors.append(f'Позиция {index}: поле "{label}" должно быть числом.')
                invalid_fields.add(key_name)
                continue

            if value is None:
                continue

            if 'min' in rules and value < rules['min']:
                message = rules.get('message_min') or f'Позиция {index}: поле "{label}" должно быть не меньше {rules["min"]}.'
                errors.append(message)
                invalid_fields.add(key_name)

            if 'max' in rules and value > rules['max']:
                message = rules.get('message_max') or f'Позиция {index}: поле "{label}" не может превышать {rules["max"]}.'
                errors.append(message)
                invalid_fields.add(key_name)

            normalized_position[field] = value
            cleaned_data[key_name] = _format_number_for_form(value, as_int=bool(rules.get('as_int')))

        if normalized_position:
            normalized_positions.append(normalized_position)

    return FormValidationResult(
        errors=errors,
        invalid_fields=sorted(invalid_fields),
        cleaned_data=cleaned_data,
        positions=normalized_positions,
    )


__all__ = ['FormValidationResult', 'validate_form_data']
