"""Оркестратор процесса генерации коммерческих предложений."""

import math
import os
from typing import Any, Dict, List, Optional, Tuple

from flask import current_app

from app.business.document_generator import create_zip_archive, generate_excel_document, generate_word_document
from app.core.exceptions import CalculationError, DocumentGenerationError, ValidationError
from app.presentation.helpers import check_templates_exist, extract_positions_from_form, extract_position_margins
from app.presentation.validators import validate_form_data
from app.services.multi_position_calculator import MultiPositionCalculator
from app.services.repositories import generation_repository
from app.services.audit_service import log_generation_created


class GenerationOrchestrator:
    """
    Координирует процесс генерации коммерческого предложения.
    
    Оркестратор управляет всеми этапами генерации КП:
    1. Валидация входных данных
    2. Расчет цен для позиций
    3. Генерация документов (Excel, Word)
    4. Сохранение истории генерации
    5. Создание ZIP-архива с документами
    
    Attributes:
        app_config: Конфигурация приложения
        calculator: Калькулятор для расчета цен
    """
    
    def __init__(self, app_config: Dict[str, Any]) -> None:
        """
        Инициализирует оркестратор.
        
        Args:
            app_config: Словарь конфигурации приложения
        """
        self.app_config = app_config
        self.calculator = MultiPositionCalculator(app_config)
    
    def validate_request(
        self, form_data: Dict[str, Any]
    ) -> Tuple[Dict[str, Any], List[Dict[str, Any]], List[str], List[str]]:
        """
        Валидирует данные запроса.

        Returns:
            Tuple[cleaned_data, positions, errors, invalid_fields]
        """
        # Обработка цены за кг
        if not form_data.get('cost_price', '').strip():
            raw_per_kg = (form_data.get('cost_price_per_kg') or '').replace(',', '.').strip()
            raw_weight = (form_data.get('weight') or '').replace(',', '.').strip()
            try:
                per_kg_value = float(raw_per_kg) if raw_per_kg else None
                weight_value = float(raw_weight) if raw_weight else None
            except ValueError:
                per_kg_value = None
                weight_value = None

            if per_kg_value is not None and weight_value is not None and weight_value > 0:
                form_data['cost_price'] = str(per_kg_value * weight_value)

        validation = validate_form_data(form_data)
        form_data = validation.cleaned_data
        invalid_fields = list(validation.invalid_fields) if validation.invalid_fields else []

        return (
            form_data,
            validation.positions or extract_positions_from_form(form_data),
            validation.errors,
            invalid_fields,
        )
    
    def calculate_prices(
        self,
        positions: List[Dict[str, Any]],
        logistics_rub: float,
        delivery_time: int,
        margin_percent: float,
        use_credit: bool = False,
        use_bank_guarantee: bool = False,
        payment_days: Optional[int] = None,
        additional_expenses: Optional[List[Dict[str, Any]]] = None,
        position_margins: Optional[Dict[int, float]] = None
    ) -> Tuple[List[Dict[str, Any]], float]:
        """
        Рассчитывает цены для позиций.
        
        Args:
            positions: Список позиций
            logistics_rub: Стоимость логистики в рублях
            delivery_time: Срок доставки в днях
            margin_percent: Целевая маржа в процентах
            use_credit: Использовать ли кредит в расчете
            use_bank_guarantee: Использовать ли банковскую гарантию в расчете
            payment_days: Количество дней оплаты (для банковской гарантии)
            additional_expenses: Список дополнительных расходов [{'name': str, 'amount': float}]
            position_margins: Словарь индивидуальных марж {position_index: margin_value}
        
        Returns:
            Tuple[position_prices, total_general_price]
        """
        # Вычисляем общую сумму дополнительных расходов в юанях
        # Дополнительные расходы в форме вводятся в юанях
        additional_expenses_total_yuan = 0
        if additional_expenses:
            for expense in additional_expenses:
                if isinstance(expense, dict):
                    amount = expense.get('amount', 0)
                    try:
                        additional_expenses_total_yuan += float(amount)
                    except (ValueError, TypeError):
                        pass
        if len(positions) == 1:
            # Для одной позиции используем старый метод
            # Если указана индивидуальная маржа, используем её
            single_margin = margin_percent
            if position_margins and 0 in position_margins:
                single_margin = position_margins[0]
            
            result = self.calculator.calculate_legacy_single_position(
                positions[0], logistics_rub, delivery_time, single_margin,
                use_credit=use_credit, use_bank_guarantee=use_bank_guarantee,
                payment_days=payment_days,
                additional_expenses_total_yuan=additional_expenses_total_yuan
            )
            position_prices = [result]
            # Используем точное значение цены без округления
            final_price = result['final_price']
            total_general_price = final_price * result['position']['quantity']
        else:
            # Для множественных позиций используем новый метод с единой маржой
            calculation_result = self.calculator.calculate_multi_position_prices(
                positions, logistics_rub, delivery_time, margin_percent,
                use_credit=use_credit, use_bank_guarantee=use_bank_guarantee,
                payment_days=payment_days,
                additional_expenses_total_yuan=additional_expenses_total_yuan,
                position_margins=position_margins
            )
            position_prices = calculation_result['positions']
            # Рассчитываем total_general_price без округления
            total_general_price = 0
            for pos_price in position_prices:
                final_price = pos_price['final_price']
                quantity = pos_price['position']['quantity']
                total_general_price += final_price * quantity
        
        # Добавляем final_price в позиции для сохранения в базе данных (без округления)
        for i, pos_price in enumerate(position_prices):
            if i < len(positions):
                positions[i]['final_price'] = pos_price['final_price']
        
        return position_prices, total_general_price
    
    def generate_documents(
        self,
        form_data: Dict[str, Any],
        positions: List[Dict[str, Any]],
        position_prices: List[Dict[str, Any]],
        final_price: float,
        total_general_price: float,
        final_price_nds: float,
        manager_fio: Optional[str],
        contact_info: Optional[str]
    ) -> Tuple[Any, Any, str]:
        """
        Генерирует Excel и Word документы и создает ZIP архив.
        
        Returns:
            Tuple[zip_buffer, file_prefix]
        """
        root_path = current_app.config.get('PROJECT_ROOT') or current_app.root_path
        excel_template_path = os.path.join(root_path, 'templates_docs', 'template.xlsx')
        word_template_path = os.path.join(root_path, 'templates_docs', 'template.docx')

        excel_file = generate_excel_document(
            excel_template_path,
            form_data,
            final_price,
            total_general_price,
            positions=positions,
            position_prices=position_prices,
            manager_fio=manager_fio,
            config=self.app_config,
        )
        word_file = generate_word_document(
            word_template_path,
            form_data,
            final_price,
            total_general_price,
            final_price_nds,
            positions=positions,
            position_prices=position_prices,
            contact_info=contact_info,
        )

        company = form_data.get('company', 'Unknown')
        tender_number = form_data.get('tender_number')
        margin_percent = form_data.get('margin_percent')

        zip_buffer, file_prefix = create_zip_archive(
            excel_file,
            word_file,
            company,
            tender_number=tender_number,
            margin_percent=margin_percent,
        )

        return zip_buffer, file_prefix
    
    def save_history(
        self,
        form_data: Dict[str, Any],
        final_price: float,
        total_general_price: float,
        user_id: Optional[int]
    ) -> bool:
        """Сохраняет историю генерации."""
        # Округляем цены до целого числа, как в Excel и Word
        rounded_final_price = round(final_price) if final_price is not None else None
        rounded_total_general_price = round(total_general_price) if total_general_price is not None else None
        saved = generation_repository.save_history(
            form_data, rounded_final_price, self.app_config, user_id, rounded_total_general_price
        )
        if saved:
            tender_number = form_data.get('tender_number', '').strip() or None
            company = form_data.get('company', '')
            log_generation_created(
                generation_id=0,  # ID будет неизвестен, но это не критично для логирования
                company=company,
                tender_number=tender_number,
            )
        return saved
    
    def orchestrate(
        self,
        form_data: Dict[str, Any],
        user_id: Optional[int],
        manager_fio: Optional[str] = None,
        contact_info: Optional[str] = None
    ) -> Tuple[Any, str]:
        """
        Основной метод оркестрации процесса генерации.
        
        Returns:
            Tuple[zip_buffer, file_prefix]
        """
        # Шаг 1: Валидация
        cleaned_data, positions, errors, invalid_fields = self.validate_request(form_data)
        if errors:
            raise ValidationError(
                'Ошибки валидации данных',
                details={
                    'errors': errors,
                    'invalid_fields': invalid_fields,
                    'cleaned_data': cleaned_data,
                },
            )
        
        if not positions:
            raise ValidationError('Не указаны позиции для расчета')
        
        # Шаг 2: Извлечение параметров
        company = cleaned_data['company'].strip()
        logistics_rub = float(cleaned_data['logistics'])
        # Для расчетов используем числовое значение маржи,
        # но для имени файла позже сохраним и строковое представление
        try:
            margin_percent = float(cleaned_data.get('margin_percent'))
        except (TypeError, ValueError):
            margin_percent = 0.0
        delivery_time = int(cleaned_data['delivery_time'])
        
        # Извлекаем флаги финансирования
        finance_credit = cleaned_data.get('finance_credit', '')
        finance_bank_guarantee = cleaned_data.get('finance_bank_guarantee', '')
        use_credit = finance_credit and str(finance_credit).lower() in ['1', 'true', 'on', 'yes']
        use_bank_guarantee = finance_bank_guarantee and str(finance_bank_guarantee).lower() in ['1', 'true', 'on', 'yes']
        
        # Извлекаем количество дней оплаты для кредита и банковской гарантии
        # В Excel формула кредита: I28*16%/365*K15, где K15 = I15 + I16 (срок поставки + условия оплаты)
        payment_terms = cleaned_data.get('payment_terms', '').strip()
        payment_days = None
        if (use_credit or use_bank_guarantee) and payment_terms:
            # Пытаемся извлечь число дней из условий оплаты
            # Используем тот же метод, что и в multi_position_processor
            from app.services.multi_position_processor import MultiPositionProcessor
            root_path = current_app.config.get('PROJECT_ROOT') or current_app.root_path
            excel_path = os.path.join(root_path, 'templates_docs', 'template.xlsx')
            processor = MultiPositionProcessor(excel_path)
            payment_days = processor._extract_days_from_payment_terms(payment_terms)
        
        # Шаг 3: Проверка шаблонов
        template_errors = check_templates_exist()
        if template_errors:
            raise DocumentGenerationError(
                'Шаблоны документов не найдены',
                details={'errors': template_errors}
            )
        
        # Шаг 4: Расчет цен
        # Извлекаем дополнительные расходы из cleaned_data
        additional_expenses = cleaned_data.get('additional_expenses', [])
        if not isinstance(additional_expenses, list):
            additional_expenses = []
        
        # Извлекаем индивидуальные маржи позиций
        position_margins = extract_position_margins(cleaned_data, len(positions))
        
        try:
            position_prices, total_general_price = self.calculate_prices(
                positions, logistics_rub, delivery_time, margin_percent,
                use_credit=use_credit, use_bank_guarantee=use_bank_guarantee,
                payment_days=payment_days,
                additional_expenses=additional_expenses,
                position_margins=position_margins if position_margins else None
            )
        except Exception as exc:
            raise CalculationError(
                f'Ошибка при расчете цен: {str(exc)}',
                calculation_type='multi_position' if len(positions) > 1 else 'single_position'
            ) from exc
        
        if not position_prices:
            raise CalculationError('Не удалось рассчитать цены для позиций')
        
        # Шаг 5: Подготовка данных для документов
        first_position = position_prices[0]
        final_price = first_position['final_price']
        # Получаем ставку НДС из конфигурации
        vat_rate = self.app_config.get('calculation_constants', {}).get('vat_rate', 0.22)
        final_price_nds = total_general_price * (1 + vat_rate)
        
        # Обновляем form_data с позициями
        cleaned_data['positions'] = positions
        
        # Шаг 6: Сохранение истории
        self.save_history(cleaned_data, final_price, total_general_price, user_id)
        
        # Шаг 7: Генерация документов
        try:
            zip_buffer, file_prefix = self.generate_documents(
                cleaned_data,
                positions,
                position_prices,
                final_price,
                total_general_price,
                final_price_nds,
                manager_fio,
                contact_info
            )
        except Exception as exc:
            raise DocumentGenerationError(
                f'Ошибка при генерации документов: {str(exc)}',
                document_type='excel_word'
            ) from exc
        
        return zip_buffer, file_prefix

