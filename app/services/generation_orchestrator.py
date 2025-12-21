"""Оркестратор процесса генерации коммерческих предложений."""

from typing import Any, Dict, List, Optional, Tuple

from flask import current_app

from app.business.document_generator import create_zip_archive, generate_excel_document, generate_word_document
from app.core.exceptions import CalculationError, DocumentGenerationError, ValidationError
from app.presentation.helpers import check_templates_exist, extract_positions_from_form
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
    
    def validate_request(self, form_data: Dict[str, Any]) -> Tuple[Dict[str, Any], List[Dict[str, Any]], List[str]]:
        """
        Валидирует данные запроса.
        
        Returns:
            Tuple[cleaned_data, positions, errors]
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

        return form_data, validation.positions or extract_positions_from_form(form_data), validation.errors
    
    def calculate_prices(
        self,
        positions: List[Dict[str, Any]],
        logistics_rub: float,
        delivery_time: int,
        margin_percent: float
    ) -> Tuple[List[Dict[str, Any]], float]:
        """
        Рассчитывает цены для позиций.
        
        Returns:
            Tuple[position_prices, total_general_price]
        """
        import math
        
        if len(positions) == 1:
            # Для одной позиции используем старый метод
            result = self.calculator.calculate_legacy_single_position(
                positions[0], logistics_rub, delivery_time, margin_percent
            )
            position_prices = [result]
            # Округляем final_price вверх до десятков, как в Excel
            final_price_rounded = math.ceil(result['final_price'] / 10.0) * 10.0
            total_general_price = final_price_rounded * result['position']['quantity']
        else:
            # Для множественных позиций используем новый метод с единой маржой
            calculation_result = self.calculator.calculate_multi_position_prices(
                positions, logistics_rub, delivery_time, margin_percent
            )
            position_prices = calculation_result['positions']
            # Пересчитываем total_general_price с округлением final_price до десятков, как в Excel
            total_general_price = 0
            for pos_price in position_prices:
                final_price_rounded = math.ceil(pos_price['final_price'] / 10.0) * 10.0
                quantity = pos_price['position']['quantity']
                total_general_price += final_price_rounded * quantity
        
        # Добавляем округленные final_price в позиции для сохранения в базе данных
        import math
        for i, pos_price in enumerate(position_prices):
            if i < len(positions):
                final_price_rounded = math.ceil(pos_price['final_price'] / 10.0) * 10.0
                positions[i]['final_price'] = final_price_rounded
        
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
        excel_template_path = 'templates_docs/template.xlsx'
        word_template_path = 'templates_docs/template.docx'

        excel_file = generate_excel_document(
            excel_template_path,
            form_data,
            final_price,
            total_general_price,
            position_prices=position_prices,
            manager_fio=manager_fio,
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
        zip_buffer, file_prefix = create_zip_archive(excel_file, word_file, company)
        
        return zip_buffer, file_prefix
    
    def save_history(
        self,
        form_data: Dict[str, Any],
        final_price: float,
        total_general_price: float,
        user_id: Optional[int]
    ) -> bool:
        """Сохраняет историю генерации."""
        saved = generation_repository.save_history(
            form_data, final_price, self.app_config, user_id, total_general_price
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
        cleaned_data, positions, errors = self.validate_request(form_data)
        if errors:
            raise ValidationError(
                'Ошибки валидации данных',
                details={'errors': errors, 'invalid_fields': cleaned_data.get('_invalid_fields', [])}
            )
        
        if not positions:
            raise ValidationError('Не указаны позиции для расчета')
        
        # Шаг 2: Извлечение параметров
        company = cleaned_data['company'].strip()
        logistics_rub = float(cleaned_data['logistics'])
        margin_percent = float(cleaned_data['margin_percent'])
        delivery_time = int(cleaned_data['delivery_time'])
        
        # Шаг 3: Проверка шаблонов
        template_errors = check_templates_exist()
        if template_errors:
            raise DocumentGenerationError(
                'Шаблоны документов не найдены',
                details={'errors': template_errors}
            )
        
        # Шаг 4: Расчет цен
        try:
            position_prices, total_general_price = self.calculate_prices(
                positions, logistics_rub, delivery_time, margin_percent
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
        final_price_nds = total_general_price * 1.2
        
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

