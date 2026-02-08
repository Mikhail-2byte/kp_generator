from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Union

from flask import Blueprint, Response, abort, flash, jsonify, redirect, render_template, request, url_for

from flask_login import current_user
from werkzeug.security import generate_password_hash

from app.presentation.forms import (
    AdminResetPasswordForm,
    AdminUserDeleteForm,
    AdminUserForm,
    AIAgentCacheForm,
    AIAgentConfigForm,
    DutyDeleteForm,
    DutyItemForm,
    GBMaterialDeleteForm,
    GBMaterialForm,
    LogisticsCityDeleteForm,
    LogisticsCityForm,
    MainCityForm,
    MainCityDeleteForm,
    EkbRfCityForm,
    EkbRfCityDeleteForm,
    TrailCityForm,
    TrailCityDeleteForm,
    TNVEDItemForm,
    TNVEDDeleteForm
)
from app.auth.security import admin_required
from app.services import datasets, datasets_validator
from app.services.audit_service import log_create, log_delete, log_update
from app.services.content_manager import ContentManager, build_manager
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[2]
from app.services.export_service import (
    create_excel_response,
    create_pdf_response,
    export_audit_logs_to_excel,
    export_audit_logs_to_pdf,
)
from app.services.repositories import admin_stats_repository, audit_log_repository, user_repository
from app.presentation.ui import build_context


admin_bp = Blueprint('admin', __name__)  # Управление справочниками через административный интерфейс


@admin_bp.route('/admin/settings', methods=['GET', 'POST'])
@admin_required
def manage_settings() -> Union[str, Response]:
    """Страница настроек системы."""
    from flask import current_app
    
    config = current_app.config.get('APP_SETTINGS', {})
    
    if request.method == 'POST':
        action = request.form.get('action', '')
        if action == 'update_settings':
            # Здесь можно добавить логику обновления настроек
            flash('Настройки обновлены.', 'success')
            return redirect(url_for('admin.manage_settings'))
    
    return render_template(
        'admin/settings.html',
        **build_context(
            'admin_settings',
            'Настройки системы',
            config=config,
        )
    )


@admin_bp.route('/admin', methods=['GET', 'POST'])
@admin_required
def admin_panel() -> Union[str, Response]:
    duty_items = datasets.load_duty_rates()
    gb_materials = datasets.load_gb_materials()
    logistics_cities = datasets.load_logistics_cities()
    user_activity = admin_stats_repository.get_user_activity()

    duty_form = DutyItemForm(prefix='duty')
    gb_form = GBMaterialForm(prefix='gb')
    logistics_form = LogisticsCityForm(prefix='logistics')

    duty_form.action.data = 'add_duty'
    gb_form.action.data = 'add_gb'
    logistics_form.action.data = 'add_city'

    manager = build_manager(actor=_current_actor())
    orders = manager.list_orders()
    templates = manager.list_templates()
    instructions = manager.list_instructions()
    versions_orders = manager.list_versions('orders_documents')
    versions_templates = manager.list_versions('task_templates')
    versions_instructions = manager.list_versions('instructions_tasks')

    if request.method == 'POST':
        action = request.form.get('action', '')

        if action == 'add_duty':
            if duty_form.validate():
                product = duty_form.product.data.strip()
                category = duty_form.category.data.strip()
                duty_percent_value = duty_form.duty_percent.data
                duty_percent = float(duty_percent_value) if duty_percent_value is not None else 0.0

                new_item = {
                    'product': product,
                    'category': category,
                    'duty_percent': duty_percent,
                    'product_search': product.lower(),
                    'category_search': category.lower(),
                    'duty_search': str(duty_percent).lower()
                }
                duty_items.append(new_item)
                datasets.save_duty_rates(duty_items, actor=_current_actor())
                datasets.refresh_duty_rates()
                log_create(
                    resource_type='duty',
                    resource_id=None,
                    description=f'Добавлена позиция пошлины: {product} ({category})',
                    data={'product': product, 'category': category, 'duty_percent': duty_percent},
                )
                flash('Позиция пошлины добавлена.', 'success')
                return redirect(url_for('admin.admin_panel'))
            flash('Исправьте ошибки в разделе пошлин.', 'danger')

        elif action == 'delete_duty':
            delete_form = DutyDeleteForm(formdata=request.form)
            if delete_form.validate():
                try:
                    index = int(delete_form.index.data)
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(duty_items):
                    duty_items.pop(index)
                    datasets.save_duty_rates(duty_items, actor=_current_actor())
                    datasets.refresh_duty_rates()
                    flash('Позиция пошлины удалена.', 'info')
                else:
                    flash('Не удалось найти позицию для удаления.', 'danger')
            else:
                flash('Не удалось подтвердить удаление.', 'danger')
            return redirect(url_for('admin.admin_panel'))

        elif action == 'add_gb':
            if gb_form.validate():
                russian_name = gb_form.russian.data.strip()
                gb_name = gb_form.gb.data.strip()
                gost = (gb_form.gost.data or '').strip() if hasattr(gb_form, 'gost') else ''
                price = (gb_form.price.data or '').strip() if hasattr(gb_form, 'price') else ''
                workpiece_type = (gb_form.workpiece_type.data or '').strip() if hasattr(gb_form, 'workpiece_type') else ''

                new_material = {
                    'russian': russian_name,
                    'gb': gb_name,
                    'notes': '',
                    'gost': gost,
                    'price': price,
                    'workpiece_type': workpiece_type
                }
                gb_materials.append(new_material)
                datasets.save_gb_materials(gb_materials, actor=_current_actor())
                datasets.refresh_gb_analogs()
                flash('Материал добавлен.', 'success')
                return redirect(url_for('admin.admin_panel'))
            flash('Исправьте ошибки в разделе аналогов.', 'danger')

        elif action == 'delete_gb':
            delete_form = GBMaterialDeleteForm(formdata=request.form)
            if delete_form.validate():
                try:
                    index = int(delete_form.index.data)
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(gb_materials):
                    gb_materials.pop(index)
                    datasets.save_gb_materials(gb_materials, actor=_current_actor())
                    datasets.refresh_gb_analogs()
                    flash('Материал удалён.', 'info')
                else:
                    flash('Не удалось найти материал для удаления.', 'danger')
            else:
                flash('Не удалось подтвердить удаление.', 'danger')
            return redirect(url_for('admin.admin_panel'))

        elif action == 'add_city':
            if logistics_form.validate():
                name = logistics_form.name.data.strip()
                region = (logistics_form.region.data or '').strip()
                truck_price_value = logistics_form.truck_price.data
                trail_price_value = logistics_form.trail_price.data
                truck_price = float(truck_price_value) if truck_price_value is not None else 0.0
                trail_price = float(trail_price_value) if trail_price_value is not None else 0.0

                logistics_cities.append({
                    'name': name,
                    'region': region,
                    'truck_price': truck_price,
                    'trail_price': trail_price
                })
                datasets.save_logistics_cities(logistics_cities, actor=_current_actor())
                flash('Город добавлен.', 'success')
                return redirect(url_for('admin.admin_panel'))
            flash('Исправьте ошибки в разделе логистики.', 'danger')

        elif action == 'delete_city':
            delete_form = LogisticsCityDeleteForm(formdata=request.form)
            if delete_form.validate():
                try:
                    index = int(delete_form.index.data)
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(logistics_cities):
                    logistics_cities.pop(index)
                    datasets.save_logistics_cities(logistics_cities, actor=_current_actor())
                    flash('Город удалён.', 'info')
                else:
                    flash('Не удалось найти город для удаления.', 'danger')
            else:
                flash('Не удалось подтвердить удаление.', 'danger')
            return redirect(url_for('admin.admin_panel'))
        elif action == 'refresh_templates':
            datasets.refresh_task_templates()
            flash('Кэш шаблонов обновлён.', 'info')
            return redirect(url_for('admin.admin_panel'))
        elif action.startswith('content:'):
            return _handle_content_action(action, manager, 'admin.admin_panel')
        else:
            flash('Неизвестное действие.', 'danger')
            return redirect(url_for('admin.admin_panel'))

    # Сводка по справочникам для health-виджета
    datasets_health = [
        result.to_dict() for result in datasets_validator.run_all_validations()
    ]

    # Статус AI агента
    try:
        from ai_agent.config import get_api_key
        from ai_agent.usage_monitor import get_api_key_status_from_db
        
        ai_agent_status = {
            'configured': bool(get_api_key()),
            'recent_error': get_api_key_status_from_db()
        }
    except Exception:
        ai_agent_status = {'configured': False, 'recent_error': None}

    return render_template(
        'admin.html',
        **build_context(
            'admin',
            'Администрирование',
            datasets_health=datasets_health,
            ai_agent_status=ai_agent_status,
        )
    )


@admin_bp.route('/admin/duty', methods=['GET', 'POST'])
@admin_required
def manage_duty() -> Union[str, Response]:
    # Загружаем все пошлины из единого файла
    all_items = datasets.load_duty_rates()
    
    # Разделяем на простые записи и ТН ВЭД
    duty_items = [item for item in all_items if not item.get('code')]
    tnved_items = [item for item in all_items if item.get('code')]
    
    duty_form = DutyItemForm(prefix='duty')
    duty_form.action.data = 'add_duty'
    
    tnved_form = TNVEDItemForm(prefix='tnved')
    tnved_form.action.data = 'add_tnved'

    if request.method == 'POST':
        action = request.form.get('action', '')

        if action == 'add_duty':
            if duty_form.validate():
                product = duty_form.product.data.strip()
                category = duty_form.category.data.strip()
                duty_percent_value = duty_form.duty_percent.data
                duty_percent = float(duty_percent_value) if duty_percent_value is not None else 0.0

                duty_items.append({
                    'product': product,
                    'category': category,
                    'duty_percent': duty_percent,
                    'product_search': product.lower(),
                    'category_search': category.lower(),
                    'duty_search': str(duty_percent).lower()
                })
                datasets.save_duty_rates(duty_items, actor=_current_actor())
                datasets.refresh_duty_rates()
                flash('Позиция пошлины добавлена.', 'success')
                return redirect(url_for('admin.manage_duty'))
            flash('Исправьте ошибки в форме.', 'danger')

        elif action == 'edit_duty':
            if duty_form.validate():
                try:
                    index = int(request.form.get('index', '-1'))
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(duty_items):
                    product = duty_form.product.data.strip()
                    category = duty_form.category.data.strip()
                    duty_percent_value = duty_form.duty_percent.data
                    duty_percent = float(duty_percent_value) if duty_percent_value is not None else 0.0

                    duty_items[index] = {
                        'product': product,
                        'category': category,
                        'duty_percent': duty_percent,
                        'product_search': product.lower(),
                        'category_search': category.lower(),
                        'duty_search': str(duty_percent).lower()
                    }
                    old_item = duty_items[index].copy()
                    datasets.save_duty_rates(duty_items, actor=_current_actor())
                    datasets.refresh_duty_rates()
                    log_update(
                        resource_type='duty',
                        resource_id=str(index),
                        description=f'Обновлена позиция пошлины: {product}',
                        data_before=old_item,
                        data_after={'product': product, 'category': category, 'duty_percent': duty_percent},
                    )
                    flash('Позиция пошлины обновлена.', 'success')
                else:
                    flash('Не удалось найти позицию для редактирования.', 'danger')
            else:
                flash('Исправьте ошибки в форме.', 'danger')
            return redirect(url_for('admin.manage_duty'))

        elif action == 'delete_duty':
            delete_form = DutyDeleteForm(formdata=request.form)
            if delete_form.validate():
                try:
                    index = int(delete_form.index.data)
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(duty_items):
                    duty_items.pop(index)
                    datasets.save_duty_rates(duty_items, actor=_current_actor())
                    datasets.refresh_duty_rates()
                    flash('Позиция пошлины удалена.', 'info')
                else:
                    flash('Не удалось найти позицию для удаления.', 'danger')
            else:
                flash('Не удалось подтвердить удаление.', 'danger')
            return redirect(url_for('admin.manage_duty'))

        elif action == 'add_tnved':
            if tnved_form.validate():
                code = tnved_form.code.data.strip()
                description = tnved_form.description.data.strip()
                keywords_display = tnved_form.keywords_display.data.strip() if tnved_form.keywords_display.data else ''
                examples = tnved_form.examples.data.strip() if tnved_form.examples.data else ''
                duty_text = tnved_form.duty_text.data.strip() if tnved_form.duty_text.data else ''
                duty_percent_value = tnved_form.duty_percent.data
                duty_percent = float(duty_percent_value) if duty_percent_value is not None else None

                # Объединяем все записи и сохраняем в единый файл
                all_items = duty_items + tnved_items
                all_items.append({
                    'code': code,
                    'description': description,
                    'keywords_display': keywords_display,
                    'examples': examples,
                    'duty_text': duty_text,
                    'duty_percent': duty_percent
                })
                datasets.save_duty_rates(all_items, actor=_current_actor())
                datasets.refresh_duty_rates()
                datasets.refresh_tnved_catalog()
                flash('Запись каталога ТН ВЭД добавлена.', 'success')
                return redirect(url_for('admin.manage_duty'))
            flash('Исправьте ошибки в форме.', 'danger')

        elif action == 'edit_tnved':
            if tnved_form.validate():
                try:
                    tnved_index = int(request.form.get('index', '-1'))
                except (TypeError, ValueError):
                    tnved_index = -1
                if 0 <= tnved_index < len(tnved_items):
                    code = tnved_form.code.data.strip()
                    description = tnved_form.description.data.strip()
                    keywords_display = tnved_form.keywords_display.data.strip() if tnved_form.keywords_display.data else ''
                    examples = tnved_form.examples.data.strip() if tnved_form.examples.data else ''
                    duty_text = tnved_form.duty_text.data.strip() if tnved_form.duty_text.data else ''
                    duty_percent_value = tnved_form.duty_percent.data
                    duty_percent = float(duty_percent_value) if duty_percent_value is not None else None

                    # Объединяем все записи и сохраняем в единый файл
                    all_items = duty_items + tnved_items
                    # Индекс в объединенном списке = индекс простых записей + индекс в списке ТН ВЭД
                    actual_index = len(duty_items) + tnved_index
                    old_item = all_items[actual_index].copy()
                    all_items[actual_index] = {
                        'code': code,
                        'description': description,
                        'keywords_display': keywords_display,
                        'examples': examples,
                        'duty_text': duty_text,
                        'duty_percent': duty_percent
                    }
                    datasets.save_duty_rates(all_items, actor=_current_actor())
                    datasets.refresh_duty_rates()
                    datasets.refresh_tnved_catalog()
                    log_update(
                        resource_type='tnved',
                        resource_id=str(tnved_index),
                        description=f'Обновлена запись каталога ТН ВЭД: {code}',
                        data_before=old_item,
                        data_after={'code': code, 'description': description, 'duty_percent': duty_percent},
                    )
                    flash('Запись каталога ТН ВЭД обновлена.', 'success')
                else:
                    flash('Не удалось найти запись для редактирования.', 'danger')
            else:
                flash('Исправьте ошибки в форме.', 'danger')
            return redirect(url_for('admin.manage_duty'))

        elif action == 'delete_tnved':
            delete_form = TNVEDDeleteForm(formdata=request.form)
            if delete_form.validate():
                try:
                    tnved_index = int(delete_form.index.data)
                except (TypeError, ValueError):
                    tnved_index = -1
                if 0 <= tnved_index < len(tnved_items):
                    # Объединяем все записи, удаляем нужную и сохраняем
                    all_items = duty_items + tnved_items
                    # Индекс в объединенном списке = индекс простых записей + индекс в списке ТН ВЭД
                    actual_index = len(duty_items) + tnved_index
                    all_items.pop(actual_index)
                    datasets.save_duty_rates(all_items, actor=_current_actor())
                    datasets.refresh_duty_rates()
                    datasets.refresh_tnved_catalog()
                    flash('Запись каталога ТН ВЭД удалена.', 'info')
                else:
                    flash('Не удалось найти запись для удаления.', 'danger')
            else:
                flash('Не удалось подтвердить удаление.', 'danger')
            return redirect(url_for('admin.manage_duty'))

        else:
            flash('Неизвестное действие.', 'danger')
            return redirect(url_for('admin.manage_duty'))

    return render_template(
        'admin/duty.html',
        **build_context(
            'admin_duty',
            'Ставки пошлин',
            duty_items=duty_items,
            duty_form=duty_form,
            tnved_items=tnved_items,
            tnved_form=tnved_form,
        )
    )


@admin_bp.route('/admin/duty/import', methods=['POST'])
@admin_required
def import_duty() -> Response:
    """Импортирует ставки пошлин из загруженного Excel файла."""
    import tempfile
    import os
    from werkzeug.utils import secure_filename
    
    # Проверяем наличие файла в запросе
    if 'excel_file' not in request.files:
        flash('Файл не выбран.', 'danger')
        return redirect(url_for('admin.manage_duty'))
    
    file = request.files['excel_file']
    
    # Проверяем, что файл выбран
    if file.filename == '':
        flash('Файл не выбран.', 'danger')
        return redirect(url_for('admin.manage_duty'))
    
    # Проверяем расширение файла
    if not (file.filename.lower().endswith('.xlsx') or file.filename.lower().endswith('.xls')):
        flash('Неверный формат файла. Требуется Excel (.xlsx или .xls).', 'danger')
        return redirect(url_for('admin.manage_duty'))
    
    temp_path = None
    try:
        # Сохраняем файл во временную директорию
        temp_dir = tempfile.gettempdir()
        filename = secure_filename(file.filename)
        temp_path = Path(temp_dir) / f'temp_import_{datetime.now().strftime("%Y%m%d%H%M%S")}_{filename}'
        file.save(str(temp_path))
        
        # Импортируем данные
        count = datasets.import_duty_rates_from_excel(temp_path, actor=_current_actor())
        
        flash(f'Импортировано ставок пошлин: {count}.', 'success')
    except Exception as exc:
        flash(f'Ошибка при импорте: {str(exc)}', 'danger')
    finally:
        # Гарантированно удаляем временный файл после закрытия
        if temp_path is not None and temp_path.exists():
            try:
                # Небольшая задержка для Windows, чтобы файл точно закрылся
                import time
                time.sleep(0.1)
                os.remove(temp_path)
            except (PermissionError, OSError):
                # Если не удалось удалить сразу, пытаемся через некоторое время
                # В реальном приложении можно добавить задачу на отложенное удаление
                pass
    
    return redirect(url_for('admin.manage_duty'))


@admin_bp.route('/admin/duty/export')
@admin_required
def export_duty() -> Response:
    """Экспортирует ставки пошлин в Excel файл."""
    from flask import Response
    from datetime import datetime
    
    try:
        excel_data = datasets.export_duty_rates_to_excel()
        # Используем ASCII имя файла для совместимости с HTTP заголовками
        filename = f'Duty_rates_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        
        return Response(
            excel_data,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={'Content-Disposition': f'attachment; filename="{filename}"'}
        )
    except Exception as exc:
        flash(f'Ошибка при экспорте: {str(exc)}', 'danger')
        return redirect(url_for('admin.manage_duty'))


@admin_bp.route('/admin/duty/tnved/import', methods=['POST'])
@admin_required
def import_tnved() -> Response:
    """Импортирует каталог ТН ВЭД из загруженного Excel файла."""
    import tempfile
    import os
    from werkzeug.utils import secure_filename
    
    # Проверяем наличие файла в запросе
    if 'excel_file' not in request.files:
        flash('Файл не выбран.', 'danger')
        return redirect(url_for('admin.manage_duty'))
    
    file = request.files['excel_file']
    
    # Проверяем, что файл выбран
    if file.filename == '':
        flash('Файл не выбран.', 'danger')
        return redirect(url_for('admin.manage_duty'))
    
    # Проверяем расширение файла
    if not (file.filename.lower().endswith('.xlsx') or file.filename.lower().endswith('.xls')):
        flash('Неверный формат файла. Требуется Excel (.xlsx или .xls).', 'danger')
        return redirect(url_for('admin.manage_duty'))
    
    temp_path = None
    try:
        # Сохраняем файл во временную директорию
        temp_dir = tempfile.gettempdir()
        filename = secure_filename(file.filename)
        temp_path = Path(temp_dir) / f'temp_import_{datetime.now().strftime("%Y%m%d%H%M%S")}_{filename}'
        file.save(str(temp_path))
        
        # Импортируем данные
        count = datasets.import_tnved_catalog_from_excel(temp_path, actor=_current_actor())
        
        flash(f'Импортировано записей каталога ТН ВЭД: {count}.', 'success')
    except Exception as exc:
        flash(f'Ошибка при импорте: {str(exc)}', 'danger')
    finally:
        # Гарантированно удаляем временный файл после закрытия
        if temp_path is not None and temp_path.exists():
            try:
                # Небольшая задержка для Windows, чтобы файл точно закрылся
                import time
                time.sleep(0.1)
                os.remove(temp_path)
            except (PermissionError, OSError):
                # Если не удалось удалить сразу, пытаемся через некоторое время
                # В реальном приложении можно добавить задачу на отложенное удаление
                pass
    
    return redirect(url_for('admin.manage_duty'))


@admin_bp.route('/admin/duty/tnved/export')
@admin_required
def export_tnved() -> Response:
    """Экспортирует каталог ТН ВЭД в Excel файл."""
    from flask import Response
    from datetime import datetime
    
    try:
        excel_data = datasets.export_tnved_catalog_to_excel()
        # Используем ASCII имя файла для совместимости с HTTP заголовками
        filename = f'TNVED_catalog_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        
        return Response(
            excel_data,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={'Content-Disposition': f'attachment; filename="{filename}"'}
        )
    except Exception as exc:
        flash(f'Ошибка при экспорте: {str(exc)}', 'danger')
        return redirect(url_for('admin.manage_duty'))


@admin_bp.route('/admin/materials', methods=['GET', 'POST'])
@admin_required
def manage_materials() -> Union[str, Response]:
    gb_materials = datasets.load_gb_materials()
    gb_form = GBMaterialForm(prefix='gb')
    gb_form.action.data = 'add_gb'

    if request.method == 'POST':
        action = request.form.get('action', '')

        if action == 'add_gb':
            if gb_form.validate():
                russian_name = gb_form.russian.data.strip()
                gb_name = gb_form.gb.data.strip()
                gost = (gb_form.gost.data or '').strip() if hasattr(gb_form, 'gost') else ''
                price = (gb_form.price.data or '').strip() if hasattr(gb_form, 'price') else ''
                workpiece_type = (gb_form.workpiece_type.data or '').strip() if hasattr(gb_form, 'workpiece_type') else ''

                gb_materials.append({
                    'russian': russian_name,
                    'gb': gb_name,
                    'notes': '',
                    'gost': gost,
                    'price': price,
                    'workpiece_type': workpiece_type
                })
                datasets.save_gb_materials(gb_materials, actor=_current_actor())
                datasets.refresh_gb_analogs()
                flash('Материал добавлен.', 'success')
                return redirect(url_for('admin.manage_materials'))
            flash('Исправьте ошибки в форме.', 'danger')

        elif action == 'edit_gb':
            if gb_form.validate():
                try:
                    index = int(request.form.get('index', '-1'))
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(gb_materials):
                    russian_name = gb_form.russian.data.strip()
                    gb_name = gb_form.gb.data.strip()
                    gost = (gb_form.gost.data or '').strip() if hasattr(gb_form, 'gost') else ''
                    price = (gb_form.price.data or '').strip() if hasattr(gb_form, 'price') else ''
                    workpiece_type = (gb_form.workpiece_type.data or '').strip() if hasattr(gb_form, 'workpiece_type') else ''

                    gb_materials[index] = {
                        'russian': russian_name,
                        'gb': gb_name,
                        'notes': '',
                        'gost': gost,
                        'price': price,
                        'workpiece_type': workpiece_type
                    }
                    datasets.save_gb_materials(gb_materials, actor=_current_actor())
                    datasets.refresh_gb_analogs()
                    flash('Материал обновлён.', 'success')
                else:
                    flash('Не удалось найти материал для редактирования.', 'danger')
            else:
                flash('Исправьте ошибки в форме.', 'danger')
            return redirect(url_for('admin.manage_materials'))

        elif action == 'delete_gb':
            delete_form = GBMaterialDeleteForm(formdata=request.form)
            if delete_form.validate():
                try:
                    index = int(delete_form.index.data)
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(gb_materials):
                    gb_materials.pop(index)
                    datasets.save_gb_materials(gb_materials, actor=_current_actor())
                    datasets.refresh_gb_analogs()
                    flash('Материал удалён.', 'info')
                else:
                    flash('Не удалось найти материал для удаления.', 'danger')
            else:
                flash('Не удалось подтвердить удаление.', 'danger')
            return redirect(url_for('admin.manage_materials'))

        else:
            flash('Неизвестное действие.', 'danger')
            return redirect(url_for('admin.manage_materials'))

    return render_template(
        'admin/materials.html',
        **build_context(
            'admin_materials',
            'Аналоги материалов',
            gb_materials=gb_materials,
            gb_form=gb_form,
        )
    )


@admin_bp.route('/admin/materials/import', methods=['POST'])
@admin_required
def import_materials() -> Response:
    """Импортирует материалы из загруженного Excel файла."""
    import tempfile
    import os
    from werkzeug.utils import secure_filename
    
    # Проверяем наличие файла в запросе
    if 'excel_file' not in request.files:
        flash('Файл не выбран.', 'danger')
        return redirect(url_for('admin.manage_materials'))
    
    file = request.files['excel_file']
    
    # Проверяем, что файл выбран
    if file.filename == '':
        flash('Файл не выбран.', 'danger')
        return redirect(url_for('admin.manage_materials'))
    
    # Проверяем расширение файла
    if not (file.filename.lower().endswith('.xlsx') or file.filename.lower().endswith('.xls')):
        flash('Неверный формат файла. Требуется Excel (.xlsx или .xls).', 'danger')
        return redirect(url_for('admin.manage_materials'))
    
    try:
        # Сохраняем файл во временную директорию
        temp_dir = tempfile.gettempdir()
        filename = secure_filename(file.filename)
        temp_path = Path(temp_dir) / f'temp_import_{datetime.now().strftime("%Y%m%d%H%M%S")}_{filename}'
        file.save(str(temp_path))
        
        # Импортируем данные
        count = datasets.import_gb_materials_from_excel(temp_path, actor=_current_actor())
        
        # Удаляем временный файл
        if temp_path.exists():
            os.remove(temp_path)
        
        flash(f'Импортировано материалов: {count}.', 'success')
    except Exception as exc:
        flash(f'Ошибка при импорте: {str(exc)}', 'danger')
        # Удаляем временный файл в случае ошибки
        if 'temp_path' in locals() and temp_path.exists():
            os.remove(temp_path)
    
    return redirect(url_for('admin.manage_materials'))


@admin_bp.route('/admin/materials/export')
@admin_required
def export_materials() -> Response:
    """Экспортирует материалы в Excel файл."""
    from flask import Response
    from datetime import datetime
    
    try:
        excel_data = datasets.export_gb_materials_to_excel()
        # Используем ASCII имя файла для совместимости с HTTP заголовками
        filename = f'Steel_prices_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        
        return Response(
            excel_data,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={'Content-Disposition': f'attachment; filename="{filename}"'}
        )
    except Exception as exc:
        flash(f'Ошибка при экспорте: {str(exc)}', 'danger')
        return redirect(url_for('admin.manage_materials'))


@admin_bp.route('/admin/logistics', methods=['GET', 'POST'])
@admin_required
def manage_logistics() -> Union[str, Response]:
    """Управление справочниками логистики (3 справочника: основные, ЕКБ+РФ, трал)."""
    active_tab = request.args.get('tab', 'main')
    
    # Загружаем данные для всех справочников
    main_cities = datasets.load_main_cities()
    ekb_rf_cities = datasets.load_ekb_rf_cities()
    trail_cities = datasets.load_trail_cities()
    
    # Группируем основные города по разделам
    main_cities_groups = []
    # Словарь для быстрого доступа к группам по имени основного города
    groups_dict = {}
    
    # Сначала создаем группы для основных городов
    for idx, city in enumerate(main_cities):
        if city.get('is_main_route'):
            main_city_name = city.get('name', '')
            groups_dict[main_city_name] = {
                'main_city_name': main_city_name,
                'main_city_index': idx,
                'main_city_data': city,
                'related_cities': []
            }
    
    # Теперь добавляем города в соответствующие группы
    for idx, city in enumerate(main_cities):
        if not city.get('is_main_route'):
            main_city_name = city.get('main_city')
            if main_city_name and main_city_name in groups_dict:
                groups_dict[main_city_name]['related_cities'].append({
                    'index': idx,
                    'data': city
                })
    
    # Порядок отображения основных городов
    MAIN_CITIES_DISPLAY_ORDER = [
        'Чита',
        'Улан-Удэ',
        'Иркутск',
        'Красноярск',
        'Новосибирск',
        'Омск',
        'Екатеринбург',
        'Москва',
        'Санкт-Петербург',
        'Минск'
    ]
    
    # Преобразуем словарь в список и сортируем по заданному порядку
    def get_city_order(city_group):
        city_name = city_group['main_city_name']
        if city_name in MAIN_CITIES_DISPLAY_ORDER:
            return MAIN_CITIES_DISPLAY_ORDER.index(city_name)
        # Если город не в списке, добавляем в конец
        return len(MAIN_CITIES_DISPLAY_ORDER)
    
    main_cities_groups = sorted(groups_dict.values(), key=get_city_order)
    
    # Инициализируем формы
    main_form = MainCityForm(prefix='main')
    main_form.action.data = 'add_main_city'
    ekb_rf_form = EkbRfCityForm(prefix='ekb_rf')
    ekb_rf_form.action.data = 'add_ekb_rf_city'
    trail_form = TrailCityForm(prefix='trail')
    trail_form.action.data = 'add_trail_city'

    if request.method == 'POST':
        action = request.form.get('action', '')
        tab = request.form.get('tab', 'main')

        # Обработка основных городов
        if action == 'add_main_city':
            if main_form.validate():
                name = main_form.name.data.strip()
                region = (main_form.region.data or '').strip()
                truck_price = float(main_form.truck_price.data) if main_form.truck_price.data is not None else 0.0
                is_main_route = bool(main_form.is_main_route.data)
                main_city = (main_form.main_city.data or '').strip() or None

                main_cities.append({
                    'name': name,
                    'region': region,
                    'truck_price': truck_price,
                    'is_main_route': is_main_route,
                    'main_city': main_city
                })
                datasets.save_main_cities(main_cities, actor=_current_actor())
                flash('Город добавлен в справочник основных городов.', 'success')
                return redirect(url_for('admin.manage_logistics', tab='main'))
            flash('Исправьте ошибки в форме.', 'danger')

        elif action == 'edit_main_city':
            if main_form.validate():
                try:
                    index = int(request.form.get('index', '-1'))
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(main_cities):
                    name = main_form.name.data.strip()
                    region = (main_form.region.data or '').strip()
                    truck_price = float(main_form.truck_price.data) if main_form.truck_price.data is not None else 0.0
                    is_main_route = bool(main_form.is_main_route.data)
                    main_city = (main_form.main_city.data or '').strip() or None

                    main_cities[index] = {
                        'name': name,
                        'region': region,
                        'truck_price': truck_price,
                        'is_main_route': is_main_route,
                        'main_city': main_city
                    }
                    datasets.save_main_cities(main_cities, actor=_current_actor())
                    flash('Город обновлён.', 'success')
                else:
                    flash('Не удалось найти город для редактирования.', 'danger')
            else:
                flash('Исправьте ошибки в форме.', 'danger')
            return redirect(url_for('admin.manage_logistics', tab='main'))

        elif action == 'delete_main_city':
            delete_form = MainCityDeleteForm(formdata=request.form)
            if delete_form.validate():
                try:
                    index = int(delete_form.index.data)
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(main_cities):
                    main_cities.pop(index)
                    datasets.save_main_cities(main_cities, actor=_current_actor())
                    flash('Город удалён.', 'info')
                else:
                    flash('Не удалось найти город для удаления.', 'danger')
            else:
                flash('Не удалось подтвердить удаление.', 'danger')
            return redirect(url_for('admin.manage_logistics', tab='main'))

        # Обработка городов ЕКБ+РФ
        elif action == 'add_ekb_rf_city':
            if ekb_rf_form.validate():
                name = ekb_rf_form.name.data.strip()
                region = (ekb_rf_form.region.data or '').strip()
                distance_from_ekb_km = ekb_rf_form.distance_from_ekb_km.data

                ekb_rf_cities.append({
                    'name': name,
                    'region': region,
                    'distance_from_ekb_km': distance_from_ekb_km
                })
                datasets.save_ekb_rf_cities(ekb_rf_cities, actor=_current_actor())
                flash('Город добавлен в справочник ЕКБ+РФ.', 'success')
                return redirect(url_for('admin.manage_logistics', tab='ekb_rf'))
            flash('Исправьте ошибки в форме.', 'danger')

        elif action == 'edit_ekb_rf_city':
            if ekb_rf_form.validate():
                try:
                    index = int(request.form.get('index', '-1'))
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(ekb_rf_cities):
                    name = ekb_rf_form.name.data.strip()
                    region = (ekb_rf_form.region.data or '').strip()
                    distance_from_ekb_km = ekb_rf_form.distance_from_ekb_km.data

                    ekb_rf_cities[index] = {
                        'name': name,
                        'region': region,
                        'distance_from_ekb_km': distance_from_ekb_km
                    }
                    datasets.save_ekb_rf_cities(ekb_rf_cities, actor=_current_actor())
                    flash('Город обновлён.', 'success')
                else:
                    flash('Не удалось найти город для редактирования.', 'danger')
            else:
                flash('Исправьте ошибки в форме.', 'danger')
            return redirect(url_for('admin.manage_logistics', tab='ekb_rf'))

        elif action == 'delete_ekb_rf_city':
            delete_form = EkbRfCityDeleteForm(formdata=request.form)
            if delete_form.validate():
                try:
                    index = int(delete_form.index.data)
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(ekb_rf_cities):
                    ekb_rf_cities.pop(index)
                    datasets.save_ekb_rf_cities(ekb_rf_cities, actor=_current_actor())
                    flash('Город удалён.', 'info')
                else:
                    flash('Не удалось найти город для удаления.', 'danger')
            else:
                flash('Не удалось подтвердить удаление.', 'danger')
            return redirect(url_for('admin.manage_logistics', tab='ekb_rf'))

        # Обработка городов трала
        elif action == 'add_trail_city':
            if trail_form.validate():
                name = trail_form.name.data.strip()
                region = (trail_form.region.data or '').strip()
                trail_price = float(trail_form.trail_price.data) if trail_form.trail_price.data is not None else 0.0

                trail_cities.append({
                    'name': name,
                    'region': region,
                    'trail_price': trail_price
                })
                datasets.save_trail_cities(trail_cities, actor=_current_actor())
                flash('Город добавлен в справочник трала.', 'success')
                return redirect(url_for('admin.manage_logistics', tab='trail'))
            flash('Исправьте ошибки в форме.', 'danger')

        elif action == 'edit_trail_city':
            if trail_form.validate():
                try:
                    index = int(request.form.get('index', '-1'))
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(trail_cities):
                    name = trail_form.name.data.strip()
                    region = (trail_form.region.data or '').strip()
                    trail_price = float(trail_form.trail_price.data) if trail_form.trail_price.data is not None else 0.0

                    trail_cities[index] = {
                        'name': name,
                        'region': region,
                        'trail_price': trail_price
                    }
                    datasets.save_trail_cities(trail_cities, actor=_current_actor())
                    flash('Город обновлён.', 'success')
                else:
                    flash('Не удалось найти город для редактирования.', 'danger')
            else:
                flash('Исправьте ошибки в форме.', 'danger')
            return redirect(url_for('admin.manage_logistics', tab='trail'))

        elif action == 'delete_trail_city':
            delete_form = TrailCityDeleteForm(formdata=request.form)
            if delete_form.validate():
                try:
                    index = int(delete_form.index.data)
                except (TypeError, ValueError):
                    index = -1
                if 0 <= index < len(trail_cities):
                    trail_cities.pop(index)
                    datasets.save_trail_cities(trail_cities, actor=_current_actor())
                    flash('Город удалён.', 'info')
                else:
                    flash('Не удалось найти город для удаления.', 'danger')
            else:
                flash('Не удалось подтвердить удаление.', 'danger')
            return redirect(url_for('admin.manage_logistics', tab='trail'))

        else:
            flash('Неизвестное действие.', 'danger')
            return redirect(url_for('admin.manage_logistics', tab=tab))

    # Рендерим страницу
    return render_template('admin/logistics.html',
                         active_tab=active_tab,
                         main_form=main_form,
                         ekb_rf_form=ekb_rf_form,
                         trail_form=trail_form,
                         main_cities_groups=main_cities_groups,
                         ekb_rf_cities=ekb_rf_cities,
                         trail_cities=trail_cities)


@admin_bp.route('/admin/logistics/main/export')
@admin_required
def export_main_cities() -> Response:
    """Экспортирует основные города в Excel файл."""
    from flask import Response
    from datetime import datetime
    
    try:
        excel_data = datasets.export_main_cities_to_excel()
        filename = f'logistics_main_cities_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        return Response(
            excel_data,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={'Content-Disposition': f'attachment; filename="{filename}"'}
        )
    except Exception as exc:
        flash(f'Ошибка при экспорте: {str(exc)}', 'danger')
        return redirect(url_for('admin.manage_logistics', tab='main'))


@admin_bp.route('/admin/logistics/main/import', methods=['POST'])
@admin_required
def import_main_cities() -> Response:
    """Импортирует основные города из Excel файла."""
    import tempfile
    import os
    import time
    from pathlib import Path
    from werkzeug.utils import secure_filename
    from datetime import datetime
    
    if 'excel_file' not in request.files:
        flash('Файл не выбран.', 'danger')
        return redirect(url_for('admin.manage_logistics', tab='main'))
    
    file = request.files['excel_file']
    
    # Проверяем, что файл выбран
    if file.filename == '':
        flash('Файл не выбран.', 'danger')
        return redirect(url_for('admin.manage_logistics', tab='main'))
    
    # Проверяем расширение файла
    if not (file.filename.lower().endswith('.xlsx') or file.filename.lower().endswith('.xls')):
        flash('Неверный формат файла. Требуется Excel (.xlsx или .xls).', 'danger')
        return redirect(url_for('admin.manage_logistics', tab='main'))
    
    temp_path = None
    try:
        # Сохраняем файл во временную директорию
        temp_dir = tempfile.gettempdir()
        filename = secure_filename(file.filename)
        temp_path = Path(temp_dir) / f'temp_import_{datetime.now().strftime("%Y%m%d%H%M%S")}_{filename}'
        file.save(str(temp_path))
        
        # Импортируем данные
        count = datasets.import_main_cities_from_excel(temp_path, actor=_current_actor())
        
        flash(f'Импортировано городов: {count}.', 'success')
    except Exception as exc:
        flash(f'Ошибка при импорте: {str(exc)}', 'danger')
    finally:
        # Гарантированно удаляем временный файл после закрытия
        if temp_path is not None and temp_path.exists():
            try:
                # Небольшая задержка для Windows, чтобы файл точно закрылся
                time.sleep(0.1)
                os.remove(temp_path)
            except (PermissionError, OSError):
                # Если не удалось удалить сразу, пытаемся через некоторое время
                pass
    
    return redirect(url_for('admin.manage_logistics', tab='main'))


@admin_bp.route('/admin/logistics/ekb-rf/export')
@admin_required
def export_ekb_rf_cities() -> Response:
    """Экспортирует города ЕКБ+РФ в Excel файл."""
    from flask import Response
    from datetime import datetime
    
    try:
        excel_data = datasets.export_ekb_rf_cities_to_excel()
        filename = f'logistics_ekb_rf_cities_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        return Response(
            excel_data,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={'Content-Disposition': f'attachment; filename="{filename}"'}
        )
    except Exception as exc:
        flash(f'Ошибка при экспорте: {str(exc)}', 'danger')
        return redirect(url_for('admin.manage_logistics', tab='ekb_rf'))


@admin_bp.route('/admin/logistics/ekb-rf/import', methods=['POST'])
@admin_required
def import_ekb_rf_cities() -> Response:
    """Импортирует города ЕКБ+РФ из Excel файла."""
    import tempfile
    import os
    import time
    from pathlib import Path
    from werkzeug.utils import secure_filename
    from datetime import datetime
    
    if 'excel_file' not in request.files:
        flash('Файл не выбран.', 'danger')
        return redirect(url_for('admin.manage_logistics', tab='ekb_rf'))
    
    file = request.files['excel_file']
    
    # Проверяем, что файл выбран
    if file.filename == '':
        flash('Файл не выбран.', 'danger')
        return redirect(url_for('admin.manage_logistics', tab='ekb_rf'))
    
    # Проверяем расширение файла
    if not (file.filename.lower().endswith('.xlsx') or file.filename.lower().endswith('.xls')):
        flash('Неверный формат файла. Требуется Excel (.xlsx или .xls).', 'danger')
        return redirect(url_for('admin.manage_logistics', tab='ekb_rf'))
    
    temp_path = None
    try:
        # Сохраняем файл во временную директорию
        temp_dir = tempfile.gettempdir()
        filename = secure_filename(file.filename)
        temp_path = Path(temp_dir) / f'temp_import_{datetime.now().strftime("%Y%m%d%H%M%S")}_{filename}'
        file.save(str(temp_path))
        
        # Импортируем данные
        count = datasets.import_ekb_rf_cities_from_excel(temp_path, actor=_current_actor())
        
        flash(f'Импортировано городов: {count}.', 'success')
    except Exception as exc:
        flash(f'Ошибка при импорте: {str(exc)}', 'danger')
    finally:
        # Гарантированно удаляем временный файл после закрытия
        if temp_path is not None and temp_path.exists():
            try:
                # Небольшая задержка для Windows, чтобы файл точно закрылся
                time.sleep(0.1)
                os.remove(temp_path)
            except (PermissionError, OSError):
                # Если не удалось удалить сразу, пытаемся через некоторое время
                pass
    
    return redirect(url_for('admin.manage_logistics', tab='ekb_rf'))


@admin_bp.route('/admin/logistics/trail/export')
@admin_required
def export_trail_cities() -> Response:
    """Экспортирует города трала в Excel файл."""
    from flask import Response
    from datetime import datetime
    
    try:
        excel_data = datasets.export_trail_cities_to_excel()
        filename = f'logistics_trail_cities_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        return Response(
            excel_data,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={'Content-Disposition': f'attachment; filename="{filename}"'}
        )
    except Exception as exc:
        flash(f'Ошибка при экспорте: {str(exc)}', 'danger')
        return redirect(url_for('admin.manage_logistics', tab='trail'))


@admin_bp.route('/admin/logistics/trail/import', methods=['POST'])
@admin_required
def import_trail_cities() -> Response:
    """Импортирует города трала из Excel файла."""
    import tempfile
    import os
    import time
    from pathlib import Path
    from werkzeug.utils import secure_filename
    from datetime import datetime
    
    if 'excel_file' not in request.files:
        flash('Файл не выбран.', 'danger')
        return redirect(url_for('admin.manage_logistics', tab='trail'))
    
    file = request.files['excel_file']
    
    # Проверяем, что файл выбран
    if file.filename == '':
        flash('Файл не выбран.', 'danger')
        return redirect(url_for('admin.manage_logistics', tab='trail'))
    
    # Проверяем расширение файла
    if not (file.filename.lower().endswith('.xlsx') or file.filename.lower().endswith('.xls')):
        flash('Неверный формат файла. Требуется Excel (.xlsx или .xls).', 'danger')
        return redirect(url_for('admin.manage_logistics', tab='trail'))
    
    temp_path = None
    try:
        # Сохраняем файл во временную директорию
        temp_dir = tempfile.gettempdir()
        filename = secure_filename(file.filename)
        temp_path = Path(temp_dir) / f'temp_import_{datetime.now().strftime("%Y%m%d%H%M%S")}_{filename}'
        file.save(str(temp_path))
        
        # Импортируем данные
        count = datasets.import_trail_cities_from_excel(temp_path, actor=_current_actor())
        
        flash(f'Импортировано городов: {count}.', 'success')
    except Exception as exc:
        flash(f'Ошибка при импорте: {str(exc)}', 'danger')
    finally:
        # Гарантированно удаляем временный файл после закрытия
        if temp_path is not None and temp_path.exists():
            try:
                # Небольшая задержка для Windows, чтобы файл точно закрылся
                time.sleep(0.1)
                os.remove(temp_path)
            except (PermissionError, OSError):
                # Если не удалось удалить сразу, пытаемся через некоторое время
                pass
    
    return redirect(url_for('admin.manage_logistics', tab='trail'))

    return render_template(
        'admin/logistics.html',
        **build_context(
            'admin_logistics',
            'Логистика',
            active_tab=active_tab,
            main_cities=main_cities,
            main_cities_groups=main_cities_groups,
            ekb_rf_cities=ekb_rf_cities,
            trail_cities=trail_cities,
            main_form=main_form,
            ekb_rf_form=ekb_rf_form,
            trail_form=trail_form,
        )
    )


@admin_bp.route('/admin/orders', methods=['GET', 'POST'])
@admin_required
def manage_orders() -> Union[str, Response]:
    return _manage_content_section('orders')


@admin_bp.route('/admin/templates', methods=['GET', 'POST'])
@admin_required
def manage_templates() -> Union[str, Response]:
    return _manage_content_section('templates')


@admin_bp.route('/admin/instructions', methods=['GET', 'POST'])
@admin_required
def manage_instructions() -> Union[str, Response]:
    return _manage_content_section('instructions')


def _handle_content_action(action: str, manager: ContentManager, redirect_endpoint: str) -> Response:
    mapping = {
        'orders': 'orders',
        'templates': 'task_templates',
        'instructions': 'instructions',
    }

    for key, collection in mapping.items():
        if action == f'content:add:{key}':
            payload = _extract_content_payload()
            manager.add_entry(collection, payload)
            flash('Запись добавлена.', 'success')
            return redirect(url_for(redirect_endpoint))

        if action == f'content:update:{key}':
            payload = _extract_content_payload()
            identifier = payload.get('id')
            if not identifier:
                flash('Не указан идентификатор записи.', 'danger')
                return redirect(url_for(redirect_endpoint))
            updated = manager.update_entry(collection, identifier, payload)
            if updated:
                flash('Запись обновлена.', 'success')
            else:
                flash('Не удалось обновить запись.', 'danger')
            return redirect(url_for(redirect_endpoint))

        if action == f'content:delete:{key}':
            identifier = request.form.get('id')
            if not identifier:
                flash('Не указан идентификатор для удаления.', 'danger')
                return redirect(url_for(redirect_endpoint))
            if manager.delete_entry(collection, identifier):
                flash('Запись удалена.', 'info')
            else:
                flash('Запись не найдена.', 'danger')
            return redirect(url_for(redirect_endpoint))

        if action == f'content:restore:{key}':
            filename = request.form.get('version')
            if not filename:
                flash('Не указана версия для восстановления.', 'danger')
                return redirect(url_for(redirect_endpoint))
            if manager.restore_version(collection, filename):
                flash('Версия восстановлена.', 'success')
            else:
                flash('Не удалось восстановить версию.', 'danger')
            return redirect(url_for(redirect_endpoint))

    flash('Неизвестное действие.', 'danger')
    return redirect(url_for(redirect_endpoint))


def _manage_content_section(section: str) -> Union[str, Response]:
    manager = build_manager(actor=_current_actor())

    if section == 'orders':
        entries_fn = manager.list_orders
        versions_key = 'orders_documents'
        template = 'admin/orders.html'
        meta = {
            'key': 'orders',
            'action_key': 'orders',
            'title': 'Распоряжения',
            'description': 'Добавляйте, редактируйте и версионируйте распоряжения.',
            'button_label': 'Новая запись',
            'modal_create_title': 'Новая запись распоряжения',
            'modal_edit_title': 'Редактирование распоряжения',
            'empty_message': 'Записи не найдены.',
            'title_field_label': 'Название распоряжения',
            'redirect_endpoint': 'admin.manage_orders',
            'active_page': 'admin_orders',
            'page_title': 'Распоряжения'
        }
    elif section == 'templates':
        entries_fn = manager.list_templates
        versions_key = 'task_templates'
        template = 'admin/templates.html'
        meta = {
            'key': 'templates',
            'action_key': 'templates',
            'title': 'Шаблоны',
            'description': 'Управляйте списком шаблонов документов и их версиями.',
            'button_label': 'Новый шаблон',
            'modal_create_title': 'Новый шаблон',
            'modal_edit_title': 'Редактирование шаблона',
            'empty_message': 'Шаблоны отсутствуют.',
            'title_field_label': 'Название шаблона',
            'redirect_endpoint': 'admin.manage_templates',
            'active_page': 'admin_templates',
            'page_title': 'Шаблоны'
        }
    elif section == 'instructions':
        entries_fn = manager.list_instructions
        versions_key = 'instructions_tasks'
        template = 'admin/instructions.html'
        meta = {
            'key': 'instructions',
            'action_key': 'instructions',
            'title': 'Инструкции',
            'description': 'Редактируйте инструкции и управляйте версионностью файлов.',
            'button_label': 'Новая инструкция',
            'modal_create_title': 'Новая инструкция',
            'modal_edit_title': 'Редактирование инструкции',
            'empty_message': 'Инструкции не найдены.',
            'title_field_label': 'Название инструкции',
            'redirect_endpoint': 'admin.manage_instructions',
            'active_page': 'admin_instructions',
            'page_title': 'Инструкции'
        }
    else:
        abort(404)

    if request.method == 'POST':
        action = request.form.get('action', '')
        return _handle_content_action(action, manager, meta['redirect_endpoint'])

    entries = entries_fn()
    versions = manager.list_versions(versions_key)

    return render_template(
        template,
        **build_context(
            meta['active_page'],
            meta['page_title'],
            content_entries=entries,
            versions=versions,
            content_meta=meta,
        )
    )


def _extract_content_payload() -> Dict[str, Any]:
    files_raw = request.form.get('files', '')
    files = datasets.parse_files_input(files_raw)
    payload = {
        'id': request.form.get('id') or request.form.get('identifier'),
        'title': (request.form.get('title') or '').strip(),
        'brief': (request.form.get('brief') or '').strip(),
        'summary': (request.form.get('summary') or '').strip(),
        'files': files,
        'updated_at': (request.form.get('updated_at') or '').strip(),
    }
    return payload


def _current_actor() -> str:
    if current_user.is_authenticated:
        return current_user.username or str(current_user.id)
    return 'system'


@admin_bp.route('/admin/users', methods=['GET', 'POST'])
@admin_required
def manage_users() -> Union[str, Response]:
    """Управление пользователями: список, создание, редактирование, удаление."""
    search = request.args.get('search', '').strip()
    role_filter = request.args.get('role', '').strip() or None
    try:
        page = int(request.args.get('page', '1'))
    except ValueError:
        page = 1

    users_data = user_repository.get_users_list(
        page=page,
        per_page=25,
        search=search if search else None,
        role_filter=role_filter
    )

    user_form = AdminUserForm(prefix='user')
    delete_form = AdminUserDeleteForm(prefix='delete')
    reset_password_form = AdminResetPasswordForm(prefix='reset')

    if request.method == 'POST':
        action = request.form.get('action', '')

        if action == 'create_user':
            if user_form.validate():
                username = user_form.username.data.strip()
                password = user_form.password.data
                
                if not password:
                    flash('Пароль обязателен при создании пользователя.', 'danger')
                elif user_repository.get_by_username(username):
                    flash('Пользователь с таким логином уже существует.', 'danger')
                else:
                    password_hash = generate_password_hash(password)
                    new_user = user_repository.create_user(
                        username=username,
                        password_hash=password_hash,
                        last_name=user_form.last_name.data.strip() if user_form.last_name.data else None,
                        first_name=user_form.first_name.data.strip() if user_form.first_name.data else None,
                        role=user_form.role.data.strip().lower() or 'user',
                        contact_info=user_form.contact_info.data.strip() if user_form.contact_info.data else None
                    )
                    if new_user:
                        log_create(
                            resource_type='user',
                            resource_id=str(new_user.id) if new_user.id else None,
                            description=f'Админ создал пользователя: {username} (роль: {user_form.role.data.strip().lower()})',
                            data={'username': username, 'role': user_form.role.data.strip().lower(), 'last_name': user_form.last_name.data, 'first_name': user_form.first_name.data},
                        )
                        flash('Пользователь успешно создан.', 'success')
                        return redirect(url_for('admin.manage_users'))
                    flash('Не удалось создать пользователя.', 'danger')

        elif action == 'edit_user':
            user_id = request.form.get('user_id')
            if not user_id:
                flash('Не указан идентификатор пользователя.', 'danger')
            else:
                user = user_repository.get_by_id(user_id)
                if not user:
                    flash('Пользователь не найден.', 'danger')
                else:
                    # Валидация без пароля (пароль опционален при редактировании)
                    username = request.form.get('user-username', '').strip()
                    password = request.form.get('user-password', '').strip()
                    confirm_password = request.form.get('user-confirm_password', '').strip()
                    
                    if not username:
                        flash('Логин обязателен.', 'danger')
                    elif password and password != confirm_password:
                        flash('Пароли не совпадают.', 'danger')
                    elif password and len(password) < 6:
                        flash('Пароль должен содержать минимум 6 символов.', 'danger')
                    else:
                        existing_user = user_repository.get_by_username(username)
                        if existing_user and str(existing_user.id) != str(user_id):
                            flash('Пользователь с таким логином уже существует.', 'danger')
                        else:
                            password_hash = None
                            if password:
                                password_hash = generate_password_hash(password)
                            
                            updated = user_repository.update_profile(
                                user_id=int(user_id),
                                username=username,
                                last_name=request.form.get('user-last_name', '').strip() or None,
                                first_name=request.form.get('user-first_name', '').strip() or None,
                                contact_info=request.form.get('user-contact_info', '').strip() or None,
                                password_hash=password_hash
                            )
                            
                            # Сохраняем данные до изменения для логирования
                            data_before = {
                                'username': user.username,
                                'last_name': user.last_name,
                                'first_name': user.first_name,
                                'role': user.role,
                                'contact_info': user.contact_info,
                            }
                            
                            # Обновляем роль отдельно, если нужно
                            new_role = request.form.get('user-role', 'user').strip().lower() or 'user'
                            role_changed = new_role != user.role
                            if role_changed:
                                from app.database.database import _session_scope
                                from app.models.models import UserRecord
                                with _session_scope() as session:
                                    user_record = session.query(UserRecord).filter(UserRecord.id == int(user_id)).one_or_none()
                                    if user_record:
                                        user_record.role = new_role
                            
                            if updated:
                                data_after = {
                                    'username': username,
                                    'last_name': request.form.get('user-last_name', '').strip() or None,
                                    'first_name': request.form.get('user-first_name', '').strip() or None,
                                    'role': new_role,
                                    'contact_info': request.form.get('user-contact_info', '').strip() or None,
                                    'password_changed': bool(password),
                                }
                                log_update(
                                    resource_type='user',
                                    resource_id=str(user_id),
                                    description=f'Админ обновил пользователя: {username}' + (f' (роль изменена: {user.role} -> {new_role})' if role_changed else ''),
                                    data_before=data_before,
                                    data_after=data_after,
                                )
                                flash('Пользователь успешно обновлён.', 'success')
                                return redirect(url_for('admin.manage_users'))
                            flash('Не удалось обновить пользователя.', 'danger')

        elif action == 'delete_user':
            if delete_form.validate():
                user_id = delete_form.user_id.data
                user = user_repository.get_by_id(user_id)
                if not user:
                    flash('Пользователь не найден.', 'danger')
                elif str(user.id) == str(current_user.id):
                    flash('Нельзя удалить самого себя.', 'danger')
                else:
                    user_data = {
                        'username': user.username,
                        'last_name': user.last_name,
                        'first_name': user.first_name,
                        'role': user.role,
                    }
                    if user_repository.delete(user_id):
                        log_delete(
                            resource_type='user',
                            resource_id=str(user_id),
                            description=f'Админ удалил пользователя: {user.username}',
                            data=user_data,
                        )
                        flash('Пользователь успешно удалён.', 'info')
                        return redirect(url_for('admin.manage_users'))
                    flash('Не удалось удалить пользователя.', 'danger')

        elif action == 'reset_password':
            if reset_password_form.validate():
                user_id = reset_password_form.user_id.data
                user = user_repository.get_by_id(user_id)
                if not user:
                    flash('Пользователь не найден.', 'danger')
                else:
                    password_hash = generate_password_hash(reset_password_form.new_password.data)
                    if user_repository.update_profile(
                        user_id=int(user_id),
                        username=user.username,
                        last_name=user.last_name,
                        first_name=user.first_name,
                        contact_info=user.contact_info,
                        password_hash=password_hash
                    ):
                        log_update(
                            resource_type='user',
                            resource_id=str(user_id),
                            description=f'Админ сбросил пароль пользователя: {user.username}',
                            data_before={'password': '***'},
                            data_after={'password': '***', 'password_reset': True},
                        )
                        flash('Пароль успешно сброшен.', 'success')
                        return redirect(url_for('admin.manage_users'))
                    flash('Не удалось сбросить пароль.', 'danger')

        elif action == 'change_role':
            user_id = request.form.get('user_id')
            new_role = request.form.get('role', '').strip().lower()
            if not user_id or new_role not in ('admin', 'user'):
                flash('Некорректные данные для изменения роли.', 'danger')
            else:
                user = user_repository.get_by_id(user_id)
                if not user:
                    flash('Пользователь не найден.', 'danger')
                elif str(user.id) == str(current_user.id):
                    flash('Нельзя изменить роль самому себе.', 'danger')
                else:
                    from app.database.database import _session_scope
                    from app.models.models import UserRecord
                    with _session_scope() as session:
                        user_record = session.query(UserRecord).filter(UserRecord.id == int(user_id)).one_or_none()
                        if user_record:
                            old_role = user_record.role
                            user_record.role = new_role
                            log_update(
                                resource_type='user',
                                resource_id=str(user_id),
                                description=f'Админ изменил роль пользователя {user.username}: {old_role} -> {new_role}',
                                data_before={'role': old_role},
                                data_after={'role': new_role},
                            )
                            flash(f'Роль пользователя изменена на "{new_role}".', 'success')
                            return redirect(url_for('admin.manage_users'))
                        flash('Не удалось изменить роль.', 'danger')

    return render_template(
        'admin/users.html',
        **build_context(
            'admin_users',
            'Управление пользователями',
            users_data=users_data,
            user_form=user_form,
            delete_form=delete_form,
            reset_password_form=reset_password_form,
            search=search,
            role_filter=role_filter,
        )
    )


@admin_bp.route('/admin/audit')
@admin_required
def audit_dashboard() -> str:
    """Дашборд с метриками активности пользователей."""
    # Получаем данные для графиков
    daily_activity = audit_log_repository.get_daily_activity(days=30)
    action_stats = audit_log_repository.get_action_stats(days=30)
    top_users = audit_log_repository.get_top_users(limit=10, days=30)
    popular_actions = audit_log_repository.get_popular_actions(limit=10, days=30)

    # Общая статистика
    total_actions = sum(item['count'] for item in daily_activity)
    unique_users = len(top_users)

    return render_template(
        'admin/audit_dashboard.html',
        **build_context(
            'admin_audit',
            'Мониторинг действий',
            daily_activity=daily_activity,
            action_stats=action_stats,
            top_users=top_users,
            popular_actions=popular_actions,
            total_actions=total_actions,
            unique_users=unique_users,
        )
    )


@admin_bp.route('/admin/audit/logs')
@admin_required
def audit_logs() -> Union[str, Response]:
    """Детальный лог действий с фильтрами и пагинацией."""
    # Проверка экспорта
    export_format = request.args.get('export', '').strip().lower()
    if export_format in ('excel', 'pdf'):
        # Применяем те же фильтры для экспорта
        user_id = request.args.get('user_id', '').strip() or None
        if user_id:
            try:
                user_id = int(user_id)
            except ValueError:
                user_id = None

        username = request.args.get('username', '').strip() or None
        action_type = request.args.get('action_type', '').strip() or None
        resource_type = request.args.get('resource_type', '').strip() or None
        date_from_str = request.args.get('date_from', '').strip() or None
        date_to_str = request.args.get('date_to', '').strip() or None
        search = request.args.get('search', '').strip() or None

        date_from = None
        date_to = None
        if date_from_str:
            try:
                date_from = datetime.strptime(date_from_str, '%Y-%m-%d')
            except ValueError:
                pass
        if date_to_str:
            try:
                date_to = datetime.strptime(date_to_str, '%Y-%m-%d')
                date_to = date_to.replace(hour=23, minute=59, second=59)
            except ValueError:
                pass

        # Получаем все логи без пагинации для экспорта
        logs_data = audit_log_repository.get_logs(
            page=1,
            per_page=10000,  # Большое значение для получения всех записей
            user_id=user_id,
            username=username,
            action_type=action_type,
            resource_type=resource_type,
            date_from=date_from,
            date_to=date_to,
            search=search,
        )

        if export_format == 'excel':
            try:
                buffer = export_audit_logs_to_excel(logs_data)
                return create_excel_response(buffer)
            except RuntimeError as e:
                flash(f'Ошибка экспорта в Excel: {str(e)}', 'danger')
                return redirect(url_for('admin.audit_logs'))
        elif export_format == 'pdf':
            try:
                buffer = export_audit_logs_to_pdf(logs_data)
                return create_pdf_response(buffer)
            except RuntimeError as e:
                flash(f'Ошибка экспорта в PDF: {str(e)}', 'danger')
                return redirect(url_for('admin.audit_logs'))

    # Параметры фильтрации
    try:
        page = int(request.args.get('page', '1'))
    except ValueError:
        page = 1

    try:
        per_page = int(request.args.get('per_page', '50'))
    except ValueError:
        per_page = 50

    user_id = request.args.get('user_id', '').strip() or None
    if user_id:
        try:
            user_id = int(user_id)
        except ValueError:
            user_id = None

    username = request.args.get('username', '').strip() or None
    action_type = request.args.get('action_type', '').strip() or None
    resource_type = request.args.get('resource_type', '').strip() or None
    date_from_str = request.args.get('date_from', '').strip() or None
    date_to_str = request.args.get('date_to', '').strip() or None
    search = request.args.get('search', '').strip() or None

    # Парсинг дат
    date_from = None
    date_to = None
    if date_from_str:
        try:
            date_from = datetime.strptime(date_from_str, '%Y-%m-%d')
        except ValueError:
            pass
    if date_to_str:
        try:
            date_to = datetime.strptime(date_to_str, '%Y-%m-%d')
            # Добавляем время конца дня
            date_to = date_to.replace(hour=23, minute=59, second=59)
        except ValueError:
            pass

    # Получаем логи
    logs_data = audit_log_repository.get_logs(
        page=page,
        per_page=per_page,
        user_id=user_id,
        username=username,
        action_type=action_type,
        resource_type=resource_type,
        date_from=date_from,
        date_to=date_to,
        search=search,
    )

    # Получаем список пользователей для фильтра
    users_list = user_repository.get_users_list(page=1, per_page=1000, search=None, role_filter=None)
    users = users_list.get('items', [])

    # Уникальные типы действий и ресурсов для фильтров
    action_types = ['login', 'logout', 'create', 'update', 'delete', 'view', 'export']
    resource_types = ['user', 'generation', 'duty', 'material', 'logistics', 'order', 'template', 'instruction']

    return render_template(
        'admin/audit_logs.html',
        **build_context(
            'admin_audit_logs',
            'Лог действий',
            logs_data=logs_data,
            users=users,
            action_types=action_types,
            resource_types=resource_types,
            filters={
                'user_id': user_id,
                'username': username,
                'action_type': action_type,
                'resource_type': resource_type,
                'date_from': date_from_str,
                'date_to': date_to_str,
                'search': search,
            },
        )
    )


@admin_bp.route('/admin/audit/api/daily-activity')
@admin_required
def api_daily_activity() -> Response:
    """API endpoint для данных графика активности по дням."""
    try:
        days = int(request.args.get('days', '30'))
    except ValueError:
        days = 30

    daily_activity = audit_log_repository.get_daily_activity(days=days)
    return jsonify(daily_activity)


@admin_bp.route('/admin/audit/api/action-distribution')
@admin_required
def api_action_distribution() -> Response:
    """API endpoint для данных распределения по типам действий."""
    try:
        days = int(request.args.get('days', '30'))
    except ValueError:
        days = 30

    action_stats = audit_log_repository.get_action_stats(days=days)
    return jsonify(action_stats)


@admin_bp.route('/admin/audit/api/top-users')
@admin_required
def api_top_users() -> Response:
    """API endpoint для списка наиболее активных пользователей."""
    try:
        limit = int(request.args.get('limit', '10'))
        days = int(request.args.get('days', '30'))
    except ValueError:
        limit = 10
        days = 30

    top_users = audit_log_repository.get_top_users(limit=limit, days=days)
    return jsonify(top_users)


@admin_bp.route('/admin/ai-agent', methods=['GET', 'POST'])
@admin_required
def manage_ai_agent() -> Union[str, Response]:
    """Страница управления AI агентом."""
    import os
    from flask import current_app
    
    config_form = AIAgentConfigForm()
    cache_form = AIAgentCacheForm()
    
    # Получаем текущие настройки (модель из app.config, заполненного из .env при load_config)
    app_settings = current_app.config.get('APP_SETTINGS') or {}
    current_settings = {
        'api_key_set': bool(os.getenv('OPENROUTER_API_KEY')),
        'model_name': app_settings.get('openrouter_model') or os.getenv('OPENROUTER_MODEL', 'xiaomi/mimo-v2-flash:free'),
        'timeout': int(os.getenv('OPENROUTER_TIMEOUT', '60')),
        'reasoning_enabled': os.getenv('OPENROUTER_REASONING_ENABLED', 'true').lower() == 'true',
        'fallback_enabled': os.getenv('AI_FALLBACK_ENABLED', 'true').lower() == 'true',
        'usage_monitoring': os.getenv('AI_USAGE_MONITORING', 'true').lower() == 'true',
        'max_history_length': int(os.getenv('AI_MAX_HISTORY_LENGTH', '20')),
        'cache_ttl': int(os.getenv('AI_CACHE_TTL', '86400')),
    }
    
    if request.method == 'POST':
        # Обработка очистки кеша
        if cache_form.validate_on_submit() and cache_form.action.data == 'clear_cache':
            try:
                from ai_agent.cache_manager import invalidate_ai_cache
                result = invalidate_ai_cache()
                if result:
                    flash('Кеш AI агента успешно очищен.', 'success')
                else:
                    flash('Не удалось очистить кеш (Redis может быть недоступен).', 'warning')
            except Exception as e:
                current_app.logger.error(f'Ошибка очистки кеша AI: {e}')
                flash(f'Ошибка при очистке кеша: {str(e)}', 'danger')
            
            return redirect(url_for('admin.manage_ai_agent'))
        
        # Обработка обновления настроек
        if config_form.validate_on_submit():
            flash('Внимание: Изменение настроек через интерфейс требует обновления .env файла на сервере. '
                  'Настройки отображаются для информации.', 'info')
            # В реальном production нужно было бы сохранять в БД или конфиг
            return redirect(url_for('admin.manage_ai_agent'))
    
    # Предзаполняем форму текущими значениями
    if request.method == 'GET':
        config_form.model_name.data = current_settings['model_name']
        config_form.timeout.data = current_settings['timeout']
        config_form.reasoning_enabled.data = current_settings['reasoning_enabled']
        config_form.fallback_enabled.data = current_settings['fallback_enabled']
        config_form.usage_monitoring.data = current_settings['usage_monitoring']
        config_form.max_history_length.data = current_settings['max_history_length']
        config_form.cache_ttl.data = current_settings['cache_ttl']
    
    # Получаем статистику кеша (если доступен)
    cache_stats = {}
    try:
        from ai_agent.cache_manager import AICacheManager
        cache_stats = AICacheManager.get_cache_stats()
    except Exception as e:
        current_app.logger.debug(f'Ошибка получения статистики кеша AI: {e}')
        cache_stats = {}
    
    # Проверяем валидность API ключа
    api_key_status = 'unknown'
    api_key_message = ''
    try:
        from ai_agent.api_validator import validate_api_key
        if current_settings['api_key_set']:
            api_key = os.getenv('OPENROUTER_API_KEY')
            validation_result = validate_api_key(api_key)
            if validation_result.is_valid:
                api_key_status = 'valid'
                api_key_message = 'API ключ валиден и работает'
            else:
                api_key_status = 'invalid'
                api_key_message = validation_result.error_message or 'API ключ невалиден'
        else:
            api_key_status = 'not_set'
            api_key_message = 'API ключ не установлен'
    except Exception as e:
        current_app.logger.error(f'Ошибка проверки API ключа: {e}')
        api_key_status = 'error'
        api_key_message = f'Ошибка при проверке: {str(e)}'
    
    return render_template(
        'admin/ai_agent_settings.html',
        **build_context(
            'admin_ai_agent',
            'Управление AI Агентом',
            config_form=config_form,
            cache_form=cache_form,
            current_settings=current_settings,
            api_key_status=api_key_status,
            api_key_message=api_key_message,
            cache_stats=cache_stats
        )
    )
