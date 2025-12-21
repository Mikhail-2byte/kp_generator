from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional, Union

from flask import Blueprint, Response, abort, flash, jsonify, redirect, render_template, request, url_for

from flask_login import current_user
from werkzeug.security import generate_password_hash

from app.presentation.forms import (
    AdminResetPasswordForm,
    AdminUserDeleteForm,
    AdminUserForm,
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
    TrailCityDeleteForm
)
from app.auth.security import admin_required
from app.services import datasets
from app.services.audit_service import log_create, log_delete, log_update
from app.services.content_manager import ContentManager, build_manager
from app.services.export_service import (
    create_excel_response,
    create_pdf_response,
    export_audit_logs_to_excel,
    export_audit_logs_to_pdf,
)
from app.services.repositories import admin_stats_repository, audit_log_repository, customer_contact_repository, user_repository
from app.presentation.ui import build_context


admin_bp = Blueprint('admin', __name__)  # Управление справочниками через административный интерфейс


@admin_bp.route('/admin/stats')
@admin_required
def manage_stats() -> str:
    user_activity = admin_stats_repository.get_user_activity()
    return render_template(
        'admin/stats.html',
        **build_context(
            'admin_stats',
            'Статистика активности',
            user_activity=user_activity,
        )
    )


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
                notes_text = (gb_form.notes.data or '').strip()
                composition_list = datasets.parse_composition_input(gb_form.composition.data)
                composition_search = ' '.join(
                    f"{comp.get('element', '')} {comp.get('content', '')}" for comp in composition_list
                ).lower()

                new_material = {
                    'russian': russian_name,
                    'gb': gb_name,
                    'notes': notes_text,
                    'composition': composition_list,
                    'composition_search': composition_search
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

    return render_template(
        'admin.html',
        **build_context(
            'admin',
            'Администрирование',
        )
    )


@admin_bp.route('/admin/duty', methods=['GET', 'POST'])
@admin_required
def manage_duty() -> Union[str, Response]:
    duty_items = datasets.load_duty_rates()
    duty_form = DutyItemForm(prefix='duty')
    duty_form.action.data = 'add_duty'

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
        )
    )


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
                notes_text = (gb_form.notes.data or '').strip()
                composition_list = datasets.parse_composition_input(gb_form.composition.data)
                composition_search = ' '.join(
                    f"{comp.get('element', '')} {comp.get('content', '')}" for comp in composition_list
                ).lower()

                gb_materials.append({
                    'russian': russian_name,
                    'gb': gb_name,
                    'notes': notes_text,
                    'composition': composition_list,
                    'composition_search': composition_search
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
                    notes_text = (gb_form.notes.data or '').strip()
                    composition_list = datasets.parse_composition_input(gb_form.composition.data)
                    composition_search = ' '.join(
                        f"{comp.get('element', '')} {comp.get('content', '')}" for comp in composition_list
                    ).lower()

                    gb_materials[index] = {
                        'russian': russian_name,
                        'gb': gb_name,
                        'notes': notes_text,
                        'composition': composition_list,
                        'composition_search': composition_search
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
        'Санкт-Петербург'
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


@admin_bp.route('/admin/customers', methods=['GET', 'POST'])
@admin_required
def manage_customers() -> Union[str, Response]:
    """Управление контактами заказчиков: список, создание, редактирование, удаление."""
    form = CustomerContactForm(prefix='customer')
    delete_form = CustomerContactDeleteForm(prefix='delete')
    search_query = request.args.get('search', '').strip()
    
    if request.method == 'POST':
        if form.submit.data and form.validate():
            contact_id = request.form.get('customer-contact_id')
            if contact_id:
                # Редактирование
                success = customer_contact_repository.update(
                    int(contact_id),
                    company_name=form.company_name.data,
                    contact_person=form.contact_person.data or None,
                    phone=form.phone.data or None,
                    email=form.email.data or None,
                    address=form.address.data or None,
                    notes=form.notes.data or None
                )
                if success:
                    log_update('customer_contact', contact_id, data_after={
                        'company_name': form.company_name.data,
                        'contact_person': form.contact_person.data,
                    })
                    flash('Контакт успешно обновлен.', 'success')
                else:
                    flash('Ошибка при обновлении контакта.', 'danger')
            else:
                # Создание
                contact_id = customer_contact_repository.create(
                    company_name=form.company_name.data,
                    contact_person=form.contact_person.data or None,
                    phone=form.phone.data or None,
                    email=form.email.data or None,
                    address=form.address.data or None,
                    notes=form.notes.data or None
                )
                if contact_id:
                    log_create('customer_contact', str(contact_id), data={
                        'company_name': form.company_name.data,
                    })
                    flash('Контакт успешно создан.', 'success')
                else:
                    flash('Ошибка при создании контакта.', 'danger')
            return redirect(url_for('admin.manage_customers'))
        
        if delete_form.submit.data and delete_form.validate():
            contact_id = delete_form.contact_id.data
            contact = customer_contact_repository.get_by_id(int(contact_id))
            if contact:
                success = customer_contact_repository.delete(int(contact_id))
                if success:
                    log_delete('customer_contact', contact_id, data=contact)
                    flash('Контакт успешно удален.', 'success')
                else:
                    flash('Ошибка при удалении контакта.', 'danger')
            else:
                flash('Контакт не найден.', 'danger')
            return redirect(url_for('admin.manage_customers'))
    
    # Получение списка контактов
    contacts = customer_contact_repository.get_all(search=search_query if search_query else None)
    
    # Получение контакта для редактирования
    edit_id = request.args.get('edit')
    edit_contact = None
    if edit_id:
        edit_contact = customer_contact_repository.get_by_id(int(edit_id))
        if edit_contact:
            form.company_name.data = edit_contact['company_name']
            form.contact_person.data = edit_contact.get('contact_person', '')
            form.phone.data = edit_contact.get('phone', '')
            form.email.data = edit_contact.get('email', '')
            form.address.data = edit_contact.get('address', '')
            form.notes.data = edit_contact.get('notes', '')
    
    return render_template(
        'admin/customers.html',
        **build_context(
            'admin_customers',
            'Управление контактами заказчиков',
            contacts=contacts,
            form=form,
            delete_form=delete_form,
            edit_contact=edit_contact,
            search_query=search_query
        )
    )
    try:
        limit = int(request.args.get('limit', '10'))
        days = int(request.args.get('days', '30'))
    except ValueError:
        limit = 10
        days = 30

    top_users = audit_log_repository.get_top_users(limit=limit, days=days)
    return jsonify(top_users)
