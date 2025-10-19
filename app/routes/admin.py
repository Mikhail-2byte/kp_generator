from flask import Blueprint, abort, flash, redirect, render_template, request, url_for

from app.forms import (
    DutyDeleteForm,
    DutyItemForm,
    GBMaterialDeleteForm,
    GBMaterialForm,
    LogisticsCityDeleteForm,
    LogisticsCityForm
)
from app.security import admin_required
from app.services import datasets
from app.services.content_manager import ContentManager, build_manager
from app.services.repositories import admin_stats_repository
from app.ui import build_context


admin_bp = Blueprint('admin', __name__)  # Управление справочниками через административный интерфейс


@admin_bp.route('/admin/stats')
@admin_required
def manage_stats():
    user_activity = admin_stats_repository.get_user_activity()
    return render_template(
        'admin/stats.html',
        **build_context(
            'admin_stats',
            'Статистика активности',
            user_activity=user_activity,
        )
    )


@admin_bp.route('/admin', methods=['GET', 'POST'])
@admin_required
def admin_panel():
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
def manage_duty():
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
def manage_materials():
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
def manage_logistics():
    logistics_cities = datasets.load_logistics_cities()
    logistics_form = LogisticsCityForm(prefix='logistics')
    logistics_form.action.data = 'add_city'

    if request.method == 'POST':
        action = request.form.get('action', '')

        if action == 'add_city':
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
                return redirect(url_for('admin.manage_logistics'))
            flash('Исправьте ошибки в форме.', 'danger')

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
            return redirect(url_for('admin.manage_logistics'))

        else:
            flash('Неизвестное действие.', 'danger')
            return redirect(url_for('admin.manage_logistics'))

    return render_template(
        'admin/logistics.html',
        **build_context(
            'admin_logistics',
            'Логистика',
            logistics_cities=logistics_cities,
            logistics_form=logistics_form,
        )
    )


@admin_bp.route('/admin/orders', methods=['GET', 'POST'])
@admin_required
def manage_orders():
    return _manage_content_section('orders')


@admin_bp.route('/admin/templates', methods=['GET', 'POST'])
@admin_required
def manage_templates():
    return _manage_content_section('templates')


@admin_bp.route('/admin/instructions', methods=['GET', 'POST'])
@admin_required
def manage_instructions():
    return _manage_content_section('instructions')


def _handle_content_action(action: str, manager: ContentManager, redirect_endpoint: str):
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


def _manage_content_section(section: str):
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


def _extract_content_payload() -> dict:
    files_raw = request.form.get('files', '')
    files = datasets.parse_files_input(files_raw)
    payload = {
        'id': request.form.get('id') or request.form.get('identifier'),
        'title': (request.form.get('title') or '').strip(),
        'summary': (request.form.get('summary') or '').strip(),
        'files': files,
        'updated_at': (request.form.get('updated_at') or '').strip(),
    }
    return payload


def _current_actor() -> str:
    from flask_login import current_user

    if current_user.is_authenticated:
        return current_user.username or str(current_user.id)
    return 'system'
