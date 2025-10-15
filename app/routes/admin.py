from flask import Blueprint, flash, redirect, render_template, request, url_for

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
from app.ui import build_context


admin_bp = Blueprint('admin', __name__)  # Управление справочниками через административный интерфейс


@admin_bp.route('/admin', methods=['GET', 'POST'])
@admin_required
def admin_panel():
    """Позволяет администраторам редактировать справочные данные приложения."""
    duty_items = datasets.load_duty_rates()
    gb_materials = datasets.load_gb_materials()
    logistics_cities = datasets.load_logistics_cities()

    duty_form = DutyItemForm(prefix='duty')
    gb_form = GBMaterialForm(prefix='gb')
    logistics_form = LogisticsCityForm(prefix='logistics')

    duty_form.action.data = 'add_duty'
    gb_form.action.data = 'add_gb'
    logistics_form.action.data = 'add_city'

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
                datasets.save_duty_rates(duty_items)
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
                    datasets.save_duty_rates(duty_items)
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
                datasets.save_gb_materials(gb_materials)
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
                    datasets.save_gb_materials(gb_materials)
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
                datasets.save_logistics_cities(logistics_cities)
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
                    datasets.save_logistics_cities(logistics_cities)
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
        else:
            flash('Неизвестное действие.', 'danger')
            return redirect(url_for('admin.admin_panel'))

    return render_template(
        'admin.html',
        **build_context(
            'admin',
            'Администрирование',
            duty_items=duty_items,
            gb_materials=gb_materials,
            logistics_cities=logistics_cities,
            duty_form=duty_form,
            gb_form=gb_form,
            logistics_form=logistics_form
        )
    )
