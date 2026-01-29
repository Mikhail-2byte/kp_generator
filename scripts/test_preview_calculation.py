#!/usr/bin/env python
"""
Скрипт для проверки предварительного расчёта.
Вводные: 1 позиция, закуп 10 ¥/кг, 10 шт, вес 1000 кг, пошлина 0,
срок поставки 150, условия оплаты 30 дн, Кредит, маржа 30%.
Логистика 10 000 руб (для сравнения с эталонным Excel).
Запуск: из корня проекта: python scripts/test_preview_calculation.py
"""
import json
import os
import sys

# Корень проекта
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

# Конфиг как в config/settings.json
CONFIG = {
    "calculation_constants": {
        "conversion_rate": 11.5,
        "logistics_cnr_ratio": 0.3,
        "logistics_rf_ratio": 0.7,
        "conversion_fee_rate": 0.032,
        "credit_rate": 0.16,
        "vat_rate": 0.22,
    }
}


def main():
    from app.services.generation_orchestrator import GenerationOrchestrator

    orchestrator = GenerationOrchestrator(CONFIG)
    positions = [{
        "quantity": 10,
        "cost_price": 10000.0,
        "weight": 1000.0,
        "duty_percent": 0,
        "product": "тест ч.1111",
        "drawing_number": "1111",
        "material": "1111",
    }]
    logistics_rub = 10000.0
    delivery_time = 150
    margin_percent = 30.0
    use_credit = True
    payment_days = 30

    position_prices, total_general_price = orchestrator.calculate_prices(
        positions,
        logistics_rub,
        delivery_time,
        margin_percent,
        use_credit=use_credit,
        payment_days=payment_days,
    )

    res = position_prices[0]
    costs = orchestrator.calculator.calculate_position_costs(
        positions[0], logistics_rub, delivery_time, 10000.0,
        use_credit=use_credit, payment_days=payment_days,
    )

    vat_rate = CONFIG["calculation_constants"]["vat_rate"]
    total_nds = total_general_price * (1 + vat_rate)

    print("=== Ожидаемые значения для сравнения с итоговым Excel ===\n")
    print("Вводные: 1 позиция, закуп 10 ¥/кг, 10 шт, вес 1000 кг, пошлина 0%,")
    print("срок поставки 150 дн, условия оплаты 30 дн, Кредит, маржа 30%, логистика 10 000 руб.\n")
    print("Позиция (расчёт):")
    print(f"  Цена без НДС за ед. (final_price):     {res['final_price']:,.2f} ¥")
    print(f"  Сумма по позиции (general_price):     {res['general_price']:,.2f} ¥")
    print(f"  Итого с НДС (22%):                   {total_nds:,.2f} ¥")
    print(f"  Фактическая маржа:                    {res['margin']:.2f}%")
    print()
    print("Составляющие затрат (всего по позиции):")
    print(f"  Закуп (стоимость контракта):          {10000 * 10:,.2f} ¥")
    print(f"  Конвертация 3,2%:                     {costs['conversion_fee_per_unit'] * 10:,.2f} ¥")
    print(f"  Таможенная пошлина:                   {costs['duty_per_unit'] * 10:,.2f} ¥")
    print(f"  Логистика КНР (итого):                {costs['logistics_cnr_per_unit'] * 10:,.2f} ¥")
    print(f"  Логистика РФ (итого):                 {costs['logistics_rf_per_unit'] * 10:,.2f} ¥")
    print(f"  Кредит (180 дн):                      {costs['credit_cost_per_unit'] * 10:,.2f} ¥")
    print()
    print("В итоговом Excel проверьте:")
    print("  — Цена без НДС за ед. ≈ 15 994,28 ¥")
    print("  — Итого по спецификации ≈ 159 940 ¥")
    print("  — Прогноз выручки по контракту (с НДС) ≈ 195 126,80 ¥")
    print("  — Кредит ≈ 7 890,41 ¥")
    print("  — Маржинальный доход % = 30%")


if __name__ == "__main__":
    main()
