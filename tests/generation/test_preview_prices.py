"""Тесты для API предварительного расчёта цен (/api/preview-prices)."""

import json
import sys
from pathlib import Path

import pytest

project_root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(project_root))


def _reference_payload():
    """Эталон: 1 поз., закуп 10 ¥/кг, 10 шт, вес 1000 кг, пошлина 0, 150/30 дн, Кредит, 30%, лог. 10000."""
    return {
        'company': 'Тест',
        'tender_number': '1111',
        'delivery_time': '150',
        'payment_terms': '100% по факту поставки в течении 30 календарных дней',
        'logistics': '10000',
        'margin_percent': '30',
        'finance_credit': '1',
        'finance_bank_guarantee': '',
        'product': 'тест ч.1111',
        'drawing_number': '1111',
        'material': '1111',
        'cost_price_per_kg': '10',
        'quantity': '10',
        'weight': '1000',
        'duty_percent': '0',
        'additional_expenses': [],
    }


class TestPreviewPricesAPI:
    """Тесты для POST /api/preview-prices."""

    def test_preview_prices_success_returns_summary_with_total_direct_costs_yuan(self, client):
        """Ответ 200, в summary есть total_direct_costs_yuan, total_final_price, conversion_rate."""
        payload = _reference_payload()
        response = client.post(
            '/api/preview-prices',
            data=json.dumps(payload),
            content_type='application/json',
        )
        assert response.status_code == 200
        data = response.get_json()
        assert data is not None
        assert 'summary' in data
        summary = data['summary']
        assert 'total_direct_costs_yuan' in summary
        assert 'total_final_price' in summary
        assert 'conversion_rate' in summary
        assert 'total_purchase_cost' in summary
        assert 'total_logistics_cost' in summary
        assert 'actual_margin_percent' in summary
        assert 'total_duty_yuan' in summary
        assert 'total_credit_yuan' in summary

    def test_preview_prices_reference_case_totals(self, client):
        """
        Эталонный кейс: итог ~159 940 ¥, прямые затраты ~111 960 ¥, прибыль ~47 980 ¥.
        """
        payload = _reference_payload()
        response = client.post(
            '/api/preview-prices',
            data=json.dumps(payload),
            content_type='application/json',
        )
        assert response.status_code == 200
        data = response.get_json()
        summary = data['summary']
        total_final = summary['total_final_price']
        total_direct = summary['total_direct_costs_yuan']
        profit = total_final - total_direct
        assert abs(total_final - 159940) < 5, f'total_final_price={total_final}'
        assert abs(total_direct - 111960) < 5, f'total_direct={total_direct}'
        assert abs(profit - 47980) < 5, f'profit={profit}'

    def test_preview_prices_single_position_returns_margin(self, client):
        """Для одной позиции в ответе есть positions[0].margin (из калькулятора)."""
        payload = _reference_payload()
        response = client.post(
            '/api/preview-prices',
            data=json.dumps(payload),
            content_type='application/json',
        )
        assert response.status_code == 200
        data = response.get_json()
        assert data['positions']
        pos = data['positions'][0]
        assert 'margin' in pos
        assert abs(pos['margin'] - 30) < 1, f"margin={pos['margin']!r}"

    def test_preview_prices_no_positions_returns_400(self, client):
        """При отсутствии позиций возвращается 400."""
        payload = {
            'company': 'Тест',
            'logistics': '10000',
            'margin_percent': '30',
            'delivery_time': '150',
        }
        response = client.post(
            '/api/preview-prices',
            data=json.dumps(payload),
            content_type='application/json',
        )
        assert response.status_code == 400
        data = response.get_json()
        assert data.get('error')
        err = data['error'].lower()
        assert 'позиц' in err or 'position' in err

    def test_preview_prices_margin_percent_affects_result(self, client):
        """Изменение margin_percent в запросе меняет total_final_price."""
        base = _reference_payload()
        base['margin_percent'] = '30'
        resp30 = client.post(
            '/api/preview-prices', data=json.dumps(base), content_type='application/json'
        )
        assert resp30.status_code == 200
        total30 = resp30.get_json()['summary']['total_final_price']

        base['margin_percent'] = '20'
        resp20 = client.post(
            '/api/preview-prices', data=json.dumps(base), content_type='application/json'
        )
        assert resp20.status_code == 200
        total20 = resp20.get_json()['summary']['total_final_price']

        assert total20 < total30, 'При марже 20% итог должен быть меньше, чем при 30%'

    def test_preview_prices_credit_affects_result(self, client):
        """Включение кредита увеличивает total_final_price и total_direct_costs_yuan."""
        base = _reference_payload()
        base['finance_credit'] = ''
        base['payment_terms'] = ''
        resp_no_credit = client.post(
            '/api/preview-prices', data=json.dumps(base), content_type='application/json'
        )
        assert resp_no_credit.status_code == 200
        no_credit = resp_no_credit.get_json()['summary']

        base['finance_credit'] = '1'
        base['payment_terms'] = '100% по факту поставки в течении 30 календарных дней'
        resp_credit = client.post(
            '/api/preview-prices', data=json.dumps(base), content_type='application/json'
        )
        assert resp_credit.status_code == 200
        with_credit = resp_credit.get_json()['summary']

        assert with_credit['total_final_price'] > no_credit['total_final_price']
        assert with_credit['total_direct_costs_yuan'] > no_credit['total_direct_costs_yuan']

    def test_preview_prices_actual_margin_matches_profit_formula(self, client):
        """actual_margin_percent = (total_final_price - total_costs) / total_final_price."""
        payload = _reference_payload()
        response = client.post(
            '/api/preview-prices',
            data=json.dumps(payload),
            content_type='application/json',
        )
        assert response.status_code == 200
        summary = response.get_json()['summary']
        total = summary['total_final_price']
        direct = summary['total_direct_costs_yuan']
        additional = summary.get('total_additional_expenses', 0) or 0
        total_costs = direct + additional
        actual = summary['actual_margin_percent']
        expected_pct = ((total - total_costs) / total * 100) if total > 0 else 0
        assert abs(actual - expected_pct) < 0.5, (
            f'actual={actual}, expected={expected_pct}'
        )


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
