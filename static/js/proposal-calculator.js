/**
 * Proposal Calculator - Reactive calculations for summary panel
 * Updates the right panel in real-time as form data changes
 */

(function() {
    'use strict';
    
    // Configuration
    const CONVERSION_RATE = 14; // 1 ¥ = 14 ₽ (should be loaded from config)
    
    // Cache DOM elements
    let summaryTotalEl, summaryProfitEl, summaryProfitabilityEl;
    let summaryTotalPurchaseCostEl, summaryLogisticsCostEl, summaryVATEl, summaryDutyEl, summaryMarginEl;
    let summaryFinalPriceRublesEl, summaryExchangeRateEl;
    let summaryExpensesListEl;
    
    // Initialize
    function init() {
        // Cache summary panel elements
        summaryTotalEl = document.getElementById('summaryTotal');
        summaryProfitEl = document.getElementById('summaryProfit');
        summaryProfitabilityEl = document.getElementById('summaryProfitability');
        summaryTotalUnitEl = document.getElementById('summaryTotalUnit');
        summaryProfitUnitEl = document.getElementById('summaryProfitUnit');
        
        summaryTotalPurchaseCostEl = document.getElementById('summaryTotalPurchaseCost');
        summaryLogisticsCostEl = document.getElementById('summaryLogisticsCost');
        summaryVATEl = document.getElementById('summaryVAT');
        summaryDutyEl = document.getElementById('summaryDuty');
        summaryMarginEl = document.getElementById('summaryMargin');
        
        summaryFinalPriceRublesEl = document.getElementById('summaryFinalPriceRubles');
        summaryExchangeRateEl = document.getElementById('summaryExchangeRate');
        summaryExpensesListEl = document.getElementById('summaryExpensesList');
        
        // Initial calculation
        calculateSummary();
        
        // Watch for form changes
        watchFormChanges();
    }
    
    /**
     * Calculate summary metrics from form data
     */
    function calculateSummary() {
        try {
            const form = document.getElementById('generateForm');
            if (!form) return;
            
            // Get budget region
            const budgetRegion = getSelectedBudgetRegion();
            const isCN = budgetRegion === 'cn';
            
            // Calculate total purchase cost
            let totalPurchaseCost = 0;
            let totalWeight = 0;
            let positionCount = 0;
            
            const positionBlocks = document.querySelectorAll('.position-block');
            positionBlocks.forEach(block => {
                const costInput = block.querySelector('[data-field="cost_price"]');
                const quantityInput = block.querySelector('[data-field="quantity"]');
                const weightInput = block.querySelector('[data-field="weight"]');
                
                if (costInput && quantityInput) {
                    const cost = parseFloat(costInput.value) || 0;
                    const qty = parseFloat(quantityInput.value) || 0;
                    totalPurchaseCost += cost * qty;
                    positionCount += qty > 0 ? 1 : 0;
                }
                
                if (weightInput && quantityInput) {
                    const weight = parseFloat(weightInput.value) || 0;
                    const qty = parseFloat(quantityInput.value) || 0;
                    totalWeight += weight * qty;
                }
            });
            
            // Get logistics cost
            const logisticsInput = document.getElementById('logistics');
            const logisticsCost = parseFloat(logisticsInput?.value) || 0;
            
            // Get margin percent
            const marginPercentInput = document.getElementById('margin_percent');
            const marginPercent = parseFloat(marginPercentInput?.value) || 30;
            
            // Calculate margin amount
            const marginAmount = totalPurchaseCost * (marginPercent / 100);
            
            // Calculate duty (simplified - sum of all position duties)
            let totalDuty = 0;
            positionBlocks.forEach(block => {
                const dutyInput = block.querySelector('[data-field="duty_percent"]');
                const costInput = block.querySelector('[data-field="cost_price"]');
                const quantityInput = block.querySelector('[data-field="quantity"]');
                
                if (dutyInput && costInput && quantityInput) {
                    const dutyPercent = parseFloat(dutyInput.value) || 0;
                    const cost = parseFloat(costInput.value) || 0;
                    const qty = parseFloat(quantityInput.value) || 0;
                    totalDuty += (cost * qty * dutyPercent / 100);
                }
            });
            
            // Get additional expenses
            const additionalExpenses = getAdditionalExpenses();
            const totalExpenses = additionalExpenses.reduce((sum, exp) => sum + (parseFloat(exp.amount) || 0), 0);
            const expensesInCurrency = isCN ? (totalExpenses / CONVERSION_RATE) : totalExpenses;
            
            // Calculate final price
            const logisticsInCurrency = isCN ? (logisticsCost / CONVERSION_RATE) : logisticsCost;
            const dutyInCurrency = isCN ? totalDuty : (totalDuty * CONVERSION_RATE);
            const finalPrice = totalPurchaseCost + logisticsInCurrency + dutyInCurrency + marginAmount + expensesInCurrency;
            
            // Calculate final price in rubles
            const finalPriceInRubles = isCN ? (finalPrice * CONVERSION_RATE) : finalPrice;
            
            // Calculate profit (simplified - margin amount)
            const profit = marginAmount;
            
            // Calculate profitability
            const profitability = totalPurchaseCost > 0 
                ? ((profit / finalPrice) * 100).toFixed(2)
                : 0;
            
            // Calculate VAT (15% of margin)
            const vatAmount = marginAmount * 0.15;
            
            // Update UI
            updateMetrics({
                total: finalPrice + expensesInCurrency,
                profit: profit,
                profitability: profitability,
                totalPurchaseCost: totalPurchaseCost,
                logistics: logisticsCost,
                vat: vatAmount,
                duty: totalDuty,
                margin: marginAmount,
                finalPriceRubles: finalPriceInRubles,
                currency: isCN ? '¥' : '₽',
                exchangeRate: CONVERSION_RATE
            });
            
            // Update expenses list in summary
            updateExpensesList(additionalExpenses, isCN);
            
        } catch (error) {
            console.error('Error calculating summary:', error);
        }
    }
    
    /**
     * Get selected budget region
     */
    function getSelectedBudgetRegion() {
        const checked = document.querySelector('input[name="budget_region"]:checked');
        return checked ? checked.value : 'cn';
    }
    
    /**
     * Get additional expenses from form
     */
    function getAdditionalExpenses() {
        const expensesList = document.getElementById('expensesList');
        if (!expensesList) return [];
        
        const expenses = [];
        const expenseItems = expensesList.querySelectorAll('.expense-item');
        
        expenseItems.forEach(item => {
            const nameEl = item.querySelector('.expense-item-name');
            const amountEl = item.querySelector('.expense-item-amount');
            
            if (nameEl && amountEl) {
                const name = nameEl.textContent.trim();
                const amountText = amountEl.textContent.replace(/[^\d.,]/g, '').replace(',', '.');
                const amount = parseFloat(amountText) || 0;
                
                if (name || amount > 0) {
                    expenses.push({ name, amount });
                }
            }
        });
        
        return expenses;
    }
    
    /**
     * Update metrics in summary panel
     */
    function updateMetrics(data) {
        const formatCurrency = (value, currency) => {
            return Number(value || 0).toLocaleString('ru-RU', { 
                minimumFractionDigits: 0, 
                maximumFractionDigits: 0 
            }) + ' ' + currency;
        };
        
        if (summaryTotalEl) {
            summaryTotalEl.textContent = formatCurrency(data.total, data.currency);
        }
        if (summaryTotalUnitEl) {
            summaryTotalUnitEl.textContent = data.currency === '¥' ? 'юаней' : 'рублей';
        }
        
        if (summaryProfitEl) {
            summaryProfitEl.textContent = formatCurrency(data.profit, data.currency);
        }
        if (summaryProfitUnitEl) {
            summaryProfitUnitEl.textContent = data.currency === '¥' ? 'юаней' : 'рублей';
        }
        
        if (summaryProfitabilityEl) {
            summaryProfitabilityEl.textContent = data.profitability + '%';
        }
        
        if (summaryTotalPurchaseCostEl) {
            summaryTotalPurchaseCostEl.textContent = formatCurrency(data.totalPurchaseCost, data.currency);
        }
        
        if (summaryLogisticsCostEl) {
            summaryLogisticsCostEl.textContent = formatCurrency(data.logistics, '₽');
        }
        
        if (summaryVATEl) {
            summaryVATEl.textContent = formatCurrency(data.vat, data.currency);
        }
        
        if (summaryDutyEl) {
            summaryDutyEl.textContent = formatCurrency(data.duty, data.currency);
        }
        
        if (summaryMarginEl) {
            summaryMarginEl.textContent = formatCurrency(data.margin, data.currency);
        }
        
        if (summaryFinalPriceRublesEl) {
            summaryFinalPriceRublesEl.textContent = formatCurrency(data.finalPriceRubles, '₽');
        }
        
        if (summaryExchangeRateEl) {
            summaryExchangeRateEl.textContent = `1 ¥ = ${data.exchangeRate} ₽`;
        }
    }
    
    /**
     * Update expenses list in summary panel
     */
    function updateExpensesList(expenses, isCN) {
        if (!summaryExpensesListEl) return;
        
        if (expenses.length === 0) {
            summaryExpensesListEl.innerHTML = '<p class="muted-text" style="font-size: 12px; margin: 0;">Нет дополнительных расходов</p>';
            return;
        }
        
        const currency = isCN ? '¥' : '₽';
        const formatCurrency = (value) => {
            return Number(value || 0).toLocaleString('ru-RU', { 
                minimumFractionDigits: 0, 
                maximumFractionDigits: 0 
            }) + ' ' + currency;
        };
        
        summaryExpensesListEl.innerHTML = expenses.map(exp => `
            <div class="expense-item-summary">
                <span class="expense-item-name">${escapeHtml(exp.name)}</span>
                <span class="expense-item-amount">${formatCurrency(exp.amount)}</span>
            </div>
        `).join('');
    }
    
    /**
     * Escape HTML
     */
    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
    
    /**
     * Watch for form changes
     */
    function watchFormChanges() {
        const form = document.getElementById('generateForm');
        if (!form) return;
        
        // Debounce function
        let updateTimeout;
        const debouncedUpdate = () => {
            clearTimeout(updateTimeout);
            updateTimeout = setTimeout(() => {
                calculateSummary();
            }, 300);
        };
        
        // Watch for changes in key fields
        const watchedSelectors = [
            '#margin_percent',
            '#logistics',
            '[data-field="cost_price"]',
            '[data-field="quantity"]',
            '[data-field="weight"]',
            '[data-field="duty_percent"]',
            'input[name="budget_region"]'
        ];
        
        watchedSelectors.forEach(selector => {
            const elements = document.querySelectorAll(selector);
            elements.forEach(el => {
                el.addEventListener('input', debouncedUpdate);
                el.addEventListener('change', debouncedUpdate);
            });
        });
        
        // Watch for position additions/removals
        const positionsContainer = document.getElementById('positionsContainer');
        if (positionsContainer) {
            const observer = new MutationObserver(() => {
                // Re-attach event listeners to new position fields
                setTimeout(() => {
                    watchedSelectors.forEach(selector => {
                        const elements = document.querySelectorAll(selector);
                        elements.forEach(el => {
                            if (!el.dataset.watched) {
                                el.dataset.watched = 'true';
                                el.addEventListener('input', debouncedUpdate);
                                el.addEventListener('change', debouncedUpdate);
                            }
                        });
                    });
                    debouncedUpdate();
                }, 100);
            });
            
            observer.observe(positionsContainer, {
                childList: true,
                subtree: true
            });
        }
        
        // Watch for expenses changes
        const expensesList = document.getElementById('expensesList');
        if (expensesList) {
            const observer = new MutationObserver(debouncedUpdate);
            observer.observe(expensesList, {
                childList: true,
                subtree: true
            });
        }
    }
    
    // Initialize when DOM is ready
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
    
    // Export for use in other scripts
    window.ProposalCalculator = {
        calculateSummary: calculateSummary,
        updateMetrics: updateMetrics
    };
})();
