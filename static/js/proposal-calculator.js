/**
 * Proposal Calculator - Reactive calculations for summary panel
 * Updates the right panel in real-time as form data changes
 */

(function() {
    'use strict';
    
    // Configuration
    let CONVERSION_RATE = null; // Будет загружен с сервера
    const DEFAULT_CONVERSION_RATE = 11.5; // Fallback значение из конфига
    const STORAGE_KEY = 'kp_generator_exchange_rate'; // Ключ для localStorage
    
    // Cache DOM elements
    let summaryTotalEl, summaryProfitEl;
    let summaryTotalUnitEl, summaryProfitUnitEl;
    let summaryTotalPurchaseCostEl, summaryLogisticsCostEl, summaryDutyEl;
    let summaryCreditEl, summaryBankGuaranteeEl;
    let summaryFinalPriceRublesEl, summaryExchangeRateEl, summaryExchangeRateResetBtn;
    let summaryExpensesListEl;
    
    /**
     * Загружает актуальный курс валют с сервера
     */
    async function loadExchangeRate() {
        const timeoutMs = 15000; // 15 секунд таймаут
        const startTime = Date.now();
        
        try {
            console.log('[ExchangeRate] Загрузка курса валют с сервера...');
            
            // Создаем промис с таймаутом через Promise.race
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
            
            let response;
            try {
                response = await fetch('/api/v1/exchange-rate', {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json',
                    },
                    signal: controller.signal
                });
                clearTimeout(timeoutId);
            } catch (fetchError) {
                clearTimeout(timeoutId);
                if (fetchError.name === 'AbortError') {
                    throw new Error(`Таймаут запроса (${timeoutMs}ms)`);
                }
                throw fetchError;
            }
            
            const elapsed = Date.now() - startTime;
            console.log(`[ExchangeRate] Ответ получен за ${elapsed}ms, статус: ${response.status}`);
            
            if (!response.ok) {
                const errorText = await response.text();
                console.error('[ExchangeRate] HTTP ошибка:', response.status, errorText.substring(0, 200));
                throw new Error(`HTTP error! status: ${response.status}, body: ${errorText.substring(0, 100)}`);
            }
            
            const data = await response.json();
            console.log('[ExchangeRate] Ответ сервера:', data);
            
            if (data && typeof data.rate === 'number') {
                const apiRate = data.rate;
                
                // Проверяем, есть ли сохраненный пользовательский курс
                const savedRate = getSavedExchangeRate();
                if (savedRate === null) {
                    // Нет сохраненного курса - используем курс из API
                    CONVERSION_RATE = apiRate;
                    console.log(`[ExchangeRate] ✓ Курс загружен из API: 1 ¥ = ${CONVERSION_RATE.toFixed(2)} ₽ (источник: ${data.source || 'unknown'})`);
                } else {
                    // Есть сохраненный курс - используем его, но показываем кнопку сброса
                    CONVERSION_RATE = savedRate;
                    console.log(`[ExchangeRate] Используется сохраненный курс: 1 ¥ = ${CONVERSION_RATE.toFixed(2)} ₽ (API курс: ${apiRate.toFixed(2)} ₽)`);
                    if (summaryExchangeRateResetBtn) {
                        summaryExchangeRateResetBtn.style.display = 'inline-flex';
                    }
                }
                
                // Обновляем отображение курса в UI
                updateExchangeRateDisplay(CONVERSION_RATE);
                return CONVERSION_RATE;
            } else {
                console.error('[ExchangeRate] Неверный формат ответа:', data);
                throw new Error('Invalid response format: rate is not a number');
            }
        } catch (error) {
            const elapsed = Date.now() - startTime;
            
            if (error.message && error.message.includes('Таймаут')) {
                console.error(`[ExchangeRate] ✗ Таймаут запроса (${elapsed}ms > ${timeoutMs}ms):`, error);
            } else if (error.name === 'TypeError' && (error.message.includes('fetch') || error.message.includes('Failed'))) {
                console.error('[ExchangeRate] ✗ Ошибка сети (нет подключения к серверу):', error);
            } else {
                console.error('[ExchangeRate] ✗ Ошибка загрузки курса:', error);
            }
            
            console.warn(`[ExchangeRate] Используется значение по умолчанию: ${DEFAULT_CONVERSION_RATE} ₽`);
            
            // Проверяем сохраненный курс
            const savedRate = getSavedExchangeRate();
            if (savedRate !== null) {
                CONVERSION_RATE = savedRate;
                console.log(`[ExchangeRate] Используется сохраненный курс: ${CONVERSION_RATE.toFixed(2)} ₽`);
            } else {
                // Используем значение по умолчанию
                CONVERSION_RATE = DEFAULT_CONVERSION_RATE;
                console.warn(`[ExchangeRate] Используется значение по умолчанию: ${DEFAULT_CONVERSION_RATE} ₽`);
            }
            
            updateExchangeRateDisplay(CONVERSION_RATE);
            return CONVERSION_RATE;
        }
    }
    
    /**
     * Получает сохраненный курс валют из localStorage
     */
    function getSavedExchangeRate() {
        try {
            const saved = localStorage.getItem(STORAGE_KEY);
            if (saved) {
                const rate = parseFloat(saved);
                if (!isNaN(rate) && rate > 0 && rate < 100) {
                    return rate;
                }
            }
        } catch (e) {
            console.warn('[ExchangeRate] Не удалось прочитать сохраненный курс:', e);
        }
        return null;
    }
    
    /**
     * Сохраняет курс валют в localStorage
     */
    function saveExchangeRate(rate) {
        try {
            localStorage.setItem(STORAGE_KEY, rate.toString());
            console.log(`[ExchangeRate] Курс сохранен: ${rate.toFixed(2)} ₽`);
        } catch (e) {
            console.warn('[ExchangeRate] Не удалось сохранить курс:', e);
        }
    }
    
    /**
     * Удаляет сохраненный курс из localStorage
     */
    function clearSavedExchangeRate() {
        try {
            localStorage.removeItem(STORAGE_KEY);
            console.log('[ExchangeRate] Сохраненный курс удален');
        } catch (e) {
            console.warn('[ExchangeRate] Не удалось удалить сохраненный курс:', e);
        }
    }
    
    /**
     * Обновляет отображение курса в UI
     */
    function updateExchangeRateDisplay(rate) {
        if (summaryExchangeRateEl) {
            summaryExchangeRateEl.value = rate.toFixed(2);
        }
    }
    
    /**
     * Получает текущий курс валют (загружает если еще не загружен)
     */
    function getConversionRate() {
        if (CONVERSION_RATE === null) {
            // Проверяем сохраненный курс
            const savedRate = getSavedExchangeRate();
            if (savedRate !== null) {
                CONVERSION_RATE = savedRate;
                return savedRate;
            }
            // Если курс еще не загружен, используем значение по умолчанию
            return DEFAULT_CONVERSION_RATE;
        }
        return CONVERSION_RATE;
    }
    
    // Initialize
    async function init() {
        // Cache summary panel elements
        summaryTotalEl = document.getElementById('summaryTotal');
        summaryProfitEl = document.getElementById('summaryProfit');
        summaryTotalUnitEl = document.getElementById('summaryTotalUnit');
        summaryProfitUnitEl = document.getElementById('summaryProfitUnit');
        
        summaryTotalPurchaseCostEl = document.getElementById('summaryTotalPurchaseCost');
        summaryLogisticsCostEl = document.getElementById('summaryLogisticsCost');
        summaryDutyEl = document.getElementById('summaryDuty');
        summaryCreditEl = document.getElementById('summaryCredit');
        summaryBankGuaranteeEl = document.getElementById('summaryBankGuarantee');
        
        summaryFinalPriceRublesEl = document.getElementById('summaryFinalPriceRubles');
        summaryExchangeRateEl = document.getElementById('summaryExchangeRate');
        summaryExchangeRateResetBtn = document.getElementById('exchangeRateReset');
        summaryExpensesListEl = document.getElementById('summaryExpensesList');
        
        // Настраиваем обработчики для редактирования курса
        setupExchangeRateHandlers();
        
        // Проверяем сохраненный курс перед загрузкой из API
        const savedRate = getSavedExchangeRate();
        if (savedRate !== null) {
            CONVERSION_RATE = savedRate;
            updateExchangeRateDisplay(savedRate);
            if (summaryExchangeRateResetBtn) {
                summaryExchangeRateResetBtn.style.display = 'inline-flex';
            }
        }
        
        // Загружаем курс валют перед первым расчетом
        await loadExchangeRate();
        
        // Initial calculation
        calculateSummary();
        
        // Watch for form changes
        watchFormChanges();
    }
    
    /**
     * Настраивает обработчики для редактирования курса валют
     */
    function setupExchangeRateHandlers() {
        if (!summaryExchangeRateEl) return;
        
        // Обработчик изменения значения курса
        summaryExchangeRateEl.addEventListener('change', function() {
            const newRate = parseFloat(this.value);
            if (!isNaN(newRate) && newRate > 0 && newRate < 100) {
                CONVERSION_RATE = newRate;
                saveExchangeRate(newRate);
                if (summaryExchangeRateResetBtn) {
                    summaryExchangeRateResetBtn.style.display = 'inline-flex';
                }
                // Пересчитываем итоги с новым курсом
                calculateSummary();
                console.log(`[ExchangeRate] Курс изменен пользователем: ${newRate.toFixed(2)} ₽`);
            } else {
                // Восстанавливаем предыдущее значение при неверном вводе
                this.value = CONVERSION_RATE ? CONVERSION_RATE.toFixed(2) : DEFAULT_CONVERSION_RATE.toFixed(2);
                alert('Введите корректное значение курса (от 0.01 до 100)');
            }
        });
        
        // Обработчик нажатия Enter
        summaryExchangeRateEl.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                this.blur(); // Убираем фокус, что вызовет событие change
            }
        });
        
        // Обработчик кнопки сброса на курс из API
        if (summaryExchangeRateResetBtn) {
            summaryExchangeRateResetBtn.addEventListener('click', async function() {
                console.log('[ExchangeRate] Сброс на курс из API...');
                clearSavedExchangeRate();
                this.style.display = 'none';
                // Загружаем актуальный курс из API
                await loadExchangeRate();
                // Пересчитываем итоги
                calculateSummary();
            });
        }
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
            
            // Get conversion rate
            const currentRate = getConversionRate();
            
            // Get additional expenses
            const additionalExpenses = getAdditionalExpenses();
            const totalExpenses = additionalExpenses.reduce((sum, exp) => sum + (parseFloat(exp.amount) || 0), 0);
            const expensesInCurrency = isCN ? (totalExpenses / currentRate) : totalExpenses;
            
            // Calculate final price
            const logisticsInCurrency = isCN ? (logisticsCost / currentRate) : logisticsCost;
            const dutyInCurrency = isCN ? totalDuty : (totalDuty * currentRate);
            const finalPrice = totalPurchaseCost + logisticsInCurrency + dutyInCurrency + marginAmount + expensesInCurrency;
            
            // Calculate final price in rubles
            const finalPriceInRubles = isCN ? (finalPrice * currentRate) : finalPrice;
            
            // Calculate profit (simplified - margin amount)
            const profit = marginAmount;
            
            // Calculate profitability
            const profitability = totalPurchaseCost > 0 
                ? ((profit / finalPrice) * 100).toFixed(2)
                : 0;
            
            // Calculate VAT (15% of margin)
            const vatAmount = marginAmount * 0.15;
            
            // Credit and bank guarantee (after duty in breakdown)
            const CREDIT_RATE = 0.16;
            const deliveryTime = parseFloat(document.getElementById('delivery_time')?.value) || 0;
            const paymentDays = 30; // из условий оплаты по умолчанию
            const financeCreditChecked = document.querySelector('input[name="finance_credit"]')?.checked;
            const financeBankGuaranteeChecked = document.querySelector('input[name="finance_bank_guarantee"]')?.checked;
            const creditBankDays = deliveryTime + paymentDays;
            
            // Кредит и банковская гарантия — в юанях для детализации
            let creditYuan = 0;
            if (financeCreditChecked && creditBankDays > 0) {
                const creditCostBase = totalPurchaseCost * (CREDIT_RATE / 365) * creditBankDays;
                creditYuan = isCN ? creditCostBase : (creditCostBase / currentRate);
            }
            let bankGuaranteeYuan = 0;
            if (financeBankGuaranteeChecked && creditBankDays > 0) {
                const revenueWithVat = finalPrice;
                const bankGuaranteeBase = revenueWithVat * (0.03 / 365) * creditBankDays;
                bankGuaranteeYuan = isCN ? bankGuaranteeBase : (bankGuaranteeBase / currentRate);
            }
            
            // Right panel is updated only from API after "Предварительный расчёт" (displayPreCalculationResults).
            // No updateMetrics / updateExpensesList here.

            // Синхронизация оценочного веса в блоке логистики
            const totalWeightValueEl = document.getElementById('totalWeightValue');
            if (totalWeightValueEl) {
                const formatted = Number.isFinite(totalWeight)
                    ? totalWeight.toLocaleString('ru-RU', { minimumFractionDigits: 0, maximumFractionDigits: 2 })
                    : '0';
                totalWeightValueEl.value = `${formatted} кг`;
            }

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
        
        if (summaryTotalPurchaseCostEl) {
            summaryTotalPurchaseCostEl.textContent = formatCurrency(data.totalPurchaseCost, data.currency);
        }
        
        if (summaryLogisticsCostEl) {
            summaryLogisticsCostEl.textContent = formatCurrency(data.logistics, '₽');
        }
        
        if (summaryDutyEl) {
            summaryDutyEl.textContent = formatCurrency(data.duty, data.currency);
        }
        
        if (summaryCreditEl) {
            summaryCreditEl.textContent = formatCurrency(data.credit ?? 0, '¥');
        }
        if (summaryBankGuaranteeEl) {
            summaryBankGuaranteeEl.textContent = formatCurrency(data.bankGuarantee ?? 0, '¥');
        }
        
        if (summaryFinalPriceRublesEl) {
            summaryFinalPriceRublesEl.textContent = formatCurrency(data.finalPriceRubles, '₽');
        }
        
        if (summaryExchangeRateEl) {
            const rate = data.exchangeRate || getConversionRate();
            updateExchangeRateDisplay(rate);
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
            '#delivery_time',
            '[data-field="cost_price"]',
            '[data-field="quantity"]',
            '[data-field="weight"]',
            '[data-field="duty_percent"]',
            'input[name="budget_region"]',
            'input[name="finance_credit"]',
            'input[name="finance_bank_guarantee"]'
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
        updateMetrics: updateMetrics,
        loadExchangeRate: loadExchangeRate,
        getConversionRate: getConversionRate
    };
})();
