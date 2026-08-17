
        let tableData = [];
        let dataTable = null;
        const reportColumns = JSON.parse(document.getElementById('report-columns-data').textContent);

        function showToast(message, type = 'info') {
            const container = document.getElementById('toastContainer');
            let alertClass = 'alert-info';
            if (type === 'success') alertClass = 'alert-success';
            if (type === 'warning') alertClass = 'alert-warning';
            if (type === 'error') alertClass = 'alert-error';

            const toastHtml = `
                <div class="alert ${alertClass} shadow-xl border-none text-xs font-semibold rounded-xl flex items-center gap-2">
                    <span>${message}</span>
                </div>
            `;
            const toastEl = $(toastHtml).appendTo(container);
            setTimeout(() => {
                toastEl.fadeOut(300, function () { $(this).remove(); });
            }, 3500);
        }

        window.CAMPAIGN_LABELS = {
            'Hiçbiri': 'Hiçbiri',
            'Avantajlı': 'Avantajlı Ürün',
            'Flaş': 'Flaş Ürün',
            'Plus': 'Plus Ürün',
            'Plus Ek İndirim %5': 'Plus Ek İndirim %5',
            'Plus Ek İndirim %10': 'Plus Ek İndirim %10',
            'Plus Ek İndirim %20': 'Plus Ek İndirim %20',
            'Karşılamalı Kampanya': 'Karşılamalı Kampanya'
        };
        var CAMPAIGN_LABELS = window.CAMPAIGN_LABELS;

        function asNumber(value) {
            if (value === null || value === undefined || value === '') return null;
            const number = Number(value);
            return Number.isFinite(number) ? number : null;
        }

        function round2(value) {
            return Math.floor((value * 100) + 0.5 + 1e-9) / 100;
        }

        function discountBetween(currentPrice, campaignPrice) {
            const current = asNumber(currentPrice);
            const campaign = asNumber(campaignPrice);
            if (current === null || campaign === null || current <= 0 || campaign <= 0) return [null, null];
            const amount = round2(Math.max(current - campaign, 0));
            return [amount, round2((amount / current) * 100)];
        }

        function selectedCampaignValues(row) {
            const selection = row.userSelection || 'Hiçbiri';
            const fields = {
                'Hiçbiri': ['Güncel Ürün Fiyatı (TL)', 'Güncel Ürün Kalan Net (TL)', 'Güncel Ürün Komisyon (%)'],
                'Avantajlı': ['Avantajlı Ürün Fiyatı (YENİ TSF) (TL)', 'Avantajlı Ürün Kalan Net (TL)', 'Avantajlı Ürün Komisyon (%)'],
                'Flaş': ['Flaş Ürün 24 Saat Fiyatı (TL)', 'Flaş Ürün Kalan Net (TL)', 'Flaş Ürün Komisyon (%)'],
                'Plus': ['Plus Fiyatı (TL)', 'Plus Net (TL)', 'Plus Komisyon (%)'],
                'Karşılamalı Kampanya': ['Karşılamalı Kampanya Fiyatı (TL)', 'Karşılamalı Kampanya Kalan Net (TL)', 'Karşılamalı Kampanya Komisyon (%)']
            };
            if (selection.startsWith('Plus Ek İndirim %')) {
                const rate = Number(selection.split('%').pop());
                return [
                    asNumber(row[`Plus Ek Fiyatı %${rate} (TL)`]),
                    asNumber(row[`Plus Ek Net %${rate} (TL)`]),
                    asNumber(row['Plus Ek Komisyon (%)'])
                ];
            }
            const counter = row.counter_evaluations?.[selection];
            if (counter) return [counter.price, counter.net, counter.rate].map(asNumber);
            const selected = fields[selection];
            return selected ? selected.map(field => asNumber(row[field])) : [null, null, null];
        }

        function buildReportRow(row) {
            const selection = row.userSelection || 'Hiçbiri';
            const currentPrice = asNumber(row['Güncel Ürün Fiyatı (TL)']);
            const [campaignPrice, campaignNet, campaignCommission] = selectedCampaignValues(row);
            const [appliedAmount, appliedPercent] = discountBetween(currentPrice, campaignPrice);
            const eligible = row['İndirim Uygulanabilir'] === 'Evet';
            const dipPrice = eligible ? asNumber(row['Düşülebilecek Dip Fiyat (TL)']) : null;
            const [availableAmount, availablePercent] = eligible
                ? discountBetween(currentPrice, dipPrice)
                : [null, null];
            const extraAmount = availableAmount === null || appliedAmount === null
                ? null
                : round2(Math.max(availableAmount - appliedAmount, 0));
            const extraPercent = extraAmount === null || currentPrice === null
                ? null
                : round2((extraAmount / currentPrice) * 100);
            const plusRate = selection.startsWith('Plus Ek İndirim %')
                ? Number(selection.split('%').pop())
                : 5;
            const plusExtraPrice = asNumber(row[`Plus Ek Fiyatı %${plusRate} (TL)`]);
            const campaignLabels = {
                'Avantajlı': 'Avantajlı Ürün',
                'Flaş': 'Flaş Ürün',
                'Plus': 'Plus Ürün'
            };
            return {
                'Barkod': row.Barkod,
                'Güncel Fiyat (TL)': currentPrice,
                'Güncel Net': asNumber(row['Güncel Ürün Kalan Net (TL)']),
                'Güncel Komisyon': asNumber(row['Güncel Ürün Komisyon (%)']),
                'Avantajlı Fiyat (TL)': asNumber(row['Avantajlı Ürün Fiyatı (YENİ TSF) (TL)']),
                'Avantajlı Net': asNumber(row['Avantajlı Ürün Kalan Net (TL)']),
                'Flaş Fiyat (TL)': asNumber(row['Flaş Ürün 24 Saat Fiyatı (TL)']),
                'Flaş Net': asNumber(row['Flaş Ürün Kalan Net (TL)']),
                'Plus Fiyat (TL)': asNumber(row['Plus Fiyatı (TL)']),
                'Plus Net': asNumber(row['Plus Net (TL)']),
                'Plus Ek İndirim Fiyat (TL)': plusExtraPrice,
                'Plus Ek İndirim Net': asNumber(row[`Plus Ek Net %${plusRate} (TL)`]),
                'Uygulanan Kampanya': campaignLabels[selection] || selection,
                'Uygulanan Kampanya Fiyat': campaignPrice,
                'Uygulanan Kampanya Net': campaignNet,
                'Uygulanan Kampanya Komisyon': campaignCommission,
                'Uygulanabilecek İndirim (TL)': availableAmount,
                'Uygulanabilecek İndirim (%)': availablePercent,
                'Uygulanan İndirim (TL)': appliedAmount,
                'Uygulanan İndirim (%)': appliedPercent,
                'Ekstra Uygulanabilir İndirim (TL)': extraAmount,
                'Ekstra Uygulanabilir İndirim (%)': extraPercent,
                'Hangisi Karlı?': row['Hangisi Daha Karlı?'],
                'Düşülebilecek Dip Fiyat (TL)': dipPrice
            };
        }

        function escapeHtml(value) {
            return $('<div>').text(value == null ? '' : String(value)).html();
        }

        function renderUploadStatuses(uploads) {
            document.querySelectorAll('input[id^="input_"][type="file"]').forEach(input => {
                const key = input.id.slice(6);
                const uploaded = uploads ? uploads[key] : null;
                input.dataset.uploaded = uploaded ? 'true' : 'false';
                const status = document.getElementById(`upload_status_${key}`);
                if (status) {
                    status.innerHTML = uploaded
                        ? `Yüklendi: ${escapeHtml(uploaded.original_name)} · ${escapeHtml(uploaded.uploaded_at_display)}`
                        : 'Henüz yüklenmedi';
                }
            });
        }

        function loadData() {
            fetch('/api/data')
                .then(res => res.json())
                .then(data => {
                    if (data.needs_calculation) {
                        $('#table-body').html(`<tr><td colspan="${reportColumns.length + 1}" style="text-align:center; padding:30px; font-size: 16px; color: #e74c3c;">
                            <strong>${data.message}</strong>
                        </td></tr>`);
                        return;
                    }
                    if (data.error) {
                        showToast(data.error, 'error');
                        return;
                    }
                    tableData = data.map((row, index) => ({
                        ...row,
                        userSelection: row['İlk Kampanya Seçimi'] || 'Hiçbiri',
                        _index: index,
                        checked: false
                    }));
                    renderTable();
                    updateBulkCampaignOptions();
                    if (data.length > 0) {
                        const hasKars = data.some(r => r['Karşılamalı Kampanya Eşleşme Durumu'] === 'Eşleşti');
                        toggleKarsilamaliOptions(hasKars);
                    }
                })
                .catch(err => {
                    showToast('Veriler yüklenirken hata oluştu: ' + err, 'error');
                });
        }

        function toggleKarsilamaliOptions(show) {
            const bulkOpt = $('#bulkCampaignSelect option[value="Karşılamalı Kampanya"]');
            if (show) bulkOpt.show();
            else bulkOpt.hide();
        }

        function calculateData() {
            const requiredInputs = Array.from(document.querySelectorAll('input[type="file"][data-required="true"]'));
            const missing = requiredInputs.find(
                input => input.files.length === 0 && input.getAttribute('data-uploaded') !== 'true'
            );
            if (missing) {
                showToast('Zorunlu iki Excel girdisini (Ürün Komisyon Tarifeleri ve Ürün Listesi) seçin.', 'warning');
                missing.focus();
                return;
            }
            const formData = new FormData();
            document.querySelectorAll('input[id^="input_"][type="file"]:not(#input_counter_multi):not(#input_plus_extra_multi)').forEach(input => {
                if (input.files.length) formData.append(input.id.slice(6), input.files[0]);
            });

            const counterConfigs = [];
            counterFilesState.forEach((item, idx) => {
                if (item.fileObj) {
                    formData.append(`counter_file_${idx}`, item.fileObj);
                }
                counterConfigs.push({
                    id: item.id,
                    filename: item.filename,
                    min_price: item.min_price,
                    discount_amount: item.discount_amount,
                    trendyol_percent: item.trendyol_percent,
                    expiry_date: item.expiry_date || ''
                });
            });
            formData.append('counter_configs_json', JSON.stringify(counterConfigs));

            const plusExtraConfigs = [];
            plusExtraFilesState.forEach((item, idx) => {
                if (item.fileObj) {
                    formData.append(`plus_extra_file_${idx}`, item.fileObj);
                }
                plusExtraConfigs.push({
                    id: item.id,
                    filename: item.filename,
                    rate: item.rate,
                    expiry_date: item.expiry_date || ''
                });
            });
            formData.append('plus_extra_configs_json', JSON.stringify(plusExtraConfigs));

            const singleExpiries = {};
            const singleKeys = ['discount', 'commission', 'current', 'advantage', 'flash', 'plus', 'muhasebe_avantaj', 'muhasebe_flas', 'muhasebe_plus'];
            singleKeys.forEach(k => {
                const el = document.getElementById('expiry_' + k);
                if (el && el.value) singleExpiries[k] = el.value;
            });
            formData.append('single_expiries_json', JSON.stringify(singleExpiries));

            document.getElementById('loadingOverlay').style.display = 'flex';

            fetch('/api/calculate', {
                method: 'POST',
                body: formData
            })
                .then(res => res.json())
                .then(data => {
                    document.getElementById('loadingOverlay').style.display = 'none';
                    if (data.success) {
                        renderUploadStatuses(data.uploads || {});
                        showToast("Veriler başarıyla hesaplandı.", 'success');
                        loadData();
                    } else {
                        showToast("Hesaplama hatası: " + data.message, 'error');
                    }
                })
                .catch(err => {
                    document.getElementById('loadingOverlay').style.display = 'none';
                    showToast("İstek başarısız oldu: " + err, 'error');
                });
        }

        function applyCampaigns(type) {
            let selections = {};
            tableData.forEach(row => {
                const stok = asNumber(row['Stok Adedi']);
                if (ignoreZeroStockChecked && stok !== null && stok === 0) {
                    selections[row.Barkod] = 'Hiçbiri';
                } else {
                    selections[row.Barkod] = row.userSelection || 'Hiçbiri';
                }
            });

            document.getElementById('loadingOverlay').style.display = 'flex';

            const payload = {
                target_type: type,
                selections: selections,
                ignore_zero_stock: ignoreZeroStockChecked,
                visibleColumns: reportColumns.filter((column, index) =>
                    !dataTable || dataTable.column(index + 1).visible()
                )
            };

            fetch('/api/apply', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            })
                .then(res => res.json())
                .then(data => {
                    if (data.success) {
                        const fileCount = data.generated_files ? data.generated_files.length : 0;
                        showToast(`<div><strong>Dosyalar Başarıyla Oluşturuldu!</strong><div class="mt-1 font-normal text-xs text-slate-100">${fileCount} adet Excel dosyası 'Çıktılar' klasörüne kaydedildi.</div></div>`, 'success');
                    } else {
                        showToast("İşlem hatası: " + data.message, 'error');
                    }
                })
                .catch(err => {
                    showToast("İstek atılırken hata oluştu: " + err, 'error');
                })
                .finally(() => {
                    document.getElementById('loadingOverlay').style.display = 'none';
                });
        }

        // Global kontrol durumu değişkenleri
        let customSearchVal = '';
        let customPageLenVal = 50;
        let customIndirimliChecked = false;
        let ignoreZeroStockChecked = true;

        // Initial load
        loadData();

        window.onStockFilterChange = function onStockFilterChange(checked) {
            ignoreZeroStockChecked = checked;
            $('#ignoreZeroStock').prop('checked', checked);
            if (dataTable) dataTable.draw();
        };

        // Remove any previous filter to prevent duplicates
        $.fn.dataTable.ext.search = $.fn.dataTable.ext.search.filter(
            fn => fn.name !== 'customTableFilter' && fn.name !== 'indirimliFilter'
        );

        const customTableFilter = function customTableFilter(settings, data, dataIndex, rowData) {
            const row = (tableData && tableData[dataIndex]) ? tableData[dataIndex] : rowData;
            if (!row) return true;

            // 1. Stok 0 kontrolü
            if (ignoreZeroStockChecked) {
                const stok = asNumber(row['Stok Adedi']);
                if (stok !== null && stok === 0) {
                    return false;
                }
            }

            // 2. Sadece İndirimli Göster (Güncel Fiyat > Dip Fiyat)
            if (customIndirimliChecked) {
                const currentPrice = asNumber(row['Güncel Ürün Fiyatı (TL)']);
                const dipPrice = asNumber(row['Düşülebilecek Dip Fiyat (TL)']);
                if (currentPrice === null || dipPrice === null || currentPrice <= dipPrice + 1e-6) {
                    return false;
                }
            }

            return true;
        };
        $.fn.dataTable.ext.search.push(customTableFilter);

        function renderColumnToggles() {
            const container = $('.top_col_container, .bottom_col_container');
            if (container.length === 0) return;
            container.empty();
            reportColumns.forEach((name, idx) => {
                const actualColIdx = idx + 1;
                const isVisible = dataTable ? dataTable.column(actualColIdx).visible() : true;
                const checkboxHtml = `
                    <label class="flex items-center gap-2 py-1 px-1.5 hover:bg-base-200 rounded-lg cursor-pointer text-base-content">
                        <input type="checkbox" class="checkbox checkbox-xs checkbox-primary col-toggle-cb" data-col="${actualColIdx}" ${isVisible ? 'checked' : ''} onchange="toggleColumnVisibility(${actualColIdx}, this.checked)">
                        <span class="font-medium text-xs">${escapeHtml(name)}</span>
                    </label>
                `;
                container.append(checkboxHtml);
            });
        }

        function toggleColumnVisibility(colIdx, isVisible) {
            if (dataTable) {
                dataTable.column(colIdx).visible(isVisible);
            }
        }

        function toggleAllColumns(isVisible) {
            if (!dataTable) return;
            for (let i = 1; i <= reportColumns.length; i++) {
                dataTable.column(i).visible(isVisible);
            }
            $('.col-toggle-cb').prop('checked', isVisible);
        }

        const initialCounterConfigs = [];
        const initialPlusExtraConfigs = [];

        let counterFilesState = Array.isArray(initialCounterConfigs) ? initialCounterConfigs.map(item => ({
            id: item.id || ('counter_' + Math.random().toString(36).substr(2, 9)),
            filename: item.filename || item.original_name || 'Karşılamalı Kampanya.xlsx',
            min_price: parseFloat(item.min_price) || 0,
            discount_amount: parseFloat(item.discount_amount) || 0,
            trendyol_percent: parseFloat(item.trendyol_percent) || 0,
            expiry_date: item.expiry_date || '',
            fileObj: null
        })) : [];

        let plusExtraFilesState = Array.isArray(initialPlusExtraConfigs) ? initialPlusExtraConfigs.map(item => ({
            id: item.id || ('plus_extra_' + Math.random().toString(36).substr(2, 9)),
            filename: item.filename || item.original_name || 'Plus Ek İndirim.xlsx',
            rate: parseFloat(item.rate) || 0,
            expiry_date: item.expiry_date || '',
            fileObj: null
        })) : [];

        $(document).ready(function () {
            renderCounterCards();
            renderPlusExtraCards();
            updateBulkCampaignOptions();
            setTimeout(renderSingleExpiryBadges, 100);
        });

        function parseCounterFilename(filename) {
            const pattern = /(\d+)\s*[-_]?tl[-_]?uzeri\s*[-_]?(\d+)\s*[-_]?tl[-_]?indirim\s*[-_]?(\d+)\s*[-_]?trendyol/i;
            const match = filename.match(pattern);
            if (match) {
                return {
                    min_price: parseFloat(match[1]) || 0,
                    discount_amount: parseFloat(match[2]) || 0,
                    trendyol_percent: parseFloat(match[3]) || 0
                };
            }
            return { min_price: 0, discount_amount: 0, trendyol_percent: 0 };
        }

        function parsePlusExtraFilename(filename) {
            const match = filename.match(/%?\s*(\d+)\s*%?/i);
            if (match) {
                return parseFloat(match[1]) || 0;
            }
            return 0;
        }

        const singleExpiriesState = {};

        window.onSingleExpiryChange = function onSingleExpiryChange(key, val) {
            singleExpiriesState[key] = val;
            renderSingleExpiryBadges();
            fetch('/api/save-expiry', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ single_expiries: { [key]: val } })
            }).catch(function () { });
        };

        function renderSingleExpiryBadges() {
            const keys = ['discount', 'commission', 'current', 'advantage', 'flash', 'plus', 'muhasebe_avantaj', 'muhasebe_flas', 'muhasebe_plus'];
            keys.forEach(k => {
                const inputEl = document.getElementById('expiry_' + k);
                const dateVal = inputEl ? inputEl.value : (singleExpiriesState[k] || '');
                if (dateVal) singleExpiriesState[k] = dateVal;
                const badgeContainer = document.getElementById('badge_single_' + k);
                if (badgeContainer) {
                    badgeContainer.innerHTML = getTimeRemainingBadge(dateVal);
                }
            });
        }

        function getTimeRemainingBadge(dateStr) {
            if (!dateStr) return '<span class="badge badge-ghost badge-xs text-base-content/40">Tarih Belirtilmedi</span>';
            const target = new Date(dateStr);
            const now = new Date();
            const diffMs = target.getTime() - now.getTime();
            if (isNaN(diffMs)) return '<span class="badge badge-ghost badge-xs text-base-content/40">Geçersiz Tarih</span>';

            const totalSecs = Math.floor(diffMs / 1000);
            const totalMins = Math.floor(diffMs / (1000 * 60));
            const totalHours = Math.floor(diffMs / (1000 * 60 * 60));
            const totalDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));

            if (diffMs <= 0 || totalSecs <= 0) return '<span class="badge badge-error badge-xs text-white font-bold animate-pulse">Süresi Doldu</span>';

            if (totalSecs < 60 && totalSecs > 0) {
                return `<span class="badge badge-error badge-xs text-white font-bold animate-pulse">Son ${totalSecs} Saniye</span>`;
            } else if (totalMins < 60 && totalSecs > 60) {
                return `<span class="badge badge-warning badge-xs font-bold">Son ${totalMins} Dakika</span>`;
            } else if (totalHours < 24 && totalMins > 60 && totalSecs > 3600) {
                return `<span class="badge badge-warning badge-xs font-bold">Son ${totalHours} Saat</span>`;
            } else if (totalDays >= 1) {
                return `<span class="badge badge-info badge-xs font-bold">Son ${totalDays} Gün</span>`;
            }
        }

        function updateAllBadgesDynamic() {
            renderSingleExpiryBadges();
            if (Array.isArray(counterFilesState)) {
                counterFilesState.forEach(item => {
                    const el = document.getElementById('badge_counter_' + item.id);
                    if (el) el.innerHTML = getTimeRemainingBadge(item.expiry_date);
                });
            }
            if (Array.isArray(plusExtraFilesState)) {
                plusExtraFilesState.forEach(item => {
                    const el = document.getElementById('badge_plus_extra_' + item.id);
                    if (el) el.innerHTML = getTimeRemainingBadge(item.expiry_date);
                });
            }
        }

        setInterval(updateAllBadgesDynamic, 1000);

        function handleMultiCounterUpload(files) {
            if (!files || !files.length) return;
            Array.from(files).forEach((file, idx) => {
                if (!file.name.toLowerCase().endsWith('.xlsx')) return;
                const parsed = parseCounterFilename(file.name);
                const item = {
                    id: 'counter_' + Date.now() + '_' + idx,
                    filename: file.name,
                    min_price: parsed.min_price,
                    discount_amount: parsed.discount_amount,
                    trendyol_percent: parsed.trendyol_percent,
                    expiry_date: '',
                    fileObj: file
                };
                counterFilesState.push(item);
            });
            renderSingleExpiryBadges();
            renderCounterCards();
            updateBulkCampaignOptions();
        }

        function removeCounterFile(id) {
            counterFilesState = counterFilesState.filter(c => c.id !== id);
            renderCounterCards();
            updateBulkCampaignOptions();
        }

        function renderCounterCards() {
            const container = $('#counterFilesContainer');
            container.empty();
            if (!counterFilesState.length) {
                container.html('<div class="text-xs text-base-content/50 italic bg-base-100/60 p-4 rounded-xl border border-dashed border-base-300 text-center">Henüz özel Karşılamalı Kampanya dosyası yüklenmedi.</div>');
                return;
            }
            counterFilesState.forEach((item) => {
                const badgeHtml = getTimeRemainingBadge(item.expiry_date);
                const cardHtml = `
                    <div class="bg-base-100 rounded-xl p-3.5 border border-base-300 shadow-2xs space-y-3">
                        <div class="flex items-center justify-between flex-wrap gap-2 pb-2 border-b border-base-200">
                            <div class="flex items-center gap-2">
                                <span class="w-2.5 h-2.5 rounded-full bg-primary"></span>
                                <span class="font-bold text-xs text-base-content">
                                    ${escapeHtml(item.filename)}
                                </span>
                                <span id="badge_counter_${item.id}">${badgeHtml}</span>
                            </div>
                            <button type="button" class="btn btn-ghost btn-xs text-error font-medium" onclick="removeCounterFile('${item.id}')">Kaldır</button>
                        </div>
                        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
                            <label class="form-control">
                                <span class="label-text text-[11px] font-semibold text-base-content/80">Kaç TL Üzeri</span>
                                <input type="number" min="0" step="0.01" value="${item.min_price}" class="input input-bordered input-xs focus:input-primary rounded-lg" onchange="updateCounterParam('${item.id}', 'min_price', this.value)">
                            </label>
                            <label class="form-control">
                                <span class="label-text text-[11px] font-semibold text-base-content/80">İndirim Tutarı (TL)</span>
                                <input type="number" min="0" step="0.01" value="${item.discount_amount}" class="input input-bordered input-xs focus:input-primary rounded-lg" onchange="updateCounterParam('${item.id}', 'discount_amount', this.value)">
                            </label>
                            <label class="form-control">
                                <span class="label-text text-[11px] font-semibold text-base-content/80">Trendyol Karşılama (%)</span>
                                <input type="number" min="0" max="100" step="0.01" value="${item.trendyol_percent}" class="input input-bordered input-xs focus:input-primary rounded-lg" onchange="updateCounterParam('${item.id}', 'trendyol_percent', this.value)">
                            </label>
                            <label class="form-control">
                                <span class="label-text text-[11px] font-semibold text-base-content/80">Bitiş Tarihi & Saati</span>
                                <input type="datetime-local" value="${item.expiry_date || ''}" class="input input-bordered input-xs focus:input-primary text-[11px] rounded-lg" onchange="updateCounterParam('${item.id}', 'expiry_date', this.value)">
                            </label>
                        </div>
                    </div>
                `;
                container.append(cardHtml);
            });
        }

        function updateCounterParam(id, key, val) {
            const item = counterFilesState.find(c => c.id === id);
            if (item) {
                item[key] = (key === 'expiry_date') ? val : (parseFloat(val) || 0);
                if (key === 'expiry_date') {
                    const el = document.getElementById('badge_counter_' + id);
                    if (el) el.innerHTML = getTimeRemainingBadge(val);
                    fetch('/api/save-expiry', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ counter_expiries: { [id]: val } })
                    }).catch(function () { });
                }
                updateBulkCampaignOptions();
            }
        }

        function handleMultiPlusExtraUpload(files) {
            if (!files || !files.length) return;
            Array.from(files).forEach((file, idx) => {
                if (!file.name.toLowerCase().endsWith('.xlsx')) return;
                const rate = parsePlusExtraFilename(file.name);
                const item = {
                    id: 'plus_extra_' + Date.now() + '_' + idx,
                    filename: file.name,
                    rate: rate,
                    expiry_date: '',
                    fileObj: file
                };
                plusExtraFilesState.push(item);
            });
            renderPlusExtraCards();
            updateBulkCampaignOptions();
        }

        function removePlusExtraFile(id) {
            plusExtraFilesState = plusExtraFilesState.filter(c => c.id !== id);
            renderPlusExtraCards();
            updateBulkCampaignOptions();
        }

        function renderPlusExtraCards() {
            const container = $('#plusExtraFilesContainer');
            container.empty();
            if (!plusExtraFilesState.length) {
                container.html('<div class="text-xs text-base-content/50 italic bg-base-100/60 p-4 rounded-xl border border-dashed border-base-300 text-center">Henüz özel Plus Ek İndirim dosyası yüklenmedi.</div>');
                return;
            }
            plusExtraFilesState.forEach((item) => {
                const badgeHtml = getTimeRemainingBadge(item.expiry_date);
                const cardHtml = `
                    <div class="bg-base-100 rounded-xl p-3.5 border border-base-300 shadow-2xs space-y-3">
                        <div class="flex items-center justify-between flex-wrap gap-2 pb-2 border-b border-base-200">
                            <div class="flex items-center gap-2">
                                <span class="w-2.5 h-2.5 rounded-full bg-secondary"></span>
                                <span class="font-bold text-xs text-base-content">
                                    ${escapeHtml(item.filename)}
                                </span>
                                <span id="badge_plus_extra_${item.id}">${badgeHtml}</span>
                            </div>
                            <button type="button" class="btn btn-ghost btn-xs text-error font-medium" onclick="removePlusExtraFile('${item.id}')">Kaldır</button>
                        </div>
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                            <label class="form-control">
                                <span class="label-text text-[11px] font-semibold text-base-content/80">Plus Müşteri Ek İndirim Oranı (%)</span>
                                <input type="number" min="0" max="100" step="0.01" value="${item.rate}" class="input input-bordered input-xs focus:input-secondary rounded-lg" onchange="updatePlusExtraParam('${item.id}', 'rate', this.value)">
                            </label>
                            <label class="form-control">
                                <span class="label-text text-[11px] font-semibold text-base-content/80">Bitiş Tarihi & Saati</span>
                                <input type="datetime-local" value="${item.expiry_date || ''}" class="input input-bordered input-xs focus:input-secondary text-[11px] rounded-lg" onchange="updatePlusExtraParam('${item.id}', 'expiry_date', this.value)">
                            </label>
                        </div>
                    </div>
                `;
                container.append(cardHtml);
            });
        }

        function updatePlusExtraParam(id, key, val) {
            const item = plusExtraFilesState.find(c => c.id === id);
            if (item) {
                item[key] = (key === 'expiry_date') ? val : (parseFloat(val) || 0);
                if (key === 'expiry_date') {
                    const el = document.getElementById('badge_plus_extra_' + id);
                    if (el) el.innerHTML = getTimeRemainingBadge(val);
                    fetch('/api/save-expiry', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ plus_extra_expiries: { [id]: val } })
                    }).catch(function () { });
                }
                updateBulkCampaignOptions();
            }
        }

        function updateBulkCampaignOptions() {
            const select = $('#bulkCampaignSelect');
            select.find('option[data-dynamic="true"]').remove();

            const dynamicCampaigns = new Set();
            if (Array.isArray(tableData)) {
                tableData.forEach(row => {
                    if (Array.isArray(row.eligible_campaigns)) {
                        row.eligible_campaigns.forEach(c => {
                            if (c !== 'Hiçbiri' && c !== 'Avantajlı' && c !== 'Flaş' && c !== 'Plus') {
                                dynamicCampaigns.add(c);
                            }
                        });
                    }
                });
            }

            if (Array.isArray(counterFilesState)) {
                counterFilesState.forEach((item, idx) => {
                    const minP = item.min_price;
                    const disc = item.discount_amount;
                    const label = minP > 0 ? `Karşılamalı (${minP} TL Üzeri / ${disc} TL İndirim)` : `Karşılamalı #${idx + 1}`;
                    dynamicCampaigns.add(label);
                });
            }

            if (Array.isArray(plusExtraFilesState)) {
                plusExtraFilesState.forEach((item, idx) => {
                    const rate = item.rate;
                    const label = rate > 0 ? `Plus Ek İndirim %${rate}` : `Plus Ek İndirim #${idx + 1}`;
                    dynamicCampaigns.add(label);
                });
            }

            const existingValues = new Set();
            select.find('option').each(function () {
                existingValues.add($(this).val());
            });

            dynamicCampaigns.forEach(cName => {
                if (!existingValues.has(cName)) {
                    const displayLabel = (CAMPAIGN_LABELS[cName] || cName) + ' Yap';
                    select.append(`<option value="${escapeHtml(cName)}" data-dynamic="true">${escapeHtml(displayLabel)}</option>`);
                    existingValues.add(cName);
                }
            });
        }

        function getUygulananKampanyaFilterValues() {
            const filterValues = new Set(['Hiçbiri', 'Avantajlı Ürün', 'Flaş Ürün', 'Plus Ürün']);
            if (Array.isArray(tableData)) {
                tableData.forEach(row => {
                    const selection = row.userSelection || 'Hiçbiri';
                    const reportLabel = CAMPAIGN_LABELS[selection] || selection;
                    if (reportLabel) filterValues.add(reportLabel);

                    if (Array.isArray(row.eligible_campaigns)) {
                        row.eligible_campaigns.forEach(cName => {
                            const label = CAMPAIGN_LABELS[cName] || cName;
                            if (label) filterValues.add(label);
                        });
                    }
                });
            }
            return Array.from(filterValues).sort();
        }

        function updateSelection(index, val) {
            const row = tableData[index];
            const eligible = Array.isArray(row.eligible_campaigns) ? row.eligible_campaigns : ['Hiçbiri'];

            if (!eligible.includes(val) && val !== 'Hiçbiri') {
                showToast(`Bu ürün (${row.Barkod}) seçilen kampanya listesinde bulunmadığı için katılamaz!`, 'warning');
                renderTable();
                return;
            }

            tableData[index].userSelection = val;
            if (dataTable) {
                const rowNode = dataTable.row(index).node();
                if (rowNode) {
                    dataTable.row(rowNode).data(tableData[index]).draw(false);
                }
            }
            saveUserSelections();
        }

        function applyBulkAction() {
            const targetCampaign = $('#bulkCampaignSelect').val();
            let appliedCount = 0;
            let skippedCount = 0;

            tableData.forEach((row) => {
                if (row.checked) {
                    const stok = asNumber(row['Stok Adedi']);
                    if (ignoreZeroStockChecked && stok !== null && stok === 0) {
                        row.userSelection = 'Hiçbiri';
                        skippedCount++;
                        return;
                    }
                    const eligible = Array.isArray(row.eligible_campaigns) ? row.eligible_campaigns : ['Hiçbiri'];
                    if (targetCampaign === 'Hiçbiri' || eligible.includes(targetCampaign)) {
                        row.userSelection = targetCampaign;
                        appliedCount++;
                    } else {
                        skippedCount++;
                    }
                }
            });

            resetUygulananKampanyaFilter();
            renderTable();
            saveUserSelections();

            if (appliedCount > 0 && skippedCount > 0) {
                showToast(`${appliedCount} seçili ürüne '${targetCampaign}' uygulandı. ${skippedCount} ürün kampanya listelerinde bulunmadığı için atlandı.`, 'info');
            } else if (appliedCount > 0) {
                showToast(`${appliedCount} ürüne '${targetCampaign}' kampanya seçimi uygulandı.`, 'success');
            } else if (skippedCount > 0) {
                showToast(`Seçilen ürünlerin hiçbirisi '${targetCampaign}' listelerinde yer almadığı için uygulanamadı.`, 'warning');
            } else {
                showToast('Lütfen önce işlem yapılacak ürünleri seçin.', 'warning');
            }
        }

        function campaignSelectionHtml(row) {
            const rowIndex = row._index;
            const eligible = Array.isArray(row.eligible_campaigns) ? row.eligible_campaigns : ['Hiçbiri'];

            let selectHtml = `<select class="select select-bordered select-xs w-full max-w-[210px] font-semibold campaign-select border-primary/40 focus:border-primary rounded-lg" onchange="updateSelection(${rowIndex}, this.value)">`;

            eligible.forEach(cName => {
                const isSelected = row.userSelection === cName;
                const displayLabel = CAMPAIGN_LABELS[cName] || cName;
                selectHtml += `<option value="${escapeHtml(cName)}" ${isSelected ? 'selected' : ''}>${escapeHtml(displayLabel)}</option>`;
            });

            selectHtml += '</select>';

            const badges = eligible
                .filter(c => c !== 'Hiçbiri')
                .map(label => `<span class="badge badge-outline badge-xs text-[10px] rounded-md">${escapeHtml(label)}</span>`)
                .join(' ');

            return `<div class="flex flex-col gap-1">${selectHtml}<div class="flex flex-wrap gap-1 max-w-[210px]">${badges || '<span class="text-[10px] text-base-content/40">Katılabilir kampanya yok</span>'}</div></div>`;
        }

        function renderTable() {
            $('.filters th').empty();
            if (dataTable) dataTable.destroy();
            const defaultHidden = new Set([
                'Güncel Fiyat (TL)', 'Güncel Net', 'Güncel Komisyon',
                'Avantajlı Fiyat (TL)', 'Avantajlı Net', 'Flaş Fiyat (TL)', 'Flaş Net',
                'Plus Fiyat (TL)', 'Plus Net', 'Plus Ek İndirim Fiyat (TL)', 'Plus Ek İndirim Net',
                'Karşılamalı Kampanya Fiyat (TL)', 'Karşılamalı Kampanya Net',
                'Uygulanan Kampanya Fiyat', 'Uygulanan Kampanya Net', 'Uygulanan Kampanya Komisyon'
            ]);
            const hiddenTargets = reportColumns
                .map((column, index) => defaultHidden.has(column) ? index + 1 : null)
                .filter(index => index !== null);
            const columns = [{
                data: null,
                orderable: false,
                render: function (data, type, row) {
                    return type === 'display'
                        ? `<input type="checkbox" class="checkbox checkbox-xs row-checkbox checkbox-primary" ${row.checked ? 'checked' : ''} onchange="updateCheck(${row._index}, this.checked)">`
                        : row.checked;
                }
            }].concat(reportColumns.map(column => ({
                data: null,
                orderable: true,
                className: column === 'Hangisi Karlı?' ? 'highlight text-center font-semibold' : '',
                render: function (data, type, row) {
                    const reportRow = buildReportRow(row);
                    if (column === 'Uygulanan Kampanya' && type === 'display') {
                        return campaignSelectionHtml(row);
                    }
                    const value = reportRow[column];
                    if (type === 'sort' || type === 'filter') return value == null ? '' : value;
                    if (value == null || value === '') return '-';
                    return typeof value === 'number' ? value.toFixed(2) : escapeHtml(value);
                }
            })));

            function buildControlBarHtml(prefix) {
                return `
                <div class="flex flex-wrap items-center justify-between gap-3 w-full">
                    <!-- Sol Taraf: Gösterim Sayısı + Arama + Sadece İndirimli Filtresi -->
                    <div class="flex items-center gap-2.5 flex-wrap">
                        <label class="flex items-center gap-1.5 text-xs font-semibold text-base-content/80">
                            <span>Göster:</span>
                            <select id="${prefix}_length" class="select select-bordered select-xs text-xs font-normal border-base-300 rounded-lg focus:select-primary" onchange="onCustomPageLengthChange(this.value)">
                                <option value="10">10</option>
                                <option value="25">25</option>
                                <option value="50" selected>50</option>
                                <option value="100">100</option>
                                <option value="250">250</option>
                                <option value="-1">Tümü</option>
                            </select>
                        </label>

                        <div class="relative min-w-[210px]">
                            <input type="text" id="${prefix}_search" class="input input-bordered input-xs w-full pl-7 text-xs border-base-300 rounded-lg focus:input-primary" placeholder="Ürün / Barkod Ara..." onkeyup="onCustomSearchKeyUp(this.value)">
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5 absolute left-2 top-2 text-base-content/40" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z" />
                            </svg>
                        </div>

                        <label class="cursor-pointer flex items-center gap-2 bg-base-200/60 px-2.5 py-1 rounded-lg border border-base-300/80 hover:bg-base-200 transition-colors">
                            <input type="checkbox" id="${prefix}_indirimli" class="checkbox checkbox-xs checkbox-primary" onchange="onCustomIndirimliFilterChange(this.checked)">
                            <span class="text-xs font-semibold text-base-content/90">Sadece İndirimli Göster</span>
                        </label>
                    </div>

                    <!-- Sağ Taraf: Sütun Seçimi + Kayıt Bilgisi + Paginasyon Düğmeleri -->
                    <div class="flex items-center gap-2.5 flex-wrap">
                        <div id="${prefix}_info" class="flex items-center px-2.5 py-1 rounded-lg border border-base-200 bg-base-200/40 text-[11px] font-semibold text-base-content/80 whitespace-nowrap">
                            -
                        </div>

                        <div class="dropdown dropdown-end">
                            <div tabindex="0" role="button" class="btn btn-xs btn-outline btn-primary gap-1.5 shadow-2xs font-semibold rounded-lg">
                                <svg xmlns="http://www.w3.org/2000/svg" class="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
                                </svg>
                                Sütun Seçimi
                            </div>
                            <div tabindex="0" class="dropdown-content z-[100] menu p-3 shadow-xl bg-base-100 rounded-xl border border-base-200 w-72 mt-2 space-y-1 text-xs">
                                <div class="font-bold text-base-content pb-2 border-b border-base-200 flex justify-between items-center">
                                    <span>Görünür Sütunlar</span>
                                    <div class="flex gap-2 text-[11px]">
                                        <button type="button" class="text-primary hover:underline font-semibold" onclick="toggleAllColumns(true)">Tümünü Seç</button>
                                        <span class="opacity-30">|</span>
                                        <button type="button" class="opacity-70 hover:underline" onclick="toggleAllColumns(false)">Temizle</button>
                                    </div>
                                </div>
                                <div class="max-h-64 overflow-y-auto space-y-1 pt-1 pr-1 ${prefix}_col_container"></div>
                            </div>
                        </div>

                        <div id="${prefix}_pagination" class="flex items-center gap-0.5"></div>
                    </div>
                </div>
            `;
            }

            function renderDualControls() {
                $('#tableTopHeaderControls').html(buildControlBarHtml('top'));
                $('#tableBottomFooterControls').html(buildControlBarHtml('bottom'));
                syncControlStates();
                renderColumnToggles();
            }

            function syncControlStates() {
                $('#top_search, #bottom_search').val(customSearchVal);
                $('#top_length, #bottom_length').val(customPageLenVal);
                $('#top_indirimli, #bottom_indirimli').prop('checked', customIndirimliChecked);
            }

            window.onCustomSearchKeyUp = function onCustomSearchKeyUp(val) {
                customSearchVal = val;
                $('#top_search, #bottom_search').val(val);
                if (dataTable) dataTable.search(val).draw();
            };

            window.onCustomPageLengthChange = function onCustomPageLengthChange(val) {
                customPageLenVal = parseInt(val, 10);
                $('#top_length, #bottom_length').val(val);
                if (dataTable) dataTable.page.len(customPageLenVal).draw();
            };

            window.onCustomIndirimliFilterChange = function onCustomIndirimliFilterChange(checked) {
                customIndirimliChecked = checked;
                $('#top_indirimli, #bottom_indirimli').prop('checked', checked);
                if (dataTable) dataTable.draw();
            };

            window.goToCustomPage = function (pageIdx) {
                if (dataTable) {
                    dataTable.page(pageIdx).draw('page');
                }
            };

            function renderCustomPagination() {
                if (!dataTable) return;
                const info = dataTable.page.info();
                const textInfo = `${info.recordsDisplay > 0 ? (info.start + 1) : 0} - ${info.end} / Toplam ${info.recordsDisplay} Ürün`;
                document.querySelectorAll('#top_info, #bottom_info').forEach(el => { el.textContent = textInfo; });

                const totalPages = info.pages;
                const currPage = info.page;

                if (totalPages <= 1) {
                    $('#top_pagination, #bottom_pagination').empty();
                    return;
                }

                let pagHtml = '';
                const prevDisabled = currPage === 0 ? 'btn-disabled opacity-40 pointer-events-none' : '';
                pagHtml += `<button type="button" class="btn btn-xs btn-ghost font-medium border border-base-300 rounded-lg ${prevDisabled}" onclick="window.goToCustomPage(${currPage - 1})">«</button>`;

                let startP = Math.max(0, currPage - 2);
                let endP = Math.min(totalPages - 1, currPage + 2);

                if (startP > 0) {
                    pagHtml += `<button type="button" class="btn btn-xs btn-ghost border border-base-300 rounded-lg" onclick="window.goToCustomPage(0)">1</button>`;
                    if (startP > 1) {
                        pagHtml += `<span class="text-xs text-base-content/40 px-1">…</span>`;
                    }
                }

                for (let i = startP; i <= endP; i++) {
                    const isCurrent = (i === currPage);
                    const btnClass = isCurrent
                        ? 'btn btn-xs btn-primary text-white font-bold rounded-lg shadow-2xs'
                        : 'btn btn-xs btn-ghost border border-base-300 rounded-lg hover:bg-base-200';
                    pagHtml += `<button type="button" class="${btnClass}" onclick="window.goToCustomPage(${i})">${i + 1}</button>`;
                }

                if (endP < totalPages - 1) {
                    if (endP < totalPages - 2) {
                        pagHtml += `<span class="text-xs text-base-content/40 px-1">…</span>`;
                    }
                    pagHtml += `<button type="button" class="btn btn-xs btn-ghost border border-base-300 rounded-lg" onclick="window.goToCustomPage(${totalPages - 1})">${totalPages}</button>`;
                }

                const nextDisabled = currPage >= totalPages - 1 ? 'btn-disabled opacity-40 pointer-events-none' : '';
                pagHtml += `<button type="button" class="btn btn-xs btn-ghost font-medium border border-base-300 rounded-lg ${nextDisabled}" onclick="window.goToCustomPage(${currPage + 1})">»</button>`;

                $('#top_pagination, #bottom_pagination').html(pagHtml);
            }

            dataTable = $('#campaignTable').DataTable({
                data: tableData,
                pageLength: customPageLenVal,
                deferRender: true,
                orderCellsTop: true,
                columnDefs: [
                    { targets: 0, orderable: false, searchable: false },
                    { targets: hiddenTargets, visible: false }
                ],
                dom: 'rt',
                language: {
                    url: 'https://cdn.datatables.net/plug-ins/1.13.6/i18n/tr.json'
                },
                createdRow: function (row, data, dataIndex) {
                    $(row).removeClass('row-avantajli row-flas row-plus');
                    if (data.userSelection === 'Avantajlı') {
                        $(row).addClass('row-avantajli');
                    } else if (data.userSelection === 'Flaş') {
                        $(row).addClass('row-flas');
                    } else if (data.userSelection === 'Plus') {
                        $(row).addClass('row-plus');
                    }
                },
                columns: columns,
                initComplete: function () {
                    renderDualControls();
                    renderCustomPagination();

                    const filterColumnNames = ['Uygulanan Kampanya', 'Hangisi Karlı?'];
                    filterColumnNames.forEach(columnName => {
                        const columnIndex = reportColumns.indexOf(columnName) + 1;
                        const column = this.api().column(columnIndex);
                        const filterCell = $('.filters th').filter(
                            (_, th) => th.dataset.reportColumn === columnName
                        );
                        const select = $('<select>', {
                            class: 'select select-bordered select-xs w-full max-w-[160px] focus:outline-none rounded-lg',
                            'aria-label': `${columnName} filtresi`
                        })
                            .append($('<option>', { value: '', text: 'Tümü' }))
                            .appendTo(filterCell)
                            .on('change', function () {
                                const value = $.fn.dataTable.util.escapeRegex($(this).val());
                                column.search(value ? `^${value}$` : '', true, false).draw();
                            });
                        const values = columnName === 'Uygulanan Kampanya'
                            ? getUygulananKampanyaFilterValues()
                            : [...new Set(tableData.map(row => row['Hangisi Daha Karlı?']).filter(Boolean))].sort();
                        values.forEach(value => select.append($('<option>', { value, text: value })));
                    });
                },
                drawCallback: function () {
                    renderCustomPagination();
                }
            });
        }

        function toggleAll(source) {
            const isChecked = source.checked;
            if (dataTable) {
                const filteredData = dataTable.rows({ search: 'applied' }).data();
                for (let i = 0; i < filteredData.length; i++) {
                    const row = filteredData[i];
                    if (row) {
                        row.checked = isChecked;
                    }
                }
                dataTable.rows().invalidate().draw(false);
            } else {
                tableData.forEach(row => { row.checked = isChecked; });
            }
        }

        function updateCheck(index, checked) {
            tableData[index].checked = checked;
        }

        function resetUygulananKampanyaFilter() {
            if (dataTable) {
                const columnIndex = reportColumns.indexOf('Uygulanan Kampanya') + 1;
                if (columnIndex > 0) {
                    dataTable.column(columnIndex).search('');
                    $('.filters th[data-report-column="Uygulanan Kampanya"] select').val('');
                }
            }
        }

        function saveUserSelections() {
            const selections = {};
            if (Array.isArray(tableData)) {
                tableData.forEach(row => {
                    if (row.Barkod && row.userSelection) {
                        selections[row.Barkod] = row.userSelection;
                    }
                });
            }
            fetch('/api/save-selections', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ selections: selections })
            }).catch(err => console.error('Seçimler kaydedilirken hata:', err));
        }

        function autoSelect() {
            let selectedCount = 0;
            tableData.forEach(row => {
                const stok = asNumber(row['Stok Adedi']);
                if (ignoreZeroStockChecked && stok !== null && stok === 0) {
                    row.userSelection = 'Hiçbiri';
                } else {
                    const rec = row['Önerilen Kampanya'] || 'Hiçbiri';
                    row.userSelection = rec;
                    if (row.userSelection !== 'Hiçbiri') selectedCount++;
                }
            });
            if (dataTable) dataTable.rows().invalidate().draw(false);
            saveUserSelections();
            showToast(`${selectedCount} ürüne en yüksek netli kampanya seçimi uygulandı.`, selectedCount ? 'success' : 'warning');
        }

        function clearFilters() {
            if (dataTable) {
                dataTable.search('').columns().search('').draw();
                $('.filters select').val('');
            }
        }

        function clearSelections() {
            tableData.forEach(row => {
                row.userSelection = 'Hiçbiri';
            });
            resetUygulananKampanyaFilter();
            if (dataTable) dataTable.rows().invalidate().draw(false);
            saveUserSelections();
            showToast("Tüm seçimler temizlendi.", 'success');
        }

        function addRuleRow() {
            const container = document.getElementById('rulesContainer');
            const rowId = 'rule_' + Date.now();

            const html = `
                <div class="flex flex-wrap items-center gap-3 p-3 bg-base-200/50 rounded-xl border border-base-200 rule-row" id="${rowId}">
                    <select class="select select-bordered select-sm rule-type focus:outline-none rounded-lg" onchange="changeRuleType('${rowId}')">
                        <option value="kiyaslama">Kampanyalar Arası Kıyaslama (Avantajlı vs Flaş)</option>
                        <option value="guncel">Güncel Satış ile Kıyaslama (Avantajlı/Flaş vs Güncel Net)</option>
                    </select>
                    
                    <div class="rule-content kiyaslama-content flex items-center gap-2 flex-wrap">
                        <span class="text-xs font-semibold">Eğer</span>
                        <select class="select select-bordered select-sm rule-target focus:outline-none rounded-lg">
                            <option value="Avantajlı">Avantajlı</option>
                            <option value="Flaş">Flaş</option>
                        </select>
                        <span class="text-xs font-semibold">kârı,</span>
                        <select class="select select-bordered select-sm rule-compare focus:outline-none rounded-lg">
                            <option value="Flaş">Flaş</option>
                            <option value="Avantajlı">Avantajlı</option>
                        </select>
                        <span class="text-xs font-semibold">kârından en fazla</span>
                        <label class="input input-bordered input-sm flex items-center gap-2 w-28 rounded-lg">
                            <input type="number" class="rule-percent w-full" placeholder="Örn: 5">
                            <span class="opacity-50 text-xs">%</span>
                        </label>
                        <span class="text-xs font-semibold">az ise seçimi</span>
                        <select class="select select-bordered select-sm rule-action focus:outline-none rounded-lg">
                            <option value="Aynı">O Kampanya (Avantajlı/Flaş) Yap</option>
                            <option value="Diğeri">Diğer Kampanya Yap</option>
                            <option value="Hiçbiri">Hiçbiri Yap</option>
                        </select>
                    </div>

                    <div class="rule-content guncel-content items-center gap-2 flex-wrap" style="display:none;">
                        <span class="text-xs font-semibold">Eğer</span>
                        <select class="select select-bordered select-sm rule-target-guncel focus:outline-none rounded-lg">
                            <option value="Avantajlı">Avantajlı</option>
                            <option value="Flaş">Flaş</option>
                        </select>
                        <span class="text-xs font-semibold">kârı, GÜNCEL NET kârdan</span>
                        <select class="select select-bordered select-sm rule-operator-guncel focus:outline-none rounded-lg">
                            <option value="fazla">Fazla (Veya Eşit) ise</option>
                            <option value="az">Maksimum % düşüşte ise</option>
                        </select>
                        <label class="input input-bordered input-sm flex items-center gap-2 w-28 rounded-lg">
                            <input type="number" class="rule-percent-guncel w-full" placeholder="Örn: 10">
                            <span class="opacity-50 text-xs">%</span>
                        </label>
                        <span class="text-xs font-semibold">ise seçimi</span>
                        <select class="select select-bordered select-sm rule-action-guncel focus:outline-none rounded-lg">
                            <option value="Aynı">O Kampanya (Avantajlı/Flaş) Yap</option>
                            <option value="Hiçbiri">Hiçbiri Yap</option>
                        </select>
                    </div>
                    
                    <div class="flex-1 flex justify-end">
                        <button class="btn btn-square btn-ghost btn-xs text-error hover:bg-error/20 rounded-lg" onclick="document.getElementById('${rowId}').remove()">
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 0-1-1h-4a1 1 0 0-1 1v3M4 7h16" /></svg>
                        </button>
                    </div>
                </div>
            `;
            container.insertAdjacentHTML('beforeend', html);
        }

        function changeRuleType(rowId) {
            const row = document.getElementById(rowId);
            const type = row.querySelector('.rule-type').value;
            if (type === 'kiyaslama') {
                row.querySelector('.kiyaslama-content').style.display = 'flex';
                row.querySelector('.guncel-content').style.display = 'none';
            } else {
                row.querySelector('.kiyaslama-content').style.display = 'none';
                row.querySelector('.guncel-content').style.display = 'flex';
            }
        }
    