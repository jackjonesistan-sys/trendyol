
        let tableData = [];
        let dataTable = null;
        const reportColumns = JSON.parse('[]');

        function showToast(message, type = 'info') {
            const container = document.getElementById('toastContainer');
            let alertClass = 'bg-blue-600 text-white border-none shadow-xl';
            if (type === 'success') alertClass = 'bg-green-600 text-white border-none shadow-xl';
            if (type === 'warning') alertClass = 'bg-yellow-500 text-white border-none shadow-xl';
            if (type === 'error') alertClass = 'bg-red-600 text-white border-none shadow-xl';

            const toastHtml = `
                <div class="alert ${alertClass} mb-2 flex flex-row items-center gap-3 transition-all duration-300 rounded-lg px-4 py-3 font-medium">
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
                const basePrice = asNumber(row[`Plus Ek Fiyatı %${rate} (TL)`]);
                return [
                    basePrice === null ? null : round2(basePrice * (1 - rate / 100)),
                    asNumber(row[`Plus Ek Net %${rate} (TL)`]),
                    asNumber(row['Plus Ek Komisyon (%)'])
                ];
            }
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
            const plusBasePrice = asNumber(row[`Plus Ek Fiyatı %${plusRate} (TL)`]);
            const plusExtraPrice = plusBasePrice === null
                ? null
                : round2(plusBasePrice * (1 - plusRate / 100));
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
                input => input.files.length === 0 && input.dataset.uploaded !== 'true'
            );
            if (missing) {
                showToast('Zorunlu üç Excel girdisini seçin.', 'warning');
                missing.focus();
                return;
            }
            const formData = new FormData();
            document.querySelectorAll('input[id^="input_"][type="file"]:not(#input_counter_multi)').forEach(input => {
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
                    trendyol_percent: item.trendyol_percent
                });
            });
            formData.append('counter_configs_json', JSON.stringify(counterConfigs));

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
                selections[row.Barkod] = row.userSelection || 'Hiçbiri';
            });

            document.getElementById('loadingOverlay').style.display = 'flex';

            const payload = {
                target_type: type,
                selections: selections,
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

        // Initial load
        loadData();

        // Remove any previous filter to prevent duplicates
        $.fn.dataTable.ext.search = $.fn.dataTable.ext.search.filter(
            fn => fn.name !== 'indirimliFilter'
        );

        const indirimliFilter = function indirimliFilter(settings, data, dataIndex, rowData) {
            let isChecked = $('.only-indirimli-filter-cb').is(':checked');
            if (!isChecked) return true;

            return rowData['İndirim Uygulanabilir'] === 'Evet';
        };
        $.fn.dataTable.ext.search.push(indirimliFilter);

        function renderColumnToggles() {
            const container = $('#columnToggleContainer');
            if (container.length === 0) return;
            container.empty();
            reportColumns.forEach((name, idx) => {
                const actualColIdx = idx + 1;
                const isVisible = dataTable ? dataTable.column(actualColIdx).visible() : true;
                const checkboxHtml = `
                    <label class="flex items-center gap-2 py-1 px-1.5 hover:bg-slate-50 rounded cursor-pointer text-gray-700">
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

        let counterFilesState = []; // [{ id, filename, label, min_price, discount_amount, trendyol_percent, fileObj }]

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
                    fileObj: file
                };
                counterFilesState.push(item);
            });
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
                container.html('<div class="text-xs text-gray-400 italic bg-gray-50 p-3 rounded-lg border border-dashed border-gray-200 text-center">Henüz özel Karşılamalı Kampanya dosyası yüklenmedi.</div>');
                return;
            }
            counterFilesState.forEach((item, idx) => {
                const cardHtml = `
                    <div class="bg-indigo-50/60 rounded-xl p-3 border border-indigo-100 space-y-2">
                        <div class="flex items-center justify-between">
                            <span class="font-semibold text-xs text-indigo-900 flex items-center gap-1.5">
                                <span class="w-2 h-2 rounded-full bg-indigo-500"></span>
                                ${escapeHtml(item.filename)}
                            </span>
                            <button type="button" class="btn btn-ghost btn-xs text-red-500 hover:bg-red-50" onclick="removeCounterFile('${item.id}')">Kaldır</button>
                        </div>
                        <div class="grid grid-cols-1 sm:grid-cols-3 gap-2">
                            <label class="form-control">
                                <span class="label-text text-[11px] text-indigo-800">Kaç TL Üzeri</span>
                                <input type="number" min="0" step="0.01" value="${item.min_price}" class="input input-bordered input-xs" onchange="updateCounterParam('${item.id}', 'min_price', this.value)">
                            </label>
                            <label class="form-control">
                                <span class="label-text text-[11px] text-indigo-800">İndirim Tutarı (TL)</span>
                                <input type="number" min="0" step="0.01" value="${item.discount_amount}" class="input input-bordered input-xs" onchange="updateCounterParam('${item.id}', 'discount_amount', this.value)">
                            </label>
                            <label class="form-control">
                                <span class="label-text text-[11px] text-indigo-800">Trendyol Karşılama (%)</span>
                                <input type="number" min="0" max="100" step="0.01" value="${item.trendyol_percent}" class="input input-bordered input-xs" onchange="updateCounterParam('${item.id}', 'trendyol_percent', this.value)">
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
                item[key] = parseFloat(val) || 0;
                updateBulkCampaignOptions();
            }
        }

        function updateBulkCampaignOptions() {
            const select = $('#bulkCampaignSelect');
            select.find('option[data-counter="true"]').remove();
            
            counterFilesState.forEach((item, idx) => {
                const minP = item.min_price;
                const disc = item.discount_amount;
                const label = minP > 0 ? `Karşılamalı (${minP} TL Üzeri / ${disc} TL İndirim)` : `Karşılamalı #${idx+1}`;
                select.append(`<option value="${escapeHtml(label)}" data-counter="true">${escapeHtml(label)} Yap</option>`);
            });
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
        }

        function applyBulkAction() {
            const targetCampaign = $('#bulkCampaignSelect').val();
            let appliedCount = 0;
            let skippedCount = 0;

            tableData.forEach((row) => {
                if (row.checked) {
                    const eligible = Array.isArray(row.eligible_campaigns) ? row.eligible_campaigns : ['Hiçbiri'];
                    if (targetCampaign === 'Hiçbiri' || eligible.includes(targetCampaign)) {
                        row.userSelection = targetCampaign;
                        appliedCount++;
                    } else {
                        skippedCount++;
                    }
                }
            });

            renderTable();

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

        function getEligibleCampaigns(row) {
            let eligible = row.eligible_campaigns;
            if (typeof eligible === 'string') {
                try { eligible = JSON.parse(eligible); } catch (e) { eligible = null; }
            }
            if (!Array.isArray(eligible) || eligible.length === 0) {
                const applicableStr = String(row['Uygulanabilir Kampanyalar'] || '');
                eligible = ['Hiçbiri'];
                if (applicableStr.includes('Avantajlı') || (row['Avantajlı Fiyat (TL)'] && Number(row['Avantajlı Fiyat (TL)']) > 0)) eligible.push('Avantajlı');
                if (applicableStr.includes('Flaş') || (row['Flaş Fiyat (TL)'] && Number(row['Flaş Fiyat (TL)']) > 0)) eligible.push('Flaş');
                if (applicableStr.includes('Plus') || (row['Plus Fiyat (TL)'] && Number(row['Plus Fiyat (TL)']) > 0)) eligible.push('Plus');
                if (applicableStr.includes('Plus Ek İndirim') || (row['Plus Ek İndirim Fiyat (TL)'] && Number(row['Plus Ek İndirim Fiyat (TL)']) > 0)) {
                    eligible.push('Plus Ek İndirim %5', 'Plus Ek İndirim %10', 'Plus Ek İndirim %20');
                }
                if (applicableStr.includes('Karşılamalı Kampanya')) eligible.push('Karşılamalı Kampanya');
            }
            if (!eligible.includes('Hiçbiri')) eligible.unshift('Hiçbiri');
            if (row.userSelection && row.userSelection !== 'Hiçbiri' && !eligible.includes(row.userSelection)) {
                eligible.push(row.userSelection);
            }
            return Array.from(new Set(eligible));
        }

        function campaignSelectionHtml(row) {
            const rowIndex = row._index;
            const eligible = getEligibleCampaigns(row);
            
            let selectHtml = `<select class="select select-bordered select-xs w-full max-w-[210px] font-semibold campaign-select border-primary/40 focus:border-primary" onchange="updateSelection(${rowIndex}, this.value)">`;
            
            eligible.forEach(cName => {
                const isSelected = row.userSelection === cName;
                const displayLabel = CAMPAIGN_LABELS[cName] || cName;
                selectHtml += `<option value="${escapeHtml(cName)}" ${isSelected ? 'selected' : ''}>${escapeHtml(displayLabel)}</option>`;
            });
            
            selectHtml += '</select>';
            
            const badges = eligible
                .filter(c => c !== 'Hiçbiri')
                .map(label => `<span class="badge badge-outline badge-xs text-[10px]">${escapeHtml(label)}</span>`)
                .join(' ');
                
            return `<div class="flex flex-col gap-1">${selectHtml}<div class="flex flex-wrap gap-1 max-w-[210px]">${badges || '<span class="text-[10px] text-gray-400">Katılabilir kampanya yok</span>'}</div></div>`;
        }

        function renderTable() {
            $('.filters th').empty();
            if (dataTable) dataTable.destroy();
            const defaultHidden = new Set([
                'Güncel Fiyat (TL)', 'Güncel Net', 'Güncel Komisyon',
                'Avantajlı Fiyat (TL)', 'Avantajlı Net', 'Flaş Fiyat (TL)', 'Flaş Net',
                'Plus Fiyat (TL)', 'Plus Net', 'Plus Ek İndirim Fiyat (TL)', 'Plus Ek İndirim Net',
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
                        ? `<input type="checkbox" class="checkbox checkbox-xs row-checkbox" ${row.checked ? 'checked' : ''} onchange="updateCheck(${row._index}, this.checked)">`
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
            dataTable = $('#campaignTable').DataTable({
                data: tableData,
                pageLength: 50,
                deferRender: true,
                orderCellsTop: true,
                columnDefs: [{ targets: hiddenTargets, visible: false }],
                dom: '<"flex flex-wrap justify-between items-center bg-gray-100 p-4 rounded-lg mb-4 border border-gray-200 shadow-sm"<"flex items-center gap-6"l f <"custom-filter-checkbox">><"custom-col-vis">>rt<"flex flex-wrap justify-between items-center bg-gray-100 p-4 rounded-lg mt-4 border border-gray-200 shadow-sm"<"flex items-center gap-6"l i><"flex items-center gap-4 custom-filter-checkbox-bottom">p>',
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
                    let checkboxHtml = `
                        <label class="cursor-pointer label gap-2 py-0">
                            <input type="checkbox" class="checkbox checkbox-sm checkbox-primary only-indirimli-filter-cb">
                            <span class="label-text font-medium text-sm">Sadece İndirim Uygulanabilecekleri Göster</span>
                        </label>
                    `;
                    $('.custom-filter-checkbox, .custom-filter-checkbox-bottom').html(checkboxHtml);

                    let colVisHtml = `
                        <div class="dropdown dropdown-end">
                            <div tabindex="0" role="button" class="btn btn-sm btn-outline border-gray-300 bg-white hover:bg-gray-100 text-gray-700 shadow-xs font-semibold gap-1.5">
                                <svg xmlns="http://www.w3.org/2000/svg" class="h-4 w-4 text-indigo-600" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
                                </svg>
                                Sütun Seçimi
                            </div>
                            <div tabindex="0" class="dropdown-content z-[100] menu p-3 shadow-2xl bg-white rounded-xl border border-gray-200 w-72 mt-2 space-y-1 text-xs">
                                <div class="font-bold text-gray-800 pb-2 border-b border-gray-100 flex justify-between items-center">
                                    <span>Görünür Sütunlar</span>
                                    <div class="flex gap-2 text-[11px]">
                                        <button type="button" class="text-indigo-600 hover:underline font-semibold" onclick="toggleAllColumns(true)">Tümünü Seç</button>
                                        <span class="text-gray-300">|</span>
                                        <button type="button" class="text-gray-500 hover:underline" onclick="toggleAllColumns(false)">Temizle</button>
                                    </div>
                                </div>
                                <div class="max-h-64 overflow-y-auto space-y-1 pt-1 pr-1" id="columnToggleContainer"></div>
                            </div>
                        </div>
                    `;
                    $('.custom-col-vis').html(colVisHtml);
                    renderColumnToggles();

                    $('.only-indirimli-filter-cb').on('change', function () {
                        let isChecked = $(this).is(':checked');
                        $('.only-indirimli-filter-cb').prop('checked', isChecked);
                        if (dataTable) dataTable.draw();
                    });

                    const filterColumnNames = ['Uygulanan Kampanya', 'Hangisi Karlı?'];
                    filterColumnNames.forEach(columnName => {
                        const columnIndex = reportColumns.indexOf(columnName) + 1;
                        const column = this.api().column(columnIndex);
                        const filterCell = $('.filters th').filter(
                            (_, th) => th.dataset.reportColumn === columnName
                        );
                        const select = $('<select>', {
                            class: 'select select-bordered select-xs w-full max-w-[160px] focus:outline-none',
                            'aria-label': `${columnName} filtresi`
                        })
                            .append($('<option>', { value: '', text: 'Tümü' }))
                            .appendTo(filterCell)
                            .on('change', function () {
                                const value = $.fn.dataTable.util.escapeRegex($(this).val());
                                column.search(value ? `^${value}$` : '', true, false).draw();
                            });
                        const values = columnName === 'Uygulanan Kampanya'
                            ? ['Hiçbiri', 'Avantajlı Ürün', 'Flaş Ürün', 'Plus Ürün', 'Plus Ek İndirim %5', 'Plus Ek İndirim %10', 'Plus Ek İndirim %20', 'Karşılamalı Kampanya']
                            : [...new Set(tableData.map(row => row['Hangisi Daha Karlı?']).filter(Boolean))].sort();
                        values.forEach(value => select.append($('<option>', { value, text: value })));
                    });
                }
            });
        }

        function toggleAll(source) {
            const isChecked = source.checked;
            tableData.forEach(row => { row.checked = isChecked; });
            if (dataTable) dataTable.rows().invalidate().draw(false);
        }

        function updateCheck(index, checked) {
            tableData[index].checked = checked;
        }

        function updateSelection(index, val) {
            tableData[index].userSelection = val;
            if (dataTable) dataTable.rows().invalidate().draw(false);
        }

        function autoSelect() {
            let selectedCount = 0;
            tableData.forEach(row => {
                row.userSelection = row['İlk Kampanya Seçimi'] || 'Hiçbiri';
                if (row.userSelection !== 'Hiçbiri') selectedCount++;
            });
            if (dataTable) dataTable.rows().invalidate().draw(false);
            showToast(`${selectedCount} ürüne en yüksek netli kampanya seçimi uygulandı.`, selectedCount ? 'success' : 'warning');
        }

        function applyBulkAction() {
            const targetAction = document.getElementById('bulkCampaignSelect').value;
            let appliedCount = 0;
            let skippedBarcodes = [];

            tableData.forEach((row, index) => {
                if (row.checked) {
                    let canApply = false;

                    if (targetAction === 'Hiçbiri') {
                        canApply = true;
                    } else {
                        const applicable = new Set(String(row['Uygulanabilir Kampanyalar'] || '').split(', '));
                        const targetGroup = targetAction.startsWith('Plus Ek İndirim')
                            ? 'Plus Ek İndirim'
                            : targetAction;
                        canApply = applicable.has(targetGroup);
                    }

                    if (canApply) {
                        row.userSelection = targetAction;
                        appliedCount++;
                    } else {
                        skippedBarcodes.push(row['Barkod']);
                    }
                }
            });

            if (dataTable) dataTable.rows().invalidate().draw(false);

            if (skippedBarcodes.length > 0) {
                showToast(`${appliedCount} seçili ürüne '${targetAction}' uygulandı. Ancak ${skippedBarcodes.length} ürüne uygulanamadı.`, 'warning');
            } else if (appliedCount > 0) {
                showToast(`Toplu eylem ${appliedCount} ürüne başarıyla uygulandı.`, 'success');
            } else {
                showToast("Lütfen işlem yapmak için listeden en az bir ürün seçin.", 'warning');
            }
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
            if (dataTable) dataTable.rows().invalidate().draw(false);
            showToast("Tüm seçimler temizlendi.", 'success');
        }

        // --- DİNAMİK KURAL MOTORU ---
        function addRuleRow() {
            const container = document.getElementById('rulesContainer');
            const rowId = 'rule_' + Date.now();

            const html = `
                <div class="flex flex-wrap items-center gap-3 p-3 bg-base-200/50 rounded-lg border border-base-200 rule-row" id="${rowId}">
                    <select class="select select-bordered select-sm rule-type focus:outline-none" onchange="changeRuleType('${rowId}')">
                        <option value="kiyaslama">Kampanyalar Arası Kıyaslama (Avantajlı vs Flaş)</option>
                        <option value="guncel">Güncel Satış ile Kıyaslama (Avantajlı/Flaş vs Güncel Net)</option>
                    </select>
                    
                    <div class="rule-content kiyaslama-content flex items-center gap-2 flex-wrap">
                        <span class="text-sm font-medium">Eğer</span>
                        <select class="select select-bordered select-sm rule-target focus:outline-none">
                            <option value="Avantajlı">Avantajlı</option>
                            <option value="Flaş">Flaş</option>
                        </select>
                        <span class="text-sm font-medium">kârı, diğerine göre</span>
                        <select class="select select-bordered select-sm rule-operator focus:outline-none">
                            <option value="fazla">Fazla (Min %)</option>
                            <option value="az">Az (Max %)</option>
                            <option value="farketmez">Fark Etmez (+-%)</option>
                        </select>
                        <label class="input input-bordered input-sm flex items-center gap-2 w-24">
                            <input type="number" class="rule-percent w-full" placeholder="Örn: 5">
                            <span class="opacity-50">%</span>
                        </label>
                        <span class="text-sm font-medium">ise</span>
                        <select class="select select-bordered select-sm rule-action focus:outline-none">
                            <option value="Avantajlı">Avantajlı Yap</option>
                            <option value="Flaş">Flaş Yap</option>
                            <option value="Hiçbiri">Hiçbiri Yap</option>
                        </select>
                    </div>
                    
                    <div class="rule-content guncel-content items-center gap-2 flex-wrap" style="display:none;">
                        <span class="text-sm font-medium">Eğer</span>
                        <select class="select select-bordered select-sm rule-target-guncel focus:outline-none">
                            <option value="Avantajlı">Avantajlı</option>
                            <option value="Flaş">Flaş</option>
                        </select>
                        <span class="text-sm font-medium">kârı, GÜNCEL NET kârdan</span>
                        <select class="select select-bordered select-sm rule-operator-guncel focus:outline-none">
                            <option value="fazla">Fazla (Veya Eşit) ise</option>
                            <option value="az">Maksimum % düşüşte ise</option>
                        </select>
                        <label class="input input-bordered input-sm flex items-center gap-2 w-24">
                            <input type="number" class="rule-percent-guncel w-full" placeholder="Örn: 10">
                            <span class="opacity-50">%</span>
                        </label>
                        <span class="text-sm font-medium">ise seçimi</span>
                        <select class="select select-bordered select-sm rule-action-guncel focus:outline-none">
                            <option value="Aynı">O Kampanya (Avantajlı/Flaş) Yap</option>
                            <option value="Hiçbiri">Hiçbiri Yap</option>
                        </select>
                    </div>
                    
                    <div class="flex-1 flex justify-end">
                        <button class="btn btn-square btn-ghost btn-sm text-error hover:bg-error/20" onclick="document.getElementById('${rowId}').remove()">
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" /></svg>
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

        function applyCustomRules() {
            const ruleRows = document.querySelectorAll('.rule-row');
            if (ruleRows.length === 0) {
                showToast("Lütfen önce kural ekleyin.", 'warning');
                return;
            }

            let matchCount = 0;

            tableData.forEach(row => {
                // Sadece 'Hiçbiri' durumunda olanlara kuralları uygula
                if (row.userSelection !== 'Hiçbiri') return;

                let ruleMatched = false;

                const n2 = row['Avantajlı Ürün Kalan Net (TL)'];
                const n3 = row['Flaş Ürün Kalan Net (TL)'];
                const f1 = row['Güncel Ürün Fiyatı (TL)'];

                const hasAv = row['Avantajlı Ürün Eşleşme Durumu'] === 'Eşleşti';
                const hasFl = row['Flaş Ürün Eşleşme Durumu'] === 'Eşleşti';

                for (let i = 0; i < ruleRows.length; i++) {
                    if (ruleMatched) break;

                    const ruleNode = ruleRows[i];
                    const type = ruleNode.querySelector('.rule-type').value;

                    if (type === 'kiyaslama') {
                        if (!hasAv || !hasFl) continue;
                        if (n2 == null || n3 == null) continue;

                        const target = ruleNode.querySelector('.rule-target').value;
                        const op = ruleNode.querySelector('.rule-operator').value;
                        const pctStr = ruleNode.querySelector('.rule-percent').value;
                        const action = ruleNode.querySelector('.rule-action').value;

                        if (!pctStr) continue;
                        const pct = parseFloat(pctStr);

                        const targetNet = (target === 'Avantajlı') ? n2 : n3;
                        const otherNet = (target === 'Avantajlı') ? n3 : n2;

                        let isMatch = false;

                        if (otherNet > 0) {
                            const diffPct = ((targetNet - otherNet) / otherNet) * 100;

                            if (op === 'fazla' && diffPct >= pct) isMatch = true;
                            if (op === 'az' && diffPct <= pct) isMatch = true;
                            if (op === 'farketmez' && Math.abs(diffPct) <= pct) isMatch = true;
                        } else if (otherNet <= 0 && targetNet > 0) {
                            if (op === 'fazla' || op === 'farketmez') isMatch = true;
                        }

                        if (isMatch) {
                            if (action === 'Hiçbiri' ||
                                (action === 'Avantajlı' && hasAv) ||
                                (action === 'Flaş' && hasFl)) {
                                row.userSelection = action;
                                ruleMatched = true;
                                matchCount++;
                            }
                        }

                    } else if (type === 'guncel') {
                        const target = ruleNode.querySelector('.rule-target-guncel').value;
                        const op = ruleNode.querySelector('.rule-operator-guncel').value;
                        const pctStr = ruleNode.querySelector('.rule-percent-guncel').value;
                        const action = ruleNode.querySelector('.rule-action-guncel').value;

                        if (!pctStr) continue;
                        const pct = parseFloat(pctStr);

                        const targetNet = (target === 'Avantajlı') ? n2 : n3;
                        const hasTarget = (target === 'Avantajlı') ? hasAv : hasFl;
                        const n1 = row['Güncel Ürün Kalan Net (TL)'];

                        if (!hasTarget || targetNet == null || n1 == null) continue;

                        let isMatch = false;

                        if (op === 'fazla' && targetNet >= n1) {
                            isMatch = true;
                        } else if (op === 'az' && n1 > 0 && targetNet < n1) {
                            const dususPct = ((n1 - targetNet) / n1) * 100;
                            if (dususPct <= pct) {
                                isMatch = true;
                            }
                        }

                        if (isMatch) {
                            if (action === 'Hiçbiri') {
                                row.userSelection = 'Hiçbiri';
                                ruleMatched = true;
                                matchCount++;
                            } else if (action === 'Aynı') {
                                row.userSelection = target;
                                ruleMatched = true;
                                matchCount++;
                            }
                        }
                    }
                }
            });

            if (dataTable) dataTable.rows().invalidate().draw(false);
            showToast(`Kurallar çalıştırıldı ve ${matchCount} adet ürünün seçimi güncellendi.`, 'success');
        }
    