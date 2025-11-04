/*
 * =================================================================
 * SCRIPT v7 (Finale Version - OHNE Fuzzy Matching)
 * - Nur noch ein Schritt: Direkte Generierung der Berichte
 * =================================================================
 */
window.onload = function() {

    // ====== KONFIGURATION ======
    const WORKER_URL = 'https://tekin-reco-backend-tool.tekin-6af.workers.dev'; // Dein Link
    // ==========================

    // Globale Variablen
    let parsedAirtableData = null;
    let parsedErpData = null;

    // DOM-Elemente
    // analyzeButton existiert nicht mehr
    const reportButton = document.getElementById('report-button');
    const airtableFileInput = document.getElementById('airtable-file');
    const erpFileInput = document.getElementById('erp-file');
    const loader = document.getElementById('loader');
    const errorMessage = document.getElementById('error-message');

    // Karten-Container (summary und fuzzy entfernt)
    const recoReportCard = document.getElementById('reconciliation-report-card');
    const reportCard = document.getElementById('final-report-card');

    // Ergebnis-Divs (summary und fuzzy entfernt)
    const teamReportDiv = document.getElementById('team-report-area');
    const personReportDiv = document.getElementById('person-report-area');
    const recoTotalsDiv = document.getElementById('reco-report-totals');
    const recoProjectsToUpdateDiv = document.getElementById('reco-report-projectsToUpdate');
    const recoUnmatchedERPProjDiv = document.getElementById('reco-report-unmatchedERP-byProject');
    const recoUnmatchedAirtableProjDiv = document.getElementById('reco-report-unmatchedAirtable-byProject');
    // const recoFuzzyMatchedDiv = document.getElementById('reco-report-fuzzyMatched'); // Optional entfernen
    const recoUnmatchedERPKVDiv = document.getElementById('reco-report-unmatchedERP-byKV');
    const recoUnmatchedAirtableNoProjDiv = document.getElementById('reco-report-unmatchedAirtable-noProj');


    // Event Listener (nur noch für Report-Button)
    reportButton.addEventListener('click', handleReport);

    const AIRTABLE_HEADER_ALIASES = {
        'Projektnummer': ['Projektnummer', 'Projekt-Nr.', 'Projekt Nr.', 'Projektnr', 'Projekt_Nr'],
        'Projekttitel': ['Projekttitel', 'Projekt Titel', 'Projektname'],
        'Agenturleistung (netto)': ['Agenturleistung (netto)', 'Agenturleistung netto', 'Agenturleistung Netto']
    };

    const ERP_HEADER_ALIASES = {
        'Projekt Projektnummer': ['Projekt Projektnummer', 'Projektnummer', 'Projekt-Nr.', 'Projekt Nr.', 'Projekt_Nr'],
        'KV-Nummer': ['KV-Nummer', 'KV Nummer', 'KV-Nr.', 'KV Nr.', 'KV_Nr']
    };

    const REQUIRED_AIRTABLE_COLUMNS = ['Projektnummer', 'Agenturleistung (netto)'];
    const REQUIRED_ERP_COLUMNS = ['Projekt Projektnummer'];

    const mapAirtableHeader = createHeaderMapper(AIRTABLE_HEADER_ALIASES);
    const mapErpHeader = createHeaderMapper(ERP_HEADER_ALIASES);


    /**
     * Parsed die hochgeladenen Dateien.
     */
    async function parseFiles() {
        errorMessage.textContent = '';
        parsedAirtableData = null;
        parsedErpData = null;
        const airtableFile = airtableFileInput.files[0];
        const erpFile = erpFileInput.files[0];
        if (!airtableFile || !erpFile) { errorMessage.textContent = 'Bitte beide Dateien auswählen.'; return false; }

        try {
            parsedAirtableData = await new Promise((resolve, reject) => {
                Papa.parse(airtableFile, {
                    header: true,
                    skipEmptyLines: 'greedy',
                    encoding: 'ISO-8859-1',
                    transformHeader: mapAirtableHeader,
                    complete: (results) => {
                        try {
                            resolve(filterParsedRows(results.data || []));
                        } catch (err) {
                            reject(err);
                        }
                    },
                    error: (e) => reject(new Error(`Airtable CSV-Fehler: ${e.message}`))
                });
            });

            if (!parsedAirtableData || parsedAirtableData.length === 0) {
                throw new Error('Airtable-Datei ist leer oder konnte nicht geparst werden.');
            }

            validateColumns(parsedAirtableData, REQUIRED_AIRTABLE_COLUMNS, 'Airtable');

            if (erpFile.name && erpFile.name.toLowerCase().endsWith('.xlsx')) {
                const rows = await readXlsxFile(erpFile);
                if (!rows || rows.length < 2) throw new Error('ERP Excel-Datei leer oder nur Header.');
                const headers = rows[0].map(cell => mapErpHeader(String(cell ?? '')));
                const erpObjects = rows.slice(1).map(row => mapArrayRowToObject(headers, row));
                parsedErpData = filterParsedRows(erpObjects);
            } else {
                parsedErpData = await new Promise((resolve, reject) => {
                    Papa.parse(erpFile, {
                        header: true,
                        skipEmptyLines: 'greedy',
                        transformHeader: mapErpHeader,
                        complete: (results) => {
                            try {
                                resolve(filterParsedRows(results.data || []));
                            } catch (err) {
                                reject(err);
                            }
                        },
                        error: (e) => reject(new Error(`ERP CSV-Fehler: ${e.message}`))
                    });
                });
            }

            if (!parsedErpData || parsedErpData.length === 0) {
                throw new Error('ERP-Datei ist leer oder konnte nicht geparst werden.');
            }

            validateColumns(parsedErpData, REQUIRED_ERP_COLUMNS, 'ERP');

            normalizeProjectNumbers(parsedAirtableData, ['Projektnummer']);
            normalizeProjectNumbers(parsedErpData, ['Projekt Projektnummer', 'KV-Nummer']);

            return true;
        } catch (error) {
            errorMessage.textContent = `Fehler beim Parsen: ${error.message || error}`;
            console.error(error);
            return false;
        }
    }

    /**
     * Hauptfunktion für die Abgleichs- und finalen Berichte
     */
    async function handleReport() {
        // Bibliotheks-Checks
        if (typeof Papa === 'undefined') { errorMessage.textContent = 'Fehler: PapaParse...'; return; }
        if (typeof readXlsxFile === 'undefined') { errorMessage.textContent = 'Fehler: Read-Excel-File...'; return; }

        setLoading(true);
        errorMessage.textContent = '';
        recoReportCard.style.display = 'none';
        reportCard.style.display = 'none';

        // 1. Dateien parsen
        const parseSuccess = await parseFiles();
        if (!parseSuccess) { setLoading(false); return; }

        try {
            // 2. Direkter Aufruf des /report Endpunkts
            const response = await fetch(`${WORKER_URL}/report`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    airtableData: parsedAirtableData,
                    erpData: parsedErpData
                    // confirmedMatches wird nicht mehr gesendet
                })
            });

             if (!response.ok) {
                 const errText = await response.text();
                 try { throw JSON.parse(errText); } catch (e) { throw new Error(errText || `Server-Fehler: ${response.statusText}`); }
            }

            const reports = await response.json();

            // 3. Abgleichs-Bericht rendern
            renderReconciliationReport_v7(reports.reconciliation); // Angepasste Render-Funktion
            recoReportCard.style.display = 'block';

            // 4. Finale Berichte rendern
            renderFinalReports_v7(reports.finalReports); // Angepasste Render-Funktion
            reportCard.style.display = 'block';

        } catch (error) {
            errorMessage.textContent = `Berichts-Fehler: ${error.error || error.message || error}`;
            console.error(error);
        } finally {
            setLoading(false);
        }
    }


    // ====== RENDER-FUNKTIONEN (ANGEPASST für v7) ======

    function setLoading(isLoading) {
        loader.style.display = isLoading ? 'block' : 'none';
        reportButton.disabled = isLoading;
    }

    // Fuzzy Match Tabelle wird nicht mehr gerendert

    // Angepasster Abgleichs-Bericht Renderer (ohne Fuzzy-Match-Bericht)
    function renderReconciliationReport_v7(reco) {
        // Finanz-Zusammenfassung
        recoTotalsDiv.innerHTML = `<div class="reco-section"><h4>Finanz-Zusammenfassung</h4><div class="reco-totals-grid">
            <div class="reco-totals-item">Gesamtsumme ERP <strong>${formatCurrency(reco.totals.totalERP)}</strong></div>
            <div class="reco-totals-item">Gesamtsumme Airtable <strong>${formatCurrency(reco.totals.totalAirtable)}</strong></div>
            <div class="reco-totals-item">Davon Zugeordnet <strong>${formatCurrency(reco.totals.totalReconciled)}</strong></div>
            <div class="reco-totals-item">Fehlende ERP-Beträge <strong>${formatCurrency(reco.totals.totalUnreconciledERP)}</strong></div>
            </div></div>`;

        // Projekte zum Aktualisieren
        recoProjectsToUpdateDiv.innerHTML = renderRecoTable(
            'To-Do: Beträge in Airtable aktualisieren (Projekt-Ebene)',
            ['Projekt-Nr.', 'Airtable-Titel', 'Airtable-Betrag', 'ERP-Gesamt (NEU)', ' Zugehörige ERP KVs'],
            reco.projectsToUpdate.map(row => `
                <tr><td>${escapeHtml(row.projNr)}</td><td>${escapeHtml(row.airtableTitle)}</td><td class="amount">${formatCurrency(row.airtableAmount)}</td>
                    <td class="amount"><strong>${formatCurrency(row.erpTotalAmount)}</strong></td><td>${escapeHtml(row.erpKVs.join(', '))}</td></tr>`)
        );

        // Projekte nur im ERP
        recoUnmatchedERPProjDiv.innerHTML = renderRecoTable(
            'To-Do: Diese Projekte fehlen in Airtable',
            ['Projekt-Nr.', 'ERP KVs (Titel, Betrag)', 'ERP-Gesamt'],
            reco.unmatchedERP_byProject.map(row => `
                <tr><td>${escapeHtml(row.projNr)}</td><td>${row.erpKVs.map(kv => `${escapeHtml(kv.kv)} (${escapeHtml(kv.title)}, ${formatCurrency(kv.amount)})`).join('<br>')}</td>
                    <td class="amount">${formatCurrency(row.erpTotalAmount)}</td></tr>`)
        );

        // Projekte nur in Airtable
        recoUnmatchedAirtableProjDiv.innerHTML = renderRecoTable(
            'Info: Diese Airtable-Projekte fehlen im ERP',
            ['Projekt-Nr.', 'Airtable-Titel', 'Airtable-Betrag'],
            reco.unmatchedAirtable_byProject.map(row => `
                <tr><td>${escapeHtml(row.projNr)}</td><td>${escapeHtml(row.airtableTitle)}</td><td class="amount">${formatCurrency(row.airtableAmount)}</td></tr>`)
        );

        // Restliche KVs nur im ERP (ohne Projektnummer)
        recoUnmatchedERPKVDiv.innerHTML = renderRecoTable(
            'Manuell zu prüfen: Einzelne ERP KVs OHNE Projektnummer',
            ['KV-Nummer', 'ERP-Titel', 'ERP-Betrag'],
            reco.unmatchedERP_byKV.map(row => `
                <tr><td>${escapeHtml(row.kv)}</td><td>${escapeHtml(row.erpTitle)}</td><td class="amount">${formatCurrency(row.erpAmount)}</td></tr>`)
        );

        // Restliche Airtable ohne Projektnummer
        recoUnmatchedAirtableNoProjDiv.innerHTML = renderRecoTable(
            'Manuell zu prüfen: Airtable-Einträge OHNE Projektnummer',
            ['Airtable-Titel', 'Airtable-Betrag'],
            reco.unmatchedAirtable_noProj.map(row => `
                <tr><td>${escapeHtml(row.airtableTitle)}</td><td class="amount">${formatCurrency(row.airtableAmount)}</td></tr>`)
        );

        // Fuzzy Match Bericht wird nicht mehr gerendert
        // const recoFuzzyMatchedDivMaybe = document.getElementById('reco-report-fuzzyMatched');
        // if (recoFuzzyMatchedDivMaybe) recoFuzzyMatchedDivMaybe.innerHTML = '';
    }

    // Finale Berichte Renderer (kleine Anpassung für leere Berichte)
    function renderFinalReports_v7(finalReports) {
        const teamItems = normalizeFinalReportCollection(finalReports?.teamReport, 'Team');
        const personItems = normalizeFinalReportCollection(finalReports?.personReport, 'Person');

        teamReportDiv.innerHTML = renderBarReport(teamItems, 'Team');
        personReportDiv.innerHTML = renderBarReport(personItems, 'Person');
    }

    function renderBarReport(items, label) {
        if (!items || items.length === 0) {
            return `<div class="report-empty">Keine Daten für diesen Bericht.</div>`;
        }

        const maxAmount = Math.max(...items.map(item => Math.abs(item.amount)));
        const safeMax = Number.isFinite(maxAmount) && maxAmount > 0 ? maxAmount : 0;

        const hasPercentage = items.some(item => item.percentage != null);

        const rows = items.map(item => {
            const width = safeMax > 0 ? Math.min(100, Math.round((Math.abs(item.amount) / safeMax) * 100)) : 0;
            const percentageText = item.percentage != null ? `${formatPercentage(item.percentage)}` : '';
            const amountText = formatCurrency(item.amount);

            return `
                <div class="bar-row">
                    <div class="bar-label">${escapeHtml(item.name || label)}</div>
                    <div class="bar-meter">
                        <div class="bar-fill" style="width: ${width}%"></div>
                        <span class="bar-value">${amountText}</span>
                    </div>
                    ${percentageText ? `<div class="bar-percentage">${percentageText}</div>` : ''}
                </div>
            `;
        }).join('');

        const reportClass = hasPercentage ? 'bar-report has-percentage' : 'bar-report';
        return `<div class="${reportClass}" aria-label="${escapeHtml(label)}-Report">${rows}</div>`;
    }

    function normalizeFinalReportCollection(reportData, fallbackLabel) {
        if (!reportData) return [];

        const collection = extractReportArray(reportData, fallbackLabel, new Set());
        if (!collection || collection.length === 0) {
            return [];
        }

        return collection
            .map(item => normalizeReportItem(item, fallbackLabel))
            .filter(item => item && item.name);
    }

    function extractReportArray(reportData, fallbackLabel, visited) {
        if (!reportData) return [];
        if (visited && visited.has(reportData)) return [];

        const coerceTabular = (rows, source) => coerceTabularArray(rows, source || reportData, fallbackLabel);

        if (Array.isArray(reportData)) {
            return coerceTabular(reportData);
        }

        if (reportData && typeof reportData === 'object') {
            if (visited) {
                visited.add(reportData);
            }

            if (Array.isArray(reportData.rows) || Array.isArray(reportData.data)) {
                const primaryRows = Array.isArray(reportData.rows) ? reportData.rows : reportData.data;
                return coerceTabular(primaryRows, reportData);
            }

            const candidateKeys = [
                'items', 'entries', 'list', 'breakdown', 'report', 'values', 'result', 'results',
                'records', 'collection', 'table', 'tableData', 'dataset', 'details', 'elements',
                'byTeam', 'teams', 'byPerson', 'people', 'groups', 'grouped', 'members'
            ];
            for (const key of candidateKeys) {
                if (!Object.prototype.hasOwnProperty.call(reportData, key)) continue;
                const candidate = reportData[key];
                if (Array.isArray(candidate)) {
                    return coerceTabular(candidate, reportData);
                }
                if (candidate && typeof candidate === 'object') {
                    const nested = extractReportArray(candidate, fallbackLabel, visited);
                    if (nested.length > 0) {
                        return nested;
                    }
                }
            }

            if (Array.isArray(reportData.headers) && Array.isArray(reportData.rows)) {
                return coerceTabular(reportData.rows, reportData);
            }

            const summaryKeys = ['total', 'sum', 'overall', 'grandtotal', 'totals', 'summe', 'gesamtsumme', 'gesamt'];
            const excludedKeys = ['headers', 'columns', 'fields', 'totals', 'summary', 'meta', 'metadata', 'title', 'type', 'unit', 'units', 'currency'];

            const nestedEntries = Object.entries(reportData)
                .filter(([key, value]) => {
                    if (excludedKeys.includes(key)) return false;
                    if (value === null || value === undefined) return false;
                    const normalizedKey = normalizeKeyForComparison(key);
                    if (normalizedKey && summaryKeys.some(summary => normalizedKey === summary || normalizedKey.startsWith(summary))) {
                        return false;
                    }
                    return true;
                });

            const arrayEntry = nestedEntries.find(([, value]) => Array.isArray(value));
            if (arrayEntry) {
                return coerceTabular(arrayEntry[1], reportData);
            }

            const objectValueEntries = nestedEntries
                .filter(([, value]) => value && typeof value === 'object' && !Array.isArray(value));
            if (objectValueEntries.length > 0) {
                const aggregated = [];
                for (const [key, value] of objectValueEntries) {
                    const nested = extractReportArray(value, fallbackLabel, visited);
                    if (nested.length > 0) {
                        nested.forEach(item => {
                            if (item && typeof item === 'object') {
                                const normalizedName = normalizeKeyForComparison(item.name || '');
                                if (!item.name || isGenericLabel(normalizedName)) {
                                    aggregated.push({ ...item, name: String(key) });
                                } else {
                                    aggregated.push(item);
                                }
                            } else {
                                aggregated.push(item);
                            }
                        });
                        continue;
                    }
                    aggregated.push({ name: key, ...value });
                }
                if (aggregated.length > 0) {
                    return aggregated;
                }
            }

            const numericEntries = nestedEntries
                .filter(([key, value]) => isNumericValue(value));
            if (numericEntries.length > 0) {
                return numericEntries.map(([key, value]) => ({ name: key, amount: value }));
            }
        }

        return [];
    }

    function coerceTabularArray(rows, sourceData, fallbackLabel) {
        if (!Array.isArray(rows)) return [];
        if (rows.length === 0) return [];

        const headers = inferHeaders(sourceData) || inferHeadersFromRows(rows);
        let dataRows = rows;

        if (headers && headers.length > 0 && Array.isArray(rows[0]) && headers.every((_, idx) => idx < rows[0].length)) {
            if (rows.length > 0 && isHeaderRow(rows[0], headers)) {
                dataRows = rows.slice(1);
            }
        }

        if (headers && headers.length > 0) {
            return dataRows.map(row => {
                if (Array.isArray(row)) {
                    return headers.reduce((obj, header, index) => {
                        if (!header) return obj;
                        obj[header] = row[index];
                        return obj;
                    }, {});
                }
                if (row && typeof row === 'object') {
                    return row;
                }
                if (row === null || row === undefined) {
                    return {};
                }
                const labelKey = headers[0] || fallbackLabel || 'Name';
                return { [labelKey]: row };
            });
        }

        if (rows.every(row => row && typeof row === 'object' && !Array.isArray(row))) {
            return rows;
        }

        return rows.map((row) => {
            if (Array.isArray(row)) {
                const [label, amount, percentage, ...rest] = row;
                const item = {};
                if (label !== undefined) {
                    item[fallbackLabel || 'Name'] = label;
                }
                if (amount !== undefined) {
                    item.amount = amount;
                }
                if (percentage !== undefined) {
                    item.percentage = percentage;
                }
                rest.forEach((value, idx) => {
                    item[`col_${idx + 3}`] = value;
                });
                return item;
            }
            if (row && typeof row === 'object') {
                return row;
            }
            return { [fallbackLabel || 'Name']: row };
        });
    }

    function inferHeaders(sourceData) {
        if (!sourceData || typeof sourceData !== 'object') return null;
        const headerKeys = ['headers', 'columns', 'fields', 'fieldNames', 'labels', 'keys'];
        for (const key of headerKeys) {
            if (Array.isArray(sourceData[key]) && sourceData[key].length > 0) {
                return sourceData[key];
            }
        }
        return null;
    }

    function inferHeadersFromRows(rows) {
        if (!Array.isArray(rows) || rows.length === 0) return null;
        const firstRow = rows[0];
        if (Array.isArray(firstRow) && firstRow.length > 0 && firstRow.every(cell => typeof cell === 'string' && cell.trim() !== '')) {
            return firstRow;
        }
        return null;
    }

    function isHeaderRow(row, headers) {
        if (!Array.isArray(row) || !Array.isArray(headers)) return false;
        if (headers.length === 0) return false;
        if (row.length < headers.length) return false;
        return headers.every((header, index) => {
            if (!header) return true;
            const cell = row[index];
            if (typeof cell !== 'string') return false;
            return normalizeKeyForComparison(cell) === normalizeKeyForComparison(header);
        });
    }

    function normalizeReportItem(rawItem, fallbackLabel) {
        if (!rawItem || typeof rawItem !== 'object') {
            return {
                name: String(rawItem ?? fallbackLabel ?? ''),
                amount: 0,
                percentage: null
            };
        }

        const normalizedMap = buildNormalizedKeyMap(rawItem);

        let name = rawItem.name || rawItem.label || rawItem.team || rawItem.person || rawItem.category || rawItem.key || rawItem.id || rawItem.title || '';
        if (!name) {
            const { value: nameCandidate } = findValueByKeyCandidates(normalizedMap, rawItem, [
                'name', 'display name', 'title', 'beschreibung', 'bezeichnung',
                'team name', 'teamname', 'team', 'gruppe', 'gruppenname', 'abteilung', 'unit', 'bereich', 'agentur', 'agent',
                'person name', 'personname', 'person', 'mitarbeiter', 'mitarbeiter/in', 'mitarbeiter name', 'mitarbeitername',
                'employee', 'salesperson', 'seller', 'berater', 'consultant', 'owner', 'kunde', 'kundenname', 'customer', 'account'
            ]);
            if (nameCandidate !== undefined) {
                name = nameCandidate;
            }
        }

        const amountMatch = findValueByKeyCandidates(normalizedMap, rawItem, [
            'amount', 'value', 'total', 'sum', 'netto', 'netto summe', 'netto gesamt', 'umsatz', 'sales', 'revenue', 'betrag', 'total netto',
            'gesamt', 'gesamtbetrag', 'net amount', 'assigned amount', 'allocated', 'commission', 'provision'
        ]);
        let amount = amountMatch.value !== undefined ? parseNumericValue(amountMatch.value) : null;

        if (!Number.isFinite(amount)) {
            amount = null;
        }

        if (amount === null) {
            const fallbackNumericEntry = Object.entries(rawItem).find(([key, value]) => {
                if (!isNumericValue(value)) return false;
                const normalizedKey = normalizeKeyForComparison(key);
                return normalizedKey && !['index', 'rank', 'position'].includes(normalizedKey);
            });
            if (fallbackNumericEntry) {
                amount = parseNumericValue(fallbackNumericEntry[1]);
                if (!name) {
                    name = fallbackNumericEntry[0];
                }
            }
        }

        const percentageMatch = findValueByKeyCandidates(normalizedMap, rawItem, ['percentage', 'percent', 'share', 'anteil', 'quote', 'ratio', 'prozent']);
        const parsedPercentage = percentageMatch.value !== undefined ? parseNumericValue(percentageMatch.value, true) : null;
        const percentage = Number.isFinite(parsedPercentage) ? parsedPercentage : null;

        if (!name) {
            const measureKeywords = ['amount', 'value', 'total', 'sum', 'netto', 'umsatz', 'sales', 'revenue', 'betrag', 'percentage', 'percent', 'share', 'anteil', 'quote', 'ratio', 'prozent'];
            const stringEntry = Object.entries(rawItem).find(([key, value]) => {
                if (typeof value !== 'string' || value.trim() === '') return false;
                const normalizedKey = normalizeKeyForComparison(key);
                if (!normalizedKey) return true;
                if (normalizedKey.startsWith('col')) return false;
                return !measureKeywords.some(keyword => normalizedKey.includes(keyword));
            });
            if (stringEntry) {
                name = stringEntry[1];
            }
        }

        const normalizedFinalName = normalizeKeyForComparison(name);
        if ((!name || isGenericLabel(normalizedFinalName)) && fallbackLabel) {
            name = fallbackLabel;
        }

        return {
            name: String(name || fallbackLabel || ''),
            amount: Number.isFinite(amount) ? amount : 0,
            percentage
        };
    }

    function buildNormalizedKeyMap(obj) {
        const map = new Map();
        Object.keys(obj || {}).forEach(key => {
            const normalized = normalizeKeyForComparison(key);
            if (normalized && !map.has(normalized)) {
                map.set(normalized, key);
            }
        });
        return map;
    }

    function findValueByKeyCandidates(normalizedMap, source, candidates) {
        for (const candidate of candidates) {
            const normalizedCandidate = normalizeKeyForComparison(candidate);
            if (normalizedCandidate && normalizedMap.has(normalizedCandidate)) {
                const actualKey = normalizedMap.get(normalizedCandidate);
                return { key: actualKey, value: source[actualKey] };
            }
        }
        return { key: null, value: undefined };
    }

    function isGenericLabel(normalizedLabel) {
        if (!normalizedLabel) return true;
        const genericLabels = ['amount', 'value', 'total', 'sum', 'result', 'entry', 'item', 'row', 'percentage', 'percent', 'share', 'ratio'];
        return genericLabels.some(generic => normalizedLabel === generic || normalizedLabel.startsWith(generic));
    }

    function parseNumericValue(value, allowNull = false) {
        if (value === null || value === undefined || value === '') {
            return allowNull ? null : 0;
        }

        if (typeof value === 'number') {
            return Number.isFinite(value) ? value : (allowNull ? null : 0);
        }

        if (typeof value === 'string') {
            const cleaned = value
                .replace(/\u00A0/g, ' ')
                .replace(/\s+/g, '')
                .replace(/[^0-9,\-\.]/g, '')
                .replace(/,(?=\d{1,2}$)/, '.')
                .replace(/\.(?=.*\.)/g, '');
            if (!cleaned) {
                return allowNull ? null : 0;
            }
            const parsed = Number(cleaned);
            if (Number.isFinite(parsed)) {
                return parsed;
            }
        }

        return allowNull ? null : 0;
    }

    function isNumericValue(value) {
        if (typeof value === 'number') {
            return Number.isFinite(value);
        }
        if (typeof value === 'string') {
            const hasDigits = /[0-9]/.test(value);
            if (!hasDigits) return false;
            const parsed = parseNumericValue(value);
            return Number.isFinite(parsed);
        }
        return false;
    }

    function formatPercentage(value) {
        if (value === null || value === undefined) {
            return '';
        }
        const num = Number(value);
        if (!Number.isFinite(num)) {
            return '';
        }
        return `${num.toLocaleString('de-DE', { minimumFractionDigits: 1, maximumFractionDigits: 1 })}%`;
    }

    // Hilfsfunktion für Tabellen-Rendering
    function renderRecoTable(title, headers, rows) {
         if (!rows || rows.length === 0) return `<div class="reco-section"><h4>${escapeHtml(title)}</h4><p>Keine Einträge.</p></div>`;
        const headerHTML = headers.map(h => h.toLowerCase().includes('betrag') || h.toLowerCase().includes('summe') ? `<th class="amount">${escapeHtml(h)}</th>` : `<th>${escapeHtml(h)}</th>`).join('');
        return `<div class="reco-section"><h4>${escapeHtml(title)} (${rows.length})</h4><table class="reco-table"><thead><tr>${headerHTML}</tr></thead><tbody>${rows.join('')}</tbody></table></div>`;
    }

    function filterParsedRows(rows) {
         if (!Array.isArray(rows)) return [];
         return rows.filter(row => row && typeof row === 'object' && !isRowEmpty(row));
    }

    function isRowEmpty(row) {
         if (!row || typeof row !== 'object') return true;
         return !Object.values(row).some(hasMeaningfulValue);
    }

    function hasMeaningfulValue(value) {
         if (value === null || value === undefined) return false;
         if (typeof value === 'string') return value.replace(/\u00A0/g, ' ').trim() !== '';
         if (typeof value === 'number') return !Number.isNaN(value);
         return true;
    }

    function sanitizeHeader(header) {
         if (header === null || header === undefined) return '';
         return String(header).replace(/^\uFEFF/, '').replace(/\u00A0/g, ' ').trim();
    }

    function normalizeKeyForComparison(key) {
         const sanitized = sanitizeHeader(key);
         if (!sanitized) return '';
         return sanitized
             .toLowerCase()
             .normalize('NFKD')
             .replace(/[\u0300-\u036f]/g, '')
             .replace(/[\s\-_().,:;\/\\]+/g, '')
             .replace(/€|eur/g, '')
             .trim();
    }

    function createHeaderMapper(aliases) {
         const directAliasMap = new Map();
         const normalizedAliasMap = new Map();

         const registerAlias = (alias, canonicalKey) => {
              const sanitizedAlias = sanitizeHeader(alias);
              if (!sanitizedAlias) return;
              directAliasMap.set(sanitizedAlias.toLowerCase(), canonicalKey);
              const normalizedAlias = normalizeKeyForComparison(sanitizedAlias);
              if (normalizedAlias) {
                   normalizedAliasMap.set(normalizedAlias, canonicalKey);
              }
         };

         Object.entries(aliases || {}).forEach(([canonical, list]) => {
              const canonicalKey = String(canonical);
              registerAlias(canonicalKey, canonicalKey);
              (list || []).forEach(alias => registerAlias(alias, canonicalKey));
         });

         return (header) => {
              const sanitized = sanitizeHeader(header);
              if (!sanitized) return sanitized;

              const lower = sanitized.toLowerCase();
              if (directAliasMap.has(lower)) {
                   return directAliasMap.get(lower);
              }

              const normalized = normalizeKeyForComparison(sanitized);
              if (normalized && normalizedAliasMap.has(normalized)) {
                   return normalizedAliasMap.get(normalized);
              }

              return sanitized;
         };
    }

    function mapArrayRowToObject(headers, row) {
         const source = Array.isArray(row) ? row : [];
         return headers.reduce((obj, header, index) => {
              if (!header) return obj;
              obj[header] = source[index];
              return obj;
         }, {});
    }

    function validateColumns(rows, requiredColumns, datasetLabel) {
         const available = new Set();
         const availableNormalized = new Map();

         (rows || []).forEach(row => {
              if (!row || typeof row !== 'object') return;
              Object.keys(row).forEach(key => {
                   if (!key) return;
                   available.add(key);
                   const normalized = normalizeKeyForComparison(key);
                   if (normalized && !availableNormalized.has(normalized)) {
                        availableNormalized.set(normalized, key);
                   }
              });
         });

         const missing = (requiredColumns || []).filter(col => {
              if (available.has(col)) return false;
              const normalizedRequired = normalizeKeyForComparison(col);
              return !availableNormalized.has(normalizedRequired);
         });

         if (missing.length > 0) {
              const availableList = Array.from(availableNormalized.values());
              const availableInfo = availableList.length > 0 ? availableList.join(', ') : 'keine';
              throw new Error(`${datasetLabel}: Fehlende Pflichtspalten (${missing.join(', ')}). Gefunden: ${availableInfo}`);
         }
    }

    function normalizeProjectNumbers(rows, keys) {
         if (!Array.isArray(rows)) return;
         rows.forEach(row => {
              if (!row || typeof row !== 'object') return;
              (keys || []).forEach(key => {
                   if (Object.prototype.hasOwnProperty.call(row, key)) {
                        row[key] = normalizeProjectKey(row[key]);
                   }
              });
         });
    }

    function normalizeProjectKey(value) {
         return (value ?? '')
             .toString()
             .replace(/\u00A0/g, ' ')
             .replace(/[\t\r\n]+/g, ' ')
             .replace(/\s+/g, ' ')
             .trim();
    }

    // Hilfsfunktion für Währungsformatierung
    function formatCurrency(value) {
         const num = Number(value);
         if (isNaN(num) || value === null || value === undefined) return 'N/A';
         return num.toLocaleString('de-DE', { style: 'currency', currency: 'EUR' });
    }

    // Hilfsfunktion zum Escapen von HTML
    function escapeHtml(unsafe) {
         if (typeof unsafe !== 'string') return unsafe;
         return unsafe.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
    }

}; // ENDE DES window.onload WRAPPERS
