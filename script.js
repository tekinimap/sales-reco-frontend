/*
 * =================================================================
 * NEUE VERSION (v4)
 * - Fügt Logik für getrennte Sammel-KV-Berichte hinzu
 * =================================================================
 */
window.onload = function() {
    
    // ====== KONFIGURATION ======
    const WORKER_URL = 'https://tekin-reco-backend-tool.tekin-6af.workers.dev'; // Dein Link ist gespeichert!
    // ==========================

    // Globale Variablen
    let parsedAirtableData = null;
    let parsedErpData = null;

    // DOM-Elemente
    const analyzeButton = document.getElementById('analyze-button');
    const reportButton = document.getElementById('report-button');
    const airtableFileInput = document.getElementById('airtable-file');
    const erpFileInput = document.getElementById('erp-file');
    const loader = document.getElementById('loader');
    const errorMessage = document.getElementById('error-message');

    // Karten-Container
    const summaryCard = document.getElementById('summary-results-card');
    const fuzzyCard = document.getElementById('fuzzy-match-card');
    const recoReportCard = document.getElementById('reconciliation-report-card');
    const reportCard = document.getElementById('final-report-card');

    // Ergebnis-Divs
    const summaryResultsDiv = document.getElementById('summary-results');
    const fuzzyMatchDiv = document.getElementById('fuzzy-match-area');
    const teamReportDiv = document.getElementById('team-report-area');
    const personReportDiv = document.getElementById('person-report-area');
    
    // NEUE Ergebnis-Divs für Abgleichs-Bericht
    const recoTotalsDiv = document.getElementById('reco-report-totals');
    const recoKVsToUpdateDiv = document.getElementById('reco-report-kvsToUpdate');
    const recoUnmatchedERPDiv = document.getElementById('reco-report-unmatchedERP');
    const recoUnmatchedAirtableDiv = document.getElementById('reco-report-unmatchedAirtable');
    const recoFuzzyMatchedDiv = document.getElementById('reco-report-fuzzyMatched');
    // ANGEPASST:
    const recoSammelKVsMatchedDiv = document.getElementById('reco-report-sammelKVs-matched');
    const recoSammelKVsFailedDiv = document.getElementById('reco-report-sammelKVs-failed');


    // Event Listeners
    analyzeButton.addEventListener('click', handleAnalyze);
    reportButton.addEventListener('click', handleReport);

    
    /**
     * Hauptfunktion für den ersten Abgleich (Phase 1-3)
     */
    async function handleAnalyze() {
        if (typeof Papa === 'undefined') {
            errorMessage.textContent = 'Fehler: PapaParse (CSV-Bibliothek) konnte nicht geladen werden. Bitte Cache leeren (Strg+Shift+R) und Seite neu laden.';
            return;
        }
        if (typeof readXlsxFile === 'undefined') {
            errorMessage.textContent = 'Fehler: Read-Excel-File (Excel-Bibliothek) konnte nicht geladen werden. Bitte Cache leeren (Strg+Shift+R) und Seite neu laden.';
            return;
        }

        setLoading(true);
        summaryCard.style.display = 'none';
        fuzzyCard.style.display = 'none';
        recoReportCard.style.display = 'none';
        reportCard.style.display = 'none';

        const parseSuccess = await parseFiles();
        if (!parseSuccess) {
            setLoading(false);
            return;
        }

        try {
            const response = await fetch(`${WORKER_URL}/analyze`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    airtableData: parsedAirtableData,
                    erpData: parsedErpData
                })
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || `Server-Fehler: ${response.statusText}`);
            }

            const results = await response.json();
            renderSummary(results.summary);
            renderFuzzyMatchTable(results.suggestions);

            summaryCard.style.display = 'block';
            fuzzyCard.style.display = 'block';
            reportButton.style.display = 'block';

        } catch (error) {
            errorMessage.textContent = `Analyse-Fehler: ${error.message}`;
            console.error(error);
        } finally {
            setLoading(false);
        }
    }
    
    /**
     * Parsed die hochgeladenen Dateien.
     */
    async function parseFiles() {
        errorMessage.textContent = '';
        const airtableFile = airtableFileInput.files[0];
        const erpFile = erpFileInput.files[0];

        if (!airtableFile || !erpFile) {
            errorMessage.textContent = 'Bitte beide Dateien auswählen.';
            return false;
        }

        try {
            parsedAirtableData = await new Promise((resolve, reject) => {
                Papa.parse(airtableFile, {
                    header: true,
                    skipEmptyLines: true,
                    encoding: "ISO-8859-1", // oft 'latin1'
                    complete: (results) => resolve(results.data),
                    error: (err) => reject(new Error(`Airtable CSV-Fehler: ${err.message}`))
                });
            });

            if (erpFile.name.endsWith('.xlsx')) {
                const rows = await readXlsxFile(erpFile);
                const headers = rows[0];
                parsedErpData = rows.slice(1).map(row => {
                    let obj = {};
                    headers.forEach((header, index) => {
                        obj[header] = row[index];
                    });
                    return obj;
                });
            } else {
                parsedErpData = await new Promise((resolve, reject) => {
                    Papa.parse(erpFile, {
                        header: true,
                        skipEmptyLines: true,
                        complete: (results) => resolve(results.data),
                        error: (err) => reject(new Error(`ERP CSV-Fehler: ${err.message}`))
                    });
                });
            }
            return true;
        } catch (error) {
            errorMessage.textContent = `Fehler beim Parsen der Dateien: ${error.message}`;
            console.error(error);
            return false;
        }
    }

    /**
     * Hauptfunktion für die finalen Berichte (Phase 4)
     */
    async function handleReport() {
        setLoading(true);
        errorMessage.textContent = '';
        recoReportCard.style.display = 'none';
        reportCard.style.display = 'none';

        const confirmedMatches = [];
        const checkboxes = document.querySelectorAll('.fuzzy-checkbox:checked');
        checkboxes.forEach(box => {
            confirmedMatches.push(JSON.parse(box.dataset.match));
        });

        try {
            const response = await fetch(`${WORKER_URL}/report`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    airtableData: parsedAirtableData,
                    erpData: parsedErpData,
                    confirmedMatches: confirmedMatches
                })
            });

            if (!response.ok) {
                const err = await response.json();
                throw new Error(err.error || `Server-Fehler: ${response.statusText}`);
            }

            const reports = await response.json();

            // 1. NEUEN Abgleichs-Bericht rendern
            renderReconciliationReport(reports.reconciliation);
            recoReportCard.style.display = 'block';

            // 2. Finale Berichte rendern
            renderFinalReports(reports.finalReports);
            reportCard.style.display = 'block';

        } catch (error) {
            errorMessage.textContent = `Berichts-Fehler: ${error.message}`;
            console.error(error);
        } finally {
            setLoading(false);
        }
    }


    // ====== RENDER-FUNKTIONEN (ALT) ======

    function setLoading(isLoading) {
        loader.style.display = isLoading ? 'block' : 'none';
        analyzeButton.disabled = isLoading;
        reportButton.disabled = isLoading;
    }

    function renderSummary(summary) {
        summaryResultsDiv.innerHTML = `
            <ul>
                <li>✅ <strong>Perfekte Treffer:</strong> ${summary.perfectMatches}</li>
                <li>⚠️ <strong>Betrags-Diskrepanzen:</strong> ${summary.discrepancies} (KV stimmt, Betrag nicht)</li>
                <li>➡️ <strong>Nur im ERP gefunden:</strong> ${summary.erpOnly} (Mögliche Kandidaten für Fuzzy Match)</li>
                <li>⬅️ <strong>Nur in Airtable gefunden:</strong> ${summary.airtableOnly} (Mit einfacher KV)</li>
            </ul>
        `;
    }

    function renderFuzzyMatchTable(suggestions) {
        if (!suggestions || suggestions.length === 0) {
            fuzzyMatchDiv.innerHTML = '<p>Keine wahrscheinlichen Zuordnungen gefunden. Sie können direkt die Berichte generieren.</p>';
            return;
        }
        let tableHTML = `<table class="fuzzy-match-table"><thead><tr>
            <th>Auswählen</th><th>Airtable-Eintrag (ohne KV)</th><th>Gefundener ERP-Eintrag</th><th>Score</th>
            </tr></thead><tbody>`;
        suggestions.forEach((match, index) => {
            const matchData = {
                airtable_id: match.airtable.Projekttitel, 
                erp_kv: match.erp['KV-Nummer']
            };
            tableHTML += `
                <tr>
                    <td><input type="checkbox" class="fuzzy-checkbox" id="match-${index}" data-match='${JSON.stringify(matchData)}'></td>
                    <td><strong>Titel:</strong> ${match.airtable.Projekttitel}<br><strong>Kunde:</strong> ${match.airtable.Auftraggeber}<br><strong>Betrag:</strong> ${formatCurrency(match.airtable.Agenturleistung_netto_cleaned)}</td>
                    <td><strong>Titel:</strong> ${match.erp.Titel}<br><strong>Kunde:</strong> ${match.erp['Projekt Etat Kunde Name']}<br><strong>Betrag:</strong> ${formatCurrency(match.erp['Agenturleistung netto'])}<br><strong>KV:</strong> ${match.erp['KV-Nummer']}</td>
                    <td><strong>${(match.score * 100).toFixed(0)}%</strong></td>
                </tr>`;
        });
        tableHTML += '</tbody></table>';
        fuzzyMatchDiv.innerHTML = tableHTML;
    }

    function renderFinalReports(reports) {
        const format = (data) => `
            <table class="report-table">
                <thead><tr><th>Kategorie</th><th>Zugewiesener Betrag (Netto)</th></tr></thead>
                <tbody>
                    ${data.map(item => `
                        <tr><td>${item.name}</td><td>${formatCurrency(item.amount)}</td></tr>
                    `).join('')}
                </tbody>
            </table>`;
        teamReportDiv.innerHTML = format(reports.teamReport);
        personReportDiv.innerHTML = format(results.personReport); // FIX: reports.personReport
    }
    
    // ====== RENDER-FUNKTIONEN (NEU/ANGEPASST) ======
    
    function renderReconciliationReport(reco) {
        // 1. Finanz-Zusammenfassung
        recoTotalsDiv.innerHTML = `
            <div class="reco-section">
                <h4>Finanz-Zusammenfassung</h4>
                <div class="reco-totals-grid">
                    <div class="reco-totals-item">Gesamtsumme ERP <strong>${formatCurrency(reco.totals.totalERP)}</strong></div>
                    <div class="reco-totals-item">Gesamtsumme Airtable <strong>${formatCurrency(reco.totals.totalAirtable)}</strong></div>
                    <div class="reco-totals-item">Davon Zugeordnet <strong>${formatCurrency(reco.totals.totalReconciled)}</strong></div>
                    <div class="reco-totals-item">Fehlende ERP-Beträge <strong>${formatCurrency(reco.totals.totalUnreconciledERP)}</strong></div>
                </div>
            </div>`;
        
        // 2. Bericht 1: KVs zum Aktualisieren
        recoKVsToUpdateDiv.innerHTML = renderRecoTable(
            'To-Do: Beträge in Airtable aktualisieren (1:1-Treffer)',
            ['KV-Nummer', 'Projekttitel', 'Airtable-Betrag', 'ERP-Betrag (NEU)'],
            reco.kvsToUpdate.map(row => `
                <tr>
                    <td>${row.kv}</td>
                    <td>${row.title}</td>
                    <td class="amount">${formatCurrency(row.airtableAmount)}</td>
                    <td class="amount"><strong>${formatCurrency(row.erpAmount)}</strong></td>
                </tr>
            `)
        );
        
        // 3. Bericht 4: Nicht zugeordnete KVs (Fehlen in Airtable)
        recoUnmatchedERPDiv.innerHTML = renderRecoTable(
            'To-Do: Diese KVs fehlen in Airtable (oder Fuzzy-Match wurde abgelehnt)',
            ['KV-Nummer', 'ERP-Titel', 'ERP-Betrag'],
            reco.unmatchedERP.map(row => `
                <tr>
                    <td>${row.kv}</td>
                    <td>${row.erpTitle}</td>
                    <td class"amount">${formatCurrency(row.erpAmount)}</td>
                </tr>
            `)
        );

        // 4. Bericht 5: Nicht zugeordnete Airtable-Einträge
        recoUnmatchedAirtableDiv.innerHTML = renderRecoTable(
            'Info: Diese Airtable-Einträge fehlen im ERP',
            ['Airtable-Titel', 'Airtable-Betrag', 'Grund'],
            reco.unmatchedAirtable.map(row => `
                <tr>
                    <td>${row.airtableTitle}</td>
                    <td class="amount">${formatCurrency(row.airtableAmount)}</td>
                    <td>${row.reason}</td>
                </tr>
            `)
        );

        // 5. Bericht 3: Erfolgreich zugeordnete Fuzzy-Matches
        recoFuzzyMatchedDiv.innerHTML = renderRecoTable(
            'Info: Erfolgreich per Fuzzy-Match zugeordnet',
            ['Airtable-Titel', 'ERP-KV', 'ERP-Titel', 'Betrag'],
            reco.fuzzyMatched.map(row => `
                <tr>
                    <td>${row.airtableTitle}</td>
                    <td>${row.erpKV}</td>
                    <td>${row.erpTitle}</td>
                    <td class="amount">${formatCurrency(row.erpAmount)}</td>
                </tr>
            `)
        );
        
        // 6. NEU: Erfolgreich zugeordnete Sammel-KVs
        recoSammelKVsMatchedDiv.innerHTML = renderRecoTable(
            'Info: Erfolgreich zugeordnete Rahmenverträge (Sammel-KVs)',
            ['Airtable-Titel', 'Gefundene ERP KVs', 'Airtable-Betrag', 'ERP-Summe'],
            reco.sammelKVsMatched.map(row => `
                <tr>
                    <td>${row.airtableTitle}</td>
                    <td>${row.foundKVs.join(', ')}</td>
                    <td class="amount">${formatCurrency(row.sumAirtable)}</td>
                    <td class="amount"><strong>${formatCurrency(row.sumERP)}</strong></td>
                </tr>
            `)
        );

        // 7. NEU: Fehlgeschlagene Sammel-KVs
        recoSammelKVsFailedDiv.innerHTML = renderRecoTable(
            'Manuell zu prüfen: Sammel-KVs mit fehlenden Abrufen',
            ['Airtable-Titel', 'Airtable KVs (Original)', 'Fehlende ERP KVs'],
            reco.sammelKVsFailed.map(row => `
                <tr>
                    <td>${row.airtableTitle}</td>
                    <td>${row.airtableKVs}</td>
                    <td>${row.missingKVs.join(', ')}</td>
                </tr>
            `)
        );
    }
    
    /**
     * Hilfsfunktion zum Erstellen der HTML-Tabellen für den Abgleichs-Bericht
     */
    function renderRecoTable(title, headers, rows) {
        if (!rows || rows.length === 0) {
            return `<div class="reco-section"><h4>${title}</h4><p>Keine Einträge.</p></div>`;
        }
        
        const headerHTML = headers.map(h => 
            h.toLowerCase().includes('betrag') || h.toLowerCase().includes('summe') ? `<th class="amount">${h}</th>` : `<th>${h}</th>`
        ).join('');
        
        return `
            <div class="reco-section">
                <h4>${title} (${rows.length})</h4>
                <table class="reco-table">
                    <thead><tr>${headerHTML}</tr></thead>
                    <tbody>${rows.join('')}</tbody>
                </table>
            </div>
        `;
    }

    function formatCurrency(value) {
        const num = Number(value);
        if (isNaN(num)) return 'N/A';
        return num.toLocaleString('de-DE', { style: 'currency', currency: 'EUR' });
    }
};
