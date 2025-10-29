 * =================================================================
 * NEUE VERSION (v5)
 * - Projektnummer-basierte Logik
 * - Angepasste Berichte
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
    
    // NEUE/ANGEPASSTE Ergebnis-Divs für Abgleichs-Bericht
    const recoTotalsDiv = document.getElementById('reco-report-totals');
    const recoProjectsToUpdateDiv = document.getElementById('reco-report-projectsToUpdate'); // Umbenannt
    const recoUnmatchedERPProjDiv = document.getElementById('reco-report-unmatchedERP-byProject'); // Neu
    const recoUnmatchedAirtableProjDiv = document.getElementById('reco-report-unmatchedAirtable-byProject'); // Neu
    const recoFuzzyMatchedDiv = document.getElementById('reco-report-fuzzyMatched');
    const recoUnmatchedERPKVDiv = document.getElementById('reco-report-unmatchedERP-byKV'); // Neu
    const recoUnmatchedAirtableNoProjDiv = document.getElementById('reco-report-unmatchedAirtable-noProj'); // Neu


    // Event Listeners
    analyzeButton.addEventListener('click', handleAnalyze);
    reportButton.addEventListener('click', handleReport);

    
    /**
     * Hauptfunktion für den ersten Abgleich (Phase 1-3) - nur noch Fuzzy Vorschläge
     */
    async function handleAnalyze() {
        // ... (Bibliotheks-Checks bleiben gleich) ...
        if (typeof Papa === 'undefined') { errorMessage.textContent = 'Fehler: PapaParse...'; return; }
        if (typeof readXlsxFile === 'undefined') { errorMessage.textContent = 'Fehler: Read-Excel-File...'; return; }

        setLoading(true);
        summaryCard.style.display = 'none';
        fuzzyCard.style.display = 'none';
        recoReportCard.style.display = 'none';
        reportCard.style.display = 'none';

        const parseSuccess = await parseFiles();
        if (!parseSuccess) { setLoading(false); return; }

        try {
            const response = await fetch(`${WORKER_URL}/analyze`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    airtableData: parsedAirtableData,
                    erpData: parsedErpData
                })
            });

            if (!response.ok) { throw await response.json(); }

            const results = await response.json();
            
            // Render nur noch Summary und Fuzzy Vorschläge
            renderSummary_v5(results.summary); // Angepasstes Summary
            renderFuzzyMatchTable(results.suggestions); // Bleibt gleich

            summaryCard.style.display = 'block';
            fuzzyCard.style.display = 'block';
            reportButton.style.display = results.suggestions.length > 0 ? 'block' : 'block'; // Button immer anzeigen

        } catch (error) {
            errorMessage.textContent = `Analyse-Fehler: ${error.error || error.message}`;
            console.error(error);
        } finally {
            setLoading(false);
        }
    }
    
    /**
     * Parsed die hochgeladenen Dateien. (Bleibt gleich)
     */
    async function parseFiles() {
        errorMessage.textContent = '';
        const airtableFile = airtableFileInput.files[0];
        const erpFile = erpFileInput.files[0];
        if (!airtableFile || !erpFile) { errorMessage.textContent = 'Bitte beide Dateien auswählen.'; return false; }
        try {
            // Airtable CSV
            parsedAirtableData = await new Promise((resolve, reject) => {
                Papa.parse(airtableFile, { header: true, skipEmptyLines: true, encoding: "ISO-8859-1", complete: (r) => resolve(r.data), error: (e) => reject(e) });
            });
            // ERP (Excel or CSV)
            if (erpFile.name.endsWith('.xlsx')) {
                const rows = await readXlsxFile(erpFile);
                const headers = rows[0];
                parsedErpData = rows.slice(1).map(row => headers.reduce((obj, h, i) => (obj[h] = row[i], obj), {}));
            } else {
                parsedErpData = await new Promise((resolve, reject) => {
                    Papa.parse(erpFile, { header: true, skipEmptyLines: true, complete: (r) => resolve(r.data), error: (e) => reject(e) });
                });
            }
            return true;
        } catch (error) {
            errorMessage.textContent = `Fehler beim Parsen: ${error.message || error}`;
            console.error(error);
            return false;
        }
    }

    /**
     * Hauptfunktion für die Abgleichs- und finalen Berichte (Phase 4)
     */
    async function handleReport() {
        setLoading(true);
        errorMessage.textContent = '';
        recoReportCard.style.display = 'none';
        reportCard.style.display = 'none';

        const confirmedMatches = [];
        document.querySelectorAll('.fuzzy-checkbox:checked').forEach(box => {
            confirmedMatches.push(JSON.parse(box.dataset.match));
        });

        try {
            // WICHTIG: Verwende die NEUE Airtable-Datei für den Report!
            const currentAirtableFile = airtableFileInput.files[0];
             if (!currentAirtableFile || currentAirtableFile.name !== 'Airtable_102025.csv') {
                 // Parse die neue Airtable-Datei erneut, falls sie nicht die aktuelle ist
                 if (airtableFileInput.files.length === 0 || airtableFileInput.files[0].name !== 'Airtable_102025.csv') {
                      errorMessage.textContent = 'Bitte die korrekte Airtable-Datei (Airtable_102025.csv) hochladen.';
                      setLoading(false);
                      return;
                 }
                const parseSuccess = await parseFiles(); // Parse beide neu zur Sicherheit
                if (!parseSuccess) { setLoading(false); return; }
             }


            const response = await fetch(`${WORKER_URL}/report`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    airtableData: parsedAirtableData, // Immer die aktuell geparsten Daten senden
                    erpData: parsedErpData,
                    confirmedMatches: confirmedMatches
                })
            });

            if (!response.ok) { throw await response.json(); }

            const reports = await response.json();

            // 1. NEUEN Abgleichs-Bericht rendern
            renderReconciliationReport_v5(reports.reconciliation);
            recoReportCard.style.display = 'block';

            // 2. Finale Berichte rendern
            renderFinalReports_v5(reports.finalReports);
            reportCard.style.display = 'block';

        } catch (error) {
            errorMessage.textContent = `Berichts-Fehler: ${error.error || error.message}`;
            console.error(error);
        } finally {
            setLoading(false);
        }
    }


    // ====== RENDER-FUNKTIONEN (ANGEPASST für v5) ======

    function setLoading(isLoading) {
        loader.style.display = isLoading ? 'block' : 'none';
        analyzeButton.disabled = isLoading;
        reportButton.disabled = isLoading;
    }

    // Angepasstes Summary für v5
    function renderSummary_v5(summary) {
        summaryResultsDiv.innerHTML = `
            <p><strong>Kurzanalyse (Details im Abgleichs-Bericht):</strong></p>
            <ul>
                <li>Airtable-Einträge: ${summary.totalAirtable}</li>
                <li>ERP KVs: ${summary.totalERP_KV}</li>
                <li>ERP Projekte (gruppiert): ${summary.totalERP_Proj}</li>
                <li>Airtable-Einträge ohne Projektnummer: ${summary.airtableWithoutProjNr} (Kandidaten für Fuzzy Match)</li>
            </ul>
        `;
    }

    // Fuzzy Match Tabelle (bleibt gleich)
    function renderFuzzyMatchTable(suggestions) {
        if (!suggestions || suggestions.length === 0) {
            fuzzyMatchDiv.innerHTML = '<p>Keine wahrscheinlichen Fuzzy-Zuordnungen für Einträge ohne Projektnummer gefunden. Sie können direkt den Bericht generieren.</p>';
            reportButton.style.display = 'block'; // Button trotzdem anzeigen
            return;
        }
        let tableHTML = `<p><strong>Vorschläge für Airtable-Einträge OHNE Projektnummer:</strong></p>
                         <table class="fuzzy-match-table"><thead><tr>
                         <th>Auswählen</th><th>Airtable-Eintrag</th><th>Gefundener ERP-Eintrag (nicht über Projekt zugeordnet)</th><th>Score</th>
                         </tr></thead><tbody>`;
        suggestions.forEach((match, index) => {
            const matchData = { airtable_id: match.airtable.Projekttitel, erp_kv: match.erp['KV-Nummer'] };
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
        reportButton.style.display = 'block';
    }
    
    // Neuer Abgleichs-Bericht Renderer für v5
    function renderReconciliationReport_v5(reco) {
        // 1. Finanz-Zusammenfassung
        recoTotalsDiv.innerHTML = `<div class="reco-section"><h4>Finanz-Zusammenfassung</h4><div class="reco-totals-grid">
            <div class="reco-totals-item">Gesamtsumme ERP <strong>${formatCurrency(reco.totals.totalERP)}</strong></div>
            <div class="reco-totals-item">Gesamtsumme Airtable <strong>${formatCurrency(reco.totals.totalAirtable)}</strong></div>
            <div class="reco-totals-item">Davon Zugeordnet <strong>${formatCurrency(reco.totals.totalReconciled)}</strong></div>
            <div class="reco-totals-item">Fehlende ERP-Beträge <strong>${formatCurrency(reco.totals.totalUnreconciledERP)}</strong></div>
            </div></div>`;
        
        // 2. Projekte zum Aktualisieren
        recoProjectsToUpdateDiv.innerHTML = renderRecoTable(
            'To-Do: Beträge in Airtable aktualisieren (Projekt-Ebene)',
            ['Projekt-Nr.', 'Airtable-Titel', 'Airtable-Betrag', 'ERP-Gesamt (NEU)', ' Zugehörige ERP KVs'],
            reco.projectsToUpdate.map(row => `
                <tr>
                    <td>${row.projNr}</td>
                    <td>${row.airtableTitle}</td>
                    <td class="amount">${formatCurrency(row.airtableAmount)}</td>
                    <td class="amount"><strong>${formatCurrency(row.erpTotalAmount)}</strong></td>
                    <td>${row.erpKVs.join(', ')}</td>
                </tr>
            `)
        );
        
        // 3. Projekte nur im ERP
        recoUnmatchedERPProjDiv.innerHTML = renderRecoTable(
            'To-Do: Diese Projekte fehlen in Airtable',
            ['Projekt-Nr.', 'ERP KVs', 'ERP-Gesamt'],
            reco.unmatchedERP_byProject.map(row => `
                <tr>
                    <td>${row.projNr}</td>
                    <td>${row.erpKVs.map(kv => `${kv.kv} (${formatCurrency(kv.amount)})`).join('<br>')}</td>
                    <td class="amount">${formatCurrency(row.erpTotalAmount)}</td>
                </tr>
            `)
        );

        // 4. Projekte nur in Airtable
        recoUnmatchedAirtableProjDiv.innerHTML = renderRecoTable(
            'Info: Diese Airtable-Projekte fehlen im ERP',
            ['Projekt-Nr.', 'Airtable-Titel', 'Airtable-Betrag'],
            reco.unmatchedAirtable_byProject.map(row => `
                <tr>
                    <td>${row.projNr}</td>
                    <td>${row.airtableTitle}</td>
                    <td class="amount">${formatCurrency(row.airtableAmount)}</td>
                </tr>
            `)
        );

        // 5. Erfolgreiche Fuzzy Matches
        recoFuzzyMatchedDiv.innerHTML = renderRecoTable(
            'Info: Erfolgreich per Fuzzy-Match zugeordnet (Airtable ohne Projektnummer)',
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

        // 6. Restliche KVs nur im ERP
        recoUnmatchedERPKVDiv.innerHTML = renderRecoTable(
            'Manuell zu prüfen: Einzelne ERP KVs ohne Projekt-Match & ohne Fuzzy-Match',
            ['KV-Nummer', 'ERP-Titel', 'ERP-Betrag'],
            reco.unmatchedERP_byKV.map(row => `
                <tr>
                    <td>${row.kv}</td>
                    <td>${row.erpTitle}</td>
                    <td class="amount">${formatCurrency(row.erpAmount)}</td>
                </tr>
            `)
        );

        // 7. Restliche Airtable ohne Projektnummer
        recoUnmatchedAirtableNoProjDiv.innerHTML = renderRecoTable(
            'Manuell zu prüfen: Airtable-Einträge ohne Projektnummer & ohne Fuzzy-Match',
            ['Airtable-Titel', 'Airtable-Betrag'],
            reco.unmatchedAirtable_noProj.map(row => `
                <tr>
                    <td>${row.airtableTitle}</td>
                    <td class="amount">${formatCurrency(row.airtableAmount)}</td>
                </tr>
            `)
        );
    }
    
    // Finale Berichte Renderer (angepasst an neue response struktur)
    function renderFinalReports_v5(finalReports) {
        const format = (data) => `
            <table class="report-table">
                <thead><tr><th>Kategorie</th><th>Zugewiesener Betrag (Netto)</th></tr></thead>
                <tbody>
                    ${data && data.length > 0 ? data.map(item => `
                        <tr><td>${item.name}</td><td>${formatCurrency(item.amount)}</td></tr>
                    `).join('') : '<tr><td colspan="2">Keine Daten für diesen Bericht.</td></tr>'}
                </tbody>
            </table>`;
        teamReportDiv.innerHTML = format(finalReports.teamReport);
        personReportDiv.innerHTML = format(finalReports.personReport);
    }
    
    // Hilfsfunktion für Tabellen-Rendering (bleibt gleich)
    function renderRecoTable(title, headers, rows) {
        if (!rows || rows.length === 0) {
            return `<div class="reco-section"><h4>${title}</h4><p>Keine Einträge.</p></div>`;
        }
        const headerHTML = headers.map(h => h.toLowerCase().includes('betrag') || h.toLowerCase().includes('summe') ? `<th class="amount">${h}</th>` : `<th>${h}</th>`).join('');
        return `<div class="reco-section"><h4>${title} (${rows.length})</h4><table class="reco-table">
                <thead><tr>${headerHTML}</tr></thead><tbody>${rows.join('')}</tbody></table></div>`;
    }

    // Hilfsfunktion für Währungsformatierung (bleibt gleich)
    function formatCurrency(value) {
        const num = Number(value);
        if (isNaN(num) || value === null) return 'N/A';
        return num.toLocaleString('de-DE', { style: 'currency', currency: 'EUR' });
    }
};
