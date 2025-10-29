/*
 * =================================================================
 * SCRIPT v6 (Angepasst an optimierten Worker)
 * - handleAnalyze verarbeitet nur noch 'suggestions'
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
    const summaryResultsDiv = document.getElementById('summary-results'); // Wird jetzt weniger genutzt
    const fuzzyMatchDiv = document.getElementById('fuzzy-match-area');
    const teamReportDiv = document.getElementById('team-report-area');
    const personReportDiv = document.getElementById('person-report-area');

    // Ergebnis-Divs für Abgleichs-Bericht (v5 Struktur)
    const recoTotalsDiv = document.getElementById('reco-report-totals');
    const recoProjectsToUpdateDiv = document.getElementById('reco-report-projectsToUpdate');
    const recoUnmatchedERPProjDiv = document.getElementById('reco-report-unmatchedERP-byProject');
    const recoUnmatchedAirtableProjDiv = document.getElementById('reco-report-unmatchedAirtable-byProject');
    const recoFuzzyMatchedDiv = document.getElementById('reco-report-fuzzyMatched');
    const recoUnmatchedERPKVDiv = document.getElementById('reco-report-unmatchedERP-byKV');
    const recoUnmatchedAirtableNoProjDiv = document.getElementById('reco-report-unmatchedAirtable-noProj');


    // Event Listeners
    analyzeButton.addEventListener('click', handleAnalyze);
    reportButton.addEventListener('click', handleReport);


    /**
     * Hauptfunktion für den ersten Abgleich (nur Fuzzy Vorschläge)
     */
    async function handleAnalyze() {
        if (typeof Papa === 'undefined') { errorMessage.textContent = 'Fehler: PapaParse...'; return; }
        if (typeof readXlsxFile === 'undefined') { errorMessage.textContent = 'Fehler: Read-Excel-File...'; return; }

        setLoading(true);
        // summaryCard.style.display = 'none'; // Summary wird nicht mehr direkt angezeigt
        fuzzyCard.style.display = 'none';
        recoReportCard.style.display = 'none';
        reportCard.style.display = 'none';
        errorMessage.textContent = ''; // Fehler löschen

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

            if (!response.ok) {
                 const errText = await response.text();
                 try { throw JSON.parse(errText); } catch (e) { throw new Error(errText || `Server-Fehler: ${response.statusText}`); }
            }

            const results = await response.json(); // Erwartet nur { suggestions: [...] }

            // Render nur Fuzzy Vorschläge
            renderFuzzyMatchTable(results.suggestions);

            // summaryCard.style.display = 'block'; // Summary nicht mehr anzeigen
            fuzzyCard.style.display = 'block';
            reportButton.style.display = 'block'; // Button immer anzeigen

        } catch (error) {
            errorMessage.textContent = `Analyse-Fehler: ${error.error || error.message || error}`;
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
        parsedAirtableData = null;
        parsedErpData = null;
        const airtableFile = airtableFileInput.files[0];
        const erpFile = erpFileInput.files[0];
        if (!airtableFile || !erpFile) { errorMessage.textContent = 'Bitte beide Dateien auswählen.'; return false; }
        try {
            // Airtable CSV
            parsedAirtableData = await new Promise((resolve, reject) => {
                Papa.parse(airtableFile, { header: true, skipEmptyLines: 'greedy', encoding: "ISO-8859-1",
                    complete: (r) => resolve(r.data.filter(row => row && Object.values(row).some(val => val !== null && val !== ''))),
                    error: (e) => reject(new Error(`Airtable CSV-Fehler: ${e.message}`))
                });
            });
            // ERP (Excel or CSV)
            if (erpFile.name.endsWith('.xlsx')) {
                const rows = await readXlsxFile(erpFile);
                 if (!rows || rows.length < 2) throw new Error('ERP Excel-Datei leer oder nur Header.');
                const headers = rows[0].map(h => String(h));
                parsedErpData = rows.slice(1).map(row => headers.reduce((obj, h, i) => (obj[h] = row[i], obj), {}))
                                  .filter(row => row && Object.values(row).some(val => val !== null && val !== undefined && val !== ''));
            } else {
                parsedErpData = await new Promise((resolve, reject) => {
                    Papa.parse(erpFile, { header: true, skipEmptyLines: 'greedy',
                        complete: (r) => resolve(r.data.filter(row => row && Object.values(row).some(val => val !== null && val !== ''))),
                        error: (e) => reject(new Error(`ERP CSV-Fehler: ${e.message}`))
                    });
                });
            }
             if (!parsedAirtableData || parsedAirtableData.length === 0) throw new Error('Airtable-Datei ist leer oder konnte nicht geparst werden.');
             if (!parsedErpData || parsedErpData.length === 0) throw new Error('ERP-Datei ist leer oder konnte nicht geparst werden.');
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
        if (!parsedAirtableData || !parsedErpData) {
            errorMessage.textContent = 'Fehler: Dateidaten nicht vorhanden. Bitte laden Sie die Dateien erneut hoch und starten Sie den Abgleich.';
            return;
        }
        setLoading(true);
        errorMessage.textContent = '';
        recoReportCard.style.display = 'none';
        reportCard.style.display = 'none';

        const confirmedMatches = [];
        document.querySelectorAll('.fuzzy-checkbox:checked').forEach(box => {
            try { confirmedMatches.push(JSON.parse(box.dataset.match)); } catch (e) { console.error("Fehler beim Parsen der Fuzzy-Match-Daten:", box.dataset.match, e); }
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
                 const errText = await response.text();
                 try { throw JSON.parse(errText); } catch (e) { throw new Error(errText || `Server-Fehler: ${response.statusText}`); }
            }

            const reports = await response.json();

            // 1. NEUEN Abgleichs-Bericht rendern
            renderReconciliationReport_v5(reports.reconciliation);
            recoReportCard.style.display = 'block';

            // 2. Finale Berichte rendern
            renderFinalReports_v5(reports.finalReports);
            reportCard.style.display = 'block';

        } catch (error) {
            errorMessage.textContent = `Berichts-Fehler: ${error.error || error.message || error}`;
            console.error(error);
        } finally {
            setLoading(false);
        }
    }


    // ====== RENDER-FUNKTIONEN (ANGEPASST für v6) ======

    function setLoading(isLoading) {
        loader.style.display = isLoading ? 'block' : 'none';
        analyzeButton.disabled = isLoading;
        reportButton.disabled = isLoading;
    }

    // Summary wird nicht mehr vom /analyze endpoint geliefert
     function renderSummary_v5(summary) {
         // Wir könnten hier eine einfachere Info anzeigen, die wir im Frontend berechnen
         summaryResultsDiv.innerHTML = `<p>Fuzzy-Match Vorschläge werden unten angezeigt (falls vorhanden). Klicken Sie auf '2. Finale Berichte generieren', um den vollständigen Abgleich durchzuführen.</p>`;
     }


    // Fuzzy Match Tabelle (bleibt gleich)
    function renderFuzzyMatchTable(suggestions) {
        // (Code von v5 bleibt hier unverändert)
        if (!suggestions || suggestions.length === 0) {
            fuzzyMatchDiv.innerHTML = '<p>Keine wahrscheinlichen Fuzzy-Zuordnungen für Einträge ohne Projektnummer gefunden. Sie können direkt den Bericht generieren.</p>';
            // reportButton.style.display = 'block'; // Wird jetzt immer angezeigt
            return;
        }
        let tableHTML = `<p><strong>Vorschläge für Airtable-Einträge OHNE Projektnummer:</strong></p>
                         <table class="fuzzy-match-table"><thead><tr>
                         <th>Auswählen</th><th>Airtable-Eintrag</th><th>Gefundener ERP-Eintrag (nicht über Projekt zugeordnet)</th><th>Score (%)</th>
                         </tr></thead><tbody>`;
        suggestions.forEach((match, index) => {
            const cleanedMatch = {
                airtable_id: match.airtable.Projekttitel?.replace(/"/g, "'") || '',
                erp_kv: match.erp['KV-Nummer']?.replace(/"/g, "'") || ''
             };
            const matchDataString = JSON.stringify(cleanedMatch).replace(/'/g, "&apos;");

            tableHTML += `
                <tr>
                    <td><input type="checkbox" class="fuzzy-checkbox" id="match-${index}" data-match='${matchDataString}'></td>
                    <td><strong>Titel:</strong> ${escapeHtml(match.airtable.Projekttitel)}<br><strong>Kunde:</strong> ${escapeHtml(match.airtable.Auftraggeber)}<br><strong>Betrag:</strong> ${formatCurrency(match.airtable.Agenturleistung_netto_cleaned)}</td>
                    <td><strong>Titel:</strong> ${escapeHtml(match.erp.Titel)}<br><strong>Kunde:</strong> ${escapeHtml(match.erp['Projekt Etat Kunde Name'])}<br><strong>Betrag:</strong> ${formatCurrency(match.erp['Agenturleistung netto'])}<br><strong>KV:</strong> ${escapeHtml(match.erp['KV-Nummer'])}</td>
                    <td><strong>${(match.score * 100).toFixed(0)}</strong></td>
                </tr>`;
        });
        tableHTML += '</tbody></table>';
        fuzzyMatchDiv.innerHTML = tableHTML;
    }

    // Abgleichs-Bericht Renderer (bleibt gleich wie v5)
    function renderReconciliationReport_v5(reco) {
        // (Code von v5 bleibt hier unverändert)
        recoTotalsDiv.innerHTML = `<div class="reco-section"><h4>Finanz-Zusammenfassung</h4><div class="reco-totals-grid">
            <div class="reco-totals-item">Gesamtsumme ERP <strong>${formatCurrency(reco.totals.totalERP)}</strong></div>
            <div class="reco-totals-item">Gesamtsumme Airtable <strong>${formatCurrency(reco.totals.totalAirtable)}</strong></div>
            <div class="reco-totals-item">Davon Zugeordnet <strong>${formatCurrency(reco.totals.totalReconciled)}</strong></div>
            <div class="reco-totals-item">Fehlende ERP-Beträge <strong>${formatCurrency(reco.totals.totalUnreconciledERP)}</strong></div>
            </div></div>`;
        recoProjectsToUpdateDiv.innerHTML = renderRecoTable(
            'To-Do: Beträge in Airtable aktualisieren (Projekt-Ebene)',
            ['Projekt-Nr.', 'Airtable-Titel', 'Airtable-Betrag', 'ERP-Gesamt (NEU)', ' Zugehörige ERP KVs'],
            reco.projectsToUpdate.map(row => `
                <tr><td>${escapeHtml(row.projNr)}</td><td>${escapeHtml(row.airtableTitle)}</td><td class="amount">${formatCurrency(row.airtableAmount)}</td>
                    <td class="amount"><strong>${formatCurrency(row.erpTotalAmount)}</strong></td><td>${escapeHtml(row.erpKVs.join(', '))}</td></tr>`)
        );
        recoUnmatchedERPProjDiv.innerHTML = renderRecoTable(
            'To-Do: Diese Projekte fehlen in Airtable',
            ['Projekt-Nr.', 'ERP KVs (Titel, Betrag)', 'ERP-Gesamt'],
            reco.unmatchedERP_byProject.map(row => `
                <tr><td>${escapeHtml(row.projNr)}</td><td>${row.erpKVs.map(kv => `${escapeHtml(kv.kv)} (${escapeHtml(kv.title)}, ${formatCurrency(kv.amount)})`).join('<br>')}</td>
                    <td class="amount">${formatCurrency(row.erpTotalAmount)}</td></tr>`)
        );
        recoUnmatchedAirtableProjDiv.innerHTML = renderRecoTable(
            'Info: Diese Airtable-Projekte fehlen im ERP',
            ['Projekt-Nr.', 'Airtable-Titel', 'Airtable-Betrag'],
            reco.unmatchedAirtable_byProject.map(row => `
                <tr><td>${escapeHtml(row.projNr)}</td><td>${escapeHtml(row.airtableTitle)}</td><td class="amount">${formatCurrency(row.airtableAmount)}</td></tr>`)
        );
        recoFuzzyMatchedDiv.innerHTML = renderRecoTable(
            'Info: Erfolgreich per Fuzzy-Match zugeordnet (Airtable ohne Projektnummer)',
            ['Airtable-Titel', 'ERP-KV', 'ERP-Titel', 'Betrag'],
            reco.fuzzyMatched.map(row => `
                <tr><td>${escapeHtml(row.airtableTitle)}</td><td>${escapeHtml(row.erpKV)}</td><td>${escapeHtml(row.erpTitle)}</td><td class="amount">${formatCurrency(row.erpAmount)}</td></tr>`)
        );
        recoUnmatchedERPKVDiv.innerHTML = renderRecoTable(
            'Manuell zu prüfen: Einzelne ERP KVs ohne Projekt-Match & ohne Fuzzy-Match',
            ['KV-Nummer', 'ERP-Titel', 'ERP-Betrag'],
            reco.unmatchedERP_byKV.map(row => `
                <tr><td>${escapeHtml(row.kv)}</td><td>${escapeHtml(row.erpTitle)}</td><td class="amount">${formatCurrency(row.erpAmount)}</td></tr>`)
        );
        recoUnmatchedAirtableNoProjDiv.innerHTML = renderRecoTable(
            'Manuell zu prüfen: Airtable-Einträge ohne Projektnummer & ohne Fuzzy-Match',
            ['Airtable-Titel', 'Airtable-Betrag'],
            reco.unmatchedAirtable_noProj.map(row => `
                <tr><td>${escapeHtml(row.airtableTitle)}</td><td class="amount">${formatCurrency(row.airtableAmount)}</td></tr>`)
        );
    }

    // Finale Berichte Renderer (bleibt gleich wie v5)
    function renderFinalReports_v5(finalReports) {
        // (Code von v5 bleibt hier unverändert)
        const format = (data, categoryLabel = 'Kategorie') => `
            <table class="report-table"><thead><tr><th>${categoryLabel}</th><th>Zugewiesener Betrag (Netto)</th></tr></thead><tbody>
                ${data && data.length > 0 ? data.map(item => `<tr><td>${escapeHtml(item.name)}</td><td>${formatCurrency(item.amount)}</td></tr>`).join('') : `<tr><td colspan="2">Keine Daten.</td></tr>`}
            </tbody></table>`;
        teamReportDiv.innerHTML = format(finalReports.teamReport, 'Team');
        personReportDiv.innerHTML = format(finalReports.personReport, 'Person');
    }

    // Hilfsfunktion für Tabellen-Rendering (bleibt gleich wie v5)
    function renderRecoTable(title, headers, rows) {
        // (Code von v5 bleibt hier unverändert)
         if (!rows || rows.length === 0) return `<div class="reco-section"><h4>${escapeHtml(title)}</h4><p>Keine Einträge.</p></div>`;
        const headerHTML = headers.map(h => h.toLowerCase().includes('betrag') || h.toLowerCase().includes('summe') ? `<th class="amount">${escapeHtml(h)}</th>` : `<th>${escapeHtml(h)}</th>`).join('');
        return `<div class="reco-section"><h4>${escapeHtml(title)} (${rows.length})</h4><table class="reco-table"><thead><tr>${headerHTML}</tr></thead><tbody>${rows.join('')}</tbody></table></div>`;
    }

    // Hilfsfunktion für Währungsformatierung (bleibt gleich wie v5)
    function formatCurrency(value) {
        // (Code von v5 bleibt hier unverändert)
         const num = Number(value);
         if (isNaN(num) || value === null || value === undefined) return 'N/A';
         return num.toLocaleString('de-DE', { style: 'currency', currency: 'EUR' });
    }

    // Hilfsfunktion zum Escapen von HTML (bleibt gleich wie v5)
    function escapeHtml(unsafe) {
        // (Code von v5 bleibt hier unverändert)
         if (typeof unsafe !== 'string') return unsafe;
         return unsafe.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
    }

}; // ENDE DES window.onload WRAPPERS
