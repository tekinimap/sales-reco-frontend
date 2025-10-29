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
        // WICHTIG: Sicherstellen, dass die NEUE Airtable-Datei verwendet wird
        if (airtableFile.name !== 'Airtable_102025.csv') {
             errorMessage.textContent = 'Falsche Airtable-Datei. Bitte "Airtable_102025.csv" hochladen.';
             airtableFileInput.value = ''; // Input zurücksetzen
             return false;
        }

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
        const format = (data, categoryLabel = 'Kategorie') => `
            <table class="report-table"><thead><tr><th>${categoryLabel}</th><th>Zugewiesener Betrag (Netto)</th></tr></thead><tbody>
                ${data && data.length > 0 ? data.map(item => `<tr><td>${escapeHtml(item.name)}</td><td>${formatCurrency(item.amount)}</td></tr>`).join('') : `<tr><td colspan="2">Keine Daten für diesen Bericht.</td></tr>`}
            </tbody></table>`;
        teamReportDiv.innerHTML = format(finalReports.teamReport, 'Team');
        personReportDiv.innerHTML = format(finalReports.personReport, 'Person');
    }

    // Hilfsfunktion für Tabellen-Rendering
    function renderRecoTable(title, headers, rows) {
         if (!rows || rows.length === 0) return `<div class="reco-section"><h4>${escapeHtml(title)}</h4><p>Keine Einträge.</p></div>`;
        const headerHTML = headers.map(h => h.toLowerCase().includes('betrag') || h.toLowerCase().includes('summe') ? `<th class="amount">${escapeHtml(h)}</th>` : `<th>${escapeHtml(h)}</th>`).join('');
        return `<div class="reco-section"><h4>${escapeHtml(title)} (${rows.length})</h4><table class="reco-table"><thead><tr>${headerHTML}</tr></thead><tbody>${rows.join('')}</tbody></table></div>`;
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
