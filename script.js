/*
 * =================================================================
 * NEUE VERSION: Der gesamte Code wird in window.onload gewrappt,
 * um sicherzustellen, dass alle Bibliotheken (PapaParse etc.)
 * geladen sind, bevor unser Code ausgeführt wird.
 * =================================================================
 */
window.onload = function() {
    
    // ====== KONFIGURATION ======
    // Trage hier die URL deines Cloudflare Workers ein
const WORKER_URL = 'https://tekin-reco-backend-tool.tekin-6af.workers.dev';     // ==========================

    // Globale Variablen zum Speichern der geparsten Daten
    let parsedAirtableData = null;
    let parsedErpData = null;

    // DOM-Elemente
    // Wir müssen sie *innerhalb* von window.onload suchen,
    // da sie erst dann garantiert existieren.
    const analyzeButton = document.getElementById('analyze-button');
    const reportButton = document.getElementById('report-button');
    const airtableFileInput = document.getElementById('airtable-file');
    const erpFileInput = document.getElementById('erp-file');
    const loader = document.getElementById('loader');
    const errorMessage = document.getElementById('error-message');

    // Karten-Container
    const summaryCard = document.getElementById('summary-results-card');
    const fuzzyCard = document.getElementById('fuzzy-match-card');
    const reportCard = document.getElementById('final-report-card');

    // Ergebnis-Divs
    const summaryResultsDiv = document.getElementById('summary-results');
    const fuzzyMatchDiv = document.getElementById('fuzzy-match-area');
    const teamReportDiv = document.getElementById('team-report-area');
    const personReportDiv = document.getElementById('person-report-area');

    
    // Event Listener für den "Abgleich starten"-Button
    analyzeButton.addEventListener('click', handleAnalyze);

    // Event Listener für den "Berichte generieren"-Button
    reportButton.addEventListener('click', handleReport);

    
    /**
     * Hauptfunktion für den ersten Abgleich (Phase 1-3)
     */
    async function handleAnalyze() {
        
        // NEUER CHECK: Wir prüfen manuell, ob die Bibliotheken geladen sind
        if (typeof Papa === 'undefined') {
            errorMessage.textContent = 'Fehler: PapaParse (CSV-Bibliothek) konnte nicht geladen werden. Bitte Ad-Blocker prüfen und Seite neu laden.';
            return;
        }
        if (typeof readXlsxFile === 'undefined') {
            errorMessage.textContent = 'Fehler: Read-Excel-File (Excel-Bibliothek) konnte nicht geladen werden. Bitte Ad-Blocker prüfen und Seite neu laden.';
            return;
        }

        // UI zurücksetzen
        setLoading(true);
        summaryCard.style.display = 'none';
        fuzzyCard.style.display = 'none';
        reportCard.style.display = 'none';

        // 1. Dateien parsen
        const parseSuccess = await parseFiles();
        if (!parseSuccess) {
            setLoading(false);
            return;
        }

        try {
            // 2. Daten an den Worker /analyze Endpunkt senden
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

            // 3. Ergebnisse rendern
            renderSummary(results.summary);
            renderFuzzyMatchTable(results.suggestions);

            // 4. UI-Karten anzeigen
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
     * Parsed die hochgeladenen Dateien und speichert sie global.
     * @returns {Promise<boolean>} True bei Erfolg, False bei Fehler.
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
            // Parse Airtable CSV
            parsedAirtableData = await new Promise((resolve, reject) => {
                // HIER WAR DER FEHLER: "Papa" war undefined.
                Papa.parse(airtableFile, {
                    header: true,
                    skipEmptyLines: true,
                    encoding: "ISO-8859-1", // oft 'latin1'
                    complete: (results) => resolve(results.data),
                    error: (err) => reject(new Error(`Airtable CSV-Fehler: ${err.message}`))
                });
            });

            // Parse ERP (Excel oder CSV)
            if (erpFile.name.endsWith('.xlsx')) {
                // read-excel-file gibt ein Array von Arrays zurück. Wir konvertieren es in ein Array von Objekten.
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
                // Parse ERP CSV
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
        reportCard.style.display = 'none';

        // 1. Bestätigte Matches aus der Tabelle sammeln
        const confirmedMatches = [];
        const checkboxes = document.querySelectorAll('.fuzzy-checkbox:checked');
        checkboxes.forEach(box => {
            confirmedMatches.push(JSON.parse(box.dataset.match));
        });

        try {
            // 2. Originaldaten + bestätigte Matches an den /report Endpunkt senden
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

            // 3. Finale Berichte rendern
            renderFinalReports(reports);
            reportCard.style.display = 'block';

        } catch (error) {
            errorMessage.textContent = `Berichts-Fehler: ${error.message}`;
            console.error(error);
        } finally {
            setLoading(false);
        }
    }


    // ====== RENDER-FUNKTIONEN ======

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

        let tableHTML = `
            <table class="fuzzy-match-table">
                <thead>
                    <tr>
                        <th>Auswählen</th>
                        <th>Airtable-Eintrag (ohne KV)</th>
                        <th>ERP-Eintrag (fehlt in Airtable)</th>
                        <th>Score</th>
                    </tr>
                </thead>
                <tbody>
        `;

        suggestions.forEach((match, index) => {
            // Daten für das Checkbox-Attribut vorbereiten
            const matchData = {
                // Wir verwenden jetzt den Projekttitel als ID, da er im Backend-Map-Key verwendet wird
                airtable_id: match.airtable.Projekttitel, 
                erp_kv: match.erp['KV-Nummer']
            };
            
            tableHTML += `
                <tr>
                    <td>
                        <input type="checkbox" class="fuzzy-checkbox" 
                               id="match-${index}" 
                               data-match='${JSON.stringify(matchData)}'>
                    </td>
                    <td>
                        <strong>Titel:</strong> ${match.airtable.Projekttitel}<br>
                        <strong>Kunde:</strong> ${match.airtable.Auftraggeber}<br>
                        <strong>Betrag:</strong> ${formatCurrency(match.airtable.Agenturleistung_netto_cleaned)}
                    </td>
                    <td>
                        <strong>Titel:</strong> ${match.erp.Titel}<br>
                        <strong>Kunde:</strong> ${match.erp['Projekt Etat Kunde Name']}<br>
                        <strong>Betrag:</strong> ${formatCurrency(match.erp['Agenturleistung netto'])}<br>
                        <strong>KV:</strong> ${match.erp['KV-Nummer']}
                    </td>
                    <td><strong>${(match.score * 100).toFixed(0)}%</strong></td>
                </tr>
            `;
        });

        tableHTML += '</tbody></table>';
        fuzzyMatchDiv.innerHTML = tableHTML;
    }

    function renderFinalReports(reports) {
        const format = (data) => `
            <table class="report-table">
                <thead>
                    <tr>
                        <th>Kategorie</th>
                        <th>Zugewiesener Betrag (Netto)</th>
                    </tr>
                </thead>
                <tbody>
                    ${data.map(item => `
                        <tr>
                            <td>${item.name}</td>
                            <td>${formatCurrency(item.amount)}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        `;

        teamReportDiv.innerHTML = format(reports.teamReport);
        personReportDiv.innerHTML = format(reports.personReport);
    }

    function formatCurrency(value) {
        const num = Number(value);
        if (isNaN(num)) return 'N/A';
        return num.toLocaleString('de-DE', { style: 'currency', currency: 'EUR' });
    }
}; // ENDE DES window.onload WRAPPERS
