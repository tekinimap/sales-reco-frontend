"use strict";

/**
 * Normalisiert Projektnummern, indem überflüssige Whitespaces entfernt werden
 * und geschützte Leerzeichen sowie Tabulatoren in reguläre Spaces umgewandelt werden.
 *
 * @param {string|undefined|null} value
 * @returns {string}
 */
function normalizeProjectKey(value) {
  return (value ?? "")
    .toString()
    .replace(/\u00A0/g, " ")
    .replace(/[\t\r\n]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Parsen von Währungswerten im (deutschen) CSV-Format. Unterstützt Tausenderpunkte,
 * Dezimalkommas sowie optionale Währungssymbole (z. B. "€").
 *
 * @param {string|number|undefined|null} value
 * @returns {number}
 */
function cleanAirtableCurrency(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number" && Number.isFinite(value)) return value;

  const raw = String(value)
    .replace(/\u00A0/g, " ")
    .replace(/€/g, "")
    .replace(/[^0-9,\-\.]/g, "")
    .trim();

  if (!raw) return 0;

  let normalized = raw;
  const hasComma = raw.includes(",");
  const hasDot = raw.includes(".");

  if (hasComma && hasDot) {
    // Tausenderpunkte entfernen, Dezimalkomma in Punkt wandeln
    normalized = raw.replace(/\./g, "").replace(/,/g, ".");
  } else if (hasComma && !hasDot) {
    normalized = raw.replace(/,/g, ".");
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
}

/**
 * Alias für ERP-Währungswerte – nutzt dieselbe Logik wie Airtable.
 *
 * @param {string|number|undefined|null} value
 * @returns {number}
 */
function cleanERPCurrency(value) {
  return cleanAirtableCurrency(value);
}

const ERP_AMOUNT_CANDIDATE_KEYS = [
  "Summe Netto",
  "Summe (Netto)",
  "Netto Summe",
  "Netto",
  "Netto Betrag",
  "Betrag Netto",
  "Gesamt Netto",
  "Gesamt (Netto)",
  "Gesamtbetrag Netto",
  "Betrag",
  "Agenturleistung (netto)",
];

const AIRTABLE_TITLE_CANDIDATE_KEYS = [
  "Projekttitel",
  "Projekt Titel",
  "Projektname",
  "Projekt",
  "Titel",
];

const ERP_TITLE_CANDIDATE_KEYS = [
  "Projekt Titel",
  "Projekt",
  "Projektname",
  "Titel",
  "Leistungsbeschreibung",
  "Bezeichnung",
];

const KV_TITLE_CANDIDATE_KEYS = [
  "KV-Titel",
  "Titel",
  "Beschreibung",
  "Leistungsbeschreibung",
  "Bezeichnung",
];

const KV_NUMBER_ALIASES = [
  "KV-Nummer",
  "KV Nummer",
  "KV-Nr.",
  "KV Nr.",
  "KV_Nr",
];

function extractFirstAvailable(row, keys) {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(row, key) && row[key] !== undefined && row[key] !== null) {
      return row[key];
    }
  }
  return "";
}

function extractErpAmount(row) {
  for (const key of ERP_AMOUNT_CANDIDATE_KEYS) {
    if (Object.prototype.hasOwnProperty.call(row, key)) {
      const amount = cleanERPCurrency(row[key]);
      if (amount !== 0) return amount;
    }
  }

  // Fallback: erstes Feld, das "netto" oder "betrag" im Namen enthält
  for (const key of Object.keys(row)) {
    const lower = key.toLowerCase();
    if (lower.includes("netto") || lower.includes("betrag")) {
      const amount = cleanERPCurrency(row[key]);
      if (amount !== 0) return amount;
    }
  }

  return 0;
}

function roundCurrency(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.round(num * 100) / 100;
}

function addCurrency(a, b) {
  return roundCurrency(roundCurrency(a) + roundCurrency(b));
}

/**
 * Gruppiert Airtable-Zeilen anhand der Projektnummer und summiert die Beträge pro Projekt.
 *
 * @param {Array<object>} rows
 * @returns {{airtableMapByProjNr: Map<string, object>, unmatchedAirtable_noProj: Array<object>, totalAirtable: number}}
 */
function cleanAndSegmentAirtable_v7(rows = []) {
  const airtableMapByProjNr = new Map();
  const unmatchedAirtable_noProj = [];
  let totalAirtable = 0;

  for (const inputRow of rows || []) {
    if (!inputRow || typeof inputRow !== "object") continue;
    const row = { ...inputRow };
    const projNr = normalizeProjectKey(row["Projektnummer"]);
    const amount = cleanAirtableCurrency(row["Agenturleistung (netto)"]);

    row.Agenturleistung_netto_cleaned = amount;
    totalAirtable = addCurrency(totalAirtable, amount);

    if (projNr) {
      const acc = airtableMapByProjNr.get(projNr) || {
        Projektnummer: projNr,
        Projekttitel: extractFirstAvailable(row, AIRTABLE_TITLE_CANDIDATE_KEYS) || "",
        AirtableRows: [],
        AirtableTotal: 0,
      };

      if (!acc.Projekttitel) {
        acc.Projekttitel = extractFirstAvailable(row, AIRTABLE_TITLE_CANDIDATE_KEYS) || acc.Projekttitel;
      }

      acc.AirtableRows.push(row);
      acc.AirtableTotal = addCurrency(acc.AirtableTotal, amount);

      airtableMapByProjNr.set(projNr, acc);
    } else {
      unmatchedAirtable_noProj.push({
        airtableTitle: extractFirstAvailable(row, AIRTABLE_TITLE_CANDIDATE_KEYS) || "",
        airtableAmount: amount,
        row,
      });
    }
  }

  return {
    airtableMapByProjNr,
    unmatchedAirtable_noProj,
    totalAirtable: roundCurrency(totalAirtable),
  };
}

/**
 * Gruppiert ERP-Zeilen anhand der Projektnummer.
 *
 * @param {Array<object>} rows
 * @returns {{erpMapByProjNr: Map<string, object>, unmatchedERP_byKV: Array<object>, totalERP: number}}
 */
function segmentErpData_v7(rows = []) {
  const erpMapByProjNr = new Map();
  const unmatchedERP_byKV = [];
  let totalERP = 0;

  for (const inputRow of rows || []) {
    if (!inputRow || typeof inputRow !== "object") continue;
    const row = { ...inputRow };
    const projNr = normalizeProjectKey(row["Projekt Projektnummer"]);
    const amount = extractErpAmount(row);
    row._erpAmount_cleaned = amount;

    totalERP = addCurrency(totalERP, amount);

    const kvNumber = normalizeProjectKey(extractFirstAvailable(row, KV_NUMBER_ALIASES)) || "";
    const kvTitle = extractFirstAvailable(row, KV_TITLE_CANDIDATE_KEYS) || "";

    if (projNr) {
      const projectTitle = extractFirstAvailable(row, ERP_TITLE_CANDIDATE_KEYS) || "";
      const acc = erpMapByProjNr.get(projNr) || {
        projNr,
        projectTitle,
        total: 0,
        kvs: [],
      };

      if (!acc.projectTitle) {
        acc.projectTitle = projectTitle;
      }

      acc.total = addCurrency(acc.total, amount);
      acc.kvs.push({
        "KV-Nummer": kvNumber,
        title: kvTitle,
        amount,
        row,
      });

      erpMapByProjNr.set(projNr, acc);
    } else {
      unmatchedERP_byKV.push({
        kv: kvNumber,
        erpTitle: kvTitle,
        erpAmount: amount,
        row,
      });
    }
  }

  return {
    erpMapByProjNr,
    unmatchedERP_byKV,
    totalERP: roundCurrency(totalERP),
  };
}

/**
 * Erstellt den Abgleichsbericht (reconciliation) basierend auf aggregierten Airtable- und ERP-Daten.
 *
 * @param {Array<object>} airtableRows
 * @param {Array<object>} erpRows
 * @returns {object}
 */
function buildReconciliationReport_v7(airtableRows = [], erpRows = []) {
  const {
    airtableMapByProjNr,
    unmatchedAirtable_noProj,
    totalAirtable,
  } = cleanAndSegmentAirtable_v7(airtableRows);

  const {
    erpMapByProjNr,
    unmatchedERP_byKV,
    totalERP,
  } = segmentErpData_v7(erpRows);

  const projectsToUpdate = [];
  const unmatchedERP_byProject = [];
  const unmatchedAirtable_byProject = [];

  let totalReconciled = 0;
  let totalUnreconciledERP = 0;

  erpMapByProjNr.forEach((erpProjectData, projNr) => {
    const airtableEntry = airtableMapByProjNr.get(projNr);
    const airtableAmount = roundCurrency(airtableEntry?.AirtableTotal || 0);
    const erpTotalAmount = roundCurrency(erpProjectData.total || 0);

    if (airtableEntry) {
      totalReconciled = addCurrency(totalReconciled, erpTotalAmount);

      if (Math.abs(airtableAmount - erpTotalAmount) >= 1) {
        projectsToUpdate.push({
          projNr,
          airtableTitle: airtableEntry.Projekttitel || "",
          airtableAmount,
          erpTotalAmount,
          erpKVs: erpProjectData.kvs.map((kv) => kv["KV-Nummer"]).filter(Boolean),
        });
      }
    } else {
      totalUnreconciledERP = addCurrency(totalUnreconciledERP, erpTotalAmount);
      unmatchedERP_byProject.push({
        projNr,
        erpKVs: erpProjectData.kvs.map((kv) => ({
          kv: kv["KV-Nummer"] || "",
          title: kv.title || "",
          amount: kv.amount || 0,
        })),
        erpTotalAmount,
      });
    }
  });

  airtableMapByProjNr.forEach((airtableEntry, projNr) => {
    if (!erpMapByProjNr.has(projNr)) {
      unmatchedAirtable_byProject.push({
        projNr,
        airtableTitle: airtableEntry.Projekttitel || "",
        airtableAmount: roundCurrency(airtableEntry.AirtableTotal || 0),
      });
    }
  });

  unmatchedERP_byKV.forEach((entry) => {
    totalUnreconciledERP = addCurrency(totalUnreconciledERP, entry.erpAmount || 0);
  });

  projectsToUpdate.sort((a, b) => a.projNr.localeCompare(b.projNr));
  unmatchedERP_byProject.sort((a, b) => a.projNr.localeCompare(b.projNr));
  unmatchedAirtable_byProject.sort((a, b) => a.projNr.localeCompare(b.projNr));

  return {
    totals: {
      totalERP: roundCurrency(totalERP),
      totalAirtable: roundCurrency(totalAirtable),
      totalReconciled: roundCurrency(totalReconciled),
      totalUnreconciledERP: roundCurrency(totalUnreconciledERP),
    },
    projectsToUpdate,
    unmatchedERP_byProject,
    unmatchedAirtable_byProject,
    unmatchedERP_byKV: unmatchedERP_byKV.map((entry) => ({
      kv: entry.kv || "",
      erpTitle: entry.erpTitle || "",
      erpAmount: roundCurrency(entry.erpAmount || 0),
    })),
    unmatchedAirtable_noProj: unmatchedAirtable_noProj.map((entry) => ({
      airtableTitle: entry.airtableTitle || "",
      airtableAmount: roundCurrency(entry.airtableAmount || 0),
    })),
  };
}

const exported = {
  normalizeProjectKey,
  cleanAirtableCurrency,
  cleanERPCurrency,
  cleanAndSegmentAirtable_v7,
  segmentErpData_v7,
  buildReconciliationReport_v7,
};

if (typeof module !== "undefined" && module.exports) {
  module.exports = exported;
}

if (typeof exports !== "undefined") {
  Object.assign(exports, exported);
}

