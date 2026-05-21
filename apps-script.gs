// ============================================================
// BITÁCORA DE TRABAJO — Google Apps Script
// Pega este código en Extensions > Apps Script de tu Google Sheet
// Luego: Deploy > New Deployment > Web App > Anyone > Deploy
// ============================================================

const SHEET_NAME = 'Bitacora';

function doGet(e) {
  const action = e.parameter.action;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  // ── GUARDAR ENTRADA ──────────────────────────────────────
  if (action === 'save') {
    const row = [
      e.parameter.date,
      e.parameter.arrival,
      e.parameter.departure,
      e.parameter.activities,
      parseFloat(e.parameter.duration)
    ];
    sheet.appendRow(row);
    return respond({ success: true, message: 'Entrada guardada' });
  }

  // ── OBTENER REPORTE ──────────────────────────────────────
  if (action === 'report') {
    const type = e.parameter.type || 'daily';
    const allData = sheet.getDataRange().getValues();

    if (allData.length <= 1) {
      return respond({ success: true, data: [], summary: emptySummary() });
    }

    const rows = allData.slice(1).filter(r => r[0] !== ''); // skip empty rows

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let filtered = [];

    if (type === 'daily') {
      const todayStr = formatDateStr(today);
      filtered = rows.filter(r => String(r[0]).trim() === todayStr);

    } else if (type === 'weekly') {
      const weekStart = new Date(today);
      weekStart.setDate(today.getDate() - today.getDay() + 1); // Lunes
      filtered = rows.filter(r => {
        const d = parseDate(r[0]);
        return d && d >= weekStart && d <= today;
      });

    } else if (type === 'monthly') {
      filtered = rows.filter(r => {
        const d = parseDate(r[0]);
        return d && d.getMonth() === today.getMonth() && d.getFullYear() === today.getFullYear();
      });
    }

    const entries = filtered.map(r => ({
      date: String(r[0]),
      arrival: String(r[1]),
      departure: String(r[2]),
      activities: String(r[3]),
      duration: parseFloat(r[4]) || 0
    }));

    const totalHours = entries.reduce((sum, e) => sum + e.duration, 0);
    const avgHours = entries.length > 0 ? totalHours / entries.length : 0;

    return respond({
      success: true,
      data: entries,
      summary: {
        total_entries: entries.length,
        total_hours: Math.round(totalHours * 100) / 100,
        avg_hours: Math.round(avgHours * 100) / 100
      }
    });
  }

  // ── PING (verificar conexión) ────────────────────────────
  if (action === 'ping') {
    return respond({ success: true, message: 'Conexión OK ✅' });
  }

  return respond({ success: false, error: 'Acción no reconocida' });
}

// ── HELPERS ─────────────────────────────────────────────────

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDateStr(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function parseDate(val) {
  if (!val) return null;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function emptySummary() {
  return { total_entries: 0, total_hours: 0, avg_hours: 0 };
}
