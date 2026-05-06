const SHEET_NAME = 'Inventario';

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, mensaje: 'Servidor activo' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '[]';
    const payload = JSON.parse(raw);
    const rows = Array.isArray(payload) ? payload : [payload];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];

    const out = rows.map((item) => {
      const ean = pick(item, ['ean', 'EAN', 'codigo_envio']) || '';
      const articulo = pick(item, ['articulo', 'Articulo', 'codigo_articulo', 'Codigo Articulo']) || '';
      const descripcion = pick(item, ['descripcion', 'Descripción', 'Descripcion']) || '';
      const cantidad = Number(pick(item, ['cantidad', 'Cantidad']) || 0);

      // Prioridad: usar vencimiento enviado desde la app; solo si no existe usar fecha actual.
      const fechaRaw = pick(item, [
        'fecha_vencimiento_iso',
        'fecha_vencimiento',
        'fecha_vencimiento_sheet',
        'fecha_vencimiento_datetime',
        'FechaVencimiento',
        'Fecha Vencimiento',
        'fecha vencimiento',
        'vencimiento',
        'Vencimiento',
        'fecha',
        'Fecha'
      ]);

      const fecha = parseFechaVenc(fechaRaw) || new Date();
      return [ean, articulo, descripcion, cantidad, fecha];
    });

    if (out.length) {
      const start = sh.getLastRow() + 1;
      sh.getRange(start, 1, out.length, 5).setValues(out);
      sh.getRange(start, 5, out.length, 1).setNumberFormat('dd-MM-yyyy');
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, filas: out.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err && err.message ? err.message : err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function pick(obj, keys) {
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    if (obj && obj[k] !== undefined && obj[k] !== null && String(obj[k]).trim() !== '') {
      return obj[k];
    }
  }
  return '';
}

function parseFechaVenc(value) {
  if (!value) return null;
  const s = String(value).trim();
  if (!s) return null;

  let m = s.match(/^(\d{4})[-\/]?(\d{2})[-\/]?(\d{2})/); // yyyy-mm-dd o yyyymmdd
  if (m) {
    const y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);
    const dt = new Date(y, mo - 1, d);
    if (isValidDateParts(dt, y, mo, d)) return dt;
  }

  m = s.match(/^(\d{2})[-\/.]?(\d{2})[-\/.]?(\d{4})$/); // dd-mm-yyyy, dd/mm/yyyy, dd.mm.yyyy, ddmmyyyy
  if (m) {
    const d = Number(m[1]);
    const mo = Number(m[2]);
    const y = Number(m[3]);
    const dt = new Date(y, mo - 1, d);
    if (isValidDateParts(dt, y, mo, d)) return dt;
  }

  const fallback = new Date(s);
  if (!isNaN(fallback.getTime())) return new Date(fallback.getFullYear(), fallback.getMonth(), fallback.getDate());
  return null;
}

function isValidDateParts(dt, y, m, d) {
  return dt.getFullYear() === y && (dt.getMonth() + 1) === m && dt.getDate() === d;
}
