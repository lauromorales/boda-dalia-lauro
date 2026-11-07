/**
 * Recibe las confirmaciones de la invitación y las guarda en la pestaña "Respuestas".
 * Se pega en la hoja "Confirmaciones · Boda Dalia y Lauro" (Extensiones > Apps Script)
 * y se implementa como aplicación web con acceso "Cualquier usuario".
 * Solo puede agregar filas: no lee ni muestra nada de la hoja.
 */
function doPost(e) {
  var p = (e && e.parameter) || {};

  // Campo trampa: si un robot lo llena, se ignora la respuesta.
  if (p.sitio) return respuesta_({ ok: true });

  var nombre = limpio_(p.nombre, 120);
  if (!nombre) return respuesta_({ ok: false, error: 'Falta el nombre' });

  var asiste = p.asiste === 'No' ? 'No' : 'Sí';
  var personas = asiste === 'Sí' ? Math.max(1, Math.min(20, parseInt(p.personas, 10) || 1)) : 0;

  var fila = [
    new Date(),
    nombre,
    asiste,
    personas,
    p.misa === 'Sí' ? 'Sí' : 'No',
    p.fiesta === 'Sí' ? 'Sí' : 'No',
    limpio_(p.telefono, 30),
    limpio_(p.alergias, 300),
    limpio_(p.mensaje, 1000)
  ];

  var candado = LockService.getScriptLock();
  candado.waitLock(10000);
  try {
    SpreadsheetApp.getActive().getSheetByName('Respuestas').appendRow(fila);
  } finally {
    candado.releaseLock();
  }
  return respuesta_({ ok: true });
}

// Recorta el texto y evita que se interprete como fórmula.
function limpio_(valor, max) {
  var t = String(valor || '').trim().slice(0, max);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function respuesta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
