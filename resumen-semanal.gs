/**
 * RESUMEN SEMANAL POR EMAIL — Reporte Semanal de Supervisores
 *
 * Corre todos los lunes a las 8:00 (disparador semanal) y envía a cada
 * responsable un email con el resumen de SUS servicios: quién reportó y
 * quién no, urgencias, avisos de cobro, ausencias, faltantes de MP,
 * cambios de menú, envíos y checklist incompleto.
 *
 * Lee la hoja "Reportes" con las columnas que escribe codigo.gs:
 * Fecha, Supervisor, Servicio, Novedades, Pendientes, Ausencias, PAE,
 * Urgencias, Cobros, Semana, PAEServicio, PAENovedades, FaltanteMP,
 * FaltanteMPDetalle, CambioMenu, CambioMenuDetalle, EnviosCoordinar,
 * EnviosPendientes, Checklist, CierreMes.
 *
 * INSTALACIÓN (una sola vez):
 *   1. Ejecutar "probarResumen" → autorizar permisos → llegan los emails de prueba.
 *   2. Ejecutar "crearDisparador" → queda programado todos los lunes 8:00.
 *   (Verificar en Configuración del proyecto que la zona horaria sea America/Montevideo.)
 */

const CONFIG = {
  responsables: [
    {
      nombre: 'Anto',
      email: 'akatzformoso@gmail.com',
      servicios: ['Ansina', 'Liceo 6 TBO', 'Liceo 4 TBO', 'Rivera', 'Vichadero',
                  'Tranqueras', 'San Gregorio', 'La Paloma', 'Paso de los Toros'],
    },
    {
      nombre: 'Ximena',
      email: 'ximenaboragno@gmail.com',
      servicios: ['Liceo 4 Melo', 'Liceo 5 Melo', 'UTU Melo', 'Liceo Treinta y Tres',
                  'Liceo Santa Clara', 'Liceo Cerro Chato', 'Liceo Pintadito',
                  'Liceo Young 2', 'Liceo San Javier'],
    },
  ],
  // Se incluyen los reportes enviados dentro de esta ventana de días
  // (9 días cubre de viernes a lunes con margen para rezagados).
  diasVentana: 9,
  nombreHoja: 'Reportes',
};

// ============================================================

function enviarResumenSemanal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.nombreHoja) || ss.getSheets()[0];
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return;

  // Mapear columnas por nombre de encabezado (tolerante a mayúsculas/espacios)
  const headers = datos[0].map(h => normalizar(h));
  const col = nombre => headers.indexOf(normalizar(nombre));
  if (col('servicio') === -1) {
    MailApp.sendEmail(CONFIG.responsables[0].email, '⚠️ Resumen semanal: error de configuración',
      'No se encontró la columna "Servicio" en la hoja. Encabezados: ' + datos[0].join(' | '));
    return;
  }
  const val = (fila, nombre) => {
    const c = col(nombre);
    if (c === -1) return '';
    const v = fila[c];
    if (v instanceof Date && !isNaN(v)) {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
    return String(v || '').trim();
  };

  // Filtrar filas de la última semana (por fecha de envío = columna Fecha)
  const ahora = new Date();
  const limite = new Date(ahora.getTime() - CONFIG.diasVentana * 24 * 60 * 60 * 1000);
  const filas = [];
  for (let i = 1; i < datos.length; i++) {
    const f = parsearFecha(datos[i][col('fecha')]);
    // Si la fecha no se puede leer, se incluye igual (mejor de más que de menos)
    if (!f || f >= limite) filas.push(datos[i]);
  }

  CONFIG.responsables.forEach(resp => {
    if (resp.email.indexOf('@') === -1) return; // email sin completar
    const propias = filas.filter(fila => resp.servicios.indexOf(val(fila, 'servicio')) !== -1);
    const html = armarHtml(resp, propias, val);
    MailApp.sendEmail({
      to: resp.email,
      subject: '📋 Resumen semanal de reportes — ' + Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      htmlBody: html,
    });
  });
}

function armarHtml(resp, filas, val) {
  const reportaron = {};
  filas.forEach(f => { reportaron[val(f, 'servicio')] = f; });
  const sinReporte = resp.servicios.filter(s => !reportaron[s]);

  let h = '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:640px;color:#1f2937">';
  h += '<h2 style="color:#1e3a5f">📋 Resumen semanal — servicios de ' + resp.nombre + '</h2>';

  // Quién reportó / quién no
  h += seccion('✅ Reportaron (' + Object.keys(reportaron).length + ' de ' + resp.servicios.length + ')');
  h += '<ul>' + Object.keys(reportaron).map(s => {
    const f = reportaron[s];
    const semana = val(f, 'semana') || val(f, 'fecha');
    return '<li><b>' + s + '</b> — ' + val(f, 'supervisor') + ' (semana ' + semana + ')</li>';
  }).join('') + '</ul>';
  if (sinReporte.length) {
    h += seccion('❌ Sin reporte');
    h += '<ul style="color:#dc2626">' + sinReporte.map(s => '<li><b>' + s + '</b></li>').join('') + '</ul>';
  }

  // Secciones de contenido (solo aparecen si hay algo para mostrar)
  h += bloque('🚨 Urgencias', filas, val, 'urgencias', '#dc2626');
  h += bloque('💰 Avisos de cobro / inconsistencias salariales', filas, val, 'cobros', '#d97706');
  h += bloque('👥 Ausencias', filas, val, 'ausencias');
  h += bloqueCondicional('⚠️ Faltante de materia prima (OC vs factura)', filas, val, 'faltantemp', 'faltantempdetalle');
  h += bloqueCondicional('🍽 Cambios de menú', filas, val, 'cambiomenu', 'cambiomenudetalle');
  h += bloque('📦 Envíos a coordinar', filas, val, 'envioscoordinar');
  h += bloque('📬 Envíos pendientes de entrega', filas, val, 'enviospendientes');

  // Visita PAE (Sí/No + servicio visitado + novedades)
  const conPAE = filas.filter(f => val(f, 'pae') === 'Sí');
  if (conPAE.length) {
    h += seccion('🏥 Visita PAE Media');
    conPAE.forEach(f => {
      const lugar = val(f, 'paeservicio');
      const det = val(f, 'paenovedades');
      h += '<p><b>' + val(f, 'servicio') + ':</b> ' + (lugar ? 'visita en ' + escaparHtml(lugar) + '. ' : '')
        + escaparHtml(det || '(sin detalle)').replace(/\n/g, '<br>') + '</p>';
    });
  }

  // Checklist incompleto
  const conPendientes = filas.filter(f => (val(f, 'checklist').match(/⬜/g) || []).length > 0);
  if (conPendientes.length) {
    h += seccion('☑️ Checklist con ítems sin completar');
    conPendientes.forEach(f => {
      const faltan = val(f, 'checklist').split('\n').filter(l => l.indexOf('⬜') !== -1)
        .map(l => l.replace('⬜', '').trim());
      h += '<p><b>' + val(f, 'servicio') + ':</b> ' + escaparHtml(faltan.join(' · ')) + '</p>';
    });
  }
  h += bloque('📅 Cierre de mes', filas, val, 'cierremes');

  // Novedades y pendientes
  h += bloque('📣 Novedades', filas, val, 'novedades');
  h += bloque('🗂 Pendientes para la próxima semana', filas, val, 'pendientes');

  // Cronograma: es UNO por supervisor (lo carga en cualquiera de sus reportes),
  // así que se agrupa por supervisor y se toma el primero con contenido.
  const cronPorSup = {};
  filas.forEach(f => {
    const sup = val(f, 'supervisor') || '(sin nombre)';
    if (!cronPorSup[sup] && tieneContenido(val(f, 'cronograma'))) cronPorSup[sup] = val(f, 'cronograma');
  });
  if (Object.keys(cronPorSup).length) {
    h += seccion('🗓 Cronograma próxima semana (por supervisor)');
    Object.keys(cronPorSup).forEach(sup => {
      h += '<p><b>' + escaparHtml(sup) + ':</b><br>' + escaparHtml(cronPorSup[sup]).replace(/\n/g, '<br>') + '</p>';
    });
  }

  h += '<hr style="border:none;border-top:1px solid #e5e7eb;margin-top:24px">';
  h += '<p style="font-size:12px;color:#6b7280">Generado automáticamente desde la planilla de reportes semanales.</p></div>';
  return h;
}

// --- helpers de armado ---

function seccion(titulo) {
  return '<h3 style="color:#1e3a5f;border-left:4px solid #2e6da4;padding-left:10px;margin-top:22px">' + titulo + '</h3>';
}

function tieneContenido(texto) {
  const t = String(texto || '').trim();
  if (!t) return false;
  const vacios = ['sin ', '(sin', '(no indicado)', '—', '-'];
  const tl = t.toLowerCase();
  return !vacios.some(v => tl.indexOf(v) === 0);
}

function bloque(titulo, filas, val, campo, color) {
  const con = filas.filter(f => tieneContenido(val(f, campo)));
  if (!con.length) return '';
  let h = seccion(titulo);
  con.forEach(f => {
    h += '<p' + (color ? ' style="color:' + color + '"' : '') + '><b>' + val(f, 'servicio') + ':</b><br>'
      + escaparHtml(val(f, campo)).replace(/\n/g, '<br>') + '</p>';
  });
  return h;
}

function bloqueCondicional(titulo, filas, val, campoSiNo, campoDetalle) {
  const con = filas.filter(f => val(f, campoSiNo) === 'Sí');
  if (!con.length) return '';
  let h = seccion(titulo);
  con.forEach(f => {
    h += '<p><b>' + val(f, 'servicio') + ':</b> '
      + escaparHtml(val(f, campoDetalle) || '(sin detalle)').replace(/\n/g, '<br>') + '</p>';
  });
  return h;
}

function normalizar(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, '');
}

function escaparHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function parsearFecha(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  const m = String(v || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), 12);
  return null;
}

// ============================================================
// EJECUTAR UNA VEZ CADA UNA:

/** Envía los resúmenes AHORA con los datos de la última semana (para probar). */
function probarResumen() {
  enviarResumenSemanal();
}

/** Programa el envío automático todos los lunes entre las 8:00 y las 9:00. */
function crearDisparador() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarResumenSemanal') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarResumenSemanal')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
}
