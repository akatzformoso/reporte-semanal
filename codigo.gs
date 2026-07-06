/**
 * CÓDIGO.GS — Receptor del formulario de Reporte Semanal (Apps Script)
 *
 * Guarda cada envío del formulario como una fila en la hoja "Reportes".
 * Versión ampliada (jul 2026): guarda TODOS los campos del formulario.
 * Las 9 columnas originales quedan en su lugar; las 11 nuevas se agregan
 * al final, así las filas viejas no se desalinean.
 *
 * DESPUÉS DE PEGAR ESTE CÓDIGO hay que publicar una versión nueva:
 * Implementar → Administrar implementaciones → ✏️ → Versión: "Nueva versión"
 * → Implementar. La URL /exec no cambia (no hay que tocar el formulario).
 */

function doGet(e) {
  const BASE = ["Fecha", "Supervisor", "Servicio", "Novedades", "Pendientes",
                "Ausencias", "PAE", "Urgencias", "Cobros"];
  const EXTRA = ["Semana", "PAEServicio", "PAENovedades", "FaltanteMP",
                 "FaltanteMPDetalle", "CambioMenu", "CambioMenuDetalle",
                 "EnviosCoordinar", "EnviosPendientes", "Checklist", "CierreMes",
                 "Cronograma"];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Reportes");
  if (!sheet) {
    sheet = ss.insertSheet("Reportes");
    sheet.appendRow(BASE.concat(EXTRA));
  } else if (sheet.getLastColumn() < BASE.length + EXTRA.length) {
    // Completa los encabezados nuevos una sola vez
    sheet.getRange(1, BASE.length + 1, 1, EXTRA.length).setValues([EXTRA]);
  }

  const p = e.parameter;
  sheet.appendRow([
    new Date(),
    p.supervisor || "",
    p.servicio || "",
    p.novedades || "",
    p.pendientes || "",
    p.ausencias || "",
    p.hayPAE || "",
    p.urgencias || "",
    p.cobros || "",
    // --- columnas nuevas ---
    p.fecha || "",            // "Semana que cierra" elegida en el formulario
    p.paeServicio || "",
    p.paeNovedades || "",
    p.hayMP || "",
    p.mpObs || "",
    p.hayMenu || "",
    p.menuDetalle || "",
    p.enviosCoordinar || "",
    p.enviosPendientes || "",
    p.checklist || "",
    p.finMes || "",
    p.cronograma || ""
  ]);
  return ContentService.createTextOutput("OK");
}
