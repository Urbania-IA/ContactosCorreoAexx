/**
 * CONFIGURACIÓN OPTIMIZADA
 */
const CONFIG = {
  SHEET_NAME_PREFIX: "Contactos",
  HOURS_BACK: 24,
  LABEL_PROCESSED: "contacto-procesado",
  MIN_SCORE: 2,
  SHEET_GENERAL_ID: "1g2vRKPrg_2Y388JHS0UrHMISmduICeLwlhzRHYpGX1o", // ← Sheet general compartido
  BATCH: 100,                  // máximo que permite Gmail
  SLEEP_MS: 0,                 // sin pausa entre lotes
  PAUSA_ENTRE_SESIONES: 30 * 1000,  // 30s entre sesiones (antes 60s)
  TIEMPO_MAX_MS: 5.2 * 60 * 1000   // ~5 min 12s por sesión
};

const RE_TEL = /(?:\+|00)?\d{2,3}[\s\-]?\d{3}[\s\-]?\d{2,3}[\s\-]?\d{2,3}|\b\d{9,12}\b/g;
const RE_WEB = /(https?:\/\/)?(www\.)?([a-zA-Z0-9-]+\.(?!gmail|outlook|hotmail|yahoo|linkedin|facebook|twitter|instagram|apple|uv|webs|wixsite)[a-zA-Z]{2,})/i;
const RE_CIERRE = /^(?:saludos|un saludo|un cordial saludo|atentamente|cordialmente|reciba un saludo|gracias|muchas gracias|best regards|kind regards|regards|cheers|thanks|sincerely)[\s,.:;!\-–]*$/i;
const RE_CITA = /(?:^On .+ wrote:$|^El .+ escribió:$|^-----Original Message-----|^De:\s.+|^From:\s.+)/m;

// ─────────────────────────────────────────────
//  FUNCIÓN PRINCIPAL: correos recientes (24h)
// ─────────────────────────────────────────────
function procesarNuevosCorreos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailsProcesados = obtenerEmailsProcesados(ss);
  const threads = GmailApp.search(
    `in:inbox -label:${CONFIG.LABEL_PROCESSED} newer_than:${CONFIG.HOURS_BACK}h`, 0, 30
  );
  _procesarThreads(ss, threads, emailsProcesados);
}

// ─────────────────────────────────────────────
//  FUNCIÓN AÑO COMPLETO
// ─────────────────────────────────────────────
function procesarCorreosDeAño() {
  const AÑO_OBJETIVO = 2025;
  _procesarPorAño(AÑO_OBJETIVO);
}

// ─────────────────────────────────────────────
//  HISTÓRICO CON CONTINUACIÓN AUTOMÁTICA
// ─────────────────────────────────────────────
function procesarHistoricoCompleto() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  if (!props.getProperty('historico_start')) {
    const confirmacion = ui.alert(
      '📬 Procesar histórico completo',
      'Esto procesará TODOS los correos en lotes automáticos.\n\n' +
      'El script se pausará y reanudará solo hasta terminar.\n\n' +
      '¿Quieres continuar?',
      ui.ButtonSet.YES_NO
    );
    if (confirmacion !== ui.Button.YES) return;
    props.setProperty('historico_start', '0');
    props.setProperty('historico_total', '0');
  }

  _continuarHistorico();
}

function _continuarHistorico() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailsProcesados = obtenerEmailsProcesados(ss);
  const label = GmailApp.getUserLabelByName(CONFIG.LABEL_PROCESSED)
              || GmailApp.createLabel(CONFIG.LABEL_PROCESSED);

  const inicio = Date.now();
  let start = parseInt(props.getProperty('historico_start') || '0');
  let totalInsertados = parseInt(props.getProperty('historico_total') || '0');
  let lotesProcesados = 0;

  // Acumulador global — escribe todo de golpe al pausar o terminar
  const acumuladoPorAño = {};

  escribirLog(ss, "▶️ Iniciando", start, totalInsertados,
    `Reanudando desde correo #${start} — ${totalInsertados} contactos acumulados`);

  while (true) {
    if (Date.now() - inicio > CONFIG.TIEMPO_MAX_MS) {
      // Volcar acumulado antes de pausar
      _volcarAcumulado(ss, acumuladoPorAño);

      props.setProperty('historico_start', String(start));
      props.setProperty('historico_total', String(totalInsertados));
      _programarContinuacion();

      const tiempoUsado = Math.round((Date.now() - inicio) / 1000);
      escribirLog(ss, "⏸️ Pausado", start, totalInsertados,
        `Tiempo usado: ${tiempoUsado}s — correos revisados: ~${start} — contactos: ${totalInsertados} — reanuda en 30s`);
      return;
    }

    const threads = GmailApp.search('in:inbox', start, CONFIG.BATCH);
    if (threads.length === 0) break;

    threads.forEach(thread => {
      const messages = thread.getMessages();
      const lastMsg = messages[messages.length - 1];
      const id = lastMsg.getId();

      if (emailsProcesados.has(id)) { thread.addLabel(label); return; }

      const metadata = analizarMensajeOptimizado(lastMsg);

      if (metadata.score >= CONFIG.MIN_SCORE) {
        const año = new Date(metadata.fecha).getFullYear();
        if (!acumuladoPorAño[año]) acumuladoPorAño[año] = [];
        acumuladoPorAño[año].push([
          metadata.fecha, metadata.nombre, metadata.email, metadata.telefono,
          metadata.web, metadata.firma, metadata.asunto, id
        ]);
        emailsProcesados.add(id);
        totalInsertados++;
      }
      thread.addLabel(label);
    });

    lotesProcesados++;

    // Log cada 10 lotes (~1000 correos)
    if (lotesProcesados % 10 === 0) {
      escribirLog(ss, "🔄 Lote OK", start, totalInsertados,
        `Correos revisados: ~${start} — contactos encontrados: ${totalInsertados}`);
    }

    if (threads.length < CONFIG.BATCH) break;
    start += CONFIG.BATCH;
  }

  // ✅ Fin — volcar todo lo acumulado
  _volcarAcumulado(ss, acumuladoPorAño);

  props.deleteProperty('historico_start');
  props.deleteProperty('historico_total');
  _eliminarTriggerContinuacion();

  const tiempoTotal = Math.round((Date.now() - inicio) / 1000);
  escribirLog(ss, "✅ Completado", start, totalInsertados,
    `🎉 Proceso terminado en ${tiempoTotal}s — Total contactos: ${totalInsertados}`);
}

// ─────────────────────────────────────────────
//  VOLCADO MASIVO — escribe todo de una vez
// ─────────────────────────────────────────────
function _volcarAcumulado(ss, acumuladoPorAño) {
  Object.entries(acumuladoPorAño).forEach(([año, registros]) => {
    if (registros.length === 0) return;

    const sheetPropio = obtenerOCrearHojaPorAño(ss, año);
    sheetPropio.getRange(
      sheetPropio.getLastRow() + 1, 1,
      registros.length, registros[0].length
    ).setValues(registros);

    escribirEnSheetGeneral(registros, parseInt(año));

    // Limpiar para no escribir doble en la siguiente pausa
    acumuladoPorAño[año] = [];
  });
}

// ─────────────────────────────────────────────
//  TRIGGER — cada 30s en vez de 60s
// ─────────────────────────────────────────────
function _programarContinuacion() {
  _eliminarTriggerContinuacion();
  ScriptApp.newTrigger('_continuarHistorico')
    .timeBased()
    .after(CONFIG.PAUSA_ENTRE_SESIONES)
    .create();
}

function _eliminarTriggerContinuacion() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === '_continuarHistorico')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function cancelarHistorico() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const start = PropertiesService.getScriptProperties().getProperty('historico_start') || '?';
  const total = PropertiesService.getScriptProperties().getProperty('historico_total') || '0';
  PropertiesService.getScriptProperties().deleteAllProperties();
  _eliminarTriggerContinuacion();
  escribirLog(ss, "🛑 Cancelado", parseInt(start), parseInt(total),
    `Proceso cancelado manualmente en correo #${start}`);
  SpreadsheetApp.getUi().alert('🛑 Proceso cancelado. Revisa la pestaña 📊 Progreso.');
}

// ─────────────────────────────────────────────
//  HOJA DE LOG / PROGRESO
// ─────────────────────────────────────────────

function obtenerOCrearHojaLog(ss) {
  const nombre = "📊 Progreso";
  let s = ss.getSheetByName(nombre);
  if (!s) {
    s = ss.insertSheet(nombre, 0);
    s.appendRow(["Timestamp", "Estado", "Lote (start)", "Contactos insertados", "Detalle"]);
    s.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#4a86e8").setFontColor("white");
    s.setColumnWidth(1, 160); // Timestamp
    s.setColumnWidth(2, 100); // Estado
    s.setColumnWidth(3, 130); // Lote (start)
    s.setColumnWidth(4, 180); // Contactos insertados
    s.setColumnWidth(5, 350); // Detalle
  }
  return s;
}

function escribirLog(ss, estado, start, totalInsertados, detalle = "") {
  const sheet = obtenerOCrearHojaLog(ss);
  const colores = {
    "▶️ Iniciando": "#e8f5e9",
    "⏸️ Pausado":   "#fff8e1",
    "🔄 Lote OK":   "#f1f8e9",
    "✅ Completado": "#e8f5e9",
    "🛑 Cancelado": "#fce8e6",
    "❌ Error":      "#fce8e6"
  };
  const fila = sheet.getLastRow() + 1;
  sheet.appendRow([new Date(), estado, start, totalInsertados, detalle]);
  sheet.getRange(fila, 1, 1, 5).setBackground(colores[estado] || "#ffffff");
  sheet.getRange(fila, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  SpreadsheetApp.getActiveSpreadsheet().toast(detalle || estado, `📬 Histórico — ${estado}`, 8);
}

// ─────────────────────────────────────────────
//  FUNCIÓN AÑO (interna)
// ─────────────────────────────────────────────
function _procesarPorAño(año) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailsProcesados = obtenerEmailsProcesados(ss);
  const label = GmailApp.getUserLabelByName(CONFIG.LABEL_PROCESSED)
              || GmailApp.createLabel(CONFIG.LABEL_PROCESSED);

  const query = `in:inbox after:${año}/01/01 before:${año + 1}/01/01`;
  const BATCH = 100;
  let start = 0;
  let totalInsertados = 0;

  while (true) {
    const threads = GmailApp.search(query, start, BATCH);
    if (threads.length === 0) break;

    const registrosLote = [];

    threads.forEach(thread => {
      const messages = thread.getMessages();
      const lastMsg = messages[messages.length - 1];
      const id = lastMsg.getId();

      if (emailsProcesados.has(id)) return;

      const metadata = analizarMensajeOptimizado(lastMsg);

      const añoMensaje = new Date(metadata.fecha).getFullYear();
      if (añoMensaje !== año) return;

      if (metadata.score >= CONFIG.MIN_SCORE) {
        registrosLote.push([
          metadata.fecha, metadata.nombre, metadata.email, metadata.telefono,
          metadata.web, metadata.firma, metadata.asunto, id
        ]);
        emailsProcesados.add(id);
      }
      thread.addLabel(label);
    });

    if (registrosLote.length > 0) {
      const sheetPropio = obtenerOCrearHojaPorAño(ss, año);
      sheetPropio.getRange(sheetPropio.getLastRow() + 1, 1, registrosLote.length, registrosLote[0].length)
                 .setValues(registrosLote);
      escribirEnSheetGeneral(registrosLote, año);
      totalInsertados += registrosLote.length;
    }

    if (threads.length < BATCH) break;
    start += BATCH;
    Utilities.sleep(500);
  }

  SpreadsheetApp.getUi().alert(
    `✅ Proceso completado\n\nAño: ${año}\nContactos insertados: ${totalInsertados}`
  );
}

// ─────────────────────────────────────────────
//  ESCRITURA EN SHEET GENERAL COMPARTIDO
// ─────────────────────────────────────────────
function escribirEnSheetGeneral(registros, año) {
  try {
    const ssGeneral = SpreadsheetApp.openById(CONFIG.SHEET_GENERAL_ID);
    const sheet = obtenerOCrearHojaGeneral(ssGeneral, año);
    const usuarioActual = Session.getActiveUser().getEmail();
    const registrosConUsuario = registros.map(fila => [usuarioActual, ...fila]);
    sheet.getRange(sheet.getLastRow() + 1, 1, registrosConUsuario.length, registrosConUsuario[0].length)
         .setValues(registrosConUsuario);
  } catch (e) {
    Logger.log(`Error escribiendo en sheet general: ${e.message}`);
  }
}

function obtenerOCrearHojaGeneral(ss, año) {
  const nombre = `Contactos-${año}`;
  let s = ss.getSheetByName(nombre);
  if (!s) {
    s = ss.insertSheet(nombre);
    s.appendRow(["Usuario", "Fecha", "Nombre", "Email", "Teléfono", "Web", "Firma", "Asunto", "ID"]);
    s.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#f3f3f3");
    s.setColumnWidth(7, 350);
    s.getRange("G:G").setWrap(true);
  }
  return s;
}

// ─────────────────────────────────────────────
//  ANÁLISIS DE MENSAJE
// ─────────────────────────────────────────────
function analizarMensajeOptimizado(msg) {
  const body = msg.getPlainBody();
  const htmlBody = msg.getBody();
  const from = msg.getFrom();

  const datos = {
    fecha: msg.getDate(),
    nombre: extraerNombreDeFrom(from),
    email: extraerEmail(from),
    telefono: "",
    web: "",
    firma: "",
    asunto: msg.getSubject(),
    score: 0
  };

  datos.firma = extraerFirma(htmlBody, body);

  const textoBusqueda = datos.firma || body.split('\n').slice(-15).join(' ');

  const tels = textoBusqueda.match(RE_TEL);
  if (tels) { datos.telefono = tels[0].trim(); datos.score += 2; }

  const webMatch = textoBusqueda.match(RE_WEB);
  if (webMatch) { datos.web = webMatch[3]; datos.score += 1; }

  if (datos.firma && datos.firma.length > 20 && datos.firma.length < 1500) datos.score += 1;
  if (body.length > 100 && body.length < 5000) datos.score += 1;

  return datos;
}

// ─────────────────────────────────────────────
//  EXTRACCIÓN DE FIRMA
// ─────────────────────────────────────────────
function extraerFirma(htmlBody, plainBody) {
  const htmlLimpio = eliminarCitasHtml(htmlBody);

  const patronesHtml = [
    /<div[^>]*class="[^"]*gmail_signature[^"]*"[^>]*>([\s\S]*)$/i,
    /<div[^>]*data-smartmail="gmail_signature"[^>]*>([\s\S]*)$/i,
    /<div[^>]*(?:id|class)="[^"]*(?:signature|firma|Signature)[^"]*"[^>]*>([\s\S]*)$/i
  ];

  for (const patron of patronesHtml) {
    const match = htmlLimpio.match(patron);
    if (match) {
      const firma = limpiarHtml(match[1]).substring(0, 1000);
      if (firma.length > 5) return firma;
    }
  }

  const textoSinCitas = plainBody.split(RE_CITA)[0];
  const lineas = textoSinCitas.split('\n').map(l => l.trim());

  for (let i = 0; i < lineas.length; i++) {
    if (RE_CIERRE.test(lineas[i])) {
      return lineas.slice(i).join('\n').replace(/\n{3,}/g, '\n\n').trim().substring(0, 1000);
    }
  }

  return lineas.filter(l => l.length > 0).slice(-8).join('\n').substring(0, 1000);
}

function eliminarCitasHtml(html) {
  return html
    .replace(/<div[^>]*class="[^"]*gmail_quote[^"]*"[\s\S]*$/i, '')
    .replace(/<blockquote[\s\S]*$/i, '');
}

function limpiarHtml(html) {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<\/td>/gi, ' ')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/[ \t]+/g, ' ')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// ─────────────────────────────────────────────
//  UTILIDADES
// ─────────────────────────────────────────────
function extraerEmail(s) { return (s.match(/<([^>]+)>/) || [null, s])[1].toLowerCase(); }
function extraerNombreDeFrom(s) { return (s.match(/^"?(.*?)"?\s*</) || [null, ""])[1]; }

function obtenerOCrearHojaPorAño(ss, año) {
  const nombre = `${CONFIG.SHEET_NAME_PREFIX}-${año}`;
  let s = ss.getSheetByName(nombre);
  if (!s) {
    s = ss.insertSheet(nombre);
    s.appendRow(["Fecha", "Nombre", "Email", "Teléfono", "Web", "Firma", "Asunto", "ID"]);
    s.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#f3f3f3");
    s.setColumnWidth(6, 350);
    s.getRange("F:F").setWrap(true);
  }
  return s;
}

function obtenerEmailsProcesados(ss) {
  const ids = new Set();
  const prefijo = CONFIG.SHEET_NAME_PREFIX + "-";
  ss.getSheets().forEach(sheet => {
    if (!sheet.getName().startsWith(prefijo)) return;
    const last = sheet.getLastRow();
    if (last < 2) return;
    sheet.getRange(2, 8, last - 1, 1).getValues()
      .flat().map(String).forEach(id => ids.add(id));
  });
  return ids;
}

// ─────────────────────────────────────────────
//  PROCESADO INTERNO COMPARTIDO (correos 24h)
// ─────────────────────────────────────────────
function _procesarThreads(ss, threads, emailsProcesados) {
  if (threads.length === 0) return;

  const label = GmailApp.getUserLabelByName(CONFIG.LABEL_PROCESSED)
              || GmailApp.createLabel(CONFIG.LABEL_PROCESSED);
  const porAño = {};

  threads.forEach(thread => {
    const messages = thread.getMessages();
    const lastMsg = messages[messages.length - 1];
    const id = lastMsg.getId();

    if (emailsProcesados.has(id)) return;

    const metadata = analizarMensajeOptimizado(lastMsg);

    if (metadata.score >= CONFIG.MIN_SCORE) {
      const año = new Date(metadata.fecha).getFullYear();
      if (!porAño[año]) porAño[año] = [];
      porAño[año].push([
        metadata.fecha, metadata.nombre, metadata.email, metadata.telefono,
        metadata.web, metadata.firma, metadata.asunto, id
      ]);
    }
    thread.addLabel(label);
  });

  Object.entries(porAño).forEach(([año, registros]) => {
    const sheetPropio = obtenerOCrearHojaPorAño(ss, año);
    sheetPropio.getRange(sheetPropio.getLastRow() + 1, 1, registros.length, registros[0].length)
               .setValues(registros);
    escribirEnSheetGeneral(registros, parseInt(año));
  });
}

// ─────────────────────────────────────────────
//  MENÚ
// ─────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📬 Procesador de Contactos')
    .addItem('▶️ Procesar últimas 24h', 'procesarNuevosCorreos')
    .addItem('📅 Procesar año completo (2025)', 'procesarCorreosDeAño')
    .addItem('🗂️ Procesar histórico completo', 'procesarHistoricoCompleto')
    .addSeparator()
    .addItem('🛑 Cancelar histórico en curso', 'cancelarHistorico')
    .addToUi();
}