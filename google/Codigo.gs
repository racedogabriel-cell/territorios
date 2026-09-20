/**
 * API de Territorios sobre Hojas de cálculo de Google.
 *
 * Uso:
 * 1. En la hoja: Extensiones → Apps Script → pegar este archivo.
 * 2. Ejecutar inicializar() una vez (autorizar acceso a la hoja).
 * 3. Implementar → Nueva implementación → Aplicación web
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquiera
 * 4. Copiar la URL en google/google.js → APPS_SCRIPT_URL
 */

var SHEET_DATOS = 'datos_semana';
var SHEET_SALIDAS = 'salidas';
var SHEET_MAPA = 'mapa_territorios';
var SHEET_ACCIONES = 'acciones';
var CANAL_PUBLICA = 'publica';
var ORIGEN_WEB = 'predicacion-github';
var CACHE_BUNDLE = 'bundle_publica';

function bustCache_() {
  try { CacheService.getScriptCache().remove(CACHE_BUNDLE); }
  catch (err) {}
}

function leerCacheBundle_() {
  try {
    var raw = CacheService.getScriptCache().get(CACHE_BUNDLE);
    if (raw) return JSON.parse(raw);
  } catch (err) {}
  return null;
}

function guardarCacheBundle_(bundle) {
  try {
    var raw = JSON.stringify(bundle);
    if (raw && raw.length < 90000) {
      CacheService.getScriptCache().put(CACHE_BUNDLE, raw, 180);
    }
  } catch (err) {}
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok_(data) {
  return jsonOutput_({ ok: true, data: data });
}

function fail_(msg) {
  return jsonOutput_({ ok: false, error: String(msg || 'Error') });
}

function parseBody_(e) {
  if (!e) return {};
  if (e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); }
    catch (err) { return {}; }
  }
  return e.parameter || {};
}

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function ensureSheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers && headers.length) {
    var existing = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    var empty = !existing.some(function (c) { return String(c || '').trim(); });
    if (empty) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function inicializar() {
  ensureSheet_(SHEET_DATOS, [
    'canal', 'exportado', 'version', 'responsables_json',
    'territorios_json', 'descripciones_json', 'mapa_extra_json'
  ]);
  ensureSheet_(SHEET_SALIDAS, [
    'NumeroTerritorio', 'PuntosDeEncuentros', 'Responsables', 'Grupos',
    'Enviado', 'Prioridad', 'ManzanasPendientes', 'Direccion', 'Descripcion'
  ]);
  ensureSheet_(SHEET_MAPA, ['numero', 'json']);
  ensureSheet_(SHEET_ACCIONES, [
    'id', 'tipo', 'payload', 'origen', 'procesado', 'creado_en', 'procesado_en'
  ]);
  var datos = ss_().getSheetByName(SHEET_DATOS);
  if (datos.getLastRow() < 2) {
    datos.appendRow([CANAL_PUBLICA, new Date().toISOString(), 2, '[]', '[]', '{}', '{}']);
  }
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) {
    inicializar();
    sh = ss_().getSheetByName(name);
  }
  return sh;
}

function parseJson_(val, fallback) {
  if (val == null || val === '') return fallback;
  if (typeof val === 'object') return val;
  try { return JSON.parse(String(val)); }
  catch (err) { return fallback; }
}

function stringify_(val) {
  return JSON.stringify(val == null ? null : val);
}

function leerTabla_(sh) {
  var lastRow = sh.getLastRow();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  if (lastRow < 1) return { header: [], rows: [] };
  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  return { header: values[0] || [], rows: values.slice(1) };
}

function idxHeader_(header) {
  var map = {};
  for (var i = 0; i < header.length; i++) {
    map[String(header[i] || '').trim().toLowerCase()] = i;
  }
  return map;
}

function cell_(row, idx, names, def) {
  for (var i = 0; i < names.length; i++) {
    var k = names[i];
    if (idx[k] == null) continue;
    var v = row[idx[k]];
    if (v != null && String(v).trim() !== '') return v;
  }
  return def;
}

function cargarBundle_() {
  var cached = leerCacheBundle_();
  if (cached) return cached;
  var bundle = cargarBundleDesdeHojas_();
  guardarCacheBundle_(bundle);
  return bundle;
}

function cargarBundleDesdeHojas_() {
  var datosSh = sheet_(SHEET_DATOS);
  var tabla = leerTabla_(datosSh);
  var idx = idxHeader_(tabla.header);
  var chosen = null;
  for (var i = 0; i < tabla.rows.length; i++) {
    var row = tabla.rows[i];
    var canal = String(cell_(row, idx, ['canal'], '') || '').trim().toLowerCase();
    if (canal === CANAL_PUBLICA) {
      chosen = row;
      break;
    }
  }
  if (!chosen && tabla.rows.length) chosen = tabla.rows[0];
  if (!chosen) {
    throw new Error('GOOGLE_SIN_DATOS: La hoja datos_semana está vacía. Publicá desde mapa.html (Enviar a Google).');
  }

  var blob = parseJson_(cell_(chosen, idx, ['datos', 'data', 'bundle'], ''), null);
  if (blob && (blob.salidas || blob.mapaSemanal)) return blob;

  var exportado = cell_(chosen, idx, ['exportado'], new Date().toISOString());
  if (exportado instanceof Date) exportado = exportado.toISOString();
  var version = Number(cell_(chosen, idx, ['version'], 2)) || 2;
  var responsables = parseJson_(cell_(chosen, idx, ['responsables', 'responsables_json'], '[]'), []);
  var territoriosCat = parseJson_(cell_(chosen, idx, ['territorios', 'territorios_json'], '[]'), []);
  var descripciones = parseJson_(cell_(chosen, idx, ['descripciones', 'descripciones_json'], '{}'), {});
  var mapaExtra = parseJson_(cell_(chosen, idx, ['mapa_extra', 'mapa_extra_json'], '{}'), {});

  var salidasSh = sheet_(SHEET_SALIDAS);
  var salTabla = leerTabla_(salidasSh);
  var sIdx = idxHeader_(salTabla.header);
  var salidas = [];
  for (var s = 0; s < salTabla.rows.length; s++) {
    var r = salTabla.rows[s];
    if (!r.some(function (c) { return String(c || '').trim(); })) continue;
    salidas.push({
      NumeroTerritorio: cell_(r, sIdx, ['numeroterritorio', 'numero'], ''),
      PuntosDeEncuentros: cell_(r, sIdx, ['puntosdeencuentros', 'punto'], ''),
      Responsables: cell_(r, sIdx, ['responsables', 'responsable'], ''),
      Grupos: cell_(r, sIdx, ['grupos', 'grupo'], ''),
      Enviado: r[sIdx.enviado] instanceof Date ? r[sIdx.enviado].toISOString() : cell_(r, sIdx, ['enviado'], ''),
      Prioridad: cell_(r, sIdx, ['prioridad'], ''),
      ManzanasPendientes: cell_(r, sIdx, ['manzanaspendientes', 'manzanas'], ''),
      Direccion: cell_(r, sIdx, ['direccion'], ''),
      Descripcion: cell_(r, sIdx, ['descripcion'], '')
    });
  }

  var mapaSh = sheet_(SHEET_MAPA);
  var mTabla = leerTabla_(mapaSh);
  var mapaTerritorios = {};
  for (var m = 0; m < mTabla.rows.length; m++) {
    var mr = mTabla.rows[m];
    var num = String(mr[0] || '').trim();
    if (!num) continue;
    mapaTerritorios[num] = parseJson_(mr[1], {});
  }

  var bundle = {
    exportado: String(exportado),
    version: version,
    salidas: salidas,
    responsables: responsables,
    territorios: territoriosCat,
    descripciones: descripciones
  };
  if (mapaExtra && typeof mapaExtra === 'object') {
    bundle.mapaSemanal = {};
    var keys = Object.keys(mapaExtra);
    for (var k = 0; k < keys.length; k++) bundle.mapaSemanal[keys[k]] = mapaExtra[keys[k]];
    bundle.mapaSemanal.territorios = mapaTerritorios;
    if (!bundle.mapaSemanal.exportado) bundle.mapaSemanal.exportado = bundle.exportado;
  } else if (Object.keys(mapaTerritorios).length) {
    bundle.mapaSemanal = { version: 1, exportado: bundle.exportado, territorios: mapaTerritorios };
  }
  return bundle;
}

function clearKeepHeader_(sh) {
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), 1)).clearContent();
}

function upsertDatosSemana_(bundle) {
  bustCache_();
  inicializar();
  var exportado = bundle.exportado || new Date().toISOString();
  var version = bundle.version || 2;
  var mapa = bundle.mapaSemanal && typeof bundle.mapaSemanal === 'object' ? bundle.mapaSemanal : {};
  var mapaTerr = mapa.territorios && typeof mapa.territorios === 'object' ? mapa.territorios : {};
  var mapaExtra = {
    version: mapa.version || 1,
    exportado: mapa.exportado || exportado,
    ajusteGlobal: mapa.ajusteGlobal || null,
    referencias: Array.isArray(mapa.referencias) ? mapa.referencias : []
  };

  var datosSh = sheet_(SHEET_DATOS);
  var tabla = leerTabla_(datosSh);
  var idx = idxHeader_(tabla.header);
  var found = 0;
  for (var i = 0; i < tabla.rows.length; i++) {
    var canal = String(cell_(tabla.rows[i], idx, ['canal'], '') || '').trim().toLowerCase();
    if (canal === CANAL_PUBLICA) {
      found = i + 2;
      break;
    }
  }
  var rowVals = [
    CANAL_PUBLICA,
    exportado,
    version,
    stringify_(bundle.responsables || []),
    stringify_(bundle.territorios || []),
    stringify_(bundle.descripciones || {}),
    stringify_(mapaExtra)
  ];
  if (found) {
    datosSh.getRange(found, 1, 1, rowVals.length).setValues([rowVals]);
  } else {
    datosSh.appendRow(rowVals);
  }

  var salidasSh = sheet_(SHEET_SALIDAS);
  clearKeepHeader_(salidasSh);
  var salidas = Array.isArray(bundle.salidas) ? bundle.salidas : [];
  if (salidas.length) {
    var out = [];
    for (var s = 0; s < salidas.length; s++) {
      var it = salidas[s] || {};
      out.push([
        it.NumeroTerritorio || '',
        it.PuntosDeEncuentros || '',
        it.Responsables || '',
        it.Grupos || '',
        it.Enviado || '',
        it.Prioridad == null ? '' : it.Prioridad,
        it.ManzanasPendientes || '',
        it.Direccion || '',
        it.Descripcion || ''
      ]);
    }
    salidasSh.getRange(2, 1, out.length, 9).setValues(out);
  }

  var mapaSh = sheet_(SHEET_MAPA);
  clearKeepHeader_(mapaSh);
  var nums = Object.keys(mapaTerr);
  if (nums.length) {
    var mOut = [];
    for (var n = 0; n < nums.length; n++) {
      mOut.push([nums[n], stringify_(mapaTerr[nums[n]])]);
    }
    mapaSh.getRange(2, 1, mOut.length, 2).setValues(mOut);
  }
  return { ok: true, salidas: salidas.length, territoriosMapa: nums.length };
}

function listarPendientes_() {
  var sh = sheet_(SHEET_ACCIONES);
  var tabla = leerTabla_(sh);
  var idx = idxHeader_(tabla.header);
  var out = [];
  for (var i = 0; i < tabla.rows.length; i++) {
    var r = tabla.rows[i];
    var proc = String(cell_(r, idx, ['procesado'], 'false')).toLowerCase();
    if (proc === 'true' || proc === '1' || proc === 'verdadero') continue;
    if (!cell_(r, idx, ['id', 'tipo'], '')) continue;
    out.push({
      id: cell_(r, idx, ['id'], i + 2),
      tipo: cell_(r, idx, ['tipo'], ''),
      payload: parseJson_(cell_(r, idx, ['payload', 'payload_json'], '{}'), {}),
      origen: cell_(r, idx, ['origen'], ORIGEN_WEB),
      creado_en: r[idx.creado_en] instanceof Date
        ? r[idx.creado_en].toISOString()
        : cell_(r, idx, ['creado_en'], '')
    });
  }
  out.sort(function (a, b) {
    return String(a.creado_en).localeCompare(String(b.creado_en));
  });
  return out;
}

function nextAccionId_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return 1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < ids.length; i++) {
    var n = Number(ids[i][0]);
    if (n > max) max = n;
  }
  return max + 1;
}

function findAccionRow_(id) {
  var sh = sheet_(SHEET_ACCIONES);
  var last = sh.getLastRow();
  if (last < 2) return { sh: sh, row: 0 };
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  var want = String(id);
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === want) return { sh: sh, row: i + 2 };
  }
  return { sh: sh, row: 0 };
}

function guardarAccion_(tipo, payload) {
  bustCache_();
  var sh = sheet_(SHEET_ACCIONES);
  var id = nextAccionId_(sh);
  var ahora = new Date().toISOString();
  sh.appendRow([id, tipo, stringify_(payload || {}), ORIGEN_WEB, false, ahora, '']);
  return { id: id, tipo: tipo, payload: payload || {}, origen: ORIGEN_WEB, creado_en: ahora };
}

function actualizarAccion_(id, tipo, payload) {
  bustCache_();
  var found = findAccionRow_(id);
  if (!found.row) throw new Error('No se encontró la acción id=' + id);
  found.sh.getRange(found.row, 2, 1, 2).setValues([[tipo, stringify_(payload || {})]]);
  return { id: id };
}

function eliminarAccion_(id) {
  bustCache_();
  var found = findAccionRow_(id);
  if (!found.row) throw new Error('No se eliminó ninguna fila (id=' + id + ').');
  var proc = String(found.sh.getRange(found.row, 5).getValue()).toLowerCase();
  if (proc === 'true' || proc === '1') {
    throw new Error('No se puede eliminar una acción ya procesada (id=' + id + ').');
  }
  found.sh.deleteRow(found.row);
  return [{ id: id }];
}

function eliminarPorNumero_(numero) {
  var num = String(numero || '').trim();
  var pendientes = listarPendientes_();
  var borradas = [];
  for (var i = pendientes.length - 1; i >= 0; i--) {
    var p = pendientes[i];
    var n = p.payload && p.payload.numero != null ? String(p.payload.numero).trim() : '';
    if (n === num) {
      eliminarAccion_(p.id);
      borradas.push(p);
    }
  }
  if (!borradas.length) {
    throw new Error('No se eliminó ninguna acción pendiente para el territorio ' + num + '.');
  }
  return borradas;
}

function marcarProcesada_(id) {
  bustCache_();
  var found = findAccionRow_(id);
  if (!found.row) throw new Error('No se encontró la acción id=' + id);
  found.sh.getRange(found.row, 5).setValue(true);
  found.sh.getRange(found.row, 7).setValue(new Date().toISOString());
  return { id: id, procesado: true };
}

function dispatch_(body) {
  var action = String((body && body.action) || '').trim();
  if (action === 'cargarDatosSemana') return ok_(cargarBundle_());
  if (action === 'cargarPagina') return ok_({ bundle: cargarBundle_(), pendientes: listarPendientes_() });
  if (action === 'listarAccionesPendientes') return ok_(listarPendientes_());
  if (action === 'upsertDatosSemana') return ok_(upsertDatosSemana_(body.bundle || {}));
  if (action === 'guardarAccion') return ok_(guardarAccion_(body.tipo, body.payload));
  if (action === 'actualizarAccion') return ok_(actualizarAccion_(body.id, body.tipo, body.payload));
  if (action === 'eliminarAccion') return ok_(eliminarAccion_(body.id));
  if (action === 'eliminarAccionesPendientesPorNumero') return ok_(eliminarPorNumero_(body.numero));
  if (action === 'marcarAccionProcesada') return ok_(marcarProcesada_(body.id));
  if (action === 'inicializar') {
    inicializar();
    return ok_({ listo: true });
  }
  throw new Error('Acción no reconocida: ' + action);
}

function doGet(e) {
  try {
    var body = (e && e.parameter) ? e.parameter : {};
    if (!body.action) body.action = 'cargarPagina';
    if (body.payload && typeof body.payload === 'string') {
      body.payload = parseJson_(body.payload, body.payload);
    }
    return dispatch_(body);
  } catch (err) {
    return fail_(err.message || err);
  }
}

function doPost(e) {
  try {
    return dispatch_(parseBody_(e));
  } catch (err) {
    return fail_(err.message || err);
  }
}
