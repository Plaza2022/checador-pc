// ============================================================
//  Bitácora Plaza China Multiplaza - Apps Script
//  Pegar en Extensiones > Apps Script dentro del archivo BitacoraPC
// ============================================================

// ---------- CONFIGURACIÓN (ajusta estos valores) ----------
const HOJA_EMPLEADOS = 'EMPLEADOS';
const HOJA_TIENDAS   = 'TIENDAS';
const HOJA_BITACORA  = 'BITACORA';   // <- pon aquí el nombre EXACTO de tu hoja de bitácora actual
const CLAVE_ADMIN    = '2126';       // ahora vive en el servidor, ya no en el HTML
const EVENTOS        = ['Entrada', 'Comer', 'Retorno', 'Salida'];

// true  = el empleado solo puede registrar en la tienda que tiene asignada en EMPLEADOS
// false = puede registrar en cualquier tienda (se guarda la tienda del QR)
const REQUIERE_MISMA_TIENDA = false;

// ---------- Columnas de EMPLEADOS (A..H) ----------
// A llave | B tienda | C nombre | D apellido | E nombre_lista | F estatus | G fecha_alta | H fecha_baja
// ---------- Columnas de TIENDAS (A..C) ----------
// A tienda (igual que el texto del QR) | B url_horario | C activa (Sí / No)

// ============================================================
//  ENTRADA WEB
// ============================================================
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    switch (p.accion) {
      case 'empleados': return json_(empleados_(p.tienda));
      case 'validar':   return json_(validar_(p.llave));
      case 'horario':   return json_(horario_(p.tienda));
      case 'registrar': return json_(registrar_(p));
      case 'admin':     return json_({ ok: String(p.clave) === CLAVE_ADMIN });
      default:          return json_({ ok: false, mensaje: 'Acción no válida' });
    }
  } catch (err) {
    return json_({ ok: false, mensaje: 'Error del servidor: ' + err.message });
  }
}

// ============================================================
//  ACCIONES
// ============================================================
function empleados_(tienda) {
  const t = buscarTienda_(tienda);
  if (!t) return { ok: false, mensaje: 'Tienda no reconocida. Escanee el QR de nuevo.' };

  const lista = datos_(HOJA_EMPLEADOS)
    .filter(r => norm_(r[1]) === norm_(t.nombre) && norm_(r[5]) === 'activo')
    .map(r => ({ llave: String(r[0]), nombre: String(r[4]) }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre));

  return { ok: true, empleados: lista };
}

function validar_(llave) {
  const emp = buscarEmpleado_(llave);
  if (!emp || norm_(emp[5]) !== 'activo') {
    return { ok: false, codigo: 'BAJA', mensaje: 'Este usuario ya no tiene acceso.' };
  }
  return { ok: true, nombre: String(emp[4]) };
}

function horario_(tienda) {
  const t = buscarTienda_(tienda);
  if (!t) return { ok: false, mensaje: 'Tienda no reconocida.' };
  if (!t.url) return { ok: false, mensaje: 'Esta tienda aún no tiene horario configurado.' };
  return { ok: true, url: t.url };
}

function registrar_(p) {
  if (EVENTOS.indexOf(p.evento) === -1) {
    return { ok: false, mensaje: 'Evento no válido.' };
  }

  const t = buscarTienda_(p.tienda);
  if (!t) return { ok: false, mensaje: 'Tienda no reconocida. Escanee el QR de nuevo.' };

  const emp = buscarEmpleado_(p.llave);
  if (!emp || norm_(emp[5]) !== 'activo') {
    return { ok: false, codigo: 'BAJA', mensaje: 'Acceso denegado: este usuario ya no está activo.' };
  }

  if (REQUIERE_MISMA_TIENDA && norm_(emp[1]) !== norm_(t.nombre)) {
    return { ok: false, mensaje: 'Este empleado no pertenece a la tienda ' + t.nombre + '.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // ORDEN DE COLUMNAS DE LA BITÁCORA: ajusta si la tuya es distinta.
    // Se agregó "llave" al FINAL para no mover las columnas que ya tienes.
    // Fecha | Tienda | Empleado | Evento | Llave
    SpreadsheetApp.getActive().getSheetByName(HOJA_BITACORA)
      .appendRow([new Date(), t.nombre, String(emp[4]), p.evento, String(emp[0])]);
  } finally {
    lock.releaseLock();
  }

  return { ok: true, mensaje: 'Registro guardado: ' + p.evento + ' (' + emp[4] + ')' };
}

// ============================================================
//  ALTA AUTOMÁTICA EN LA HOJA EMPLEADOS
//  Al capturar tienda/nombre en una fila nueva: genera la llave,
//  pone estatus Activo y fecha de alta. Al cambiar a Baja: fecha de baja.
// ============================================================
function onEdit(e) {
  if (!e || !e.range) return;
  const hoja = e.range.getSheet();
  if (hoja.getName() !== HOJA_EMPLEADOS) return;

  const fila = e.range.getRow();
  if (fila < 2) return;

  const r = hoja.getRange(fila, 1, 1, 8).getValues()[0];

  // Alta: hay datos pero todavía no hay llave
  if (!r[0] && (r[1] || r[2] || r[4])) {
    hoja.getRange(fila, 1).setValue('E-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase());
    if (!r[5]) hoja.getRange(fila, 6).setValue('Activo');
    if (!r[6]) hoja.getRange(fila, 7).setValue(new Date());
    return;
  }

  // Baja / reactivación
  if (norm_(r[5]) === 'baja' && !r[7]) hoja.getRange(fila, 8).setValue(new Date());
  if (norm_(r[5]) === 'activo' && r[7]) hoja.getRange(fila, 8).clearContent();
}

// ============================================================
//  AUXILIARES
// ============================================================
function norm_(s) {
  return String(s == null ? '' : s).trim().toLowerCase();
}

function datos_(nombreHoja) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(nombreHoja);
  if (!hoja) throw new Error('No existe la hoja "' + nombreHoja + '"');
  return hoja.getDataRange().getValues().slice(1); // sin encabezados
}

function buscarTienda_(tienda) {
  const fila = datos_(HOJA_TIENDAS).find(r => norm_(r[0]) === norm_(tienda));
  if (!fila) return null;
  if (norm_(fila[2]) === 'no') return null; // tienda desactivada
  return { nombre: String(fila[0]).trim(), url: String(fila[1] || '').trim() };
}

function buscarEmpleado_(llave) {
  if (!llave) return null;
  return datos_(HOJA_EMPLEADOS).find(r => String(r[0]) === String(llave)) || null;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
