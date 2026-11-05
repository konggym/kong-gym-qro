/**
 * KONG GYM · Conexión de la app con esta hoja (SOLO LECTURA)
 * ----------------------------------------------------------
 * La app pregunta por un socio con su ID (KG-008) y los últimos 4 dígitos
 * de su WhatsApp. Este script responde solo con lo necesario:
 * nombre, membresía, plan, fechas y estado. Nunca modifica la hoja y
 * nunca envía montos, teléfonos ni contactos de emergencia.
 *
 * Instalación: Extensiones → Apps Script → pega este código → Implementar →
 * Nueva implementación → Aplicación web → Ejecutar como: Yo →
 * Quién tiene acceso: Cualquier persona → Implementar → copia la URL.
 */

const HOJA_SOCIOS = 'Socios';
const MAX_INTENTOS = 5;   // intentos fallidos permitidos por ID
const BLOQUEO_MIN = 15;   // minutos de bloqueo tras rebasarlos

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.accion === 'ping') return json_({ ok: true, gym: 'Kong Gym' });
    if (p.accion === 'socio') return json_(buscarSocio_(p.id, p.tel));
    return json_({ ok: false, error: 'accion' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'servidor' });
  }
}

function buscarSocio_(idRaw, telRaw) {
  const id = String(idRaw || '').trim().toUpperCase().replace(/\s+/g, '');
  const tel = String(telRaw || '').replace(/\D/g, '');
  if (!/^KG-\d{1,5}$/.test(id) || tel.length !== 4) return { ok: false, error: 'datos' };

  const cache = CacheService.getScriptCache();
  const clave = 'fallos_' + id;
  const fallos = Number(cache.get(clave) || 0);
  if (fallos >= MAX_INTENTOS) return { ok: false, error: 'bloqueado', minutos: BLOQUEO_MIN };
  const fallo = () => {
    cache.put(clave, String(fallos + 1), BLOQUEO_MIN * 60);
    return { ok: false, error: 'no_coincide', restantes: Math.max(0, MAX_INTENTOS - fallos - 1) };
  };

  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_SOCIOS);
  if (!hoja) return { ok: false, error: 'hoja' };
  const datos = hoja.getDataRange().getValues();

  // Busca la fila de encabezados (la que dice "ID" en la columna A)
  let h = -1;
  for (let i = 0; i < Math.min(20, datos.length); i++) {
    if (String(datos[i][0]).trim() === 'ID') { h = i; break; }
  }
  if (h < 0) return { ok: false, error: 'hoja' };
  const col = {};
  datos[h].forEach((t, j) => { col[String(t).trim()] = j; });
  const v = (fila, nombre) => (col[nombre] === undefined ? '' : fila[col[nombre]]);

  for (let i = h + 1; i < datos.length; i++) {
    const fila = datos[i];
    if (String(fila[0]).trim().toUpperCase() !== id) continue;
    const nombre = String(v(fila, 'Nombre completo') || '').trim();
    if (!nombre) return fallo();
    const wa = String(v(fila, 'WhatsApp (10 dígitos)') || '').replace(/\D/g, '');
    if (!wa) return { ok: false, error: 'sin_whatsapp' };
    if (wa.slice(-4) !== tel) return fallo();
    cache.remove(clave);
    const dias = v(fila, 'Días restantes');
    return {
      ok: true,
      actualizado: new Date().toISOString(),
      socio: {
        id: id,
        nombre: nombre,
        membresia: String(v(fila, 'Membresía') || ''),
        plan: String(v(fila, 'Plan') || ''),
        inscripcion: fecha_(v(fila, 'Fecha de inscripción')),
        nacimiento: fecha_(v(fila, 'Fecha de nacimiento')),
        activo: String(v(fila, 'Activo') || '').trim().toLowerCase() !== 'no',
        vence: fecha_(v(fila, 'Vence el')),
        dias: dias === '' || dias === null ? null : Number(dias),
        estado: String(v(fila, 'Estado') || '')
      }
    };
  }
  return fallo();
}

function fecha_(x) {
  if (x instanceof Date && !isNaN(x)) return Utilities.formatDate(x, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const m = String(x || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return '';
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Para probar desde el editor: cambia el ID y los 4 dígitos y presiona ▶ Ejecutar */
function probar() {
  console.log(JSON.stringify(buscarSocio_('KG-002', '5513'), null, 2));
}
