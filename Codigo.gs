/**
 * AULA DE PROTECCIÓN CIVIL · Registro de resultados en Google Sheets
 * IES Gregorio Prieto
 *
 * INSTALACIÓN (una sola vez):
 *  1. Crea una hoja de cálculo nueva en Google Drive (p. ej. "Protección Civil · Registro").
 *  2. Menú Extensiones > Apps Script. Borra el contenido y pega este archivo completo. Guarda.
 *  3. Botón Implementar > Nueva implementación > tipo "Aplicación web".
 *       - Ejecutar como: Yo
 *       - Quién tiene acceso: Cualquier usuario
 *  4. Autoriza los permisos y copia la URL que termina en /exec.
 *  5. Pega esa URL en index.html, en CONFIG.SCRIPT_URL.
 *
 * Si cambias este código, usa Implementar > Gestionar implementaciones > Editar > Nueva versión
 * para que la URL siga siendo la misma.
 *
 * Hojas que se crean solas:
 *  - Registros: cada entrega de módulo (todas las veces que se repite).
 *  - Resumen: una fila por alumno/a con la MEJOR nota de cada módulo, media y si ha completado.
 *  - PlanFamiliar: cada registro de la tarea práctica (solo nº de apartados, sin datos de la familia).
 */

const HOJA_REG = 'Registros';
const HOJA_RES = 'Resumen';
const HOJA_PLAN = 'PlanFamiliar';
const NOTA_MINIMA = 5;

const CAB_REG = ['Fecha', 'Alumno/a', 'Grupo', 'Módulo', 'Título', 'Nota', 'Test', 'Actividad', 'Intento', 'Reflexión', 'Código'];
const CAB_RES = ['Alumno/a', 'Grupo', 'M1', 'M2', 'M3', 'M4', 'M5', 'Plan familiar', 'Media módulos', 'Completado', 'Última actividad'];
const CAB_PLAN = ['Fecha', 'Alumno/a', 'Grupo', 'Apartados', 'Apartados completos', 'Código'];

function doGet() {
  return ContentService.createTextOutput('Aula de Protección Civil: registro activo.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!d.nombre || !d.grupo) return json({ ok: false, error: 'Faltan nombre o grupo' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (d.tipo === 'modulo') {
      const sh = hoja(ss, HOJA_REG, CAB_REG);
      sh.appendRow([new Date(), txt(d.nombre), txt(d.grupo), txt(d.modulo), txt(d.titulo),
        num(d.nota), txt(d.test), txt(d.actividad), num(d.intento) || 1, txt(d.reflexion, 1500), txt(d.codigo)]);
      actualizarResumen(ss, d);
    } else if (d.tipo === 'plan') {
      const sh = hoja(ss, HOJA_PLAN, CAB_PLAN);
      sh.appendRow([new Date(), txt(d.nombre), txt(d.grupo), txt(d.apartados), txt(d.detalle), txt(d.codigo)]);
      actualizarResumen(ss, d);
    } else {
      return json({ ok: false, error: 'Tipo desconocido' });
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function actualizarResumen(ss, d) {
  const sh = hoja(ss, HOJA_RES, CAB_RES);
  const datos = sh.getDataRange().getValues();
  const clave = (limpia(d.nombre) + '|' + limpia(d.grupo)).toLowerCase();
  let fila = -1;
  for (let i = 1; i < datos.length; i++) {
    if ((limpia(datos[i][0]) + '|' + limpia(datos[i][1])).toLowerCase() === clave) { fila = i + 1; break; }
  }
  if (fila === -1) {
    sh.appendRow([txt(d.nombre), txt(d.grupo), '', '', '', '', '', '', '', 'NO', new Date()]);
    fila = sh.getLastRow();
  }

  if (d.tipo === 'modulo') {
    const n = Number(String(d.modulo).replace(/\D/g, ''));
    if (n >= 1 && n <= 5) {
      const celda = sh.getRange(fila, 2 + n); // M1 está en la columna 3
      const actual = celda.getValue();
      const nota = num(d.nota);
      if (actual === '' || nota > Number(actual)) celda.setValue(nota);
    }
  } else if (d.tipo === 'plan') {
    sh.getRange(fila, 8).setValue(txt(d.apartados));
  }

  const notas = sh.getRange(fila, 3, 1, 5).getValues()[0];
  const hechas = notas.filter(v => v !== '');
  const media = hechas.length ? Math.round(hechas.reduce((a, b) => a + Number(b), 0) / hechas.length * 10) / 10 : '';
  const completado = notas.every(v => v !== '' && Number(v) >= NOTA_MINIMA) ? 'SÍ' : 'NO';
  sh.getRange(fila, 9, 1, 3).setValues([[media, completado, new Date()]]);
}

/* ---------- utilidades ---------- */
function hoja(ss, nombre, cabecera) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.appendRow(cabecera);
    sh.getRange(1, 1, 1, cabecera.length).setFontWeight('bold').setBackground('#0f4a8a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}
function limpia(v) { return String(v == null ? '' : v).trim().replace(/\s+/g, ' '); }
// Evita que un texto que empiece por = + - @ se interprete como fórmula.
function txt(v, max) {
  let s = limpia(v).slice(0, max || 300);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
function num(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
