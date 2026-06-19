// ═══════════════════════════════════════════════════════════════════
//  MADRYN NORTE — Backend Google Apps Script
//  1. Abrí el Sheet → Extensiones → Apps Script
//  2. Borrá todo y pegá este archivo → Guardá
//  3. Ejecutá "inicializarHojas" → autorizá permisos
//  4. Implementar → Nueva implementación → Aplicación web
//     · Ejecutar como: Yo  · Acceso: Cualquier usuario
//  5. Copiá la URL y pegala en la app
// ═══════════════════════════════════════════════════════════════════

const SS_ID = '15BswSBX6aaKWoRyNaRfqha5JPji8LylgG01wEYjlxnA';

const ESQUEMA = {
  usuarios:      ['id','nombre','username','email','password','rol','activo','foto'],
  actividades:   ['id','nombre','profesor_id','dias_horarios','valor_mensual','cupo_maximo','estado'],
  alumnos:       ['id','nombre','apellido','dni','fecha_nacimiento','telefono',
                  'es_menor','es_sponsor','responsable_nombre','responsable_dni','responsable_vinculo','responsable_telefono',
                  'estado',
                  'tipo_descuento','descuento_pct','motivo_descuento',
                  'observaciones','salud','contacto_emergencia',
                  'fecha_alta','fecha_baja','foto'],
  inscripciones: ['id','alumno_id','actividad_id','estado','fecha_inscripcion'],
  cuotas:        ['id','alumno_id','actividad_id','tipo','mes_anio','monto','descuento','estado','medio_pago','fecha_pago','observaciones'],
  asistencia:    ['id','alumno_id','actividad_id','fecha','estado','observacion','registrado_por'],
  nuevos:        ['id','nombre','apellido','dni','fecha_nacimiento','telefono','es_menor',
                  'responsable_nombre','responsable_dni','responsable_vinculo','responsable_telefono',
                  'actividad','salud','contacto_emergencia','acepta_condiciones',
                  'foto_url','estado','fecha_envio','fuente'],
};

// ── Respuestas ────────────────────────────────────────────────────

function respOk(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data: data || null }))
    .setMimeType(ContentService.MimeType.JSON);
}
function respErr(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Sheet ─────────────────────────────────────────────────────────

function getSS() { return SpreadsheetApp.openById(SS_ID); }

function inicializarHojas() {
  const ss = getSS();
  for (const nombre in ESQUEMA) {
    let hoja = ss.getSheetByName(nombre);
    if (!hoja) hoja = ss.insertSheet(nombre);
    const cols = ESQUEMA[nombre];
    if (hoja.getLastRow() === 0) {
      hoja.getRange(1, 1, 1, cols.length).setValues([cols]);
      hoja.getRange(1, 1, 1, cols.length)
        .setFontWeight('bold').setBackground('#1a6636').setFontColor('#ffffff');
      hoja.setFrozenRows(1);
    } else {
      agregarColumnasNuevas(hoja, cols);
    }
  }
  obtenerCarpetaFotos();
  // Asegurar que _config tenga la clave valor_cuota_socio
  leerConfigSheet();
  Logger.log('Hojas inicializadas');
}

function agregarColumnasNuevas(hoja, cols) {
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h).trim(); });
  cols.forEach(function(col) {
    if (!encabezados.includes(col)) {
      const nuevaCol = hoja.getLastColumn() + 1;
      hoja.getRange(1, nuevaCol).setValue(col)
        .setFontWeight('bold').setBackground('#1a6636').setFontColor('#ffffff');
      Logger.log('Columna agregada: ' + col + ' en ' + hoja.getName());
    }
  });
}

// ── Drive / Fotos ─────────────────────────────────────────────────

function obtenerCarpetaFotos() {
  const NOMBRE = 'Madryn Norte - Fotos';
  const ss = getSS();
  let cfg = ss.getSheetByName('_config');
  if (!cfg) { cfg = ss.insertSheet('_config'); cfg.hideSheet(); }
  const it = DriveApp.getFoldersByName(NOMBRE);
  if (it.hasNext()) {
    const c = it.next();
    if (cfg.getLastRow() === 0) cfg.appendRow(['foto_folder_id', c.getId()]);
    return c;
  }
  const padreIt = DriveApp.getFileById(SS_ID).getParents();
  const padre = padreIt.hasNext() ? padreIt.next() : DriveApp.getRootFolder();
  const nueva = padre.createFolder(NOMBRE);
  cfg.appendRow(['foto_folder_id', nueva.getId()]);
  return nueva;
}

function guardarFotoEnDrive(base64Data, nombreArchivo) {
  try {
    if (!base64Data || !base64Data.includes(',')) return null;
    const partes = base64Data.split(',');
    const mime   = partes[0].match(/:(.*?);/)[1] || 'image/jpeg';
    const bytes  = Utilities.base64Decode(partes[1]);
    const blob   = Utilities.newBlob(bytes, mime, nombreArchivo + '.jpg');
    const carpeta = obtenerCarpetaFotos();
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/thumbnail?id=' + archivo.getId() + '&sz=w400';
  } catch (e) {
    Logger.log('Error foto: ' + e.message);
    return null;
  }
}

// ── Config sheet (clave-valor) ────────────────────────────────────

function leerConfigSheet() {
  const ss = getSS();
  let sh = ss.getSheetByName('_config');
  if (!sh) { sh = ss.insertSheet('_config'); sh.hideSheet(); }
  const cfg = {};
  if (sh.getLastRow() > 0) {
    const vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
    vals.forEach(function(r) { if (r[0]) cfg[String(r[0])] = r[1]; });
  }
  // Valor por defecto
  if (cfg['valor_cuota_socio'] === undefined) {
    sh.appendRow(['valor_cuota_socio', 0]);
    cfg['valor_cuota_socio'] = 0;
  }
  return cfg;
}

function guardarConfigSheet(clave, valor) {
  const ss = getSS();
  let sh = ss.getSheetByName('_config');
  if (!sh) { sh = ss.insertSheet('_config'); sh.hideSheet(); }
  if (sh.getLastRow() > 0) {
    const claves = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    for (let i = 0; i < claves.length; i++) {
      if (String(claves[i][0]) === String(clave)) {
        sh.getRange(i + 1, 2).setValue(valor);
        return;
      }
    }
  }
  sh.appendRow([clave, valor]);
}

// ── CRUD ──────────────────────────────────────────────────────────

function leerHoja(nombre) {
  const ss   = getSS();
  const hoja = ss.getSheetByName(nombre);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const cols   = ESQUEMA[nombre];
  const nCols  = Math.max(cols.length, hoja.getLastColumn());
  const hdrs   = hoja.getRange(1, 1, 1, nCols).getValues()[0].map(String);
  const valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, nCols).getValues();
  return valores
    .filter(function(f) { return f[0] !== '' && f[0] !== null; })
    .map(function(fila) {
      const obj = {};
      hdrs.forEach(function(h, i) {
        let v = fila[i];
        if (v instanceof Date) {
          // mes_anio se guarda como "2025-06" pero Sheets lo lee como fecha → volvemos al formato correcto
          if (h === 'mes_anio') {
            v = Utilities.formatDate(v, 'UTC', 'yyyy-MM');
          } else {
            // 'UTC' evita que el offset de Argentina (-3h) corra la fecha al día anterior
            v = Utilities.formatDate(v, 'UTC', 'yyyy-MM-dd');
          }
        }
        obj[h] = (v === '' || v === null) ? null : v;
      });
      return obj;
    });
}

function insertarFila(nombre, datos) {
  const ss   = getSS();
  const hoja = ss.getSheetByName(nombre);
  if (!hoja) throw new Error('Hoja no encontrada: ' + nombre);
  // Autosincroniza columnas del esquema antes de escribir, así campos nuevos
  // (ej: username/password) no se pierden si la hoja se creó antes del cambio.
  if (ESQUEMA[nombre] && hoja.getLastRow() > 0) agregarColumnasNuevas(hoja, ESQUEMA[nombre]);
  const lastCol = hoja.getLastColumn();
  const hdrs = lastCol > 0
    ? hoja.getRange(1, 1, 1, lastCol).getValues()[0].map(String)
    : ESQUEMA[nombre];
  const fila = hdrs.map(function(h) {
    const v = datos[h];
    return (v !== undefined && v !== null) ? v : '';
  });
  hoja.appendRow(fila);
}

function actualizarFila(nombre, id, datos) {
  const ss   = getSS();
  const hoja = ss.getSheetByName(nombre);
  if (!hoja) throw new Error('Hoja no encontrada: ' + nombre);
  if (ESQUEMA[nombre] && hoja.getLastRow() > 0) agregarColumnasNuevas(hoja, ESQUEMA[nombre]);
  const nCols = Math.max(ESQUEMA[nombre].length, hoja.getLastColumn());
  const hdrs  = hoja.getRange(1, 1, 1, nCols).getValues()[0].map(String);
  const ult   = hoja.getLastRow();
  if (ult < 2) throw new Error('Sin datos en ' + nombre);
  const ids = hoja.getRange(2, 1, ult - 1, 1).getValues();
  let numFila = -1;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) { numFila = i + 2; break; }
  }
  if (numFila < 0) throw new Error('ID no encontrado: ' + id);
  const actual = hoja.getRange(numFila, 1, 1, nCols).getValues()[0];
  const actualObj = {};
  hdrs.forEach(function(h, i) { actualObj[h] = actual[i]; });
  const merged   = Object.assign({}, actualObj, datos);
  const nuevaFila = hdrs.map(function(h) {
    const v = merged[h];
    return (v !== undefined && v !== null) ? v : '';
  });
  hoja.getRange(numFila, 1, 1, nCols).setValues([nuevaFila]);
}

function eliminarFila(nombre, id) {
  const ss   = getSS();
  const hoja = ss.getSheetByName(nombre);
  if (!hoja) return;
  const ult = hoja.getLastRow();
  if (ult < 2) return;
  const ids = hoja.getRange(2, 1, ult - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) { hoja.deleteRow(i + 2); return; }
  }
}

// ── Autenticación ─────────────────────────────────────────────────

function verificarLogin(identifier, pwdHash) {
  if (!identifier) return null;
  const usuarios = leerHoja('usuarios');
  const id = String(identifier).toLowerCase();
  const u = usuarios.find(function(u) {
    return (u.username && String(u.username).toLowerCase() === id) ||
           (u.email && String(u.email).toLowerCase() === id);
  });
  if (!u) return null;
  if (!u.password) return u;
  if (u.password === pwdHash) return u;
  return null;
}

function verificarUsuario(identifier) {
  if (!identifier) return null;
  const usuarios = leerHoja('usuarios');
  const id = String(identifier).toLowerCase();
  return usuarios.find(function(u) {
    return (u.email && String(u.email).toLowerCase() === id) ||
           (u.username && String(u.username).toLowerCase() === id);
  }) || null;
}

// ── GET ───────────────────────────────────────────────────────────

function doGet(e) {
  try {
    const accion = e.parameter.accion;

    if (accion === 'init') {
      inicializarHojas();
      return respOk({ message: 'Hojas creadas y listas' });
    }

    if (accion === 'todo') {
      const data = {};
      for (const nombre in ESQUEMA) {
        data[nombre] = leerHoja(nombre);
      }
      data.nuevos = (data.nuevos || []).map(function(n) {
        n.foto = n.foto_url; return n;
      });
      data.usuarios = (data.usuarios || []).map(function(u) {
        const c = Object.assign({}, u); delete c.password; return c;
      });
      // Incluir config general (valor_cuota_socio, etc.)
      data.config = leerConfigSheet();
      return respOk(data);
    }

    return respErr('Acción GET no reconocida: ' + accion);
  } catch (ex) {
    return respErr(ex.message);
  }
}

// ── POST ──────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const accion = body.accion;
    const hoja   = body.hoja;
    const id     = body.id;
    const datos  = body.datos;
    const email  = body.email;
    const pwd    = body.pwd;

    // ── Login ──
    if (accion === 'login') {
      const identifier = body.username || body.email || '';
      const u = verificarLogin(identifier, pwd);
      if (!u) return respErr('Usuario o contraseña incorrectos');
      const safe = Object.assign({}, u);
      delete safe.password;
      return respOk({ usuario: safe });
    }

    // ── Registro inicial admin ──
    if (accion === 'registrar_usuario') {
      const usuarios = leerHoja('usuarios');
      const existe = usuarios.find(function(u) {
        return u.email && u.email.toLowerCase() === (datos.email || '').toLowerCase();
      });
      if (!existe) insertarFila('usuarios', datos);
      return respOk({ message: 'Usuario registrado' });
    }

    // ── Inscripción pública (sin auth) ──
    if (accion === 'inscripcion_nueva') {
      if (!datos) return respErr('Sin datos');
      let fotoUrl = null;
      if (datos.foto && datos.foto.length > 100) {
        const nom = (datos.nombre || 'alumno') + '_' + (datos.dni || Date.now());
        fotoUrl = guardarFotoEnDrive(datos.foto, nom);
      }
      const reg = {
        id:                   datos.id || ('x' + new Date().getTime()),
        nombre:               datos.nombre || '',
        apellido:             datos.apellido || '',
        dni:                  datos.dni || '',
        fecha_nacimiento:     datos.fecha_nacimiento || '',
        telefono:             datos.telefono || '',
        es_menor:             datos.es_menor || '',
        responsable_nombre:   datos.responsable_nombre || '',
        responsable_dni:      datos.responsable_dni || '',
        responsable_vinculo:  datos.responsable_vinculo || '',
        responsable_telefono: datos.responsable_telefono || '',
        actividad:            datos.actividad || '',
        salud:                datos.salud || '',
        contacto_emergencia:  datos.contacto_emergencia || '',
        acepta_condiciones:   datos.acepta_condiciones || '',
        foto_url:             fotoUrl || '',
        estado:               'pendiente',
        fecha_envio:          new Date().toISOString(),
        fuente:               datos.fuente || 'formulario',
      };
      insertarFila('nuevos', reg);
      return respOk({ message: 'Inscripción recibida' });
    }

    // ── Acciones autenticadas ──
    const usuario = verificarUsuario(email);
    if (!usuario) return respErr('Usuario no autorizado: ' + email);

    if (accion === 'insertar') {
      if (!hoja || !datos) return respErr('Faltan parámetros');
      insertarFila(hoja, datos);
      return respOk({ message: 'Insertado' });
    }

    if (accion === 'actualizar') {
      if (!hoja || !id || !datos) return respErr('Faltan parámetros');
      actualizarFila(hoja, id, datos);
      return respOk({ message: 'Actualizado' });
    }

    if (accion === 'eliminar') {
      if (!hoja || !id) return respErr('Faltan parámetros');
      eliminarFila(hoja, id);
      return respOk({ message: 'Eliminado' });
    }

    // ── Guardar configuración general ──
    if (accion === 'guardar_config') {
      if (!datos) return respErr('Sin datos');
      Object.keys(datos).forEach(function(k) { guardarConfigSheet(k, datos[k]); });
      return respOk({ message: 'Configuración guardada' });
    }

    return respErr('Acción no reconocida: ' + accion);
  } catch (ex) {
    return respErr(ex.message);
  }
}

// ── Google Forms (trigger onFormSubmit) ───────────────────────────

const CAMPOS_FORM = {
  'A que actividad queres inscribirte?':                                           'actividad',
  'NOMBRE Y APELLIDO DEL PARTICIPANTE:':                                           '_nombre_completo',
  'FECHA DE NACIMIENTO DEL PARTICIPANTE:':                                         'fecha_nacimiento',
  'EDAD DEL PARTICIPANTE :':                                                       '_ignorar',
  'DNI DEL PARTICIPANTE':                                                          'dni',
  'NUMERO DE CONTACTO (whatsapp) DEL PARTICIPANTE. (No necesario en menores)':    'telefono',
  'El participante es menor de edad?':                                             'es_menor',
  'Nombre de Padre, Madre o Tutor que autoriza':                                   'responsable_nombre',
  'DNI de adulto responsable':                                                     'responsable_dni',
  'Vinculo con el participante':                                                   'responsable_vinculo',
  'Numero de contacto (whatsapp) del responsable (Obligatorio)':                   'responsable_telefono',
  '¿Tiene alguna lesión, enfermedad o condición médica que debamos conocer? SI / NO - DETALLAR': 'salud',
  'Contacto de emergencia:':                                                        'contacto_emergencia',
  'Declaro haber leído y acepto las condiciones de inscripción, participación y deslinde de responsabilidad.': 'acepta_condiciones',
};

function onFormSubmit(e) {
  try {
    const respuestas = e.response.getItemResponses();
    const datos = { id:'gf_'+new Date().getTime(), nombre:'', apellido:'', dni:'',
      fecha_nacimiento:'', telefono:'', es_menor:'', responsable_nombre:'',
      responsable_dni:'', responsable_vinculo:'', responsable_telefono:'',
      actividad:'', salud:'', contacto_emergencia:'', acepta_condiciones:'',
      foto_url:'', estado:'pendiente', fecha_envio:new Date().toISOString(), fuente:'Google Forms' };
    respuestas.forEach(function(r) {
      const pregunta = r.getItem().getTitle().trim();
      const campo    = CAMPOS_FORM[pregunta];
      if (!campo || campo === '_ignorar') return;
      if (campo === '_nombre_completo') {
        const val = (r.getResponse() || '').trim(), partes = val.split(' ');
        datos.nombre = partes[0] || ''; datos.apellido = partes.slice(1).join(' ') || '';
      } else {
        const resp = r.getResponse();
        datos[campo] = Array.isArray(resp) ? resp.join(', ') : (resp || '');
      }
    });
    if (datos.fecha_nacimiento instanceof Date)
      datos.fecha_nacimiento = Utilities.formatDate(datos.fecha_nacimiento, 'UTC', 'yyyy-MM-dd');
    insertarFila('nuevos', datos);
    Logger.log('Form submit: ' + datos.nombre + ' ' + datos.apellido);
  } catch (ex) { Logger.log('Error onFormSubmit: ' + ex.message); }
}

// ═══════════════════════════════════════════════════════════════════
//  UTILIDADES — Ejecutar desde el editor, una sola vez
// ═══════════════════════════════════════════════════════════════════

// OPCIONAL / MANUAL: ejecutar una sola vez desde el editor de Apps Script.
// Renombra el header 'tipo_socio' (columna ya retirada del ESQUEMA) a
// '_legado_tipo_socio' para conservar el dato histórico sin que vuelva a
// aparecer como columna activa de alumnos.
function migrarEsquemaAlumnos() {
  const ss = getSS();
  const hoja = ss.getSheetByName('alumnos');
  if (!hoja || hoja.getLastColumn() === 0) { Logger.log('Hoja alumnos vacía o inexistente'); return; }
  const hdrs = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  const idx = hdrs.indexOf('tipo_socio');
  if (idx < 0) { Logger.log('No hay columna tipo_socio para migrar'); return; }
  hoja.getRange(1, idx + 1).setValue('_legado_tipo_socio');
  Logger.log('Columna tipo_socio renombrada a _legado_tipo_socio');
}

// Convierte alumnos con IDs viejos → A1001+
// Convierte actividades con IDs UUID → slug del nombre
// Actualiza todas las referencias en inscripciones / cuotas / asistencia
function arreglarIds() {
  const ss = getSS();

  function slugAct(nombre) {
    return String(nombre).toLowerCase()
      .replace(/[áàäâ]/g,'a').replace(/[éèëê]/g,'e').replace(/[íìïî]/g,'i')
      .replace(/[óòöô]/g,'o').replace(/[úùüû]/g,'u').replace(/ñ/g,'n')
      .replace(/[^a-z0-9]+/g,'-').replace(/^-|-$/g,'');
  }

  const shAct = ss.getSheetByName('actividades');
  const actVals = shAct.getDataRange().getValues();
  const actH = actVals[0];
  const actIdIdx = actH.indexOf('id'); const actNomIdx = actH.indexOf('nombre');
  const actIdMap = {};
  const usedSlugs = {};
  for (let i = 1; i < actVals.length; i++) {
    const viejoId = String(actVals[i][actIdIdx]);
    const nombre  = String(actVals[i][actNomIdx]);
    let slug = slugAct(nombre);
    if (!slug) slug = 'actividad-' + i;
    if (usedSlugs[slug]) slug = slug + '-' + i;
    usedSlugs[slug] = true;
    if (viejoId !== slug) {
      actIdMap[viejoId] = slug;
      shAct.getRange(i + 1, actIdIdx + 1).setValue(slug);
    }
  }

  const shAl = ss.getSheetByName('alumnos');
  const alVals = shAl.getDataRange().getValues();
  const alH = alVals[0]; const alIdIdx = alH.indexOf('id');
  const alIdMap = {};
  let maxNum = 1000;
  for (let i = 1; i < alVals.length; i++) {
    const id = String(alVals[i][alIdIdx]);
    const m = id.match(/^A(\d+)$/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1]));
  }
  for (let i = 1; i < alVals.length; i++) {
    const viejoId = String(alVals[i][alIdIdx]);
    if (!viejoId.match(/^A\d{4,}$/)) {
      maxNum++;
      const nuevoId = 'A' + String(maxNum).padStart(4,'0');
      alIdMap[viejoId] = nuevoId;
      shAl.getRange(i + 1, alIdIdx + 1).setValue(nuevoId);
    }
  }

  function actualizarRefs(sheetName, alCol, actCol) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    const vals = sh.getDataRange().getValues();
    const h = vals[0];
    const alIdx  = alCol  !== null ? h.indexOf(alCol)  : -1;
    const actIdx = actCol !== null ? h.indexOf(actCol) : -1;
    for (let i = 1; i < vals.length; i++) {
      if (alIdx >= 0) {
        const v = String(vals[i][alIdx]);
        if (alIdMap[v]) sh.getRange(i+1, alIdx+1).setValue(alIdMap[v]);
      }
      if (actIdx >= 0) {
        const v = String(vals[i][actIdx]);
        if (actIdMap[v]) sh.getRange(i+1, actIdx+1).setValue(actIdMap[v]);
      }
    }
  }

  actualizarRefs('inscripciones', 'alumno_id', 'actividad_id');
  actualizarRefs('cuotas',        'alumno_id', 'actividad_id');
  actualizarRefs('asistencia',    'alumno_id', 'actividad_id');

  Logger.log('IDs corregidos. Alumnos: ' + Object.keys(alIdMap).length +
             '. Actividades: ' + Object.keys(actIdMap).length);
}

// Agrega 10 alumnos ficticios (observaciones="PRUEBA") a cada actividad sin inscriptos
function agregarAlumnosPrueba() {
  const ss = getSS();
  const shAl  = ss.getSheetByName('alumnos');
  const shIns = ss.getSheetByName('inscripciones');
  const shAct = ss.getSheetByName('actividades');

  const alH   = shAl.getRange(1,1,1,shAl.getLastColumn()).getValues()[0];
  const insH  = shIns.getRange(1,1,1,shIns.getLastColumn()).getValues()[0];

  const actVals = shAct.getDataRange().getValues();
  const actH = actVals[0];
  const actIdIdx  = actH.indexOf('id');
  const actNomIdx = actH.indexOf('nombre');

  const insVals = shIns.getLastRow() > 1 ? shIns.getRange(2,1,shIns.getLastRow()-1,shIns.getLastColumn()).getValues() : [];
  const actIdInsIdx = insH.indexOf('actividad_id');
  const actsConAlumnos = new Set(insVals.map(function(r) { return String(r[actIdInsIdx]); }));

  const alVals = shAl.getLastRow() > 1 ? shAl.getRange(2,1,shAl.getLastRow()-1,1).getValues() : [];
  let maxNum = 1000;
  alVals.forEach(function(r) { const m = String(r[0]).match(/^A(\d+)$/); if(m) maxNum = Math.max(maxNum, parseInt(m[1])); });

  const NOMBRES = ['Lucas','Valentina','Santiago','Florencia','Mateo','Sofía','Nicolás','Camila','Diego','Martina'];
  const APELLS  = ['García','Rodríguez','González','Fernández','López','Martínez','Pérez','Sánchez','Romero','Torres'];
  const FNACS   = ['1990-03-15','1991-07-22','1993-11-08','1994-05-30','1995-02-14','1996-09-01','1998-04-19','2000-12-25','2002-06-07','2004-08-11'];

  let insertados = 0;
  for (let a = 1; a < actVals.length; a++) {
    const actId  = String(actVals[a][actIdIdx]);
    const actNom = String(actVals[a][actNomIdx]);
    if (actsConAlumnos.has(actId)) continue;

    for (let k = 0; k < 10; k++) {
      maxNum++;
      const alId = 'A' + String(maxNum).padStart(4,'0');
      const alumno = {
        id: alId, nombre: NOMBRES[k], apellido: APELLS[k],
        dni: String(44000000 + maxNum), fecha_nacimiento: FNACS[k],
        telefono: '2804' + String(100000 + maxNum), es_menor: 'NO',
        responsable_nombre:'', responsable_dni:'', responsable_vinculo:'', responsable_telefono:'',
        estado:'activo', tipo_socio:'activo', tipo_descuento:'ninguno', descuento_pct:'0', motivo_descuento:'',
        observaciones:'PRUEBA · ' + actNom, salud:'No', contacto_emergencia:'2804000000',
        fecha_alta:'2025-01-01', fecha_baja:'', foto:''
      };
      shAl.appendRow(alH.map(function(h) { return alumno[h] !== undefined ? alumno[h] : ''; }));

      const ins = { id: Utilities.getUuid(), alumno_id: alId, actividad_id: actId, estado:'activo', fecha_inscripcion:'2025-01-01' };
      shIns.appendRow(insH.map(function(h) { return ins[h] !== undefined ? ins[h] : ''; }));
      insertados++;
    }
    Logger.log('10 alumnos de prueba → ' + actNom);
  }
  Logger.log('Total insertados: ' + insertados);
}
