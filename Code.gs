/** =====================================================================
 *  SISTEMA DE CRUCE DE INFORMACIÓN - BASE SIATH / BASE LSI / BASE DATA
 *  Policía Nacional de Colombia - Jefatura Nacional de Servicio de Policía
 *  API de Google Apps Script (Web App)
 *  ---------------------------------------------------------------------
 *  Funciones públicas (consumidas desde Index.html vía google.script.run):
 *    doGet()                      -> sirve la interfaz institucional
 *    apiEstado()                  -> diagnóstico del servicio
 *    apiListarHojas(fileId)       -> hojas disponibles de un xlsx subido
 *    apiSubirArchivo(nombre, b64) -> carga un xlsx/csv a la carpeta temporal
 *    apiAnalizar(params)          -> cruce SIATH/LSI/DATA + reporte
 *    apiCruzarNovedades(params)   -> valida novedades contra las 3 bases
 *    apiGenerarWordReporte(id)    -> .docx oficial del reporte
 *    apiLimpiarTemporal()         -> borra la carpeta temporal
 * ===================================================================== */

/* ======================= CONFIGURACIÓN =============================== */

var CFG = {
  VERSION: 'JESEP-CRUCE-1.0.0',

  // Palabra usada para el filtro transversal de todas las bases.
  PALABRA_FILTRO: 'JESEP',

  // Marcador de "novedades adicionales de nómina" (conteo solicitado).
  MARCADOR_ADICIONALES: 'NOVEDADES_ADICIONALES_NOMINA',

  // Carpetas temporales de Drive (se crean solas la primera vez).
  CARPETA_TEMP: 'JESEP_CRUCE_TMP',

  // Hoja preferida por base (vacío = detección automática).
  HOJA_SIATH_PREF: 'JESEP',
  HOJA_LSI_PREF: '',
  HOJA_DATA_PREF: '',

  // Límites de seguridad (evitan agotar memoria/hoja de cálculo).
  MAX_BYTES_PKG: 90 * 1024 * 1024,   // tamaño máximo del .xlsx
  MAX_BYTES_HOJA: 34 * 1024 * 1024,  // tamaño máximo del XML de una hoja
  MAX_FILAS: 200000,                 // filas máximas leídas por hoja
  MAX_DETALLE: 3000,                 // filas de detalle máxima por concepto
  MAX_NOVEDADES: 60000,

  // Número de páginas del Word (se recalcula al construir el documento).
  PAGINAS_ESTIMADAS: '1'
};

/* ======================= UTILIDADES DE TEXTO ======================== */

/** Quita tildes y lleva a mayúsculas. */
function norm_(v) {
  if (v === null || v === undefined) return '';
  var s = String(v);
  if (s === '' || s === 'null' || s === 'undefined') return '';
  s = s.toUpperCase();
  s = s.replace(/[ÁÀÄÂ]/g, 'A').replace(/[ÉÈËÊ]/g, 'E')
       .replace(/[ÍÌÏÎ]/g, 'I').replace(/[ÓÒÖÔ]/g, 'O')
       .replace(/[ÚÙÜÛ]/g, 'U').replace(/Ñ/g, 'N').replace(/Ç/g, 'C');
  return s;
}

/** Normaliza un encabezado: solo A-Z, 0-9 y guion bajo. */
function normCab_(v) {
  return norm_(v).replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

/** Normaliza un valor de texto (trim + colapsa espacios). */
function txt_(v) {
  if (v === null || v === undefined) return '';
  var s = String(v);
  if (s === 'null' || s === 'undefined' || s === 'NaN') return '';
  return s.replace(/\s+/g, ' ').trim();
}

/**
 * Normaliza una identificación: deja solo dígitos.
 * Conserva el cero inicial porque en la Policía es significativo.
 */
function idNorm_(v) {
  var s = txt_(v);
  if (!s) return '';
  if (/^\d+(\.0+)?$/.test(s)) s = s.replace(/\.0+$/, '');
  var d = s.replace(/\D/g, '');
  return d;
}

/** Convierte a número devolviendo 0 si no es numérico. */
function num_(v) {
  if (v === null || v === undefined) return 0;
  var s = String(v).replace(/\s/g, '').replace(/\.(?=\d{3}\b)/g, '');
  s = s.replace(/[^0-9,.\-]/g, '').replace(/,/g, '.');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/** Formatea un número como entero con separador de miles. */
function fmtInt_(n) {
  var neg = n < 0;
  n = Math.round(Math.abs(n));
  var s = String(n), out = '', c = 0;
  for (var i = s.length - 1; i >= 0; i--) {
    out = s[i] + out;
    c++;
    if (c % 3 === 0 && i > 0) out = '.' + out;
  }
  return (neg ? '-' : '') + out;
}

/** Formatea un número como moneda sin decimales. */
function fmtMon_(n) {
  var neg = n < 0;
  var s = fmtInt_(Math.abs(n));
  return (neg ? '-$ ' : '$ ') + s;
}

/** Divide "APELLIDO1 APELLIDO2 NOMBRE1 NOMBRE2" en sus dos partes. */
function separarNombre_(completo) {
  var s = txt_(completo);
  if (!s) return { nombres: '', apellidos: '' };
  var p = s.split(' ');
  if (p.length <= 1) return { nombres: s, apellidos: '' };
  var nNombres = p.length >= 4 ? 2 : 1;
  if (p.length === 2) nNombres = 1;
  if (p.length === 3) nNombres = 2;
  return {
    apellidos: p.slice(0, p.length - nNombres).join(' '),
    nombres: p.slice(p.length - nNombres).join(' ')
  };
}

/** Escapa texto para insertarlo en XML/HTML. */
function escXml_(v) {
  return txt_(v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

/** Escapa para HTML. */
function escHtml_(v) {
  return escXml_(v).replace(/\r?\n/g, '<br>');
}

/** Decodifica entidades XML básicas. */
function decXml_(s) {
  if (!s) return '';
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>')
          .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
          .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(n); })
          .replace(/&#x([0-9a-fA-F]+);/g, function (_, h) { return String.fromCharCode(parseInt(h, 16)); })
          .replace(/&amp;/g, '&');
}

/** ¿El valor indica "verdadero"? */
function esVerdad_(v) {
  var s = norm_(txt_(v));
  if (!s) return false;
  return ['SI', 'S', '1', 'X', 'VERDADERO', 'TRUE', 'V', 'Y', 'YES'].indexOf(s) >= 0;
}

/** Genera id de sesión/reporte. */
function nuevoId_(prefijo) {
  return prefijo + '-' + Utilities.getUuid().slice(0, 13);
}

/* ======================= LECTURA DE ZIP / XLSX ====================== */

/** Devuelve la carpeta temporal de Drive, creándola si hace falta. */
function carpetaTemp_(crear) {
  var it = DriveApp.getFoldersByName(CFG.CARPETA_TEMP);
  if (it.hasNext()) return it.next();
  if (!crear) return null;
  return DriveApp.createFolder(CFG.CARPETA_TEMP);
}

/** Elimina la carpeta temporal completa. */
function apiLimpiarTemporal() {
  var f = carpetaTemp_(false);
  if (!f) return { ok: true, eliminados: 0 };
  var n = f.getFiles().size();
  var archivos = f.getFiles();
  while (archivos.hasNext()) {
    try { archivos.next().setTrashed(true); } catch (e) {}
  }
  try { f.setTrashed(true); } catch (e2) {}
  return { ok: true, eliminados: n };
}

/** Recibe un archivo del cliente (base64) y lo guarda en la carpeta temporal. */
function apiSubirArchivo(nombre, base64, hojasPreferidas) {
  try {
    if (!base64) return { ok: false, error: 'No se recibió contenido del archivo.' };
    var limpio = String(base64).replace(/^data:[^;]+;base64,/, '');
    var bytes = Utilities.base64Decode(limpio);
    if (!bytes || !bytes.length) return { ok: false, error: 'El archivo llegó vacío.' };
    if (bytes.length > CFG.MAX_BYTES_PKG) {
      return { ok: false, error: 'El archivo supera el tamaño máximo permitido (' +
        Math.round(CFG.MAX_BYTES_PKG / 1048576) + ' MB).' };
    }
    var blob = Utilities.newBlob(bytes, MimeType.ZIP, nombre || 'archivo.xlsx');
    var f = carpetaTemp_(true).createFile(blob);
    return { ok: true, fileId: f.getId(), nombre: nombre, bytes: bytes.length,
             hojasPreferidas: hojasPreferidas || null };
  } catch (e) {
    return { ok: false, error: 'Error al guardar el archivo: ' + e.message };
  }
}

/** Descomprime un .xlsx y devuelve un mapa nombre->blob. */
function abrirPaquete_(blob) {
  var intentos = [];
  intentos.push(function () { return Utilities.unzip(blob); });
  intentos.push(function () {
    return Utilities.unzip(Utilities.newBlob(blob.getBytes(), MimeType.ZIP, 'p.xlsx'));
  });
  intentos.push(function () { return Utilities.unzip(blob.getBytes()); });

  var ultimo = null;
  for (var i = 0; i < intentos.length; i++) {
    try {
      var partes = intentos[i]();
      if (partes && partes.length) {
        var mapa = {};
        for (var j = 0; j < partes.length; j++) mapa[partes[j].getName()] = partes[j];
        return mapa;
      }
    } catch (e) { ultimo = e; }
  }
  throw new Error('No se pudo leer el archivo .xlsx. ' +
    (ultimo ? ultimo.message : 'Verifique que no esté corrupto ni protegido con contraseña.'));
}

/** Lee xl/workbook.xml y devuelve [{nombre, ruta, id}]. */
function listarHojasDePaquete_(paquete) {
  var wb = paquete['xl/workbook.xml'];
  if (!wb) return [];
  var xml = wb.getDataAsString('UTF-8');

  // rId -> target
  var rels = {};
  var rb = paquete['xl/_rels/workbook.xml.rels'];
  if (rb) {
    var rx = rb.getDataAsString('UTF-8');
    var reR = /<Relationship\b([^>]*)\/>/g, m;
    while ((m = reR.exec(rx))) {
      var at = {};
      var reA = /(\w+)="([^"]*)"/g, a;
      while ((a = reA.exec(m[1]))) at[a[1]] = a[2];
      if (at.Id && at.Target) rels[at.Id] = at.Target.replace(/^\/?xl\//, '').replace(/^\//, '');
    }
  }

  var hojas = [], reS = /<sheet\b([^>]*)\/>/g;
  while ((m = reS.exec(xml))) {
    var at2 = {}, reA2 = /([\w:]+)="([^"]*)"/g, b;
    while ((b = reA2.exec(m[1]))) at2[b[1]] = b[2];
    var rid = at2['r:id'] || at2.id;
    var ruta = rels[rid] ? ('xl/' + rels[rid]) : '';
    hojas.push({ nombre: decXml_(at2.name || ''), ruta: ruta, id: rid });
  }
  return hojas;
}

/** Elige la hoja: coincidencia exacta > nombre con la palabra de filtro > primera. */
function elegirHoja_(paquete, hojas, preferencia) {
  if (!hojas || !hojas.length) throw new Error('El libro no contiene hojas.');

  var pref = normCab_(preferencia || '');
  if (pref) {
    for (var i = 0; i < hojas.length; i++) {
      if (normCab_(hojas[i].nombre) === pref) {
        return { nombre: hojas[i].nombre, ruta: hojas[i].ruta };
      }
    }
  }
  // Una hoja cuyo nombre contenga la palabra de filtro (p. ej. "JESEP") ya viene
  // filtrada y suele ser la ligera: se prefiere para no cargar la hoja completa.
  for (var k = 0; k < hojas.length; k++) {
    if (norm_(hojas[k].nombre).indexOf(norm_(CFG.PALABRA_FILTRO)) >= 0) {
      return { nombre: hojas[k].nombre, ruta: hojas[k].ruta };
    }
  }
  return { nombre: hojas[0].nombre, ruta: hojas[0].ruta };
}

/** Lee sharedStrings.xml -> arreglo de textos. */
function leerSharedStrings_(paquete) {
  var b = paquete['xl/sharedStrings.xml'];
  if (!b) return [];
  var xml = b.getDataAsString('UTF-8');
  var out = [], i = 0;
  var re = /<si\b[^>]*\/>|<si\b[^>]*>([\s\S]*?)<\/si>/g, m;
  while ((m = re.exec(xml))) {
    if (m[1] === undefined) { out.push(''); i++; continue; }
    var frag = m[1], res = '', re2 = /<t\b[^>]*\/>|<t\b[^>]*>([\s\S]*?)<\/t>/g, t;
    while ((t = re2.exec(frag))) {
      res += (t[1] === undefined) ? '' : decXml_(t[1]);
    }
    out.push(res);
    i++;
  }
  return out;
}

/** Convierte "BC12" -> 55 (índice de columna base 1). */
function colIdx_(ref) {
  var letras = String(ref).replace(/[^A-Za-z]/g, '').toUpperCase();
  var n = 0;
  for (var i = 0; i < letras.length; i++) n = n * 26 + (letras.charCodeAt(i) - 64);
  return n;
}

/**
 * Lee una hoja .xlsx y devuelve {cabeceras:[], filas:[[]], celdas:[[]], truncado, bytes}
 * Se procesa fila por fila para no retener el XML completo en memoria.
 */
function leerHojaXlsx_(paquete, hojaElegida) {
  var b = paquete[hojaElegida.ruta];
  if (!b) throw new Error('No se encontró el XML de la hoja "' + hojaElegida.nombre + '".');

  var bytes = 0;
  try {
    bytes = b.getBytes().length;
  } catch (eTam) {
    throw new Error('No se pudo dimensionar la hoja "' + hojaElegida.nombre +
      '". Seleccione una hoja más liviana, por ejemplo la hoja "' + CFG.PALABRA_FILTRO + '".');
  }
  if (bytes > CFG.MAX_BYTES_HOJA) {
    throw new Error('La hoja "' + hojaElegida.nombre + '" ocupa ' +
      Math.round(bytes / 1048576) + ' MB y supera el límite de ' +
      Math.round(CFG.MAX_BYTES_HOJA / 1048576) + ' MB. ' +
      'Seleccione una hoja más liviana (por ejemplo la hoja "' + CFG.PALABRA_FILTRO + '").');
  }

  var ss = leerSharedStrings_(paquete);

  var xml;
  try {
    xml = b.getDataAsString('UTF-8');
  } catch (eTxt) {
    throw new Error('No se pudo leer la hoja "' + hojaElegida.nombre +
      '" (' + Math.round(bytes / 1048576) + ' MB): se agotó la memoria del servidor. ' +
      'Seleccione la hoja "' + CFG.PALABRA_FILTRO + '" o una hoja con menos columnas.');
  }

  var filas = [];
  var truncado = false;
  var reFila = /<row\b[^>]*\/>|<row\b[^>]*>([\s\S]*?)<\/row>/g, mF;

  while ((mF = reFila.exec(xml))) {
    if (mF[1] === undefined) continue;
    var frag = mF[1];

    var fila = [], idx = 1, vFila = [];
    var reC = /<c\b([^>]*)\/>|<c\b([^>]*)>([\s\S]*?)<\/c>/g, mC;
    while ((mC = reC.exec(frag))) {
      var attrs = mC[1] !== undefined ? mC[1] : (mC[2] || '');
      var cuerpo = mC[3] !== undefined ? mC[3] : '';

      var ref = '';
      var reR = /\br="([^"]*)"/.exec(attrs);
      if (reR) ref = reR[1];
      var col = ref ? colIdx_(ref) : idx;
      idx = col + 1;

      var tipo = '';
      var reT = /\bt="([^"]*)"/.exec(attrs);
      if (reT) tipo = reT[1];

      var valor = '';
      if (tipo === 'inlineStr') {
        var reI = /<t\b[^>]*>([\s\S]*?)<\/t>/g, mI, acc = '';
        while ((mI = reI.exec(cuerpo))) acc += decXml_(mI[1]);
        valor = acc;
      } else {
        var reV = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(cuerpo);
        if (reV) {
          var crudo = decXml_(reV[1]);
          if (tipo === 's') {
            var n = parseInt(crudo, 10);
            valor = (isNaN(n) || ss[n] === undefined) ? '' : ss[n];
          } else {
            valor = crudo;
          }
        }
      }
      fila[col - 1] = valor;
    }
    // Rellena huecos y recorta
    var ancho = 0;
    for (var q = 0; q < fila.length; q++) if (fila[q] !== undefined && fila[q] !== '') ancho = q + 1;
    fila.length = ancho;
    for (var z = 0; z < ancho; z++) if (fila[z] === undefined) fila[z] = '';

    if (!fila.join('')) continue;
    filas.push(fila);
    if (filas.length >= CFG.MAX_FILAS) { truncado = true; break; }
  }

  if (!filas.length) return { cabeceras: [], filas: [], truncado: truncado, bytes: bytes };

  var cab = filas[0].map(function (h) { return txt_(h); });
  var cuerpo = filas.slice(1);

  // Detecta si la primera fila es un título y la segunda es el encabezado real.
  var cabNorm = cab.map(normCab_);
  var tieneEnc = function (c) {
    return c === 'IDENTIFICACION' || c === 'IDENTIFICACION_NUEVA' ||
           c === 'DESCRIPCION' || c === 'NOMBRES' || c === 'NOMBRE' ||
           c === 'UNIDAD' || c === 'APELLIDOS' || c === 'CATEGORIA';
  };
  if (cuerpo.length && !tieneEnc(cabNorm[0]) && cuerpo.length) {
    var sig = cuerpo[0].map(normCab_);
    if (sig.some(tieneEnc)) {
      cab = cuerpo[0].map(function (h) { return txt_(h); });
      cuerpo = cuerpo.slice(1);
    }
  }

  return { cabeceras: cab, filas: cuerpo, truncado: truncado, bytes: bytes };
}

/* ======================= LECTURA CSV / TXT ========================== */

/** Parsea CSV/TSV respetando comillas. */
function parseDelimitado_(texto, sep) {
  var out = [], fila = [], val = '', enComillas = false;
  if (!sep) {
    sep = (texto.indexOf(';') >= 0 && texto.indexOf(',') < 0) ? ';' :
         (texto.indexOf('\t') >= 0 && texto.indexOf(',') < 0) ? '\t' : ',';
  }
  for (var i = 0; i < texto.length; i++) {
    var ch = texto[i];
    if (enComillas) {
      if (ch === '"') {
        if (texto[i + 1] === '"') { val += '"'; i++; }
        else enComillas = false;
      } else val += ch;
    } else {
      if (ch === '"') enComillas = true;
      else if (ch === sep) { fila.push(val); val = ''; }
      else if (ch === '\n') {
        fila.push(val); out.push(fila); fila = []; val = '';
        if (out.length >= CFG.MAX_NOVEDADES) break;
      } else if (ch === '\r') { /* ignorar */ }
      else val += ch;
    }
  }
  if (val.length || fila.length) { fila.push(val); out.push(fila); }
  var res = [];
  for (var k = 0; k < out.length; k++) {
    if (out[k].join('').trim() !== '') res.push(out[k]);
  }
  return { separador: sep, filas: res };
}

/** Punto de entrada unificado: xlsx o csv -> {cabeceras, filas}. */
function leerTabla_(blob, nombre, preferencia) {
  var n = norm_(nombre || '');
  if (/\.CSV$/.test(n) || blob.getContentType() === 'text/csv') {
    var t = parseDelimitado_(blob.getDataAsString('UTF-8'));
    return {
      cabeceras: (t.filas[0] || []).map(function (h) { return txt_(h); }),
      filas: t.filas.slice(1), truncado: false, bytes: blob.getBytes().length,
      hoja: 'CSV'
    };
  }
  var paquete = abrirPaquete_(blob);
  var hojas = listarHojasDePaquete_(paquete);
  if (!hojas.length) throw new Error('El archivo no parece ser un libro de Excel válido.');
  var elegida = elegirHoja_(paquete, hojas, preferencia);
  var r = leerHojaXlsx_(paquete, elegida);
  r.hoja = elegida.nombre;
  r.hojasDisponibles = hojas.map(function (h) { return h.nombre; });
  return r;
}

/** Devuelve las hojas de un archivo ya subido. */
function apiListarHojas(fileId) {
  try {
    var blob = DriveApp.getFileById(fileId).getBlob();
    if (/\.CSV$/i.test(blob.getName())) {
      return { ok: true, hojas: ['CSV'], tipo: 'csv' };
    }
    var paquete = abrirPaquete_(blob);
    var hojas = listarHojasDePaquete_(paquete);
    return { ok: true, hojas: hojas.map(function (h) { return { nombre: h.nombre }; }), tipo: 'xlsx' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/* ======================= MAPEO DE COLUMNAS ========================== */

var ALIAS = {
  /* ---- BASE SIATH ---- */
  SIATH: {
    id:          ['IDENTIFICACION', 'IDENTIFICACION_NUEVA', 'ID', 'CEDULA', 'NUMERO_IDENTIFICACION', 'NUI'],
    nombre:      ['NOMBRE', 'NOMBRES', 'NOMBRE_COMPLETO', 'NOMBRE_Y_APELLIDOS', 'APELLIDOS_Y_NOMBRES', 'NOMBRE_DEL_FUNCIONARIO'],
    unidad:      ['UNIDAD', 'FUERZA', 'UNDE_FUERZA', 'DEPENDENCIA', 'UNIDAD_DE_TRABAJO'],
    grado:       ['GRADO', 'GRAD_ALFABETICO', 'GRADO_ALFABETICO'],
    estado:      ['ESTADO_EMPLEADO', 'ESTADO', 'SITUACION'],
    devengado:   ['TOTAL_DEVENGADO', 'DEVENGADO', 'VALOR_DEVENGADO'],
    neto:        ['NETO', 'VALOR_NETO', 'NETO_PAGADO'],
    descripcion: ['DESCRIPCION_DEPENDENCIA', 'DESCRIPCION', 'DEPENDENCIA'],
    cargo:       ['CARGO', 'CARGO_PRESENTE', 'DENOMINACION_CARGO'],
    ordenPublico:['ORDEN_PUBLICO_AUTOMATICO', 'ORDEN_PUBLICO', 'OPA'],
    embargos:    ['EMBARGOS_APLICADOS', 'EMBARGO_APLICADO', 'EMBARGOS'],
    embFamilia:  ['EMB_FAMILIA', 'EMBARGO_FAMILIA', 'EMB_FAMILIAR'],
    embEjecutivo:['EMB_EJECUTIVO', 'EMBARGOS_EJECUTIVOS'],
    embRemanente:['EMB_REMANENTES', 'EMBARGOS_REMANENTES'],
    otrosDocs:   ['OTROS_DCTOS_PENDIENTES', 'OTROS_DOCUMENTOS'],
    menos50:     ['DEVENGA_MENOS_50_SALARIO', 'MENOS_50_SALARIO']
  },
  /* ---- BASE LSI (novedades generales) ---- */
  LSI: {
    id:          ['IDENTIFICACION', 'IDENTIFICACION_NUEVA', 'ID', 'CEDULA', 'NUMERO_IDENTIFICACION'],
    nombres:     ['NOMBRES', 'NOMBRE'],
    apellidos:   ['APELLIDOS', 'APELLIDO', 'NOMBRES (#1)'],
    descripcion: ['DESCRIPCION', 'DESCRIPCION_NOVEDAD', 'TIPO_NOVEDAD', 'NOVEDAD'],
    claseNov:    ['NOVE_ID_CLASE_NOVEDAD', 'CLASE_NOVEDAD', 'ID_CLASE_NOVEDAD'],
    anioMes:     ['ANO_MES', 'ANOMES', 'PERIODO', 'ANO_MES_NOVEDAD'],
    fisica:      ['FISICA_ACTUAL', 'FISICA'],
    unidad:      ['UNDE_FUERZA', 'FUERZA', 'UNIDAD'],
    disp:        ['NUMERO_DISPOSICION', 'DISP_ID_DISPOSICION', 'DISPOSICION'],
    fechaDisp:   ['FECHA_DISPOSICION', 'FECHA_DISPOSICION'],
    fechaFis:    ['FECHA_FISCAL'],
    diasLic:     ['DIAS_LICENCIA', 'DIAS_NO_LABORAR', 'NUM_DIAS_NO_LABORAR'],
    diasVac:     ['DIAS_VACACIONES', 'DIAS_PEND_VACACIONES', 'DIAS_VACACIONES_DEROGA'],
    diasCom:     ['NUMERO_DIAS_COMISION', 'NUMERO_DIAS_COMISION'],
    tipoComExt:  ['TIPO_COMISION_EXTERIOR', 'TIPO_COMISION', 'LUGAR_COMISION'],
    proc:        ['NUMERO_PROCESO', 'F_PROCESO_INTERFAZ'],
    cond:        ['CONDICION_REGISTRO', 'CONDICION'],
    tipoInc:     ['TIPO_INCAPACIDAD', 'NUMERO_INCAPACIDAD'],
    estado:      ['DESTADO_INTERFAZ', 'ESTADO'],
    adic:        ['NOVEDADES_ADICIONALES_NOMINA', 'ADICIONALES_NOMINA', 'NOVEDAD_ADICIONAL', 'ADICIONAL'],

    // --- Datos de traslado: unidad de origen y de destino ---
    consecutivoActual: ['EMPL_CONSECUTIVO', 'UNDE_CONSECUTIVO_EMPLEADO', 'CONSECUTIVO_EMPLEADO'],
    fuerzaActual:     ['EMPL_UNDE_FUERZA', 'FUERZA_EMPLEADO'],
    destinoNomina:    ['UNDE_CONSECUTIVO_TRASLA_NOMI', 'UNDE_CONSEC_TRASLA_NOMI', 'CONSECUTIVO_TRASLADO_NOMI'],
    destinoLabor:     ['UNDE_CONSECUTIVO_TRASLA_LABOR', 'UNDE_CONSEC_TRASLA_LABOR', 'CONSECUTIVO_TRASLADO_LABOR'],
    fuerzaDestino:    ['UNDE_FUERZA_TRASLA_LABOR', 'UNDE_FUERZA_TRASLA_NOMI', 'FUERZA_TRASLADO'],
    creadoPor:        ['CREADO_POR', 'USUARIO_CREACION']
  },
  /* ---- BASE DATA (reporte de incapacidades) ---- */
  DATA: {
    id:          ['IDENTIFICACION', 'ID', 'CEDULA', 'NUMERO_IDENTIFICACION'],
    nombres:     ['NOMBRES', 'NOMBRE'],
    apellidos:   ['APELLIDOS', 'APELLIDO'],
    categoria:   ['CATEGORIA', 'DESCRIPCION', 'TIPO_EXCUSA', 'CONCEPTO'],
    sigla:       ['SIGLA_FISICA', 'SIGLA', 'ALFABETICO'],
    dias:        ['NUMERO_DIAS', 'DIAS', 'DIAS_LICENCIA'],
    fechaInicio: ['ID_FECHFISC_AUS_LABORAL', 'FECHA_INICIO', 'FECHA_FISCAL'],
    fechaFin:    ['FECHA_FINAL', 'FECHA_TERMINO'],
    tipoExcusa:  ['TIPO_EXCUSA'],
    descripcion: ['DESCRIPCION_DEPENDENCIA', 'DESCRIPCION'],
    cargo:       ['CARGO'],
    disposicion: ['NUMERO_DISPOSICION', 'DISPOSICION'],
    ultimaOp:    ['ULTIMA_OPERACION']
  },
  /* ---- NOVEDADES SUBIDAS (validación) ---- */
  NOV: {
    id:          ['IDENTIFICACION', 'ID', 'CEDULA', 'NUMERO_IDENTIFICACION', 'IDENTIFICACION_NUEVA'],
    nombres:     ['NOMBRES', 'NOMBRE'],
    apellidos:   ['APELLIDOS', 'APELLIDO'],
    descripcion: ['DESCRIPCION', 'NOVEDAD', 'TIPO_NOVEDAD', 'DESCRIPCION_NOVEDAD'],
    anioMes:     ['ANO_MES', 'PERIODO'],
    fecha:       ['FECHA', 'FECHA_DISPOSICION', 'FECHA_REGISTRO', 'FECHA_NOVEDAD'],
    referencia:  ['NUMERO_PROCESO', 'NUMERO_DISPOSICION', 'RADICADO', 'REFERENCIA', 'CONSECUTIVO'],
    claseNov:    ['NOVE_ID_CLASE_NOVEDAD', 'CLASE_NOVEDAD'],
    unidad:      ['UNIDAD', 'FUERZA', 'UNDE_FUERZA', 'DEPENDENCIA'],
    cargo:       ['CARGO']
  }
};

/**
 * Resuelve índices de columna a partir de los encabezados y una tabla de alias.
 * Devuelve {clave: indiceBase0}
 */
function mapearColumnas_(cabeceras, alias) {
  var mapa = {};
  var norm = cabeceras.map(normCab_);
  for (var clave in alias) {
    if (!alias.hasOwnProperty(clave)) continue;
    mapa[clave] = -1;
    for (var i = 0; i < norm.length; i++) {
      if (norm[i] === alias[clave][0]) { mapa[clave] = i; break; }
    }
  }
  // Segunda pasada: coincidencia por "contiene" en orden de alias
  for (var clave2 in alias) {
    if (!alias.hasOwnProperty(clave2)) continue;
    if (mapa[clave2] >= 0) continue;
    for (var a = 0; a < alias[clave2].length; a++) {
      var objetivo = alias[clave2][a];
      for (var j = 0; j < norm.length; j++) {
        if (norm[j] === objetivo || norm[j].indexOf(objetivo) >= 0 || objetivo.indexOf(norm[j]) >= 0) {
          if (norm[j].length >= 3) { mapa[clave2] = j; break; }
        }
      }
      if (mapa[clave2] >= 0) break;
    }
  }
  return mapa;
}

/** Valor de una fila por el mapa de columnas. */
function val_(fila, mapa, clave) {
  var i = mapa[clave];
  if (i === undefined || i < 0 || i >= fila.length) return '';
  return txt_(fila[i]);
}

/** Detecta si alguna celda de la fila contiene el texto buscado. */
function filaContiene_(fila, palabra) {
  var p = norm_(palabra);
  if (!p) return false;
  for (var i = 0; i < fila.length; i++) {
    if (fila[i] !== undefined && fila[i] !== '' && norm_(fila[i]).indexOf(p) >= 0) return true;
  }
  return false;
}

/* ======================= MODELADO DE LAS BASES ====================== */

/** Normaliza una fila de la BASE SIATH. */
function leerSiath_(tabla, config) {
  var m = mapearColumnas_(tabla.cabeceras, ALIAS.SIATH);
  if (m.id < 0) throw new Error('BASE SIATH: no se encontró la columna IDENTIFICACION. Encabezados: ' +
    tabla.cabeceras.slice(0, 25).join(', '));

  var palabra = norm_(config.palabraFiltro || CFG.PALABRA_FILTRO);
  var soloPalabra = config.filtrarPorPalabra !== false;
  var salida = [], stats = { leidas: tabla.filas.length, conPalabra: 0, fueraDeFiltro: 0 };

  for (var i = 0; i < tabla.filas.length; i++) {
    var f = tabla.filas[i];
    var id = idNorm_(val_(f, m, 'id'));
    if (!id) continue;

    var coincidePalabra = filaContiene_(f, palabra);
    if (coincidePalabra) stats.conPalabra++;

    // Filtro: la persona debe pertenecer a la unidad/palabra buscada.
    if (soloPalabra && !coincidePalabra) { stats.fueraDeFiltro++; continue; }

    var nombreCompleto = val_(f, m, 'nombre');
    var partes = separarNombre_(nombreCompleto);

    salida.push({
      id: id,
      nombreCompleto: nombreCompleto,
      nombres: partes.nombres,
      apellidos: partes.apellidos,
      unidad: val_(f, m, 'unidad'),
      grado: val_(f, m, 'grado'),
      estado: val_(f, m, 'estado'),
      devengado: num_(val_(f, m, 'devengado')),
      neto: num_(val_(f, m, 'neto')),
      descripcion: val_(f, m, 'descripcion'),
      cargo: val_(f, m, 'cargo'),
      ordenPublico: txt_(val_(f, m, 'ordenPublico')),
      embargos: num_(val_(f, m, 'embargos')),
      embFamilia: num_(val_(f, m, 'embFamilia')),
      embEjecutivo: num_(val_(f, m, 'embEjecutivo')),
      embRemanente: num_(val_(f, m, 'embRemanente')),
      otrosDocs: num_(val_(f, m, 'otrosDocs')),
      menos50: txt_(val_(f, m, 'menos50')),
      tienePalabra: coincidePalabra
    });
  }

  // Red de seguridad: si la columna buscada no marcaba nada (hoja completa sin
  // la palabra), se reintenta sin el filtro para no entregar un padrón vacío.
  if (soloPalabra && salida.length === 0 && stats.leidas > 0) {
    stats.aviso = 'Ninguna fila de BASE SIATH contiene la palabra "' + palabra +
      '". Se reprocesó la hoja sin ese filtro: use la hoja "' + CFG.PALABRA_FILTRO +
      '" para obtener el padrón filtrado.';
    salida = releerSiathSinFiltro_(tabla, m, stats.leidas);
  }
  return { filas: salida, mapa: m, stats: stats };
}

/** Reconstruye las filas SIATH sin aplicar el filtro por palabra. */
function releerSiathSinFiltro_(tabla, m, limite) {
  var salida = [];
  var n = Math.min(limite, tabla.filas.length);
  for (var i = 0; i < n; i++) {
    var f = tabla.filas[i];
    var id = idNorm_(val_(f, m, 'id'));
    if (!id) continue;
    var nombreCompleto = val_(f, m, 'nombre');
    var partes = separarNombre_(nombreCompleto);
    salida.push({
      id: id, nombreCompleto: nombreCompleto,
      nombres: partes.nombres, apellidos: partes.apellidos,
      unidad: val_(f, m, 'unidad'), grado: val_(f, m, 'grado'),
      estado: val_(f, m, 'estado'), devengado: num_(val_(f, m, 'devengado')),
      neto: num_(val_(f, m, 'neto')), descripcion: val_(f, m, 'descripcion'),
      cargo: val_(f, m, 'cargo'), ordenPublico: txt_(val_(f, m, 'ordenPublico')),
      embargos: num_(val_(f, m, 'embargos')), embFamilia: num_(val_(f, m, 'embFamilia')),
      embEjecutivo: num_(val_(f, m, 'embEjecutivo')), embRemanente: num_(val_(f, m, 'embRemanente')),
      otrosDocs: num_(val_(f, m, 'otrosDocs')), menos50: txt_(val_(f, m, 'menos50')),
      tienePalabra: false
    });
  }
  return salida;
}

/** Normaliza una fila de la BASE LSI (novedades generales). */
function leerLsi_(tabla, config) {
  var m = mapearColumnas_(tabla.cabeceras, ALIAS.LSI);
  if (m.id < 0) throw new Error('BASE LSI: no se encontró la columna IDENTIFICACION. Encabezados: ' +
    tabla.cabeceras.slice(0, 25).join(', '));

  var salida = [], stats = { leidas: tabla.filas.length, conPalabra: 0 };
  var periodo = '';

  for (var i = 0; i < tabla.filas.length; i++) {
    var f = tabla.filas[i];
    var id = idNorm_(val_(f, m, 'id'));
    if (!id) continue;
    if (filaContiene_(f, config.palabraFiltro || CFG.PALABRA_FILTRO)) stats.conPalabra++;
    if (!periodo && m.anioMes >= 0) periodo = val_(f, m, 'anioMes');

    var nombres = val_(f, m, 'nombres');
    var apellidos = val_(f, m, 'apellidos');
    if (!apellidos && !nombres) {
      var sp = separarNombre_(val_(f, m, 'nombre'));
      nombres = sp.nombres; apellidos = sp.apellidos;
    }

    salida.push({
      id: id,
      nombres: nombres,
      apellidos: apellidos,
      nombreCompleto: txt_([apellidos, nombres].filter(String).join(' ')),
      descripcion: val_(f, m, 'descripcion'),
      descripcionNorm: norm_(val_(f, m, 'descripcion')),
      clase: val_(f, m, 'claseNov'),
      anioMes: val_(f, m, 'anioMes'),
      fisica: val_(f, m, 'fisica'),
      unidad: val_(f, m, 'unidad'),
      disposicion: val_(f, m, 'disp'),
      fechaDisposicion: val_(f, m, 'fechaDisp'),
      fechaFiscal: val_(f, m, 'fechaFis'),
      diasLicencia: num_(val_(f, m, 'diasLic')),
      diasVacaciones: num_(val_(f, m, 'diasVac')),
      diasComision: num_(val_(f, m, 'diasCom')),
      comisionExterior: txt_(val_(f, m, 'tipoComExt')),
      proceso: val_(f, m, 'proc'),
      condicion: val_(f, m, 'cond'),
      tipoIncapacidad: val_(f, m, 'tipoInc'),
      estado: val_(f, m, 'estado'),
      adicional: esVerdad_(val_(f, m, 'adic')),
      columnaAdicional: txt_(val_(f, m, 'adic')),

      // Datos de traslado (origen vs destino)
      fisicaActual: txt_(val_(f, m, 'fisica')),
      consecutivoActual: txt_(val_(f, m, 'consecutivoActual')),
      fuerzaActual: txt_(val_(f, m, 'fuerzaActual')),
      destinoNomina: txt_(val_(f, m, 'destinoNomina')),
      destinoLabor: txt_(val_(f, m, 'destinoLabor')),
      fuerzaDestino: txt_(val_(f, m, 'fuerzaDestino')),
      creadoPor: txt_(val_(f, m, 'creadoPor'))
    });
  }
  return { filas: salida, mapa: m, stats: stats, periodo: periodo };
}

/** Normaliza una fila de la BASE DATA (reporte de incapacidades). */
function leerData_(tabla, config) {
  var m = mapearColumnas_(tabla.cabeceras, ALIAS.DATA);
  if (m.id < 0) throw new Error('BASE DATA: no se encontró la columna IDENTIFICACION. Encabezados: ' +
    tabla.cabeceras.slice(0, 25).join(', '));

  var salida = [], stats = { leidas: tabla.filas.length, conPalabra: 0 };
  for (var i = 0; i < tabla.filas.length; i++) {
    var f = tabla.filas[i];
    var id = idNorm_(val_(f, m, 'id'));
    if (!id) continue;
    if (filaContiene_(f, config.palabraFiltro || CFG.PALABRA_FILTRO)) stats.conPalabra++;

    salida.push({
      id: id,
      nombres: val_(f, m, 'nombres'),
      apellidos: val_(f, m, 'apellidos'),
      nombreCompleto: txt_([val_(f, m, 'apellidos'), val_(f, m, 'nombres')].filter(String).join(' ')),
      categoria: val_(f, m, 'categoria'),
      categoriaNorm: norm_(val_(f, m, 'categoria')),
      sigla: val_(f, m, 'sigla'),
      dias: num_(val_(f, m, 'dias')),
      fechaInicio: val_(f, m, 'fechaInicio'),
      fechaFin: val_(f, m, 'fechaFin'),
      tipoExcusa: val_(f, m, 'tipoExcusa'),
      descripcion: val_(f, m, 'descripcion'),
      cargo: val_(f, m, 'cargo'),
      disposicion: val_(f, m, 'disposicion'),
      ultimaOperacion: val_(f, m, 'ultimaOp')
    });
  }
  return { filas: salida, mapa: m, stats: stats };
}

/* ======================= MOTOR DE ANÁLISIS ========================== */

/** Conjuntos de identificadores. */
function construirIndices_(arr) {
  var s = {};
  for (var i = 0; i < arr.length; i++) s[arr[i].id] = true;
  return s;
}

/** Cuantifica un conjunto de filas de detalle. */
function cuantificar_(registros) {
  var ids = {};
  var dias = 0;
  for (var i = 0; i < registros.length; i++) {
    ids[registros[i].id] = true;
    dias += registros[i]._dias || 0;
  }
  var n = 0;
  for (var k in ids) if (ids.hasOwnProperty(k)) n++;
  return {
    registros: registros.length,
    funcionarios: n,
    dias: dias,
    detalle: registros.slice(0, CFG.MAX_DETALLE),
    truncado: registros.length > CFG.MAX_DETALLE
  };
}

/** Busca filas LSI cuya descripción cumple un predicado. */
function buscarEnLsi_(lsi, predicado) {
  var out = [];
  for (var i = 0; i < lsi.length; i++) {
    if (predicado(lsi[i])) out.push(lsi[i]);
  }
  return out;
}

/**
 * Detecta novedades "ADICIONALES DE NÓMINA".
 * Señales: columna que contenga ADICIONAL con valor verdadero, o
 * descripción que contenga el marcador configurado.
 */
function detectarAdicionales_(lsi, marcador) {
  var marcas = [];
  var out = [];
  var mk = norm_(marcador || CFG.MARCADOR_ADICIONALES);
  for (var i = 0; i < lsi.length; i++) {
    var r = lsi[i];
    var d = norm_(r.descripcion);
    var esAdicional = false, motivo = '';
    if (d.indexOf('ADICIONAL') >= 0) { esAdicional = true; motivo = 'DESCRIPCION contiene "ADICIONAL"'; }
    if (r.adicional) { esAdicional = true; motivo = motivo || 'Columna de adicionales = verdadero'; }
    if (mk && d.indexOf(mk) >= 0) { esAdicional = true; motivo = motivo || 'Marcador ' + marcador; }
    if (esAdicional) {
      r._dias = 0;
      r._motivo = motivo;
      out.push(r);
      marcas[motivo] = (marcas[motivo] || 0) + 1;
    }
  }
  return { filas: out, senales: marcas };
}

/** Analiza las 3 bases y arma el reporte completo. */
function analizarTablas_(tSiath, tLsi, tData, config) {
  config = config || {};
  var avisos = [];

  var siath = leerSiath_(tSiath, config);
  var lsiAll = leerLsi_(tLsi, config);
  var dataAll = leerData_(tData, config);

  var idxSiath = construirIndices_(siath.filas);
  var idsSiath = siath.filas.map(function (r) { return r.id; });
  var porIdSiath = {};
  for (var pi = 0; pi < siath.filas.length; pi++) porIdSiath[siath.filas[pi].id] = siath.filas[pi];

  // --- Filtro transversal por la palabra + pertenencia al padrón SIATH ---
  function filtrarPorPadrón_(lista) {
    var dentro = [], fuera = [];
    for (var i = 0; i < lista.length; i++) {
      if (idxSiath[lista[i].id]) dentro.push(lista[i]);
      else fuera.push(lista[i]);
    }
    return { dentro: dentro, fuera: fuera };
  }

  var fLsi = filtrarPorPadrón_(lsiAll.filas);
  var fData = filtrarPorPadrón_(dataAll.filas);
  var lsi = fLsi.dentro;
  var data = fData.dentro;

  avisos.push('BASE SIATH: ' + siath.stats.leidas + ' filas leidas, ' +
    siath.filas.length + ' funcionarios dentro del padron filtrado (' +
    (config.palabraFiltro || CFG.PALABRA_FILTRO) + '), ' + siath.stats.fueraDeFiltro + ' excluidas.');
  avisos.push('BASE LSI: ' + lsiAll.filas.length + ' filas leidas, ' + lsi.length +
    ' pertenecen al padron SIATH, ' + fLsi.fuera.length + ' no pertenecen.');
  avisos.push('BASE DATA: ' + dataAll.filas.length + ' filas leidas, ' + data.length +
    ' pertenecen al padron SIATH, ' + fData.fuera.length + ' no pertenecen.');
  if (siath.stats.aviso) avisos.push(siath.stats.aviso);
  if (tSiath.truncado) avisos.push('BASE SIATH: la hoja supero el limite de ' + CFG.MAX_FILAS + ' filas y fue truncada.');
  if (tLsi.truncado) avisos.push('BASE LSI: la hoja supero el limite de ' + CFG.MAX_FILAS + ' filas y fue truncada.');
  if (tData.truncado) avisos.push('BASE DATA: la hoja supero el limite de ' + CFG.MAX_FILAS + ' filas y fue truncada.');

  /* ---------------- Conteos solicitados ---------------- */

  // 1) Adicionales de nómina
  var ad = detectarAdicionales_(lsi, config.marcadorAdicional);
  for (i = 0; i < ad.filas.length; i++) ad.filas[i]._tipo = 'LSI';
  if (!ad.filas.length) {
    avisos.push('Adicionales de nómina: 0 registros. El archivo LSI no trae la columna "' +
      (config.marcadorAdicional || CFG.MARCADOR_ADICIONALES) +
      '" ni descripciones con la palabra ADICIONAL.');
  }

  // 2) Incapacidades
  var incLsi = buscarEnLsi_(lsi, function (r) { return r.descripcionNorm.indexOf('INCAPACIDAD') >= 0; });
  var incData = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i].categoriaNorm.indexOf('INCAPACIDAD') >= 0) incData.push(data[i]);
  }
  for (i = 0; i < incLsi.length; i++) { incLsi[i]._dias = incLsi[i].diasLicencia; incLsi[i]._tipo = 'LSI'; }
  for (i = 0; i < incData.length; i++) { incData[i]._dias = incData[i].dias; incData[i]._tipo = 'DATA'; }
  var incapacidad = combinar_(incLsi, incData);

  // 3) Disfrute de vacaciones
  var vacLsi = buscarEnLsi_(lsi, function (r) { return r.descripcionNorm.indexOf('VACACION') >= 0; });
  var vacData = [];
  for (i = 0; i < data.length; i++) {
    if (data[i].categoriaNorm.indexOf('VACACION') >= 0) vacData.push(data[i]);
  }
  for (i = 0; i < vacLsi.length; i++) { vacLsi[i]._dias = vacLsi[i].diasVacaciones; vacLsi[i]._tipo = 'LSI'; }
  for (i = 0; i < vacData.length; i++) { vacData[i]._dias = vacData[i].dias; vacData[i]._tipo = 'DATA'; }
  var vacaciones = combinar_(vacLsi, vacData);

  // 4) Embargos aplicados (SIATH)
  var embargos = [];
  for (i = 0; i < siath.filas.length; i++) {
    if (siath.filas[i].embargos > 0) {
      siath.filas[i]._dias = siath.filas[i].embargos;
      siath.filas[i]._tipo = 'SIATH';
      embargos.push(siath.filas[i]);
    }
  }

  // 5) Embargos a favor de familia (SIATH)
  var embFam = [];
  for (i = 0; i < siath.filas.length; i++) {
    if (siath.filas[i].embFamilia > 0) {
      siath.filas[i]._dias = siath.filas[i].embFamilia;
      siath.filas[i]._tipo = 'SIATH';
      embFam.push(siath.filas[i]);
    }
  }

  // 6) NO en orden público automático (SIATH)
  var noOrden = [];
  for (i = 0; i < siath.filas.length; i++) {
    var op = norm_(siath.filas[i].ordenPublico);
    if (op === 'NO' || op.indexOf('NO') === 0) {
      siath.filas[i]._tipo = 'SIATH';
      noOrden.push(siath.filas[i]);
    }
  }

  // 7) Comisiones en el exterior (LSI)
  var comExt = buscarEnLsi_(lsi, function (r) {
    return r.descripcionNorm.indexOf('EXTERIOR') >= 0 || norm_(r.comisionExterior).indexOf('EXTERIOR') >= 0;
  });
  for (i = 0; i < comExt.length; i++) { comExt[i]._dias = comExt[i].diasComision; comExt[i]._tipo = 'LSI'; }

  /* ---------------- Detalle por concepto ---------------- */

  function grupo_(clave, etiqueta, universo, predicado,opts) {
    opts = opts || {};
    var filas = buscarEnLsi_(universo, predicado);
    var base = [];
    for (var j = 0; j < filas.length; j++) base.push(filas[j]);
    var det = base.map(function (r) {
      return {
        id: r.id,
        nombres: r.nombres,
        apellidos: r.apellidos,
        nombreCompleto: r.nombreCompleto,
        descripcion: r.descripcion,
        clase: r.clase,
        anioMes: r.anioMes,
        dias: r._dias || 0,
        diasVacaciones: r.diasVacaciones,
        diasLicencia: r.diasLicencia,
        diasComision: r.diasComision,
        comisionExterior: r.comisionExterior,
        disposicion: r.disposicion,
        fechaDisposicion: r.fechaDisposicion,
        proceso: r.proceso,
        categoria: r.categoria,
        fisica: r.fisica,
        unidad: r.unidad,
        grado: r.grado,
        cargo: r.cargo,
        descripcionDep: r.descripcion || ''
      };
    });
    var q = cuantificar_(base);
    q.clave = clave;
    q.etiqueta = etiqueta;
    q.detalle = det.slice(0, CFG.MAX_DETALLE);
    if (opts.nota) q.nota = opts.nota;
    return q;
  }

  // Bloque 1 - Subsidios y bonificaciones (identificar quién)
  var gruposSubsidio = [];
  var bloquesSubsidio = [
    { clave: 'DISMINUCION_SUBSIDIO', etiqueta: 'DISMINUCION DEL SUBSIDIO FAMILIAR',
      pred: function (r) { return r.descripcionNorm.indexOf('DISMINUCION DEL SUBSIDIO FAMILIAR') >= 0; } },
    { clave: 'EXTINCION_SUBSIDIO', etiqueta: 'EXTINCION DEL SUBSIDIO FAMILIAR',
      pred: function (r) { return r.descripcionNorm.indexOf('EXTINCION DEL SUBSIDIO FAMILIAR') >= 0; } },
    { clave: 'EXTINCION_BONIF', etiqueta: 'EXTINCION DEL DE LA BONIFICACION PARA LA ASISTENCIA FAMILIAR',
      pred: function (r) { return r.descripcionNorm.indexOf('EXTINCION DEL DE LA BONIFICACION PARA LA ASISTENCIA FAMILIAR') >= 0; } },
    { clave: 'BONIF_ASIST_FAM_ADIC', etiqueta: 'BONIFICACION PARA LA ASISTENCIA FAMILIAR (incluye NOVEDADES_ADICIONALES_NOMINA)',
      pred: function (r) {
        if (r.descripcionNorm.indexOf('BONIFICACION PARA LA ASISTENCIA FAMILIAR') < 0) return false;
        if (r.descripcionNorm.indexOf('EXTINCION') >= 0) return false;   // las extinciones van aparte
        if (r.descripcionNorm.indexOf('DISMINUCION') >= 0) return false;
        return true;
      } }
  ];
  for (i = 0; i < bloquesSubsidio.length; i++) {
   gruposSubsidio.push(grupo_(bloquesSubsidio[i].clave, bloquesSubsidio[i].etiqueta, lsi, bloquesSubsidio[i].pred));
  }

  // Bloque 2 - Operativos (quién y cuántos)
  var bloquesOperativos = [
    { clave: 'TRASLADO', etiqueta: 'TRASLADO',
      pred: function (r) { return r.descripcionNorm.indexOf('TRASLADO') >= 0; } },
    { clave: 'ALTA_PENSIONADO', etiqueta: 'ALTA TITULAR PENSIONADO / ALTA SUSTITUTO PENSIONADO',
      pred: function (r) { return r.descripcionNorm.indexOf('ALTA') >= 0 && r.descripcionNorm.indexOf('PENSIONADO') >= 0; },
      nota: 'Incluye tanto el alta de titular como la de sustituto pensionado.' },
    { clave: 'LICENCIA_MARIA', etiqueta: 'LICENCIA LEY MARIA (LICENCIA PATERNIDAD)',
      pred: function (r) { return r.descripcionNorm.indexOf('LEY MARIA') >= 0 || r.descripcionNorm.indexOf('LICENCIA PATERNIDAD') >= 0; } },
    { clave: 'LICENCIA_LUTO', etiqueta: 'LICENCIA REMUNERADA POR LUTO',
      pred: function (r) { return r.descripcionNorm.indexOf('LUTO') >= 0; } },
    { clave: 'BONIF_ASIST_FAM', etiqueta: 'BONIFICACION PARA LA ASISTENCIA FAMILIAR',
      pred: function (r) { return r.descripcionNorm.indexOf('BONIFICACION PARA LA ASISTENCIA FAMILIAR') >= 0; } }
  ];
  var bloquesOps = [];
  for (i = 0; i < bloquesOperativos.length; i++) {
    bloquesOps.push(grupo_(bloquesOperativos[i].clave, bloquesOperativos[i].etiqueta, lsi,
      bloquesOperativos[i].pred, { nota: bloquesOperativos[i].nota }));
  }

  /* ---------------- Armar reporte ---------------- */

  var porPersona = {};
  for (i = 0; i < siath.filas.length; i++) {
    var p = siath.filas[i];
    porPersona[p.id] = p.id;
  }

  var reporte = {
    ok: true,
    version: CFG.VERSION,
    generadoEn: new Date().toISOString(),
    configuracion: {
      palabraFiltro: config.palabraFiltro || CFG.PALABRA_FILTRO,
      marcadorAdicional: config.marcadorAdicional || CFG.MARCADOR_ADICIONALES,
      cruzaBases: config.cruzaBases !== false
    },
    fuentes: {
      siath: { hoja: tSiath.hoja, filasLeidas: tSiath.filas.length, filas: siath.filas.length,
               conPalabra: siath.stats.conPalabra, bytes: tSiath.bytes, truncado: !!tSiath.truncado },
      lsi: { hoja: tLsi.hoja, filasLeidas: tLsi.filas.length, filas: lsi.length,
             fueraDelPadron: fLsi.fuera.length, conPalabra: lsiAll.stats.conPalabra,
             periodo: lsiAll.periodo, bytes: tLsi.bytes, truncado: !!tLsi.truncado },
      data: { hoja: tData.hoja, filasLeidas: tData.filas.length, filas: data.length,
              fueraDelPadron: fData.fuera.length, conPalabra: dataAll.stats.conPalabra,
              bytes: tData.bytes, truncado: !!tData.truncado }
    },
    funcionarios: {
      // CANTIDAD TOTAL DE PERSONAL NOMINADO según BASE SIATH
      nominados: siath.filas.length,
      identificaciones: idsSiath.length,
      filasLeidasSiath: siath.stats.leidas,
      hojasSiath: tSiath.hoja,
      porDependencia: resumenDependencias_(siath.filas, 25)
    },
    traslados: analizarTraslados_(lsi, porIdSiath, config.palabraFiltro || CFG.PALABRA_FILTRO,
                                  construirMapaUnidades_(config.mapaUnidades)),
    conteos: {
      adicionales: (function () {
        var q = objetoDetalle_(ad.filas, 'LSI');
        q.senales = ad.senales;
        return q;
      })(),
      incapacidades: objetoDetalle_(incapacidad),
      vacaciones: objetoDetalle_(vacaciones),
      embargosAplicados: objetoDetalle_(embargos, 'SIATH'),
      embargosFamilia: objetoDetalle_(embFam, 'SIATH'),
      noOrdenPublico: objetoDetalle_(noOrden, 'SIATH'),
      comisionesExterior: objetoDetalle_(comExt)
    },
    grupos: {
      Subsidios: gruposSubsidio,
      Operativos: bloquesOps
    },
    avisos: avisos,
    fueraDelPadron: {
      lsi: fLsi.fuera.slice(0, 500).map(detalleLsi_),
      data: fData.fuera.slice(0, 500).map(detalleData_)
    }
  };
  return reporte;
}

/** Combina filas LSI y DATA en un solo arreglo etiquetado. */
function combinar_(a, b) {
  var out = [];
  for (var i = 0; i < a.length; i++) out.push(a[i]);
  for (var j = 0; j < b.length; j++) out.push(b[j]);
  return out;
}

/** Top dependencias del padrón SIATH, para caracterizar el personal nominado. */
function resumenDependencias_(filas, limite) {
  var c = {};
  for (var i = 0; i < filas.length; i++) {
    var d = filas[i].descripcion || filas[i].cargo || 'SIN DEPENDENCIA';
    c[d] = (c[d] || 0) + 1;
  }
  var arr = [];
  for (var k in c) if (c.hasOwnProperty(k)) arr.push({ dependencia: k, total: c[k] });
  arr.sort(function (a, b) { return b.total - a.total; });
  return arr.slice(0, limite || 20);
}

/** Detalle uniforme para una fila LSI. */
function detalleLsi_(r) {
  return {
    id: r.id,
    nombres: r.nombres,
    apellidos: r.apellidos,
    nombreCompleto: r.nombreCompleto,
    descripcion: r.descripcion,
    categoria: r.categoria || '',
    clase: r.clase,
    anioMes: r.anioMes,
    dias: r._dias || r.diasLicencia || r.diasVacaciones || r.diasComision || 0,
    diasVacaciones: r.diasVacaciones,
    diasLicencia: r.diasLicencia,
    diasComision: r.diasComision,
    comisionExterior: r.comisionExterior,
    disposicion: r.disposicion,
    fechaDisposicion: r.fechaDisposicion,
    proceso: r.proceso,
    unidad: r.unidad,
    fisica: r.fisica,
    motivo: r._motivo || ''
  };
}

/** Detalle uniforme para una fila DATA. */
function detalleData_(r) {
  return {
    id: r.id,
    nombres: r.nombres,
    apellidos: r.apellidos,
    nombreCompleto: r.nombreCompleto,
    categoria: r.categoria,
    dias: r.dias,
    sigla: r.sigla,
    fechaInicio: r.fechaInicio,
    fechaFin: r.fechaFin,
    tipoExcusa: r.tipoExcusa,
    descripcion: r.descripcion,
    cargo: r.cargo,
    disposicion: r.disposicion,
    unidad: r.descripcion
  };
}

/**
 * Cuantifica y arma el detalle de un conjunto de filas.
 * Cada fila debe traer la marca _tipo = 'LSI' | 'DATA' | 'SIATH'.
 */
function objetoDetalle_(filas, origen) {
  var q = cuantificar_(filas);
  var det = [];
  for (var i = 0; i < filas.length && det.length < CFG.MAX_DETALLE; i++) {
    var r = filas[i];
    var base;
    if (r._tipo === 'DATA') base = detalleData_(r);
    else if (r._tipo === 'SIATH') base = detalleSiath_(r);
    else base = detalleLsi_(r);
    base.origen = origen || r._tipo || '';
    det.push(base);
  }
  q.detalle = det;
  return q;
}

/** Detalle uniforme para una fila SIATH. */
function detalleSiath_(r) {
  return {
    id: r.id,
    nombres: r.nombres,
    apellidos: r.apellidos,
    nombreCompleto: r.nombreCompleto,
    unidad: r.unidad,
    grado: r.grado,
    cargo: r.cargo,
    descripcion: r.descripcion,
    estado: r.estado,
    devengado: r.devengado,
    neto: r.neto,
    ordenPublico: r.ordenPublico,
    embargos: r.embargos,
    embFamilia: r.embFamilia,
    embEjecutivo: r.embEjecutivo,
    dias: r._dias || 0
  };
}

/**
 * Construye el índice código de unidad -> nombre de unidad a partir de un texto
 * con líneas "CODIGO=NOMBRE" o "CODIGO, NOMBRE".
 * Permite que el reporte muestre nombres de unidad y no solo códigos.
 */
function construirMapaUnidades_(texto) {
  var mapa = {};
  if (!texto) return mapa;
  var lineas = String(texto).split(/\r?\n/);
  for (var i = 0; i < lineas.length; i++) {
    var ln = txt_(lineas[i]);
    if (!ln || ln.charAt(0) === '#') continue;
    var sep = ln.indexOf('=');
    if (sep < 0) sep = ln.indexOf(',');
    if (sep < 0) continue;
    var codigo = txt_(ln.substring(0, sep));
    var nombre = txt_(ln.substring(sep + 1));
    if (codigo && nombre) mapa[codigo] = nombre;
  }
  return mapa;
}

/** Nombre legible de una unidad: código -> nombre, o el código si no está mapeado. */
function nombreUnidad_(codigo, mapa) {
  var c = txt_(codigo);
  if (!c) return 'NO INFORMADO';
  if (mapa && mapa[c]) return mapa[c] + ' (cód. ' + c + ')';
  return 'Unidad código ' + c;
}

/**
 * Analiza los TRASLADOS del padrón y determina quién está nominalmente en la
 * unidad filtrada (JESEP) pero quedó asignado a otra unidad.
 */
function analizarTraslados_(lsi, siathPorId, palabra, mapaUnidades) {
  var origen = [];
  var otros = [], iguales = 0, nominaDistintaLabor = 0, sinDestino = 0;

  for (var i = 0; i < lsi.length; i++) {
    var r = lsi[i];
    if (r.descripcionNorm.indexOf('TRASLADO') < 0) continue;

    var base = siathPorId[r.id];
    var consActual = txt_(r.consecutivoActual);
    var destLab = txt_(r.destinoLabor);
    var destNom = txt_(r.destinoNomina);

    if (!destLab && !destNom) { sinDestino++; continue; }

    var destino = destLab || destNom;
    var cambia = (consActual && destino && consActual !== destino);
    var nominaVsLabor = (destLab && destNom && destLab !== destNom);
    if (nominaVsLabor) nominaDistintaLabor++;
    if (!cambia) iguales++;

    origen.push({
      id: r.id,
      nombres: r.nombres,
      apellidos: r.apellidos,
      nombreCompleto: r.nombreCompleto,
      unidadNominal: txt_(base ? (base.unidad || palabra) : palabra),
      dependenciaNominal: base ? base.descripcion : '',
      grado: base ? base.grado : '',
      cargo: base ? base.cargo : '',
      consecutivoActual: consActual,
      unidadActual: nombreUnidad_(consActual, mapaUnidades),
      destinoNomina: destNom,
      destinoLaboral: destLab,
      unidadDestino: nombreUnidad_(destino, mapaUnidades),
      fuerzaDestino: txt_(r.fuerzaDestino),
      cambiaUnidad: !!cambia,
      nominaDistintaLabor: nominaVsLabor,
      anioMes: r.anioMes,
      disposicion: r.disposicion,
      fechaDisposicion: r.fechaDisposicion,
      proceso: r.proceso,
      creadoPor: r.creadoPor
    });
  }

  var fueraDeJESEP = [];
  for (i = 0; i < origen.length; i++) if (origen[i].cambiaUnidad) fueraDeJESEP.push(origen[i]);

  // Agrupación por unidad de destino
  var porDestino = {};
  for (i = 0; i < fueraDeJESEP.length; i++) {
    var k = fueraDeJESEP[i].unidadDestino;
    porDestino[k] = (porDestino[k] || 0) + 1;
  }
  var destinos = [];
  for (var d in porDestino) if (porDestino.hasOwnProperty(d)) destinos.push({ unidad: d, total: porDestino[d] });
  destinos.sort(function (a, b) { return b.total - a.total; });

  var_ids = fueraDeJESEP.map(function (x) { return x.id; });
  var unicos = {};
  for (i = 0; i < var_ids.length; i++) unicos[var_ids[i]] = true;
  var nUnicos = 0;
  for (var u in unicos) if (unicos.hasOwnProperty(u)) nUnicos++;

  return {
    total: origen.length,
    fueraDeUnidad: fueraDeJESEP.length,
    funcionariosFuera: nUnicos,
    sinCambioDeUnidad: iguales,
    nominaDistintaLabor: nominaDistintaLabor,
    sinDestinoInformado: sinDestino,
    porDestino: destinos,
    detalle: fueraDeJESEP.slice(0, CFG.MAX_DETALLE),
    detalleCompleto: origen.slice(0, CFG.MAX_DETALLE),
    truncado: fueraDeJESEP.length > CFG.MAX_DETALLE
  };
}

/* ======================= CRUCE DE NOVEDADES SUBIDAS ================ */

/** Lee un archivo de novedades y lo cruza contra las 3 bases. */
function cruzarNovedades_(tablaNov, tSiath, tLsi, tData, config) {
  config = config || {};
  var m = mapearColumnas_(tablaNov.cabeceras, ALIAS.NOV);

  // Índices de referencia: padrón SIATH y subconjuntos cruzados de LSI / DATA.
  var siathFilas = leerSiath_(tSiath, config).filas;
  var lsiFilas = leerLsi_(tLsi, config).filas;
  var dataFilas = leerData_(tData, config).filas;
  var idxSiath = construirIndices_(siathFilas);
  var idxLsi = {}, idxData = {};
  var i;
  for (i = 0; i < lsiFilas.length; i++) if (idxSiath[lsiFilas[i].id]) idxLsi[lsiFilas[i].id] = true;
  for (i = 0; i < dataFilas.length; i++) if (idxSiath[dataFilas[i].id]) idxData[dataFilas[i].id] = true;

  var lista = [], vistos = {};
  for (i = 0; i < tablaNov.filas.length && lista.length < CFG.MAX_NOVEDADES; i++) {
    var f = tablaNov.filas[i];
    var id = m.id >= 0 ? idNorm_(val_(f, m, 'id')) : '';
    if (!id) continue;
    var nombres = val_(f, m, 'nombres');
    var apellidos = val_(f, m, 'apellidos');
    if (!nombres && !apellidos) {
      var sp = separarNombre_(val_(f, m, 'nombre'));
      nombres = sp.nombres; apellidos = sp.apellidos;
    }
    var enSiath = !!idxSiath[id], enLsi = !!idxLsi[id], enData = !!idxData[id];
    var clave = id + '|' + norm_(nombres) + '|' + norm_(apellidos);
    if (vistos[clave]) continue;
    vistos[clave] = true;

    lista.push({
      id: id,
      nombres: nombres,
      apellidos: apellidos,
      nombreCompleto: txt_([apellidos, nombres].filter(String).join(' ')),
      descripcion: val_(f, m, 'descripcion'),
      anioMes: val_(f, m, 'anioMes'),
      fecha: val_(f, m, 'fecha'),
      referencia: val_(f, m, 'referencia'),
      clase: val_(f, m, 'claseNov'),
      enSiath: enSiath,
      enLsi: enLsi,
      enData: enData,
      grabado: enSiath,
      estado: enSiath ? (enLsi ? 'GRABADO Y CRUZADO' : 'GRABADO EN SIATH, SIN NOVEDAD EN LSI')
                       : 'NO ESTA EN LA BASE SIATH'
    });
  }

  var grabadas = [], noGrabadas = [], parciales = [];
  for (i = 0; i < lista.length; i++) {
    if (lista[i].enSiath) grabadas.push(lista[i]);
    else noGrabadas.push(lista[i]);
    if (lista[i].enSiath && !lista[i].enLsi && !lista[i].enData) parciales.push(lista[i]);
  }

  // Detalle por tipo de novedad en las que sí quedaron grabadas
  var porTipo = {};
  for (i = 0; i < grabadas.length; i++) {
    var t = norm_(grabadas[i].descripcion) || 'SIN CLASIFICACION';
    if (!porTipo[t]) porTipo[t] = { tipo: grabadas[i].descripcion || 'SIN CLASIFICACION', total: 0, ids: [] };
    porTipo[t].total++;
    if (porTipo[t].ids.length < 500) porTipo[t].ids.push(grabadas[i].id);
  }
  var tipos = [];
  for (var k in porTipo) if (porTipo.hasOwnProperty(k)) tipos.push(porTipo[k]);
  tipos.sort(function (a, b) { return b.total - a.total; });

  // Nombres master del padrón SIATH: resuelve datos incompletos y
  // permite detectar diferencias de nombre frente a la novedad subida.
  var maestro = {};
  for (i = 0; i < siathFilas.length; i++) maestro[siathFilas[i].id] = siathFilas[i];
  var nombresDistintos = 0;
  for (i = 0; i < lista.length; i++) {
    var ma = maestro[lista[i].id];
    if (ma) {
      lista[i].nombresSIATH = ma.nombres;
      lista[i].apellidosSIATH = ma.apellidos;
      lista[i].coincideNombre = norm_([lista[i].nombres, lista[i].apellidos].join(' ')) === norm_(ma.nombreCompleto);
      if (!lista[i].nombres) {
        lista[i].nombres = ma.nombres;
        lista[i].apellidos = ma.apellidos;
        lista[i].nombreCompleto = ma.nombreCompleto;
      }
    }
  }
  for (i = 0; i < lista.length; i++) if (lista[i].coincideNombre === false) nombresDistintos++;

  return {
    ok: true,
    generadoEn: new Date().toISOString(),
    archivo: {
      hoja: tablaNov.hoja,
      filasLeidas: tablaNov.filas.length,
      registrosUnicos: lista.length,
      columnas: tablaNov.cabeceras.length,
      columnasDetectadas: Object.keys(m).filter(function (k) { return m[k] >= 0; })
    },
    totales: {
      novedades: lista.length,
      grabadas: grabadas.length,
      noGrabadas: noGrabadas.length,
      porcentajeGrabado: lista.length ? Math.round(grabadas.length / lista.length * 1000) / 10 : 0,
      enLsi: lista.filter(function (x) { return x.enLsi; }).length,
      enData: lista.filter(function (x) { return x.enData; }).length,
      soloSiathSinNovelty: parciales.length,
      nombresDistintos: nombresDistintos
    },
    porTipo: tipos,
    detalle: lista.slice(0, CFG.MAX_DETALLE),
    noEncontradas: noGrabadas.slice(0, CFG.MAX_DETALLE),
    truncado: lista.length > CFG.MAX_DETALLE
  };
}

/* ======================= GENERADOR DE ZIP / DOCX ==================== */

var CRC_TABLA = (function () {
  var t = new Array(256), c, n, k;
  for (n = 0; n < 256; n++) {
    c = n;
    for (k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32_(bytes) {
  var c = 0xFFFFFFFF;
  for (var i = 0; i < bytes.length; i++) c = CRC_TABLA[(c ^ (bytes[i] & 0xFF)) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

/** Convierte texto a arreglo de bytes sin signo. */
function utf8Bytes_(texto) {
  var b = Utilities.newBlob(String(texto), MimeType.UTF8).getBytes();
  var out = new Array(b.length);
  for (var i = 0; i < b.length; i++) out[i] = b[i] & 0xFF;
  return out;
}

/** Escribe un entero de 2 o 4 bytes en little-endian. */
function pushLE_(arr, v, n) {
  for (var i = 0; i < n; i++) arr.push((v >> (i * 8)) & 0xFF);
}

/**
 * Crea un ZIP (método STORE, sin compresión) a partir de
 * [{nombre, texto}] y devuelve un Blob .zip/.docx.
 */
function crearZip_(archivos, nombre, mime) {
  var locales = [], centrales = [], offset = 0, i;

  for (i = 0; i < archivos.length; i++) {
    var nombreBytes = utf8Bytes_(archivos[i].nombre);
    var datos = utf8Bytes_(archivos[i].texto);
    var crc = crc32_(datos);

    var lh = [].concat([0x50, 0x4B, 0x03, 0x04]);
    pushLE_(lh, 20, 2);            // versión necesaria
    pushLE_(lh, 0, 2);             // flags
    pushLE_(lh, 0, 2);             // método = store
    pushLE_(lh, 0, 2);             // hora
    pushLE_(lh, 0, 2);             // fecha
    pushLE_(lh, crc, 4);
    pushLE_(lh, datos.length, 4);  // comprimido
    pushLE_(lh, datos.length, 4);  // original
    pushLE_(lh, nombreBytes.length, 2);
    pushLE_(lh, 0, 2);             // extra
    locales.push(lh, nombreBytes, datos);

    var ch = [].concat([0x50, 0x4B, 0x01, 0x02]);
    pushLE_(ch, 20, 2);            // versión creador
    pushLE_(ch, 20, 2);            // versión necesaria
    pushLE_(ch, 0, 2);
    pushLE_(ch, 0, 2);
    pushLE_(ch, 0, 2);
    pushLE_(ch, 0, 2);
    pushLE_(ch, crc, 4);
    pushLE_(ch, datos.length, 4);
    pushLE_(ch, datos.length, 4);
    pushLE_(ch, nombreBytes.length, 2);
    pushLE_(ch, 0, 2);             // extra
    pushLE_(ch, 0, 2);             // comentario
    pushLE_(ch, 0, 2);             // disco
    pushLE_(ch, 0, 2);             // atributos internos
    pushLE_(ch, 0, 4);             // atributos externos
    pushLE_(ch, offset, 4);
    centrales.push(ch, nombreBytes);

    offset += 30 + nombreBytes.length + datos.length;
  }

  var centralSize = centrales.reduce(function (a, x) { return a + x.length; }, 0);
  var eocd = [].concat([0x50, 0x4B, 0x05, 0x06]);
  pushLE_(eocd, 0, 2);
  pushLE_(eocd, 0, 2);
  pushLE_(eocd, archivos.length, 2);
  pushLE_(eocd, archivos.length, 2);
  pushLE_(eocd, centralSize, 4);
  pushLE_(eocd, offset, 4);
  pushLE_(eocd, 0, 2);

  var todos = locales.concat(centrales, eocd);
  var total = todos.reduce(function (a, x) { return a + x.length; }, 0);
  var plano = new Array(total);
  var p = 0;
  for (i = 0; i < todos.length; i++) {
    for (var j = 0; j < todos[i].length; j++) plano[p++] = todos[i][j];
  }
  var signed = new Array(plano.length);
  for (i = 0; i < plano.length; i++) signed[i] = plano[i] > 127 ? plano[i] - 256 : plano[i];

  return Utilities.newBlob(signed, mime, nombre);
}

/* ---------- Ayudantes de construcción DOCX ---------- */

var W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function p_(texto, opts) {
  opts = opts || {};
  var pPr = '<w:pPr>' +
    (opts.align ? '<w:jc w:val="' + opts.align + '"/>' : '') +
    (opts.espacio ? '<w:spacing w:before="' + opts.espacio + '" w:after="' + (opts.espacioDespues || 60) + '"/>' : '') +
    (opts.sangria ? '<w:ind w:left="' + opts.sangria + '"/>' : '') +
    '<w:pBdr>' +
      (opts.borde ? '<w:top w:val="single" w:sz="12" w:space="4" w:color="' + opts.borde + '"/>' : '') +
      (opts.borde ? '<w:bottom w:val="single" w:sz="12" w:space="4" w:color="' + opts.borde + '"/>' : '') +
    '</w:pBdr>' +
    '</w:pPr>';
  var rPr = '<w:rPr>' +
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>' +
    '<w:b/>' +
    (opts.tam ? '<w:sz w:val="' + opts.tam + '"/><w:szCs w:val="' + opts.tam + '"/>' : '<w:sz w:val="20"/><w:szCs w:val="20"/>') +
    '<w:color w:val="' + (opts.color || '1F3864') + '"/>' +
    (opts.mayus ? '<w:caps/>' : '') +
    '</w:rPr>';
  return '<w:p>' + pPr + '<w:r>' + rPr +
    '<w:t xml:space="preserve">' + escXml_(texto) + '</w:t></w:r></w:p>';
}

function celda_(texto, opts) {
  opts = opts || {};
  var tcPr = '<w:tcPr>' +
    '<w:tcW w:w="' + (opts.ancho || 2000) + '" w:type="dxa"/>' +
    (opts.sombra ? '<w:shd w:val="clear" w:color="auto" w:fill="' + opts.sombra + '"/>' : '') +
    '<w:vAlign w:val="center"/>' +
    '</w:tcPr>';
  var rPr = '<w:rPr>' +
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>' +
    '<w:sz w:val="' + (opts.tam || 16) + '"/><w:szCs w:val="' + (opts.tam || 16) + '"/>' +
    (opts.negrita ? '<w:b/>' : '') +
    '<w:color w:val="' + (opts.color || '000000') + '"/>' +
    '</w:rPr>';
  return '<w:tc>' + tcPr +
    '<w:p><w:pPr><w:spacing w:before="20" w:after="20"/>' +
    '<w:jc w:val="' + (opts.align || 'left') + '"/></w:pPr>' +
    '<w:r>' + rPr + '<w:t xml:space="preserve">' + escXml_(texto) + '</w:t></w:r></w:p></w:tc>';
}

function tabla_(encabezados, anchos, filas, opciones) {
  opciones = opciones || {};
  var bordes = '<w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(function (b) {
      return '<w:' + b + ' w:val="single" w:sz="4" w:space="0" w:color="1F3864"/>';
    }).join('') + '</w:tblBorders>';

  var xml = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>' +
    '<w:tblW w:w="5000" w:type="pct"/>' + bordes +
    '<w:tblLayout w:type="fixed"/></w:tblPr>';

  // Encabezado
  xml += '<w:tr><w:trPr><w:tblHeader/></w:trPr>';
  for (var i = 0; i < encabezados.length; i++) {
    xml += celda_(encabezados[i], { ancho: anchos[i], sombra: opciones.colorEncabezado || '1F3864',
      color: 'FFFFFF', negrita: true, align: i === 0 ? 'center' : 'left', tam: 16 });
  }
  xml += '</w:tr>';

  // Cuerpo
  for (var r = 0; r < filas.length; r++) {
    xml += '<w:tr>';
    var alterna = (r % 2 === 1) ? (opciones.colorAlterno || 'EDF1F8') : null;
    for (var c = 0; c < encabezados.length; c++) {
      xml += celda_(filas[r][c] === undefined ? '' : filas[r][c],
        { ancho: anchos[c], sombra: alterna, align: c === 0 ? 'center' : 'left' });
    }
    xml += '</w:tr>';
  }
  xml += '</w:tbl>' + p_('', { espacio: 40 });
  return xml;
}

function saltoPagina_() {
  return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
}

function docxContenido_(cuerpo) {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document ' + W_NS + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<w:body>' + cuerpo +
    '<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter"/>' +
    '<w:pgSz w:w="12240" w:h="15840"/>' +
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>' +
    '</w:sectPr></w:body></w:document>';
}

function docxEstilos_() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:styles ' + W_NS + '>' +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>' +
    '<w:sz w:val="20"/><w:szCs w:val="20"/>' +
    '<w:lang w:val="es-CO"/>' +
    '</w:rPr></w:rPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
    '<w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>' +
    '<w:tblPr><w:tblBorders>' +
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
    '</w:tblBorders></w:tblPr></w:style>' +
    '</w:styles>';
}

function docxFooter_(paginas) {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:ftr ' + W_NS + '><w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
    '<w:r><w:rPr><w:sz w:val="14"/><w:color w:val="666666"/></w:rPr>' +
    '<w:t xml:space="preserve">Policía Nacional de Colombia - Jefatura Nacional de Servicio de Policía (JESEP) | Página </w:t></w:r>' +
    '<w:r><w:rPr><w:sz w:val="14"/><w:color w:val="666666"/></w:rPr>' +
    '<w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:t>1</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '<w:r><w:rPr><w:sz w:val="14"/><w:color w:val="666666"/></w:rPr>' +
    '<w:t xml:space="preserve"> de ' + paginas + '</w:t></w:r>' +
    '</w:p></w:ftr>';
}

function docxCrear_(cuerpo, nombreBase) {
  var partes = [
    {
      nombre: '[Content_Types].xml',
      texto: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
        '</Types>'
    },
    {
      nombre: '_rels/.rels',
      texto: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
        '</Relationships>'
    },
    {
      nombre: 'word/_rels/document.xml.rels',
      texto: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '<Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
        '</Relationships>'
    },
    { nombre: 'word/styles.xml', texto: docxEstilos_() },
    { nombre: 'word/document.xml', texto: docxContenido_(cuerpo) },
    {
      nombre: 'word/footer1.xml',
      texto: docxFooter_('' + CFG.PAGINAS_ESTIMADAS)
    },
    {
      nombre: 'docProps/core.xml',
      texto: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
        'xmlns:dc="http://purl.org/dc/elements/1.1/">' +
        '<dc:title>Reporte de cruce SIATH - LSI - DATA</dc:title>' +
        '<dc:creator>Jefatura Nacional de Servicio de Policía - JESEP</dc:creator>' +
        '</cp:coreProperties>'
    },
    {
      nombre: 'docProps/app.xml',
      texto: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
        '<Application>Google Apps Script</Application></Properties>'
    }
  ];
  return crearZip_(partes, nombreBase + '.docx',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
}

/* ======================= ARMADO DEL REPORTE WORD ==================== */

function construirReporteWord_(rep) {
  var b = [];
  var hoy = new Date();
  var fechaTxt = hoy.toLocaleDateString('es-CO', { day: '2-digit', month: 'long', year: 'numeric' });

  /* --- Portada / encabezado institucional --- */
  b.push(p_('POLICÍA NACIONAL DE COLOMBIA', { align: 'center', tam: 32, color: '0B2A5B', mayus: true, espacio: 0 }));
  b.push(p_('Jefatura Nacional de Servicio de Policía - JESEP', { align: 'center', tam: 24, color: '1F3864', espacio: 0 }));
  b.push(p_('Reporte de cruce de información - BASE SIATH / BASE LSI / BASE DATA', { align: 'center', tam: 22, color: '8B6914', borde: 'C9A227', espacio: 120 }));
  b.push(p_('Palabra de filtro aplicada: "' + rep.configuracion.palabraFiltro + '"', { align: 'center', tam: 18, color: '555555', espacio: 0 }));
  b.push(p_('Generado: ' + fechaTxt + ' | Versión ' + rep.version, { align: 'center', tam: 16, color: '777777', espacio: 200 }));

  b.push(p_('PERSONAL NOMINADO SEGÚN BASE SIATH: ' + fmtInt_(rep.funcionarios.nominados) + ' FUNCIONARIOS',
    { align: 'center', tam: 26, color: '8B6914', borde: 'C9A227', espacio: 200 }));

  b.push(p_('1. RESUMEN EJECUTIVO', { align: 'left', tam: 24, color: '0B2A5B', borde: 'C9A227', espacio: 200 }));
  b.push(p_('Total de personal nominado (BASE SIATH, hoja "' + rep.funcionarios.hojasSiath + '", filtro "' +
    rep.configuracion.palabraFiltro + '"): ' + fmtInt_(rep.funcionarios.nominados) +
    ' funcionarios. Filas leídas en el archivo: ' + fmtInt_(rep.funcionarios.filasLeidasSiath) + '.',
    { tam: 17, color: '333333', espacio: 60 }));

  b.push(tabla_(
    ['Concepto', 'Valor'],
    [5200, 3800],
    [
      ['PERSONAL NOMINADO (funcionarios en el padrón)', fmtInt_(rep.funcionarios.nominados)],
      ['Filas BASE SIATH leídas / filtradas', fmtInt_(rep.fuentes.siath.filasLeidas) + ' / ' + fmtInt_(rep.fuentes.siath.filas)],
      ['Filas BASE LSI leídas / cruzadas', fmtInt_(rep.fuentes.lsi.filasLeidas) + ' / ' + fmtInt_(rep.fuentes.lsi.filas)],
      ['Filas BASE DATA leídas / cruzadas', fmtInt_(rep.fuentes.data.filasLeidas) + ' / ' + fmtInt_(rep.fuentes.data.filas)],
      ['Traslados con cambio de unidad', fmtInt_(rep.traslados.fueraDeUnidad) + ' de ' + fmtInt_(rep.traslados.total)],
      ['Periodo novedades (ANO_MES)', rep.fuentes.lsi.periodo || 'No informado'],
      ['Hojas utilizadas', 'SIATH: ' + rep.fuentes.siath.hoja + ' | LSI: ' + rep.fuentes.lsi.hoja + ' | DATA: ' + rep.fuentes.data.hoja]
    ]
  ));

  if (rep.funcionarios.porDependencia && rep.funcionarios.porDependencia.length) {
    b.push(p_('1.1 Composición del personal nominado por dependencia (top ' +
      rep.funcionarios.porDependencia.length + ')', { tam: 20, color: '1F3864', espacio: 120 }));
    b.push(tabla_(
      ['Dependencia / cargo en BASE SIATH', 'Funcionarios'],
      [7400, 1600],
      rep.funcionarios.porDependencia.map(function (d) {
        return [d.dependencia, fmtInt_(d.total)];
      })
    ));
  }

  /* --- 2. Conteos --- */
  b.push(p_('2. CONTEOS SOLICITADOS', { align: 'left', tam: 24, color: '0B2A5B', borde: 'C9A227', espacio: 200 }));
  var c = rep.conteos;
  b.push(tabla_(
    ['Concepto', 'Registros', 'Funcionarios', 'Días / valor'],
    [4200, 1400, 1600, 1800],
    [
      ['Adicionales de nómina (' + rep.configuracion.marcadorAdicional + ')', fmtInt_(c.adicionales.registros), fmtInt_(c.adicionales.funcionarios), '-'],
      ['Incapacidades', fmtInt_(c.incapacidades.registros), fmtInt_(c.incapacidades.funcionarios), fmtInt_(c.incapacidades.dias)],
      ['Disfrute de vacaciones', fmtInt_(c.vacaciones.registros), fmtInt_(c.vacaciones.funcionarios), fmtInt_(c.vacaciones.dias)],
      ['Embargos aplicados', fmtInt_(c.embargosAplicados.registros), fmtInt_(c.embargosAplicados.funcionarios), fmtInt_(c.embargosAplicados.dias)],
      ['Embargos a favor de familia', fmtInt_(c.embargosFamilia.registros), fmtInt_(c.embargosFamilia.funcionarios), fmtInt_(c.embargosFamilia.dias)],
      ['NO en orden público automático', fmtInt_(c.noOrdenPublico.registros), fmtInt_(c.noOrdenPublico.funcionarios), '-'],
      ['Comisiones en el exterior', fmtInt_(c.comisionesExterior.registros), fmtInt_(c.comisionesExterior.funcionarios), fmtInt_(c.comisionesExterior.dias)]
    ]
  ));

  /* --- 3. Detalle de embargos y orden público (BASE SIATH) --- */
  b.push(p_('3. EMBARGOS, ORDEN PÚBLICO Y SITUACIÓN EN BASE SIATH', { align: 'left', tam: 24, color: '0B2A5B', borde: 'C9A227', espacio: 200 }));
  b.push(tablaConceptoSiath_('3.1 Funcionarios con EMBARGOS_APLICADOS', c.embargosAplicados));
  b.push(tablaConceptoSiath_('3.2 Funcionarios con EMB_FAMILIA', c.embargosFamilia));
  b.push(tablaConceptoSiath_('3.3 Funcionarios que NO están en ORDEN_PUBLICO_AUTOMATICO', c.noOrdenPublico));

  /* --- 4. Incapacidades y vacaciones --- */
  b.push(p_('4. INCAPACIDADES Y DISFRUTE DE VACACIONES', { align: 'left', tam: 24, color: '0B2A5B', borde: 'C9A227', espacio: 200 }));
  b.push(tablaDetalleMixto_('4.1 Incapacidades', c.incapacidades));
  b.push(tablaDetalleMixto_('4.2 Disfrute de vacaciones', c.vacaciones));

  /* --- 5. Comisiones en el exterior --- */
  b.push(p_('5. COMISIONES EN EL EXTERIOR', { align: 'left', tam: 24, color: '0B2A5B', borde: 'C9A227', espacio: 200 }));
  b.push(tablaDetalleMixto_('5.1 Detalle', c.comisionesExterior));

  /* --- 6. Subsidios --- */
  b.push(saltoPagina_());
  b.push(p_('6. SUBSIDIOS FAMILIARES Y BONIFICACIONES', { align: 'left', tam: 24, color: '0B2A5B', borde: 'C9A227', espacio: 200 }));
  b.push(p_('Identificación, nombres y apellidos de los funcionarios con novedades de este bloque.',
    { tam: 16, color: '555555', espacio: 60 }));
  for (var i = 0; i < rep.grupos.Subsidios.length; i++) {
    b.push(tablaGrupo_(rep.grupos.Subsidios[i]));
  }

  /* --- 7. Operativos --- */
  b.push(saltoPagina_());
  b.push(p_('7. TRASLADOS, PENSIÓN, LICENCIAS Y BONIFICACIONES', { align: 'left', tam: 24, color: '0B2A5B', borde: 'C9A227', espacio: 200 }));
  b.push(p_('Quiénes y cuántos, por cada concepto.', { tam: 16, color: '555555', espacio: 60 }));
  for (var j = 0; j < rep.grupos.Operativos.length; j++) {
    b.push(tablaGrupo_(rep.grupos.Operativos[j]));
  }

  /* --- 7A. Traslados: nominally in the filter unit but now in another unit --- */
  b.push(saltoPagina_());
  b.push(p_('7A. TRASLADOS: FUNCIONARIOS NOMINALMENTE EN ' + rep.configuracion.palabraFiltro +
    ' QUE QUEDARON EN OTRAS UNIDADES', { align: 'left', tam: 24, color: '0B2A5B', borde: 'C9A227', espacio: 200 }));
  b.push(tablaTraslados_(rep.traslados, rep.configuracion.palabraFiltro));

  /* --- 8. Adicionales --- */
  b.push(p_('8. NOVEDADES ADICIONALES DE NÓMINA', { align: 'left', tam: 24, color: '0B2A5B', borde: 'C9A227', espacio: 200 }));
  b.push(p_('Criterio de detección: columna "' + rep.configuracion.marcadorAdicional +
    '" con valor verdadero, o descripción que contenga la palabra ADICIONAL.',
    { tam: 16, color: '555555', espacio: 60 }));
  b.push(tablaDetalleMixto_('8.1 Detalle', c.adicionales));
  if (!c.adicionales.registros) {
    b.push(p_('No se registraron novedades adicionales de nómina en el archivo LSI para el padrón filtrado.',
      { tam: 18, color: '8B0000', espacio: 60 }));
  }

  /* --- 9. Cruce de novedades (si existe) --- */
  if (rep.cruce) {
    b.push(saltoPagina_());
    b.push(p_('9. VALIDACIÓN DE NOVEDADES SUBIDAS CONTRA LAS BASES', { align: 'left', tam: 24, color: '0B2A5B', borde: 'C9A227', espacio: 200 }));
    var cr = rep.cruce;
    b.push(p_('Archivo analizado: ' + (cr.archivo.nombre || '') + ' | Filas leídas: ' +
      fmtInt_(cr.archivo.filasLeidas) + ' | Registros únicos: ' + fmtInt_(cr.archivo.registrosUnicos),
      { tam: 16, color: '555555', espacio: 60 }));
    b.push(tabla_(
      ['Indicador', 'Cantidad'],
      [6000, 3000],
      [
        ['Novedades analizadas', fmtInt_(cr.totales.novedades)],
        ['Grabadas (encontradas en BASE SIATH)', fmtInt_(cr.totales.grabadas)],
        ['NO grabadas (no existen en ninguna base)', fmtInt_(cr.totales.noGrabadas)],
        ['Porcentaje grabado', String(cr.totales.porcentajeGrabado).replace('.', ',') + ' %'],
        ['Además encontradas en BASE LSI', fmtInt_(cr.totales.enLsi)],
        ['Además encontradas en BASE DATA', fmtInt_(cr.totales.enData)],
        ['Solo en SIATH, sin novedad en LSI ni DATA', fmtInt_(cr.totales.soloSiathSinNovelty)]
      ]
    ));

    b.push(p_('9.1 Novedades por tipo y total', { tam: 20, color: '1F3864', espacio: 120 }));
    b.push(tabla_(
      ['Tipo de novedad', 'Total', 'Identificaciones (muestra)'],
      [3600, 900, 4500],
      cr.porTipo.map(function (t) { return [t.tipo, fmtInt_(t.total), t.ids.join(', ')]; })
    ));

    if (cr.noEncontradas.length) {
      b.push(p_('9.2 NOVEDADES QUE NO ESTÁN GRABADAS EN NINGUNA BASE', { tam: 20, color: '8B0000', espacio: 120 }));
      b.push(tabla_(
        ['Identificación', 'Nombres', 'Apellidos', 'Novedad', 'Referencia', 'Estado'],
        [1300, 1700, 1700, 1900, 1100, 1300],
        cr.noEncontradas.map(function (x) {
          return [x.id, x.nombres, x.apellidos, x.descripcion, x.referencia, x.estado];
        })
      ));
    } else {
      b.push(p_('9.2 Todas las novedades cargadas se encuentran grabadas en las bases.',
        { tam: 18, color: '0B5D1E', espacio: 60 }));
    }

    b.push(p_('9.3 Detalle completo de novedades', { tam: 20, color: '1F3864', espacio: 120 }));
    b.push(tabla_(
      ['Identificación', 'Nombres', 'Apellidos', 'Novedad', 'SIATH', 'LSI', 'DATA', 'Estado'],
      [1200, 1500, 1500, 2000, 700, 600, 700, 1800],
      cr.detalle.map(function (x) {
        return [x.id, x.nombres, x.apellidos, x.descripcion,
          x.enSiath ? 'SÍ' : 'NO', x.enLsi ? 'SÍ' : 'NO', x.enData ? 'SÍ' : 'NO', x.estado];
      })
    ));
  }

  /* --- 10. Avisos técnicos --- */
  b.push(saltoPagina_());
  b.push(p_('10. OBSERVACIONES TÉCNICAS DEL PROCESO', { align: 'left', tam: 24, color: '0B2A5B', borde: 'C9A227', espacio: 200 }));
  var filasAviso = [];
  for (var k = 0; k < rep.avisos.length; k++) filasAviso.push([fmtInt_(k + 1), rep.avisos[k]]);
  b.push(tabla_(['#', 'Observación'], [700, 8300], filasAviso));

  var fuera = (rep.fueraDelPadron.lsi.length + rep.fueraDelPadron.data.length);
  if (fuera) {
    b.push(p_('Registros de las bases LSI y DATA que no pertenecen al padrón filtrado y por lo tanto ' +
      'quedaron excluidos del cruce: ' + fmtInt_(fuera) + '.', { tam: 18, color: '8B6914', espacio: 120 }));
  }

  CFG.PAGINAS_ESTIMADAS = String(Math.max(1, Math.ceil(b.join('').length / 24000)));
  return b.join('');
}

function tablaConceptoSiath_(titulo, q) {
  var filas = q.detalle.map(function (r) {
    return [r.id, r.nombres, r.apellidos, r.grado, r.cargo || r.descripcion,
            r.ordenPublico || '', r.embargos || 0, r.embFamilia || 0];
  });
  return p_(titulo + ' — total ' + fmtInt_(q.funcionarios) + ' funcionario(s), ' +
    fmtInt_(q.registros) + ' registro(s)', { tam: 20, color: '1F3864', espacio: 120 }) +
    tabla_(['Identificación', 'Nombres', 'Apellidos', 'Grado', 'Dependencia / Cargo', 'Orden púb.', 'Emb.', 'Emb. Fam.'],
      [1200, 1500, 1500, 700, 2200, 900, 600, 700], filas);
}

function tablaDetalleMixto_(titulo, q) {
  var filas = q.detalle.map(function (r) {
    return [r.id, r.nombres, r.apellidos, r.descripcion || r.categoria || '',
            r.origen || '', r.dias || 0, r.fechaInicio || r.fechaDisposicion || '', r.disposicion || ''];
  });
  var cab = ['Identificación', 'Nombres', 'Apellidos', 'Concepto', 'Origen', 'Días', 'Fecha', 'Documento'];
  return p_(titulo + ' — total ' + fmtInt_(q.funcionarios) + ' funcionario(s), ' +
    fmtInt_(q.registros) + ' registro(s)', { tam: 20, color: '1F3864', espacio: 120 }) +
    tabla_(cab, [1200, 1500, 1500, 2200, 900, 600, 900, 1200], filas);
}

/** Sección de traslados: origen nominal vs unidad de destino. */
function tablaTraslados_(t, palabra) {
  var b = [];

  b.push(p_('Se cruzaron las cédulas de todos los traslados contra el padrón de personal nominado de ' +
    'BASE SIATH. Un traslado se considera "cambio de unidad" cuando el consecutivo de la unidad de ' +
    'destino (UNDE_CONSECUTIVO_TRASLA_NOMI / _LABOR) es distinto del consecutivo de la unidad ' +
    'actual del funcionario (EMPL_CONSECUTIVO).',
    { tam: 16, color: '555555', espacio: 80 }));

  b.push(tabla_(
    ['Indicador', 'Cantidad'],
    [6400, 2600],
    [
      ['Total de traslados del padrón ' + palabra, fmtInt_(t.total)],
      ['Funcionarios que quedaron en OTRAS unidades', fmtInt_(t.fueraDeUnidad)],
      ['Funcionarios distintos afectados', fmtInt_(t.funcionariosFuera)],
      ['Traslados que no cambian de unidad', fmtInt_(t.sinCambioDeUnidad)],
      ['Traslados con destino nomina distinto del laboral', fmtInt_(t.nominaDistintaLabor)],
      ['Traslados sin unidad de destino informada', fmtInt_(t.sinDestinoInformado)]
    ]
  ));

  if (!t.porDestino.length) {
    b.push(p_('No se identificaron traslados con cambio de unidad dentro del padrón ' + palabra + '.',
      { tam: 18, color: '8B0000', espacio: 60 }));
    return b.join('');
  }

  b.push(p_('7A.1 Distribución por unidad de destino', { tam: 20, color: '1F3864', espacio: 120 }));
  b.push(tabla_(
    ['Unidad de destino', 'Funcionarios'],
    [7400, 1600],
    t.porDestino.map(function (d) { return [d.unidad, fmtInt_(d.total)]; })
  ));

  b.push(p_('7A.2 Detalle de funcionarios que estaban nominalmente en ' + palabra +
    ' y ahora están en otras unidades — ' + fmtInt_(t.fueraDeUnidad) + ' registro(s)',
    { tam: 20, color: '1F3864', espacio: 120 }));

  b.push(tabla_(
    ['Identificación', 'Nombres', 'Apellidos', 'Unidad de origen', 'Unidad de destino',
     'Cód. actual', 'Cód. destino', '¿Cambia?', 'Periodo', 'Documento'],
    [1150, 1300, 1300, 1250, 1750, 850, 850, 700, 700, 850],
    t.detalle.map(function (r) {
      return [r.id, r.nombres, r.apellidos, r.unidadNominal, r.unidadDestino,
              r.consecutivoActual, (r.destinoLaboral || r.destinoNomina),
              r.cambiaUnidad ? 'SÍ' : 'NO', r.anioMes, r.disposicion];
    })
  ));

  if (t.truncado) {
    b.push(p_('Nota: se muestran los primeros ' + CFG.MAX_DETALLE + ' registros de ' +
      fmtInt_(t.fueraDeUnidad) + '.', { tam: 15, color: '8B6914', espacio: 40 }));
  }
  return b.join('');
}

function tablaGrupo_(g) {
  var filas = g.detalle.map(function (r) {
    return [r.id, r.nombres, r.apellidos, r.descripcion || '',
            r.clase || '', r.anioMes || '', r.dias || 0, r.disposicion || '', r.fechaDisposicion || ''];
  });
  var cab = ['Identificación', 'Nombres', 'Apellidos', 'Novedad', 'Clase', 'Periodo', 'Días', 'Documento', 'Fecha doc.'];
  var cab2 = p_(g.etiqueta.toUpperCase() + ' — ' + fmtInt_(g.registros) + ' registro(s), ' +
    fmtInt_(g.funcionarios) + ' funcionario(s)', { tam: 20, color: '1F3864', espacio: 120 }) +
    (g.nota ? p_(g.nota, { tam: 15, color: '777777', espacio: 20 }) : '') +
    tabla_(cab, [1200, 1400, 1400, 1900, 600, 800, 500, 900, 1000], filas);
  if (g.truncado) {
    return cab2 + p_('Nota: se muestran los primeros ' + CFG.MAX_DETALLE + ' registros de ' +
      fmtInt_(g.registros) + '.', { tam: 15, color: '8B6914', espacio: 40 });
  }
  return cab2;
}

/* ======================= PERSISTENCIA DE REPORTES =================== */

function guardarReporte_(obj) {
  var f = carpetaTemp_(true).createFile(
    Utilities.newBlob(JSON.stringify(obj), MimeType.PLAIN_TEXT, 'reporte.json'));
  return f.getId();
}

function leerReporte_(id) {
  var f = DriveApp.getFileById(id);
  if (f.isTrashed()) throw new Error('El reporte ya fue eliminado de la carpeta temporal.');
  return JSON.parse(f.getDataAsString('UTF-8'));
}

/* ======================= API PÚBLICA =============================== */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Cruce SIATH · LSI · DATA — JESEP')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function apiEstado() {
  return {
    ok: true,
    version: CFG.VERSION,
    palabraFiltro: CFG.PALABRA_FILTRO,
    marcadorAdicional: CFG.MARCADOR_ADICIONALES,
    servidor: 'Google Apps Script',
    hora: new Date().toISOString(),
    hojaSiathPreferida: CFG.HOJA_SIATH_PREF,
    limites: {
      maxArchivoMB: Math.round(CFG.MAX_BYTES_PKG / 1048576),
      maxHojaMB: Math.round(CFG.MAX_BYTES_HOJA / 1048576),
      maxDetalle: CFG.MAX_DETALLE
    }
  };
}

/**
 * Ejecuta el cruce completo.
 * params = { siath:{fileId,nombre,hoja}, lsi:{...}, data:{...}, opciones:{...} }
 */
function apiAnalizar(params) {
  try {
    if (!params || !params.siath || !params.lsi || !params.data) {
      return { ok: false, error: 'Debe cargar los tres archivos: BASE SIATH, BASE LSI y BASE DATA.' };
    }
    var opciones = params.opciones || {};
    var t0 = Date.now();

    var tSiath = leerArchivoSubido_(params.siath);
    var tLsi = leerArchivoSubido_(params.lsi);
    var tData = leerArchivoSubido_(params.data);

    var config = {
      palabraFiltro: opciones.palabraFiltro || CFG.PALABRA_FILTRO,
      marcadorAdicional: opciones.marcadorAdicional || CFG.MARCADOR_ADICIONALES,
      mapaUnidades: opciones.mapaUnidades || '',
      filtrarPorPalabra: opciones.filtrarPorPalabra !== false,
      cruzaBases: opciones.cruzaBases !== false
    };

    var rep = analizarTablas_(tSiath, tLsi, tData, config);

    // Cruce opcional de novedades, si se subieron en el mismo paso.
    if (params.novedades && params.novedades.fileId) {
      try {
        var tNov = leerArchivoSubido_(params.novedades);
        rep.cruce = cruzarNovedades_(tNov, tSiath, tLsi, tData, config);
        rep.cruce.archivo.nombre = params.novedades.nombre || '';
      } catch (eNov) {
        rep.avisos.push('No se pudo procesar el archivo de novedades: ' + eNov.message);
      }
    }

    rep.duracionMs = Date.now() - t0;
    rep.reportId = guardarReporte_(rep);
    return rep;
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

/** Lee un archivo previamente subido a la carpeta temporal. */
function leerArchivoSubido_(ref) {
  var blob = DriveApp.getFileById(ref.fileId).getBlob();
  return leerTabla_(blob, ref.nombre || blob.getName(), ref.hoja || '');
}

/** Cruza un archivo de novedades previamente subido. */
function apiCruzarNovedades(params) {
  try {
    if (!params || !params.novedades || !params.novedades.fileId) {
      return { ok: false, error: 'Debe subir el archivo de novedades realizadas.' };
    }
    if (!params.siath || !params.lsi || !params.data) {
      return { ok: false, error: 'Debe tener cargadas las tres bases para poder cruzar.' };
    }
    var opciones = params.opciones || {};
    var config = {
      palabraFiltro: opciones.palabraFiltro || CFG.PALABRA_FILTRO,
      marcadorAdicional: opciones.marcadorAdicional || CFG.MARCADOR_ADICIONALES
    };
    var tSiath = leerArchivoSubido_(params.siath);
    var tLsi = leerArchivoSubido_(params.lsi);
    var tData = leerArchivoSubido_(params.data);
    var tNov = leerArchivoSubido_(params.novedades);

    var cruce = cruzarNovedades_(tNov, tSiath, tLsi, tData, config);
    cruce.archivo.nombre = params.novedades.nombre || '';
    return cruce;
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

/** Genera el .docx y lo devuelve en base64. */
function apiGenerarWordReporte(reportId) {
  try {
    var rep = leerReporte_(reportId);
    var cuerpo = construirReporteWord_(rep);
    var blob = docxCrear_(cuerpo, 'Reporte_Cruce_SIATH_LSI_DATA_JESEP');
    var b64 = Utilities.base64Encode(blob.getBytes());
    return {
      ok: true,
      base64: b64,
      nombre: 'Reporte_Cruce_SIATH_LSI_DATA_JESEP.docx',
      bytes: blob.getBytes().length
    };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

/** Texto plano del reporte, útil para copiar/pegar. */
function apiGenerarTextoReporte(reportId) {
  try {
    var rep = leerReporte_(reportId);
    var L = [];
    L.push('POLICÍA NACIONAL DE COLOMBIA — JESEP');
    L.push('Reporte de cruce SIATH / LSI / DATA — filtro "' + rep.configuracion.palabraFiltro + '"');
    L.push('Generado: ' + rep.generadoEn);
    L.push('');
    L.push('FUNCIONARIOS NOMINADOS: ' + rep.funcionarios.nominados);
    var c = rep.conteos;
    L.push('Adicionales de nómina: ' + c.adicionales.registros);
    L.push('Incapacidades: ' + c.incapacidades.registros + ' registros / ' + c.incapacidades.funcionarios + ' funcionarios');
    L.push('Vacaciones: ' + c.vacaciones.registros + ' registros / ' + c.vacaciones.funcionarios + ' funcionarios');
    L.push('Embargos aplicados: ' + c.embargosAplicados.funcionarios);
    L.push('Embargos familia: ' + c.embargosFamilia.funcionarios);
    L.push('NO en orden público automático: ' + c.noOrdenPublico.funcionarios);
    L.push('Comisiones en el exterior: ' + c.comisionesExterior.registros);
    return { ok: true, texto: L.join('\n') };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}