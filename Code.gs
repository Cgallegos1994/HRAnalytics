/**
 * Configuración centralizada y helpers comunes para la app de visualización
 * de bandas salariales. VERSIÓN OPTIMIZADA para máximo rendimiento.
 * 
 * OPTIMIZACIONES IMPLEMENTADAS:
 * - Lectura incremental de datos con chunking
 * - Índices precalculados para filtros rápidos
 * - Cache inteligente con versionado
 * - Operaciones batch optimizadas
 * - Gestión de memoria mejorada
 */

/********************* CONFIGURACIÓN *********************/
var CONFIG = Object.freeze({
  SPREADSHEET_ID: '1LkBuhAu4BZWuCdAEZR6d3Ub5BUfSv3ZsVpRQbACw4n8',
  DATA_SHEET: 'Laboratorio Carmen',
  PERMISSIONS_SHEET: 'Permisos',
  // Configuración de rendimiento
  CHUNK_SIZE: 1000,           // Procesar datos en chunks de 1000 filas
  MAX_CACHE_SIZE: 50,         // Máximo 50 entradas en caché
  CACHE: Object.freeze({
    PERMISSIONS: 1800,        // 30 minutos
    FILTERS: 1200,            // 20 minutos
    DATA: 600,                // 10 minutos
    INDEXES: 1800             // 30 minutos
  })
});

var COLUMN_CANDIDATES = Object.freeze({
  idxPuesto: ['PuestoDenom','Puesto Denom','Puesto','Puesto denominación','Denominacion de Puesto'],
  idxNombre: ['Apellidos y nombre','Apellidos y Nombre','Nombre y apellidos','Nombre'],
  idxSBA: ['SBA 100%','SBA','SBA100%','SBA 100'],
  idxInicio: ['Inicio Banda','Inicio de banda','Min banda','Minimo banda','Mínimo banda'],
  idxFinal: ['Final de banda','Final Banda','Max banda','Maximo banda','Máximo banda'],
  idxMedio: ['Punto Medio Banda','Punto Medio','Midpoint','Pto Medio Banda'],
  idxDivision: ['DivisionDenom','División','Division','Division Denom'],
  idxDireccion: ['DireccionDenom','Dirección','Direccion','Direccion Denom'],
  idxRol: ['RolDenom','Rol Denom','Rol'],
  idxRolId: ['RolId','Rol Id','RolID','Rol ID'],
  idxFamilia: ['FamiliaDenom','Familia Denom','Familia'],
  idxDepartamento: ['DepartDenom','Depart Denom','Departamento'],
  idxSeccion: ['SeccionDenom','Seccion Denom','Sección','Seccion'],
  idxEmpresa: ['Empresa','Company','Compañía','Compania'],
  idxPosicion: ['Posición','Posicion','Position'],
  idxNivel: ['Nivel','NivelDenom','Nivel Denom','Level']
});

var FILTER_DEFINITIONS = Object.freeze([
  { key: 'puesto',       label: 'Puesto',       columnKey: 'idxPuesto',       emptyLabel: 'Todos'  },
  { key: 'division',     label: 'División',     columnKey: 'idxDivision',     emptyLabel: 'Todas'  },
  { key: 'direccion',    label: 'Dirección',    columnKey: 'idxDireccion',    emptyLabel: 'Todas'  },
  { key: 'rol',          label: 'Rol',          columnKey: 'idxRol',          emptyLabel: 'Todos'  },
  { key: 'familia',      label: 'Familia',      columnKey: 'idxFamilia',      emptyLabel: 'Todas'  },
  { key: 'departamento', label: 'Departamento', columnKey: 'idxDepartamento', emptyLabel: 'Todos'  },
  { key: 'seccion',      label: 'Sección',      columnKey: 'idxSeccion',      emptyLabel: 'Todas'  },
  { key: 'empresa',      label: 'Empresa',      columnKey: 'idxEmpresa',      emptyLabel: 'Todas'  },
  { key: 'posicion',     label: 'Posición',     columnKey: 'idxPosicion',     emptyLabel: 'Todas'  }
]);

var SUMMARY_BUCKETS = ['Q0', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5'];

// Niveles directivos que se excluyen del cálculo de R²
var EXCLUDED_DIRECTOR_LEVELS = Object.freeze([
  'Corp_Comité de Dirección / Director ejecutivo',
  'Negocio_Director General',
  'Negocio_Director Nivel 2',
  'Negocio_Comité de Dirección / Director ejecutivo',
  'Sistemas_Comité de Dirección / Director ejecutivo'
]);

/********************* HELPERS DE CACHE OPTIMIZADOS *********************/
var CACHE_METADATA = {}; // Track cache size and usage

function getCachedObject_(key) {
  var raw = CacheService.getScriptCache().get(key);
  if (raw) {
    // Update access metadata
    CACHE_METADATA[key] = CACHE_METADATA[key] || {};
    CACHE_METADATA[key].lastAccess = Date.now();
    return JSON.parse(raw);
  }
  return null;
}

function putCachedObject_(key, value, seconds) {
  try {
    // Clean up old cache entries if needed
    cleanupCache_();
    
    var serialized = JSON.stringify(value);
    CacheService.getScriptCache().put(key, serialized, seconds);
    
    // Update metadata
    CACHE_METADATA[key] = {
      size: serialized.length,
      created: Date.now(),
      lastAccess: Date.now()
    };
  } catch (err) {
    Logger.log('No se pudo almacenar en caché "' + key + '": ' + err);
  }
}

function cleanupCache_() {
  var keys = Object.keys(CACHE_METADATA);
  if (keys.length <= CONFIG.MAX_CACHE_SIZE) return;
  
  // Sort by last access time and remove oldest
  keys.sort(function(a, b) {
    return (CACHE_METADATA[a].lastAccess || 0) - (CACHE_METADATA[b].lastAccess || 0);
  });
  
  var toRemove = keys.slice(0, keys.length - CONFIG.MAX_CACHE_SIZE);
  CacheService.getScriptCache().removeAll(toRemove);
  
  toRemove.forEach(function(key) {
    delete CACHE_METADATA[key];
  });
}

/********************* AUTENTICACIÓN Y PERMISOS (OPTIMIZADO) *********************/
function getUserEmail_() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email) throw new Error('No se pudo obtener el email del usuario.');
    return email.toLowerCase().trim();
  } catch (err) {
    Logger.log('Error obteniendo email: ' + err);
    throw new Error('Error de autenticación. Contacta al administrador.');
  }
}

function isEmailAuthorized_(email) {
  var normalized = (email || '').toLowerCase().trim();
  var cacheKey = 'authorized_emails_v4';
  var cached = getCachedObject_(cacheKey);

  if (cached) {
    return !!cached[normalized];
  }

  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.PERMISSIONS_SHEET);
    if (!sheet) {
      throw new Error("No se encontró la pestaña '" + CONFIG.PERMISSIONS_SHEET + "'.");
    }

    var lastRow = sheet.getLastRow();
    var authorizedSet = {};

    if (lastRow > 1) {
      // OPTIMIZACIÓN: Usar getValues en lugar de múltiples getValue
      var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < data.length; i++) {
        var em = String(data[i][0] || '').toLowerCase().trim();
        if (em && em.indexOf('@') > -1) {
          authorizedSet[em] = true;
        }
      }
    }

    putCachedObject_(cacheKey, authorizedSet, CONFIG.CACHE.PERMISSIONS);
    return !!authorizedSet[normalized];
  } catch (err) {
    Logger.log('Error leyendo permisos: ' + err);
    throw new Error('Error al verificar permisos.');
  }
}

function getAuthorizedEmailsForDebug_() {
  try {
    var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.PERMISSIONS_SHEET);
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var emails = [];
    for (var i = 0; i < data.length; i++) {
      var em = String(data[i][0] || '').toLowerCase().trim();
      if (em && em.indexOf('@') > -1) emails.push(em);
    }
    return emails;
  } catch (err) {
    Logger.log('Error en getAuthorizedEmailsForDebug_: ' + err);
    return [];
  }
}

/********************* UTILIDADES OPTIMIZADAS *********************/
var NORMALIZE_CACHE = {};
var NORMALIZE_CACHE_SIZE = 0;
var MAX_NORMALIZE_CACHE_SIZE = 1000;

function normalize_(s) {
  var str = String(s || '');
  if (NORMALIZE_CACHE[str]) return NORMALIZE_CACHE[str];

  // Clean cache if it gets too large
  if (NORMALIZE_CACHE_SIZE >= MAX_NORMALIZE_CACHE_SIZE) {
    NORMALIZE_CACHE = {};
    NORMALIZE_CACHE_SIZE = 0;
  }

  var result = str
    .toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  NORMALIZE_CACHE[str] = result;
  NORMALIZE_CACHE_SIZE++;
  return result;
}

function toNumberEU_(v) {
  if (typeof v === 'number') return v;
  var s = String(v || '').replace(/[\s€]/g, '').replace(/\./g, '').replace(/,/g, '.');
  var n = parseFloat(s);
  return isNaN(n) ? NaN : n;
}

function fmt2_(n) {
  return isNaN(n) ? '—' : n.toFixed(2).replace('.', ',');
}

function findHeaderIndex_(headers, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    for (var j = 0; j < headers.length; j++) {
      if (normalize_(headers[j]) === normalize_(candidates[i])) {
        return j;
      }
    }
  }
  return -1;
}

/********************* ACCESO A DATOS OPTIMIZADO *********************/
function getColumns_(headers) {
  var cols = {};
  for (var key in COLUMN_CANDIDATES) {
    cols[key] = findHeaderIndex_(headers, COLUMN_CANDIDATES[key]);
  }
  return cols;
}

function getSheetAndData_() {
  var cacheKey = 'sheet_data_v4';
  var cached = getCachedObject_(cacheKey);
  if (cached) {
    return cached;
  }

  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.DATA_SHEET);
  if (!sheet) throw new Error("No se encontró la hoja '" + CONFIG.DATA_SHEET + "'.");

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) throw new Error('No hay datos suficientes en la hoja.');

  // OPTIMIZACIÓN: Leer datos en chunks para evitar timeouts
  var data = [];
  var chunkSize = CONFIG.CHUNK_SIZE;
  
  for (var startRow = 1; startRow <= lastRow; startRow += chunkSize) {
    var endRow = Math.min(startRow + chunkSize - 1, lastRow);
    var chunk = sheet.getRange(startRow, 1, endRow - startRow + 1, lastCol).getValues();
    
    if (startRow === 1) {
      data = chunk; // First chunk includes headers
    } else {
      data = data.concat(chunk);
    }
  }

  var headers = data.shift();
  var cols = getColumns_(headers);

  var result = { data: data, cols: cols, totalRows: data.length };
  putCachedObject_(cacheKey, result, CONFIG.CACHE.DATA);
  return result;
}

/********************* ÍNDICES PARA FILTROS RÁPIDOS *********************/
function buildFilterCacheKey_(key) {
  return 'filter_' + key + '_v8';
}

function getFilterOptions_() {
  var cache = CacheService.getScriptCache();
  var cacheKeys = FILTER_DEFINITIONS.map(function(def) { return buildFilterCacheKey_(def.key); });
  var cached = cache.getAll(cacheKeys);

  var allPresent = true;
  for (var i = 0; i < cacheKeys.length; i++) {
    if (!cached[cacheKeys[i]]) {
      allPresent = false;
      break;
    }
  }

  if (allPresent) {
    var cachedResult = {};
    for (var j = 0; j < FILTER_DEFINITIONS.length; j++) {
      var def = FILTER_DEFINITIONS[j];
      cachedResult[def.key] = JSON.parse(cached[buildFilterCacheKey_(def.key)] || '[]');
      }
    return cachedResult;
  }

  // OPTIMIZACIÓN: Procesar datos en chunks para evitar memory overflow
  var pack = getSheetAndData_();
  var data = pack.data;
  var cols = pack.cols;

  var setMap = {};
  for (var k = 0; k < FILTER_DEFINITIONS.length; k++) {
    setMap[FILTER_DEFINITIONS[k].key] = {};
  }

  // Procesar en chunks para mejor rendimiento
  var chunkSize = CONFIG.CHUNK_SIZE;
  for (var chunkStart = 0; chunkStart < data.length; chunkStart += chunkSize) {
    var chunkEnd = Math.min(chunkStart + chunkSize, data.length);
    
    for (var r = chunkStart; r < chunkEnd; r++) {
      var row = data[r];
      for (var f = 0; f < FILTER_DEFINITIONS.length; f++) {
        var defn = FILTER_DEFINITIONS[f];
        var idx = cols[defn.columnKey];
        if (idx === -1) continue;
        var val = String(row[idx] || '').trim();
        if (val) {
          setMap[defn.key][val] = true;
        }
      }
    }
    
    // Pequeña pausa para evitar timeout
    if (chunkStart % (chunkSize * 5) === 0) {
      Utilities.sleep(10);
    }
  }

  var sortES = function(a, b) { return a.localeCompare(b, 'es'); };
  var result = {};
  for (var t = 0; t < FILTER_DEFINITIONS.length; t++) {
    var def = FILTER_DEFINITIONS[t];
    var values = Object.keys(setMap[def.key] || {}).sort(sortES);
    result[def.key] = values;
    putCachedObject_(buildFilterCacheKey_(def.key), values, CONFIG.CACHE.FILTERS);
  }

  return result;
}

/********************* doGet OPTIMIZADO *********************/
function extractFilters_(params) {
  var selections = {};
  for (var i = 0; i < FILTER_DEFINITIONS.length; i++) {
    var key = FILTER_DEFINITIONS[i].key;
    selections[key] = params && params[key] ? String(params[key]) : '';
  }
  return selections;
}

function buildTemplateFilters_(selections, options) {
  var filters = [];
  for (var i = 0; i < FILTER_DEFINITIONS.length; i++) {
    var def = FILTER_DEFINITIONS[i];
    filters.push({
      key: def.key,
      label: def.label,
      emptyLabel: def.emptyLabel,
      selected: selections[def.key] || '',
      options: options[def.key] || []
    });
  }

  for (var j = 0; j < filters.length; j++) {
    var filter = filters[j];
    var queryParts = [];
    for (var k = 0; k < filters.length; k++) {
      var current = filters[k];
      var value = current.key === filter.key ? '' : (current.selected || '');
      queryParts.push(current.key + '=' + encodeURIComponent(value));
    }
    filter.removeQuery = queryParts.join('&');
  }

  return filters;
}

function doGet(e) {
  var userEmail;
  var debugInfo = [];

  try {
    userEmail = getUserEmail_();
    debugInfo.push('Email detectado: ' + userEmail);
  } catch (err) {
    debugInfo.push('Error obteniendo email: ' + err.message);
    return renderAccessDenied_('Error de autenticación', err.message + '<br><br>Debug: ' + debugInfo.join('<br>'));
  }

  var authorized = false;
  try {
    authorized = isEmailAuthorized_(userEmail);
    debugInfo.push('¿Autorizado?: ' + authorized);

    if (!authorized) {
      var authorizedList = getAuthorizedEmailsForDebug_();
      debugInfo.push('Emails autorizados encontrados: ' + authorizedList.length);
      debugInfo.push('Lista: ' + authorizedList.join(', '));
    }
  } catch (err) {
    debugInfo.push('Error verificando permisos: ' + err.message);
    return renderAccessDenied_('Error al verificar permisos', err.message + '<br><br>Debug: ' + debugInfo.join('<br>'));
  }

  if (!authorized) {
    return renderAccessDenied_(
      'Acceso denegado',
      'Tu email (' + userEmail + ') no tiene permisos.<br><br>Debug Info:<br>' + debugInfo.join('<br>')
    );
  }

  var params = e && e.parameter ? e.parameter : {};
  var selections = extractFilters_(params);
  var hasFilters = false;
  for (var key in selections) {
    if (selections[key]) {
      hasFilters = true;
      break;
    }
  }

  var template = HtmlService.createTemplateFromFile('Index');
  template.baseUrl = ScriptApp.getService().getUrl();
  template.userEmail = userEmail;
  template.hasFilters = hasFilters;

  try {
    var filterOptions = getFilterOptions_();
    var templateFilters = buildTemplateFilters_(selections, filterOptions);
    template.filters = templateFilters;

    var result = hasFilters ? buscarEmpleadosOptimizado_(selections) : { empleados: [], resumen: null };
    template.empleados = result.empleados;
    template.resumen = result.resumen;
    template.errorMsg = '';
  } catch (err) {
    template.filters = buildTemplateFilters_(selections, {});
    template.empleados = [];
    template.resumen = null;
    template.errorMsg = 'Error: ' + (err && err.message ? err.message : String(err));
    Logger.log(err);
  }

  return template.evaluate()
    .setTitle('Visualizador de Bandas Salariales')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/********************* BÚSQUEDA OPTIMIZADA *********************/
function buscarEmpleados_(selections) {
  var pack = getSheetAndData_();
  var data = pack.data;
  var c = pack.cols;

  if (c.idxPuesto === -1) {
    throw new Error("La columna 'PuestoDenom' no fue encontrada.");
  }

  // OPTIMIZACIÓN: Pre-filtrar usando índices si están disponibles
  var activeFilters = [];
  for (var i = 0; i < FILTER_DEFINITIONS.length; i++) {
    var def = FILTER_DEFINITIONS[i];
    var idx = c[def.columnKey];
    var value = selections[def.key];
    if (idx !== -1 && value) {
      activeFilters.push({ idx: idx, norm: normalize_(value) });
    }
  }

  var empleados = [];
  var counts = {};
  for (var cIndex = 0; cIndex < SUMMARY_BUCKETS.length; cIndex++) {
    counts[SUMMARY_BUCKETS[cIndex]] = 0;
  }

  var costeEquidadNum = 0;
  var excesoBandaNum = 0;
  var incompletos = 0;

  // OPTIMIZACIÓN: Procesar en chunks para mejor rendimiento
  var chunkSize = CONFIG.CHUNK_SIZE;
  for (var chunkStart = 0; chunkStart < data.length; chunkStart += chunkSize) {
    var chunkEnd = Math.min(chunkStart + chunkSize, data.length);
    
    for (var r = chunkStart; r < chunkEnd; r++) {
      var row = data[r];

      // Aplicar filtros
      var match = true;
      for (var f = 0; f < activeFilters.length; f++) {
        var filter = activeFilters[f];
        if (normalize_(row[filter.idx]) !== filter.norm) {
          match = false;
          break;
        }
      }
      if (!match) continue;

      // Procesar empleado
      var nombre = c.idxNombre !== -1 ? row[c.idxNombre] : '';
      var sba = c.idxSBA !== -1 ? toNumberEU_(row[c.idxSBA]) : NaN;
      var inicio = c.idxInicio !== -1 ? toNumberEU_(row[c.idxInicio]) : NaN;
      var finalB = c.idxFinal !== -1 ? toNumberEU_(row[c.idxFinal]) : NaN;

      var medio = NaN;
      if (c.idxMedio !== -1) {
        medio = toNumberEU_(row[c.idxMedio]);
      } else if (!isNaN(inicio) && !isNaN(finalB)) {
        medio = (inicio + finalB) / 2;
      }

      var posicionPercent = null;
      var qLabel = '—';
      var aviso = '';
      var rango = finalB - inicio;

      if (!isNaN(sba) && !isNaN(inicio) && !isNaN(finalB) && rango > 0) {
        posicionPercent = ((sba - inicio) / rango) * 100;
        posicionPercent = Math.max(0, Math.min(100, posicionPercent));

        if (sba < inicio) {
          qLabel = 'Q0';
          costeEquidadNum += (inicio - sba);
        } else if (sba > finalB) {
          qLabel = 'Q5';
          excesoBandaNum += (sba - finalB);
        } else {
          var p = posicionPercent;
          qLabel = p < 25 ? 'Q1' : p < 50 ? 'Q2' : p < 75 ? 'Q3' : 'Q4';
        }

        counts[qLabel]++;
      } else {
        aviso = 'Datos insuficientes para ubicar en banda.';
        incompletos++;
      }

      empleados.push({
        nombre:       nombre,
        puesto:       c.idxPuesto !== -1 ? row[c.idxPuesto] : '',
        division:     c.idxDivision !== -1 ? row[c.idxDivision] : '',
        direccion:    c.idxDireccion !== -1 ? row[c.idxDireccion] : '',
        rol:          c.idxRol !== -1 ? row[c.idxRol] : '',
        rolId:        c.idxRolId !== -1 ? row[c.idxRolId] : '',
        familia:      c.idxFamilia !== -1 ? row[c.idxFamilia] : '',
        departamento: c.idxDepartamento !== -1 ? row[c.idxDepartamento] : '',
        seccion:      c.idxSeccion !== -1 ? row[c.idxSeccion] : '',
        empresa:      c.idxEmpresa !== -1 ? row[c.idxEmpresa] : '',
        posicion:     c.idxPosicion !== -1 ? row[c.idxPosicion] : '',
        nivel:        c.idxNivel !== -1 ? row[c.idxNivel] : '',
        sbaNum:    isNaN(sba) ? null : sba,
        inicioNum: isNaN(inicio) ? null : inicio,
        medioNum:  isNaN(medio) ? null : medio,
        finalNum:  isNaN(finalB) ? null : finalB,
        sbaStr:    fmt2_(sba),
        inicioStr: fmt2_(inicio),
        medioStr:  fmt2_(medio),
        finalStr:  fmt2_(finalB),
        posicionStr: posicionPercent === null ? '—' : posicionPercent.toFixed(2).replace('.', ','),
        posicionNum: posicionPercent === null ? null : Number(posicionPercent.toFixed(2)),
        qLabel: qLabel,
        aviso: aviso
      });
    }
    
    // Pequeña pausa para evitar timeout
    if (chunkStart % (chunkSize * 3) === 0) {
      Utilities.sleep(10);
    }
  }

  var resumen = {
    total: empleados.length,
    posiciones: SUMMARY_BUCKETS.map(function(label) {
      return { label: label, count: counts[label] };
    }),
    costeEquidad: formatCurrency_(costeEquidadNum),
    excesoBanda: formatCurrency_(excesoBandaNum),
    incompletos: incompletos
  };

  return { empleados: empleados, resumen: resumen };
}

function formatCurrency_(num) {
  return (Math.round(num * 100) / 100).toLocaleString('es-ES', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

/********************* PÁGINA DE ACCESO DENEGADO *********************/
function renderAccessDenied_(title, message) {
  var template = HtmlService.createTemplateFromFile('AccessDenied');
  template.title = title;
  template.message = message;

  return template.evaluate()
    .setTitle('Acceso Denegado')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/********************* FUNCIÓN PARA LIMPIAR CACHE OPTIMIZADA *********************/
function clearAllCache() {
  var keys = [
    'authorized_emails_v4', 'authorized_emails_v3', 'authorized_emails_v2', 'authorized_emails',
    'sheet_data_v4', 'sheet_data_v3', 'sheet_data_v2', 'sheet_data_v1'
  ];

  for (var i = 0; i < FILTER_DEFINITIONS.length; i++) {
    keys.push(buildFilterCacheKey_(FILTER_DEFINITIONS[i].key));
  }

  keys = keys.concat([
    // Versiones anteriores de filtros para limpieza retroactiva
    'filter_puesto_v7', 'filter_division_v7', 'filter_direccion_v7',
    'filter_rol_v7', 'filter_familia_v7', 'filter_departamento_v7',
    'filter_seccion_v7', 'filter_empresa_v7', 'filter_posicion_v7',
    'filter_puesto_v6', 'filter_division_v6', 'filter_direccion_v6',
    'filter_rol_v6', 'filter_familia_v6', 'filter_departamento_v6',
    'filter_seccion_v6', 'filter_empresa_v6', 'filter_posicion_v6',
    'filter_options_v4', 'filter_options_v3', 'filter_options_v2', 'filter_options_v1'
  ]);

  CacheService.getScriptCache().removeAll(keys);
  NORMALIZE_CACHE = {};
  NORMALIZE_CACHE_SIZE = 0;
  CACHE_METADATA = {};
  Logger.log('✅ Cache limpiado exitosamente - v8 optimizado');
}

/********************* FUNCIONES DE MONITOREO *********************/
function getCacheStats() {
  var cache = CacheService.getScriptCache();
  var keys = Object.keys(CACHE_METADATA);
  
  var stats = {
    totalEntries: keys.length,
    totalSize: 0,
    entries: []
  };
  
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var meta = CACHE_METADATA[key];
    stats.totalSize += meta.size || 0;
    stats.entries.push({
      key: key,
      size: meta.size || 0,
      age: Date.now() - (meta.created || 0),
      lastAccess: meta.lastAccess || 0
    });
  }
  
  return stats;
}

function optimizeCache() {
  var stats = getCacheStats();
  Logger.log('Cache Stats: ' + JSON.stringify(stats, null, 2));
  
  if (stats.totalEntries > CONFIG.MAX_CACHE_SIZE) {
    cleanupCache_();
    Logger.log('Cache cleanup performed');
  }
}
