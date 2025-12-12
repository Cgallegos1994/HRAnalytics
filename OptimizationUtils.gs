/**
 * Utilidades de optimización avanzada para la aplicación de bandas salariales
 * Incluye funciones para indexación, paginación y procesamiento asíncrono
 */

/********************* ÍNDICES AVANZADOS *********************/
function buildAdvancedIndexes_() {
  var cacheKey = 'advanced_indexes_v1';
  var cached = getCachedObject_(cacheKey);
  if (cached) {
    return cached;
  }

  var pack = getSheetAndData_();
  var data = pack.data;
  var cols = pack.cols;

  var indexes = {
    byPuesto: {},
    byDivision: {},
    byDireccion: {},
    byRol: {},
    byFamilia: {},
    byDepartamento: {},
    bySeccion: {},
    byEmpresa: {},
    byPosicion: {},
    byRolId: {},
    salaryRanges: {
      byRol: {},
      byPuesto: {},
      byDivision: {}
    }
  };

  // Construir índices en chunks
  var chunkSize = CONFIG.CHUNK_SIZE;
  for (var chunkStart = 0; chunkStart < data.length; chunkStart += chunkSize) {
    var chunkEnd = Math.min(chunkStart + chunkSize, data.length);
    
    for (var r = chunkStart; r < chunkEnd; r++) {
      var row = data[r];
      var rowIndex = r;

      // Índices básicos
      if (cols.idxPuesto !== -1 && row[cols.idxPuesto]) {
        var puesto = normalize_(row[cols.idxPuesto]);
        if (!indexes.byPuesto[puesto]) indexes.byPuesto[puesto] = [];
        indexes.byPuesto[puesto].push(rowIndex);
      }

      if (cols.idxDivision !== -1 && row[cols.idxDivision]) {
        var division = normalize_(row[cols.idxDivision]);
        if (!indexes.byDivision[division]) indexes.byDivision[division] = [];
        indexes.byDivision[division].push(rowIndex);
      }

      if (cols.idxDireccion !== -1 && row[cols.idxDireccion]) {
        var direccion = normalize_(row[cols.idxDireccion]);
        if (!indexes.byDireccion[direccion]) indexes.byDireccion[direccion] = [];
        indexes.byDireccion[direccion].push(rowIndex);
      }

      if (cols.idxRol !== -1 && row[cols.idxRol]) {
        var rol = normalize_(row[cols.idxRol]);
        if (!indexes.byRol[rol]) indexes.byRol[rol] = [];
        indexes.byRol[rol].push(rowIndex);
      }

      if (cols.idxFamilia !== -1 && row[cols.idxFamilia]) {
        var familia = normalize_(row[cols.idxFamilia]);
        if (!indexes.byFamilia[familia]) indexes.byFamilia[familia] = [];
        indexes.byFamilia[familia].push(rowIndex);
      }

      if (cols.idxDepartamento !== -1 && row[cols.idxDepartamento]) {
        var departamento = normalize_(row[cols.idxDepartamento]);
        if (!indexes.byDepartamento[departamento]) indexes.byDepartamento[departamento] = [];
        indexes.byDepartamento[departamento].push(rowIndex);
      }

      if (cols.idxSeccion !== -1 && row[cols.idxSeccion]) {
        var seccion = normalize_(row[cols.idxSeccion]);
        if (!indexes.bySeccion[seccion]) indexes.bySeccion[seccion] = [];
        indexes.bySeccion[seccion].push(rowIndex);
      }

      if (cols.idxEmpresa !== -1 && row[cols.idxEmpresa]) {
        var empresa = normalize_(row[cols.idxEmpresa]);
        if (!indexes.byEmpresa[empresa]) indexes.byEmpresa[empresa] = [];
        indexes.byEmpresa[empresa].push(rowIndex);
      }

      if (cols.idxPosicion !== -1 && row[cols.idxPosicion]) {
        var posicion = normalize_(row[cols.idxPosicion]);
        if (!indexes.byPosicion[posicion]) indexes.byPosicion[posicion] = [];
        indexes.byPosicion[posicion].push(rowIndex);
      }

      if (cols.idxRolId !== -1 && row[cols.idxRolId]) {
        var rolId = normalize_(row[cols.idxRolId]);
        if (!indexes.byRolId[rolId]) indexes.byRolId[rolId] = [];
        indexes.byRolId[rolId].push(rowIndex);
      }

      // Índices de rangos salariales
      var sba = toNumberEU_(row[cols.idxSBA]);
      var inicio = toNumberEU_(row[cols.idxInicio]);
      var final = toNumberEU_(row[cols.idxFinal]);

      if (!isNaN(sba) && !isNaN(inicio) && !isNaN(final)) {
        var rawRol = cols.idxRol !== -1 ? normalize_(row[cols.idxRol]) : '';
        var rawPuesto = cols.idxPuesto !== -1 ? normalize_(row[cols.idxPuesto]) : '';
        var rawDivision = cols.idxDivision !== -1 ? normalize_(row[cols.idxDivision]) : '';

        if (rawRol) {
          if (!indexes.salaryRanges.byRol[rawRol]) {
            indexes.salaryRanges.byRol[rawRol] = { min: Infinity, max: -Infinity, count: 0, total: 0 };
          }
          indexes.salaryRanges.byRol[rawRol].min = Math.min(indexes.salaryRanges.byRol[rawRol].min, inicio);
          indexes.salaryRanges.byRol[rawRol].max = Math.max(indexes.salaryRanges.byRol[rawRol].max, final);
          indexes.salaryRanges.byRol[rawRol].count++;
          indexes.salaryRanges.byRol[rawRol].total += sba;
        }

        if (rawPuesto) {
          if (!indexes.salaryRanges.byPuesto[rawPuesto]) {
            indexes.salaryRanges.byPuesto[rawPuesto] = { min: Infinity, max: -Infinity, count: 0, total: 0 };
          }
          indexes.salaryRanges.byPuesto[rawPuesto].min = Math.min(indexes.salaryRanges.byPuesto[rawPuesto].min, inicio);
          indexes.salaryRanges.byPuesto[rawPuesto].max = Math.max(indexes.salaryRanges.byPuesto[rawPuesto].max, final);
          indexes.salaryRanges.byPuesto[rawPuesto].count++;
          indexes.salaryRanges.byPuesto[rawPuesto].total += sba;
        }

        if (rawDivision) {
          if (!indexes.salaryRanges.byDivision[rawDivision]) {
            indexes.salaryRanges.byDivision[rawDivision] = { min: Infinity, max: -Infinity, count: 0, total: 0 };
          }
          indexes.salaryRanges.byDivision[rawDivision].min = Math.min(indexes.salaryRanges.byDivision[rawDivision].min, inicio);
          indexes.salaryRanges.byDivision[rawDivision].max = Math.max(indexes.salaryRanges.byDivision[rawDivision].max, final);
          indexes.salaryRanges.byDivision[rawDivision].count++;
          indexes.salaryRanges.byDivision[rawDivision].total += sba;
        }
      }
    }

    // Pequeña pausa para evitar timeout
    if (chunkStart % (chunkSize * 2) === 0) {
      Utilities.sleep(10);
    }
  }

  putCachedObject_(cacheKey, indexes, CONFIG.CACHE.INDEXES);
  return indexes;
}

/********************* BÚSQUEDA OPTIMIZADA CON ÍNDICES *********************/
function buscarEmpleadosOptimizado_(selections) {
  var pack = getSheetAndData_();
  var data = pack.data;
  var c = pack.cols;

  if (c.idxPuesto === -1) {
    throw new Error("La columna 'PuestoDenom' no fue encontrada.");
  }

  // Obtener índices
  var indexes = buildAdvancedIndexes_();

  // Determinar filas candidatas usando índices
  var candidateRows = null;
  var hasActiveFilters = false;

  for (var i = 0; i < FILTER_DEFINITIONS.length; i++) {
    var def = FILTER_DEFINITIONS[i];
    var value = selections[def.key];
    if (!value) continue;

    hasActiveFilters = true;
    var normalizedValue = normalize_(value);
    var indexKey = def.columnKey.replace('idx', 'by');
    
    if (indexes[indexKey] && indexes[indexKey][normalizedValue]) {
      var rowsForFilter = indexes[indexKey][normalizedValue];
      
      if (candidateRows === null) {
        candidateRows = new Set(rowsForFilter);
      } else {
        var newCandidates = new Set();
        for (var j = 0; j < rowsForFilter.length; j++) {
          if (candidateRows.has(rowsForFilter[j])) {
            newCandidates.add(rowsForFilter[j]);
          }
        }
        candidateRows = newCandidates;
      }
    } else {
      return { empleados: [], resumen: createEmptyResumen_() };
    }
  }

  if (!hasActiveFilters) {
    candidateRows = null;
  }

  var rowsToProcess = candidateRows ? Array.from(candidateRows) : null;

  return procesarEmpleadosConCandidatos_(data, c, rowsToProcess);
}

function procesarEmpleadosConCandidatos_(data, cols, candidateRows) {
  var empleados = [];
  var counts = {};
  for (var cIndex = 0; cIndex < SUMMARY_BUCKETS.length; cIndex++) {
    counts[SUMMARY_BUCKETS[cIndex]] = 0;
  }

  var costeEquidadNum = 0;
  var excesoBandaNum = 0;
  var incompletos = 0;

  var rowsToProcess = candidateRows || [];
  var shouldProcessAll = candidateRows === null;

  if (shouldProcessAll) {
    for (var r = 0; r < data.length; r++) {
      var result = procesarEmpleado_(data[r], r, cols);
      if (result) {
        empleados.push(result.empleado);
        counts[result.qLabel]++;
        if (result.costeEquidad > 0) costeEquidadNum += result.costeEquidad;
        if (result.excesoBanda > 0) excesoBandaNum += result.excesoBanda;
        if (result.incompleto) incompletos++;
      }
    }
  } else {
    for (var i = 0; i < rowsToProcess.length; i++) {
      var r = rowsToProcess[i];
      var result = procesarEmpleado_(data[r], r, cols);
      if (result) {
        empleados.push(result.empleado);
        counts[result.qLabel]++;
        if (result.costeEquidad > 0) costeEquidadNum += result.costeEquidad;
        if (result.excesoBanda > 0) excesoBandaNum += result.excesoBanda;
        if (result.incompleto) incompletos++;
      }
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

function procesarEmpleado_(row, rowIndex, cols) {
  var nombre = cols.idxNombre !== -1 ? row[cols.idxNombre] : '';
  var sba = cols.idxSBA !== -1 ? toNumberEU_(row[cols.idxSBA]) : NaN;
  var inicio = cols.idxInicio !== -1 ? toNumberEU_(row[cols.idxInicio]) : NaN;
  var finalB = cols.idxFinal !== -1 ? toNumberEU_(row[cols.idxFinal]) : NaN;

  var medio = NaN;
  if (cols.idxMedio !== -1) {
    medio = toNumberEU_(row[cols.idxMedio]);
  } else if (!isNaN(inicio) && !isNaN(finalB)) {
    medio = (inicio + finalB) / 2;
  }

  var posicionPercent = null;
  var qLabel = '—';
  var aviso = '';
  var rango = finalB - inicio;
  var costeEquidad = 0;
  var excesoBanda = 0;
  var incompleto = false;

  if (!isNaN(sba) && !isNaN(inicio) && !isNaN(finalB) && rango > 0) {
    posicionPercent = ((sba - inicio) / rango) * 100;
    posicionPercent = Math.max(0, Math.min(100, posicionPercent));

    if (sba < inicio) {
      qLabel = 'Q0';
      costeEquidad = inicio - sba;
    } else if (sba > finalB) {
      qLabel = 'Q5';
      excesoBanda = sba - finalB;
    } else {
      var p = posicionPercent;
      qLabel = p < 25 ? 'Q1' : p < 50 ? 'Q2' : p < 75 ? 'Q3' : 'Q4';
    }
  } else {
    aviso = 'Datos insuficientes para ubicar en banda.';
    incompleto = true;
  }

  var empleado = {
    nombre:       nombre,
    puesto:       cols.idxPuesto !== -1 ? row[cols.idxPuesto] : '',
    division:     cols.idxDivision !== -1 ? row[cols.idxDivision] : '',
    direccion:    cols.idxDireccion !== -1 ? row[cols.idxDireccion] : '',
    rol:          cols.idxRol !== -1 ? row[cols.idxRol] : '',
    rolId:        cols.idxRolId !== -1 ? row[cols.idxRolId] : '',
    familia:      cols.idxFamilia !== -1 ? row[cols.idxFamilia] : '',
    departamento: cols.idxDepartamento !== -1 ? row[cols.idxDepartamento] : '',
    seccion:      cols.idxSeccion !== -1 ? row[cols.idxSeccion] : '',
    empresa:      cols.idxEmpresa !== -1 ? row[cols.idxEmpresa] : '',
    posicion:     cols.idxPosicion !== -1 ? row[cols.idxPosicion] : '',
    nivel:        cols.idxNivel !== -1 ? row[cols.idxNivel] : '',
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
  };

  return {
    empleado: empleado,
    qLabel: qLabel,
    costeEquidad: costeEquidad,
    excesoBanda: excesoBanda,
    incompleto: incompleto
  };
}

function createEmptyResumen_() {
  return {
    total: 0,
    posiciones: SUMMARY_BUCKETS.map(function(label) {
      return { label: label, count: 0 };
    }),
    costeEquidad: formatCurrency_(0),
    excesoBanda: formatCurrency_(0),
    incompletos: 0
  };
}

/********************* PAGINACIÓN *********************/
function buscarEmpleadosPaginado_(selections, page, pageSize) {
  page = page || 1;
  pageSize = pageSize || 50;
  
  var result = buscarEmpleadosOptimizado_(selections);
  var empleados = result.empleados;
  
  var totalPages = Math.ceil(empleados.length / pageSize);
  var startIndex = (page - 1) * pageSize;
  var endIndex = Math.min(startIndex + pageSize, empleados.length);
  
  var paginatedEmpleados = empleados.slice(startIndex, endIndex);
  
  return {
    empleados: paginatedEmpleados,
    resumen: result.resumen,
    pagination: {
      currentPage: page,
      totalPages: totalPages,
      pageSize: pageSize,
      totalItems: empleados.length,
      hasNext: page < totalPages,
      hasPrevious: page > 1
    }
  };
}

/********************* FUNCIONES DE MONITOREO AVANZADO *********************/
function getPerformanceStats() {
  var cache = CacheService.getScriptCache();
  var stats = {
    cache: getCacheStats(),
    memory: {
      normalizeCacheSize: NORMALIZE_CACHE_SIZE,
      maxNormalizeCacheSize: MAX_NORMALIZE_CACHE_SIZE
    },
    config: {
      chunkSize: CONFIG.CHUNK_SIZE,
      maxCacheSize: CONFIG.MAX_CACHE_SIZE
    },
    timestamp: new Date().toISOString()
  };
  
  return stats;
}

function clearAllCacheAdvanced() {
  clearAllCache();
  
  var indexKeys = [
    'advanced_indexes_v1',
    'salary_ranges_v1',
    'filter_indexes_v1'
  ];
  
  CacheService.getScriptCache().removeAll(indexKeys);
  
  Logger.log('✅ Cache avanzado limpiado exitosamente');
}

/********************* FUNCIÓN DE PRUEBA DE RENDIMIENTO *********************/
function testPerformance() {
  var startTime = Date.now();
  
  var dataStart = Date.now();
  var pack = getSheetAndData_();
  var dataTime = Date.now() - dataStart;
  
  var indexStart = Date.now();
  var indexes = buildAdvancedIndexes_();
  var indexTime = Date.now() - indexStart;
  
  var filterStart = Date.now();
  var filters = getFilterOptions_();
  var filterTime = Date.now() - filterStart;
  
  var totalTime = Date.now() - startTime;
  
  var results = {
    dataLoadTime: dataTime + 'ms',
    indexBuildTime: indexTime + 'ms',
    filterBuildTime: filterTime + 'ms',
    totalTime: totalTime + 'ms',
    dataRows: pack.data.length,
    indexEntries: Object.keys(indexes.byPuesto).length + Object.keys(indexes.byDivision).length,
    cacheEntries: Object.keys(CACHE_METADATA).length
  };
  
  Logger.log('Performance Test Results: ' + JSON.stringify(results, null, 2));
  return results;
}
