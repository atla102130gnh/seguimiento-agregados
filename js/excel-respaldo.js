/*
  ============================================================================
  Generador del Excel de respaldo — mismo formato que el archivo original
  ----------------------------------------------------------------------------
  Lo usan:
    • La página (botones "Exportar" y "Descargar respaldo ahora")
    • GitHub Actions (scripts/respaldo.mjs), cada viernes
  Así el respaldo automático y el manual son idénticos.

  Uso: ExcelRespaldo.libro(XLSX, { agregados, catalogos, generado })
       → devuelve un workbook de SheetJS listo para XLSX.write / writeFile
  ============================================================================
*/
(function (root) {
  'use strict';

  var ENCABEZADOS = [
    'CLIENTE', 'PROVEEDOR', 'CONTRARRECIBO', 'NUR', 'CPT', 'DESCRIPCION',
    'TIPO DE SERVICIO', 'SOLICITANTE', 'FECHA SOLICITUD', 'FECHA DE CARGA',
    'ESTATUS', 'ESTATUS OPERACIONES', 'NOTAS', 'AÑO SURTIDO'
  ];
  var ANCHOS = [18, 42, 15, 16, 16, 52, 22, 20, 15, 15, 16, 30, 40, 12];

  // 'AAAA-MM-DD' → número de serie de Excel (fecha real, no texto)
  function serieExcel(iso) {
    if (!iso) return null;
    var p = String(iso).slice(0, 10).split('-');
    var utc = Date.UTC(+p[0], +p[1] - 1, +p[2]);
    return (utc - Date.UTC(1899, 11, 30)) / 86400000;
  }

  function celdaTexto(v) {
    return { t: 's', v: v == null ? '' : String(v) };
  }
  function celdaFecha(iso) {
    var n = serieExcel(iso);
    return n == null ? { t: 's', v: '' } : { t: 'n', v: n, z: 'dd/mm/yyyy' };
  }
  function celdaNumero(n) {
    return n == null ? { t: 's', v: '' } : { t: 'n', v: Number(n) };
  }

  // NUR y contrarrecibo SIEMPRE como texto: así Excel no los convierte a
  // notación científica (que es como se perdieron 3 NUR en el archivo original).
  function filaAgregado(r, extra) {
    var fila = [
      celdaTexto(r.cliente), celdaTexto(r.proveedor), celdaTexto(r.contrarrecibo),
      celdaTexto(r.nur), celdaTexto(r.cpt), celdaTexto(r.descripcion),
      celdaTexto(r.tipo_servicio), celdaTexto(r.solicitante),
      celdaFecha(r.fecha_solicitud), celdaFecha(r.fecha_carga),
      celdaTexto(r.estatus), celdaTexto(r.estatus_operaciones), celdaTexto(r.notas),
      celdaNumero(r.anio_surtido)
    ];
    return extra ? fila.concat(extra) : fila;
  }

  function hoja(XLSX, encabezados, filas, anchos) {
    var aoa = [encabezados.map(function (h) { return { t: 's', v: h }; })].concat(filas);
    // Se arma celda por celda para controlar el tipo (texto / fecha / número)
    var ws = {};
    var maxC = 0;
    aoa.forEach(function (fila, r) {
      fila.forEach(function (celda, c) {
        ws[XLSX.utils.encode_cell({ r: r, c: c })] = celda;
        if (c > maxC) maxC = c;
      });
    });
    var rango = { s: { r: 0, c: 0 }, e: { r: Math.max(aoa.length - 1, 0), c: Math.max(maxC, encabezados.length - 1) } };
    ws['!ref'] = XLSX.utils.encode_range(rango);
    ws['!cols'] = (anchos || []).map(function (w) { return { wch: w }; });
    if (aoa.length > 1) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: encabezados.length - 1 } }) };
    return ws;
  }

  function ordenOriginal(a, b) {
    // Primero el orden del Excel original, luego lo capturado en el sistema
    var fa = a.fila_excel == null ? Infinity : a.fila_excel;
    var fb = b.fila_excel == null ? Infinity : b.fila_excel;
    if (fa !== fb) return fa - fb;
    return String(a.created_at || '').localeCompare(String(b.created_at || '')) || (a.id - b.id);
  }

  function libro(XLSX, datos) {
    var agregados = (datos.agregados || []).slice().sort(ordenOriginal);
    var cat = datos.catalogos || {};
    var generado = datos.generado || new Date();
    var wb = XLSX.utils.book_new();

    // Una hoja por año surtido (más reciente primero): "AGREGADOS 2026", "AGREGADOS 2027"…
    var activos = agregados.filter(function (r) { return !r.eliminado; });
    var anios = Array.from(new Set(activos.map(function (r) { return r.anio_surtido; }))).sort(function (a, b) { return b - a; });
    if (!anios.length) anios = [new Date().getFullYear()];
    anios.forEach(function (anio) {
      var filas = activos.filter(function (r) { return r.anio_surtido === anio; }).map(function (r) { return filaAgregado(r); });
      XLSX.utils.book_append_sheet(wb, hoja(XLSX, ENCABEZADOS, filas, ANCHOS), 'AGREGADOS ' + anio);
    });

    // Bajas (nunca se borran: aquí quedan con su motivo)
    var bajas = agregados.filter(function (r) { return r.eliminado; }).map(function (r) {
      return filaAgregado(r, [celdaTexto(r.eliminado_en ? String(r.eliminado_en).slice(0, 19).replace('T', ' ') : ''), celdaTexto(r.eliminado_por), celdaTexto(r.motivo_baja)]);
    });
    XLSX.utils.book_append_sheet(wb, hoja(XLSX, ENCABEZADOS.concat(['FECHA DE BAJA', 'BAJA POR', 'MOTIVO DE BAJA']), bajas, ANCHOS.concat([20, 28, 40])), 'BAJAS');

    // Registros que requieren revisión (de la migración o marcados después)
    var revision = activos.filter(function (r) { return r.requiere_revision; }).map(function (r) {
      return filaAgregado(r, [celdaTexto(r.motivo_revision), celdaNumero(r.fila_excel)]);
    });
    XLSX.utils.book_append_sheet(wb, hoja(XLSX, ENCABEZADOS.concat(['MOTIVO DE REVISION', 'FILA EXCEL ORIGINAL']), revision, ANCHOS.concat([60, 12])), 'REVISION');

    // Catálogos
    function nombres(lista) { return (lista || []).map(function (x) { return x.nombre + (x.activo === false ? ' (INACTIVO)' : ''); }).sort(); }
    var cols = [nombres(cat.clientes), nombres(cat.proveedores), nombres(cat.solicitantes), nombres(cat.tipos_servicio)];
    var largo = Math.max.apply(null, cols.map(function (c) { return c.length; }).concat([0]));
    var filasCat = [];
    for (var i = 0; i < largo; i++) filasCat.push(cols.map(function (c) { return celdaTexto(c[i] || ''); }));
    XLSX.utils.book_append_sheet(wb, hoja(XLSX, ['CLIENTES', 'PROVEEDORES', 'SOLICITANTES', 'TIPOS DE SERVICIO'], filasCat, [26, 48, 26, 28]), 'CATALOGOS');

    // Resumen del respaldo
    var cargados = activos.filter(function (r) { return r.estatus === 'CARGADO'; }).length;
    var info = [
      ['Generado', generado.toLocaleString('es-MX', { timeZone: 'America/Mexico_City' })],
      ['Registros activos', activos.length],
      ['Cargados', cargados],
      ['No cargados', activos.length - cargados],
      ['Con más de 4 días sin cargar', activos.filter(function (r) { return r.estatus_operaciones === 'CAMBIO DE ESTATUS OPERACIONES'; }).length],
      ['Bajas', bajas.length],
      ['Requieren revisión', revision.length]
    ].map(function (f) { return [celdaTexto(f[0]), typeof f[1] === 'number' ? celdaNumero(f[1]) : celdaTexto(f[1])]; });
    XLSX.utils.book_append_sheet(wb, hoja(XLSX, ['CONCEPTO', 'VALOR'], info, [32, 26]), 'RESUMEN');

    return wb;
  }

  // Exportación de lo que se ve en la consulta (una sola hoja)
  function consulta(XLSX, filas) {
    var wb = XLSX.utils.book_new();
    var datos = (filas || []).map(function (r) {
      return filaAgregado(r, r.eliminado ? [celdaTexto('BAJA: ' + (r.motivo_baja || ''))] : [celdaTexto('')]);
    });
    XLSX.utils.book_append_sheet(wb, hoja(XLSX, ENCABEZADOS.concat(['OBSERVACION']), datos, ANCHOS.concat([30])), 'CONSULTA');
    return wb;
  }

  function nombreArchivo(fecha) {
    var d = fecha || new Date();
    var iso = new Intl.DateTimeFormat('en-CA', { timeZone: 'America/Mexico_City', year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
    return 'RESPALDO_AGREGADOS_' + iso + '.xlsx';
  }

  root.ExcelRespaldo = { libro: libro, consulta: consulta, nombreArchivo: nombreArchivo, ENCABEZADOS: ENCABEZADOS };
})(typeof globalThis !== 'undefined' ? globalThis : this);
