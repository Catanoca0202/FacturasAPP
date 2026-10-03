// var spreadsheet = SpreadsheetApp.getActive();

// a cambiar cuando se pregunte y agg los otros porcinetos
function obtenerInformacionProducto(producto) {
    let spreadsheet = SpreadsheetApp.getActive();
    let hojaProductos = spreadsheet.getSheetByName('Productos');
    let ultimaFila = hojaProductos.getLastRow();

    Logger.log("producto dentro de obtener " + producto);

    let fila = -1;
    let identificadores = hojaProductos.getRange(2, PRODUCT_COLUMNS.IDENTIFICADOR_UNICO, ultimaFila - 1, 1).getValues();
    for (let i = 0; i < identificadores.length; i++) {
      if (String(identificadores[i][0]).trim() === String(producto).trim()) {
        fila = i + 2;
        break;
      }
    }
    if (fila === -1) {
      throw new Error("Producto no encontrado: " + producto);
    }

    let codigoProducto = hojaProductos.getRange(fila, PRODUCT_COLUMNS.CODIGO_REFERENCIA).getValue();
    let regimen = hojaProductos.getRange(fila, PRODUCT_COLUMNS.REGIMEN).getValue();
    let operacionExenta = hojaProductos.getRange(fila, PRODUCT_COLUMNS.OPERACION_EXENTA).getValue();
    let valorUnitario = hojaProductos.getRange(fila, PRODUCT_COLUMNS.VALOR_UNITARIO).getValue();
    let porcientoIva = hojaProductos.getRange(fila, PRODUCT_COLUMNS.TARIFA_IMPUESTO).getDisplayValue();
    let precioConIva = hojaProductos.getRange(fila, PRODUCT_COLUMNS.PRECIO_CON_IMPUESTO).getValue();
    let tipoImpuesto = hojaProductos.getRange(fila, PRODUCT_COLUMNS.TIPO_IMPUESTO).getValue();
    let checkRecargo = hojaProductos.getRange(fila, PRODUCT_COLUMNS.CHECK_RECARGO).getValue();
    let tarifaRecargo = hojaProductos.getRange(fila, PRODUCT_COLUMNS.TARIFA_RECARGO).getDisplayValue();
    let checkRetencion = hojaProductos.getRange(fila, PRODUCT_COLUMNS.CHECK_RETENCION).getValue();
    let tarifaRetencionDisplay = hojaProductos.getRange(fila, PRODUCT_COLUMNS.TARIFA_RETENCION).getDisplayValue();
    let porcentajeRetencionDisplay = hojaProductos.getRange(fila, PRODUCT_COLUMNS.PORCENTAJE_RETENCION).getDisplayValue();
    // With "Otros" (RFC 466) the effective rate lives in its own column
    let tarifaRetencion = tarifaRetencionEfectiva_(tarifaRetencionDisplay, porcentajeRetencionDisplay);
    let descripcionRetencion = hojaProductos.getRange(fila, PRODUCT_COLUMNS.DESCRIPCION_RETENCION).getValue();
    let estado = hojaProductos.getRange(fila, PRODUCT_COLUMNS.ESTADO).getValue();
    let tipoProducto = hojaProductos.getRange(fila, PRODUCT_COLUMNS.TIPO_PRODUCTO).getValue();

    // Read Calificación Operación and Exento (checkbox boolean)
    let calificacionOperacion = hojaProductos.getRange(fila, PRODUCT_COLUMNS.CALIFICACION_OPERACION).getValue() || '';
    let exento = hojaProductos.getRange(fila, PRODUCT_COLUMNS.EXENTO).getValue() === true;

    let informacionProducto = {
      "codigo Producto": codigoProducto,
      "regimen": regimen,
      "operacionExenta": operacionExenta || "",
      "valor Unitario": valorUnitario,
      "IVA": porcientoIva,
      "precio Con Iva": precioConIva,
      "impuestos": tipoImpuesto,
      "tipoProducto": String(tipoProducto || '').trim(),
      "Recargo de equivalencia": checkRecargo === true || checkRecargo === "TRUE" ? tarifaRecargo : "",
      "retencion": checkRetencion === true || checkRetencion === "TRUE" ? tarifaRetencion : "",
      "descripcionRetencion": checkRetencion === true || checkRetencion === "TRUE" ? String(descripcionRetencion || "") : "",
      "Estado": estado,
      "calificacionOperacion": String(calificacionOperacion),
      "exento": exento
    };

    return informacionProducto;
  }

  function buscarProductos(terminoBusqueda) {
    var spreadsheet = SpreadsheetApp.getActive();
    var hojaProductos = spreadsheet.getSheetByName('Productos');
    var ultimaFila = hojaProductos.getLastRow();
    if (ultimaFila <= 1) return [];

    var identificadores = hojaProductos.getRange(2, PRODUCT_COLUMNS.IDENTIFICADOR_UNICO, ultimaFila - 1, 1).getValues();
    // Columna A (1): Estado ("Valido"/"No Valido")
    var estados = hojaProductos.getRange(2, 1, ultimaFila - 1, 1).getValues();

    var normalizar = function(s) {
      return String(s || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '');
    };
    var query = normalizar(terminoBusqueda || '');

    var resultados = [];
    for (var i = 0; i < identificadores.length; i++) {
      var estado = String(estados[i][0] || '');
      if (estado !== 'Valido') continue; // Solo productos válidos

      var prod = identificadores[i][0];
      if (!prod) continue;

      var prodNorm = normalizar(prod);
      if (query === '' || prodNorm.includes(query)) {
        resultados.push(String(prod));
      }
    }

    // Limitar resultados para respuestas más ligeras
    return resultados.slice(0, 50);
  }

  /**
   * Batch-fetch product metadata for a list of Identificador Unico values.
   * Returns a map: { identifierKey: { ...product info dict } }
   * Uses 2 bulk sheet reads (getValues + getDisplayValues) for efficiency.
   */
  function obtenerInformacionProductosBatch(listaIds) {
    let result = {};
    if (!listaIds || listaIds.length === 0) return result;

    let spreadsheet = SpreadsheetApp.getActive();
    let hojaProductos = spreadsheet.getSheetByName('Productos');
    let ultimaFila = hojaProductos.getLastRow();
    if (ultimaFila <= 1) return result;

    let numCols = PRODUCT_COLUMNS.IDENTIFICADOR_UNICO;
    let allData = hojaProductos.getRange(2, 1, ultimaFila - 1, numCols).getValues();
    let allDisplay = hojaProductos.getRange(2, 1, ultimaFila - 1, numCols).getDisplayValues();

    // Build a set of requested IDs for fast lookup
    let searchSet = {};
    for (let p = 0; p < listaIds.length; p++) {
      searchSet[String(listaIds[p]).trim()] = true;
    }

    for (let i = 0; i < allData.length; i++) {
      let idKey = String(allData[i][PRODUCT_COLUMNS.IDENTIFICADOR_UNICO - 1]).trim();
      if (!idKey || !searchSet[idKey]) continue;

      let row = allData[i];
      let displayRow = allDisplay[i];
      let checkRecargo = row[PRODUCT_COLUMNS.CHECK_RECARGO - 1];
      let checkRetencion = row[PRODUCT_COLUMNS.CHECK_RETENCION - 1];

      result[idKey] = {
        "codigo Producto": row[PRODUCT_COLUMNS.CODIGO_REFERENCIA - 1],
        "regimen": row[PRODUCT_COLUMNS.REGIMEN - 1],
        "operacionExenta": row[PRODUCT_COLUMNS.OPERACION_EXENTA - 1] || "",
        "valor Unitario": row[PRODUCT_COLUMNS.VALOR_UNITARIO - 1],
        "IVA": displayRow[PRODUCT_COLUMNS.TARIFA_IMPUESTO - 1],
        "precio Con Iva": row[PRODUCT_COLUMNS.PRECIO_CON_IMPUESTO - 1],
        "impuestos": row[PRODUCT_COLUMNS.TIPO_IMPUESTO - 1],
        "tipoProducto": String(row[PRODUCT_COLUMNS.TIPO_PRODUCTO - 1] || '').trim(),
        "Recargo de equivalencia": (checkRecargo === true || checkRecargo === "TRUE") ? displayRow[PRODUCT_COLUMNS.TARIFA_RECARGO - 1] : "",
        "retencion": (checkRetencion === true || checkRetencion === "TRUE")
          ? tarifaRetencionEfectiva_(displayRow[PRODUCT_COLUMNS.TARIFA_RETENCION - 1], displayRow[PRODUCT_COLUMNS.PORCENTAJE_RETENCION - 1])
          : "",
        "descripcionRetencion": (checkRetencion === true || checkRetencion === "TRUE")
          ? String(row[PRODUCT_COLUMNS.DESCRIPCION_RETENCION - 1] || "")
          : "",
        "Estado": row[PRODUCT_COLUMNS.ESTADO - 1],
        "calificacionOperacion": String(row[PRODUCT_COLUMNS.CALIFICACION_OPERACION - 1] || ''),
        "exento": row[PRODUCT_COLUMNS.EXENTO - 1] === true
      };
    }
    return result;
  }

