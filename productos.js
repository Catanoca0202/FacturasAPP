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
    let valorUnitario = hojaProductos.getRange(fila, PRODUCT_COLUMNS.VALOR_UNITARIO).getValue();
    let porcientoIva = hojaProductos.getRange(fila, PRODUCT_COLUMNS.TARIFA_IMPUESTO).getDisplayValue();
    let precioConIva = hojaProductos.getRange(fila, PRODUCT_COLUMNS.PRECIO_CON_IMPUESTO).getValue();
    let tipoImpuesto = hojaProductos.getRange(fila, PRODUCT_COLUMNS.TIPO_IMPUESTO).getValue();
    let checkRecargo = hojaProductos.getRange(fila, PRODUCT_COLUMNS.CHECK_RECARGO).getValue();
    let tarifaRecargo = hojaProductos.getRange(fila, PRODUCT_COLUMNS.TARIFA_RECARGO).getDisplayValue();
    let checkRetencion = hojaProductos.getRange(fila, PRODUCT_COLUMNS.CHECK_RETENCION).getValue();
    let tarifaRetencion = hojaProductos.getRange(fila, PRODUCT_COLUMNS.TARIFA_RETENCION).getDisplayValue();
    let estado = hojaProductos.getRange(fila, PRODUCT_COLUMNS.ESTADO).getValue();

    let informacionProducto = {
      "codigo Producto": codigoProducto,
      "regimen": regimen,
      "valor Unitario": valorUnitario,
      "IVA": porcientoIva,
      "precio Con Iva": precioConIva,
      "impuestos": tipoImpuesto,
      "Recargo de equivalencia": checkRecargo === true || checkRecargo === "TRUE" ? tarifaRecargo : "",
      "retencion": checkRetencion === true || checkRetencion === "TRUE" ? tarifaRetencion : "",
      "Estado": estado
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
  
   

  
