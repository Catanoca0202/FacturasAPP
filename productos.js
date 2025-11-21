// var spreadsheet = SpreadsheetApp.getActive();

// a cambiar cuando se pregunte y agg los otros porcinetos
function obtenerInformacionProducto(producto) {
    let spreadsheet = SpreadsheetApp.getActive();
    let datos_sheet = spreadsheet.getSheetByName('Datos');
    let celdaProducto = datos_sheet.getRange("I11");
    Logger.log("producto dentro de obtener "+producto)
    celdaProducto.setValue(producto);
  
  
  
    let codigoProducto = datos_sheet.getRange("H11").getValue();
    let valorUnitario = datos_sheet.getRange("J11").getValue();
    let porcientoIva = datos_sheet.getRange("K11").getValue();
    let precioConIva = datos_sheet.getRange("L11").getValue();
    let impuestos = datos_sheet.getRange("M11").getValue();
    let descunetos=datos_sheet.getRange("N11").getValue();
    let retencion=datos_sheet.getRange("O11").getValue();
    let RecgEquivalencia=datos_sheet.getRange("P11").getValue();
    let estado=datos_sheet.getRange("Q11").getValue();
    // Logger.log("Dentro de funcion dict porcientoIva "+ porcientoIva)
    // Logger.log("Dentro de funcion dict porcientoIva sin string"+ datos_sheet.getRange("K11").getValue())
    

    let informacionProducto = {
      "codigo Producto": codigoProducto,
      "valor Unitario": valorUnitario,
      "IVA": porcientoIva,
      "precio Con Iva": precioConIva,
      "impuestos": impuestos,
      "descuentos": descunetos,
      "retencion":retencion,
      "Recargo de equivalencia":RecgEquivalencia,
      "Estado":estado

    };
  
    return informacionProducto;
  }

  function buscarProductos(terminoBusqueda) {
    var spreadsheet = SpreadsheetApp.getActive();
    var hojaProductos = spreadsheet.getSheetByName('Productos');
    var ultimaFila = hojaProductos.getLastRow();
    if (ultimaFila <= 1) return [];

    // Columna N (14): Identificador único "Nombre-Código"
    var identificadores = hojaProductos.getRange(2, 14, ultimaFila - 1, 1).getValues();
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
  
   

  
