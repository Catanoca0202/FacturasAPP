// var spreadsheet = SpreadsheetApp.getActive();

// a cambiar cuando se pregunte y agg los otros porcinetos
function obtenerInformacionProducto(producto) {
  let spreadsheet = SpreadsheetApp.getActive();
  let datos_sheet = spreadsheet.getSheetByName('Datos');
  let celdaProducto = datos_sheet.getRange("I11");
  Logger.log("producto dentro de obtener " + producto)
  celdaProducto.setValue(producto);



  let codigoProducto = datos_sheet.getRange("H11").getValue();
  let valorUnitario = datos_sheet.getRange("J11").getValue();
  let porcientoIva = datos_sheet.getRange("K11").getValue();
  let precioConIva = datos_sheet.getRange("L11").getValue();
  let impuestos = datos_sheet.getRange("M11").getValue();
  let descunetos = datos_sheet.getRange("N11").getValue();
  let retencion = datos_sheet.getRange("O11").getValue();
  let RecgEquivalencia = datos_sheet.getRange("P11").getValue();
  let estado = datos_sheet.getRange("Q11").getValue();
  // Logger.log("Dentro de funcion dict porcientoIva "+ porcientoIva)
  // Logger.log("Dentro de funcion dict porcientoIva sin string"+ datos_sheet.getRange("K11").getValue())


  // Normalizar valores vacíos/nulos a 0 para evitar errores en fórmulas
  function normalizeNumberOrZero(raw) {
    try {
      if (raw === null || typeof raw === 'undefined') return 0;
      if (typeof raw === 'number') return raw;
      const s = String(raw).replace('%', '').replace(',', '.').trim();
      if (s === '') return 0;
      const n = Number(s);
      if (isNaN(n)) return 0;
      return n; // Mantener tal cual (0..1, 0..100) según origen; luego se transforma donde se usa
    } catch (_) { return 0; }
  }

  descunetos = normalizeNumberOrZero(descunetos);
  retencion = normalizeNumberOrZero(retencion);
  RecgEquivalencia = normalizeNumberOrZero(RecgEquivalencia);

  // Si quedaron en 0, devolver como cadena "0,0" para fórmulas locales
  if (descunetos === 0) descunetos = "0,0";
  if (retencion === 0) retencion = "0,0";
  if (RecgEquivalencia === 0) RecgEquivalencia = "0,0";

  let informacionProducto = {
    "codigo Producto": codigoProducto,
    "valor Unitario": valorUnitario,
    "IVA": porcientoIva,
    "precio Con Iva": precioConIva,
    "impuestos": impuestos,
    "descuentos": descunetos,
    "retencion": retencion,
    "Recargo de equivalencia": RecgEquivalencia,
    "Estado": estado

  };

  return informacionProducto;
}

function buscarProductos(terminoBusqueda) {
  var spreadsheet = SpreadsheetApp.getActive();
  var hojaProductos = spreadsheet.getSheetByName('Productos');
  var ultimaFila = hojaProductos.getLastRow();
  var valores = hojaProductos.getRange(2, 10, ultimaFila - 1, 1).getValues();

  // Filtrar los productos que coincidan con el término de búsqueda
  var productosFiltrados = valores
    .map(function (row) { return row[0]; })
    .filter(function (producto) {
      // Verificar que 'producto' es una cadena antes de llamar a 'toLowerCase'
      return typeof producto === 'string' && producto.toLowerCase().includes(terminoBusqueda.toLowerCase());
    });

  return productosFiltrados;
}



