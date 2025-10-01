PREFACTURA_ROW = 3;
PREFACTURA_COLUMN = 2;
COL_TOTALES_PREFACTURA = 11;// K
FILA_INICIAL_PREFACTURA = 8;
COLUMNA_FINAL = 50;
ADDITIONAL_ROWS = 3 + 3; //(Personalizacion)


// var spreadsheet = SpreadsheetApp.getActive();
// var prefactura_sheet = spreadsheet.getSheetByName('Factura');
// var unidades_sheet = spreadsheet.getSheetByName('Unidades');
// var listadoestado_sheet = spreadsheet.getSheetByName('ListadoEstado');
// var hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
// var folderId = hojaDatosEmisor.getRange("B14").getValue();
var paisesCodigos = {
  "Afganistán": "AF",
  "Albania": "AL",
  "Alemania": "DE",
  "Andorra": "AD",
  "Angola": "AO",
  "Antigua y Barbuda": "AG",
  "Arabia Saudita": "SA",
  "Argelia": "DZ",
  "Argentina": "AR",
  "Armenia": "AM",
  "Australia": "AU",
  "Austria": "AT",
  "Azerbaiyán": "AZ",
  "Bahamas": "BS",
  "Bangladés": "BD",
  "Barbados": "BB",
  "Baréin": "BH",
  "Bélgica": "BE",
  "Belice": "BZ",
  "Benín": "BJ",
  "Bielorrusia": "BY",
  "Birmania": "MM",
  "Bolivia": "BO",
  "Bosnia y Herzegovina": "BA",
  "Botsuana": "BW",
  "Brasil": "BR",
  "Brunéi": "BN",
  "Bulgaria": "BG",
  "Burkina Faso": "BF",
  "Burundi": "BI",
  "Bután": "BT",
  "Cabo Verde": "CV",
  "Camboya": "KH",
  "Camerún": "CM",
  "Canadá": "CA",
  "Catar": "QA",
  "Chad": "TD",
  "Chile": "CL",
  "China": "CN",
  "Chipre": "CY",
  "Ciudad del Vaticano": "VA",
  "Colombia": "CO",
  "Comoras": "KM",
  "Corea del Norte": "KP",
  "Corea del Sur": "KR",
  "Costa de Marfil": "CI",
  "Costa Rica": "CR",
  "Croacia": "HR",
  "Cuba": "CU",
  "Dinamarca": "DK",
  "Dominica": "DM",
  "Ecuador": "EC",
  "Egipto": "EG",
  "El Salvador": "SV",
  "Emiratos Árabes Unidos": "AE",
  "Eritrea": "ER",
  "Eslovaquia": "SK",
  "Eslovenia": "SI",
  "España": "ES",
  "Estados Unidos": "US",
  "Estonia": "EE",
  "Etiopía": "ET",
  "Filipinas": "PH",
  "Finlandia": "FI",
  "Fiyi": "FJ",
  "Francia": "FR",
  "Gabón": "GA",
  "Gambia": "GM",
  "Georgia": "GE",
  "Ghana": "GH",
  "Granada": "GD",
  "Grecia": "GR",
  "Guatemala": "GT",
  "Guyana": "GY",
  "Guinea": "GN",
  "Guinea ecuatorial": "GQ",
  "Guinea-Bisáu": "GW",
  "Haití": "HT",
  "Honduras": "HN",
  "Hungría": "HU",
  "India": "IN",
  "Indonesia": "ID",
  "Irak": "IQ",
  "Irán": "IR",
  "Irlanda": "IE",
  "Islandia": "IS",
  "Islas Marshall": "MH",
  "Islas Salomón": "SB",
  "Israel": "IL",
  "Italia": "IT",
  "Jamaica": "JM",
  "Japón": "JP",
  "Jordania": "JO",
  "Kazajistán": "KZ",
  "Kenia": "KE",
  "Kirguistán": "KG",
  "Kiribati": "KI",
  "Kosovo": "XK",
  "Kuwait": "KW",
  "Laos": "LA",
  "Lesoto": "LS",
  "Letonia": "LV",
  "Líbano": "LB",
  "Liberia": "LR",
  "Libia": "LY",
  "Liechtenstein": "LI",
  "Lituania": "LT",
  "Luxemburgo": "LU",
  "Macedonia del Norte": "MK",
  "Madagascar": "MG",
  "Malasia": "MY",
  "Malaui": "MW",
  "Maldivas": "MV",
  "Malí": "ML",
  "Malta": "MT",
  "Marruecos": "MA",
  "Mauricio": "MU",
  "Mauritania": "MR",
  "México": "MX",
  "Micronesia": "FM",
  "Moldavia": "MD",
  "Mónaco": "MC",
  "Mongolia": "MN",
  "Montenegro": "ME",
  "Mozambique": "MZ",
  "Namibia": "NA",
  "Nauru": "NR",
  "Nepal": "NP",
  "Nicaragua": "NI",
  "Níger": "NE",
  "Nigeria": "NG",
  "Noruega": "NO",
  "Nueva Zelanda": "NZ",
  "Omán": "OM",
  "Países Bajos": "NL",
  "Pakistán": "PK",
  "Palaos": "PW",
  "Panamá": "PA",
  "Papúa Nueva Guinea": "PG",
  "Paraguay": "PY",
  "Perú": "PE",
  "Polonia": "PL",
  "Portugal": "PT",
  "Reino Unido": "GB",
  "República Centroafricana": "CF",
  "República Checa": "CZ",
  "República del Congo": "CG",
  "República Democrática del Congo": "CD",
  "República Dominicana": "DO",
  "Ruanda": "RW",
  "Rumania": "RO",
  "Rusia": "RU",
  "Samoa": "WS",
  "San Cristóbal y Nieves": "KN",
  "San Marino": "SM",
  "San Vicente y las Granadinas": "VC",
  "Santa Lucía": "LC",
  "Santo Tomé y Príncipe": "ST",
  "Senegal": "SN",
  "Serbia": "RS",
  "Seychelles": "SC",
  "Sierra Leona": "SL",
  "Singapur": "SG",
  "Siria": "SY",
  "Somalia": "SO",
  "Sri Lanka": "LK",
  "Suazilandia": "SZ",
  "Sudáfrica": "ZA",
  "Sudán": "SD",
  "Sudán del Sur": "SS",
  "Suecia": "SE",
  "Suiza": "CH",
  "Surinam": "SR",
  "Tailandia": "TH",
  "Tanzania": "TZ",
  "Tayikistán": "TJ",
  "Timor Oriental": "TL",
  "Togo": "TG",
  "Tonga": "TO",
  "Trinidad y Tobago": "TT",
  "Túnez": "TN",
  "Turkmenistán": "TM",
  "Turquía": "TR",
  "Tuvalu": "TV",
  "Ucrania": "UA",
  "Uganda": "UG",
  "Uruguay": "UY",
  "Uzbekistán": "UZ",
  "Vanuatu": "VU",
  "Venezuela": "VE",
  "Vietnam": "VN",
  "Yemen": "YE",
  "Yibuti": "DJ",
  "Zambia": "ZM",
  "Zimbabue": "ZW"
};

var diccionarioCaluclarIva = {
  "0.21": "21,00",
  "0.1": "1,00",
  "0.05": "5,00",
  "0.04": "4,00",
  "0": 0
}

function verificarEstadoValidoFactura() {
  const spreadsheet = SpreadsheetApp.getActive();
  const hojaFactura = spreadsheet.getSheetByName('Factura');

  let estaValido = { success: true, message: "" };

  // Campos a verificar
  const clienteActual = hojaFactura.getRange("B2").getValue();
  const numFactura = hojaFactura.getRange("G2").getValue();
  const fechaPago = hojaFactura.getRange("G3").getValue();
  const fechaEmision = hojaFactura.getRange("G4").getValue();
  const formaPago = hojaFactura.getRange("G5").getValue();


  // Verificar cliente
  if (!clienteActual || clienteActual === "") {
    estaValido.success = false;
    estaValido.message = "El cliente actual no está definido.";
    return estaValido;
  }

  // Verificar número de factura
  if (!numFactura || numFactura === "") {
    estaValido.success = false;
    estaValido.message = "El número de factura no está definido.";
    return estaValido;
  }

  // Verificar fecha de pago
  if (!fechaPago || fechaPago === "") {
    estaValido.success = false;
    estaValido.message = "La fecha de pago no está definida.";
    return estaValido;
  }

  // Verificar fecha de emisión
  if (!fechaEmision || fechaEmision === "") {
    estaValido.success = false;
    estaValido.message = "La fecha de emisión no está definida.";
    return estaValido;
  }

  // Verificar que la fecha de emisión no sea posterior a la fecha de pago
  if (new Date(fechaEmision) > new Date(fechaPago)) {
    estaValido.success = false;
    estaValido.message = "La fecha de emisión no puede ser posterior a la fecha de pago.";
    return estaValido;
  }

  // Verificar forma de pago
  if (!formaPago || formaPago === "") {
    estaValido.success = false;
    estaValido.message = "La forma de pago no está definida.";
    return estaValido;
  }



  // Verificar productos
  const totalProductos = hojaFactura.getRange("A16").getValue();
  if (totalProductos === "Total filas") {
    const valorTotalProductos = hojaFactura.getRange("B16").getValue();
    if (valorTotalProductos === 0 || valorTotalProductos === "") {
      estaValido.success = false;
      estaValido.message = "No se han agregado productos a la factura.";
      return estaValido;
    }
  }

  // Si pasa todas las validaciones, está válido
  estaValido.success = true;
  estaValido.message = "Factura válida para guardar.";
  return estaValido;
}



function verificarEstadoValidoFactura() {
  const spreadsheet = SpreadsheetApp.getActive();
  const hojaFactura = spreadsheet.getSheetByName('Factura');

  let estaValido = { success: true, message: "" };

  // Campos a verificar
  const clienteActual = hojaFactura.getRange("B2").getValue();
  const numFactura = hojaFactura.getRange("G2").getValue();
  const fechaPago = hojaFactura.getRange("G3").getValue();
  const fechaEmision = hojaFactura.getRange("G4").getValue();
  const formaPago = hojaFactura.getRange("G5").getValue();
  const asesor = hojaFactura.getRange("G8").getValue();


  // Verificar cliente
  if (!clienteActual || clienteActual === "") {
    estaValido.success = false;
    estaValido.message = "El cliente actual no está definido.";
    return estaValido;
  }

  // Verificar número de factura
  if (!numFactura || numFactura === "") {
    estaValido.success = false;
    estaValido.message = "El número de factura no está definido.";
    return estaValido;
  }

  // Verificar fecha de pago
  if (!fechaPago || fechaPago === "") {
    estaValido.success = false;
    estaValido.message = "La fecha de pago no está definida.";
    return estaValido;
  }

  // Verificar fecha de emisión
  if (!fechaEmision || fechaEmision === "") {
    estaValido.success = false;
    estaValido.message = "La fecha de emisión no está definida.";
    return estaValido;
  }

  // Verificar que la fecha de emisión no sea posterior a la fecha de pago
  if (new Date(fechaEmision) > new Date(fechaPago)) {
    estaValido.success = false;
    estaValido.message = "La fecha de emisión no puede ser posterior a la fecha de pago.";
    return estaValido;
  }

  // Verificar forma de pago
  if (!formaPago || formaPago === "") {
    estaValido.success = false;
    estaValido.message = "La forma de pago no está definida.";
    return estaValido;
  }

  if (!asesor || asesor === "") {
    estaValido.success = false;
    estaValido.message = "Asesor no está definida. Si no tienes asesor, escribe el nombre del contacto en la casilla correspodiente ";
    return estaValido;
  }



  // Verificar productos
  const totalProductos = hojaFactura.getRange("A16").getValue();
  if (totalProductos === "Total filas") {
    const valorTotalProductos = hojaFactura.getRange("B16").getValue();
    if (valorTotalProductos === 0 || valorTotalProductos === "") {
      estaValido.success = false;
      estaValido.message = "No se han agregado productos a la factura.";
      return estaValido;
    }
  }

  // Si pasa todas las validaciones, está válido
  estaValido.success = true;
  estaValido.message = "Factura válida para guardar.";
  return estaValido;
}


function guardarFactura(){
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let estadoVinculacion = hojaDatosEmisor.getRange("B16").getValue();
  let estadoFactura = verificarEstadoValidoFactura();
  try {
    if (estadoVinculacion == "Desvinculado") {
      SpreadsheetApp.getUi().alert("Recuerda que antes de poder generar una factura es necesario haber vinculado tu cuenta de FacturasApp");
      try { showVincularParaEnviar(); } catch (_) {}
      return;
    }
    if (estadoFactura.success) {
      // Validaciones previas a guardar: solo validar consecutivo
      let consecutivoOk = verificarEstadoConsecutivo();
      if (consecutivoOk) {
        guardarYGenerarInvoice();
        guardarFacturaHistorial();
        Logger.log("guardar factura");
        enviarFactura();
        limpiarHojaFactura();
        return;
      } else {
        SpreadsheetApp.getUi().alert("No fue posible guardar la factura. Verifica el consecutivo configurado.");
        return;
      }
    } else {
      SpreadsheetApp.getUi().alert("Error al generar factura. " + estadoFactura.message);
      return;
    }
  } catch (error) {
    let mensaje = String(error && error.message ? error.message : error);
    if (/TypePerson/i.test(mensaje)) {
      mensaje = "Error en los datos del cliente: Tipo de persona inválido. Debe ser 'Autonomo' o 'Empresa'. Verifica la columna 'Tipo de persona' en la hoja Clientes.";
    }
    SpreadsheetApp.getUi().alert("No se pudo guardar/enviar la factura. " + mensaje);
    Logger.log("guardarFactura error: " + mensaje);
  }
}

function agregarFilaNueva() {
  // 1) Obtener el candado
  const lock = LockService.getScriptLock();
  try {
    // Esperar hasta 5s para obtener el candado
    lock.waitLock(6000);

    // --- AQUÍ PONES TU LÓGICA ---
    var spreadsheet = SpreadsheetApp.getActive();
    let hojaFactura = spreadsheet.getSheetByName('Factura');
    let numeroFilasParaAgregar = hojaFactura.getRange("B13").getValue();

    // Verificar si numeroFilasParaAgregar es nulo, vacío o no es un número
    if (numeroFilasParaAgregar == 0 || numeroFilasParaAgregar == "" || isNaN(numeroFilasParaAgregar)) {
      SpreadsheetApp.getUi().alert("Error: Por favor, ingresa un número válido de filas para agregar.");
      return; // Detener la ejecución si hay error
    }

    let taxSectionStartRow = getTaxSectionStartRow(hojaFactura);
    const productStartRow = 15;
    const lastProductRow = getLastProductRow(hojaFactura, productStartRow, taxSectionStartRow);

    Logger.log("Agregar fila nueva");
    hojaFactura.insertRows(lastProductRow, numeroFilasParaAgregar);

    // Si quieres asegurar que la hoja ya refleje los cambios antes de salir:
    SpreadsheetApp.flush();

  } catch (err) {
    // Si no se pudo conseguir el lock en 5 seg o hay otro error
    Logger.log("Error en agregarFilaNueva: " + err);
  } finally {
    // 2) Liberar el candado
    lock.releaseLock();
  }
}

function agregarProductoDesdeFactura(cantidad, producto) {
  var spreadsheet = SpreadsheetApp.getActive();
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let factura_sheet = spreadsheet.getSheetByName('Factura');
  let taxSectionStartRow = getTaxSectionStartRow(hojaFactura);//recordar este devuelve el lugar en donde deberian estar base imponible, toca restar -1
  const productStartRow = 15;
  const lastProductRow = getLastProductRow(hojaFactura, productStartRow, taxSectionStartRow);

  let dictInformacionProducto = {}
  if (producto === "" || cantidad === "" || cantidad === 0) {
    throw new Error('Porfavor elige un producto y un cantidad adecuado');
  } else {
    Logger.log("entra a dictInformacionProducto")
    dictInformacionProducto = obtenerInformacionProducto(producto);
  }
  Logger.log("dictInformacionProducto " + dictInformacionProducto["codigo Producto"])
  let rowParaDatos = lastProductRow
  let rowParaTotalTaxes = taxSectionStartRow
  let cantidadProductos = hojaFactura.getRange("B16").getValue()//estado defaul de total productos
  if (cantidadProductos === 0 || cantidadProductos === "") {
    factura_sheet.getRange("A15").setValue(dictInformacionProducto["codigo Producto"])
    factura_sheet.getRange("B15").setValue(producto)
    factura_sheet.getRange("C15").setValue(cantidad)
    factura_sheet.getRange("D15").setValue(dictInformacionProducto["valor Unitario"])
    // Completar fórmulas base con separador decimal coma
    factura_sheet.getRange("E15").setValue("=(D15*C15)") // Subtotal
    const ivaForFormula = String(dictInformacionProducto["IVA"]).replace(".", ",");
    const retForFormula = String(dictInformacionProducto["retencion"]).replace(".", ",");
    const recForFormula = String(dictInformacionProducto["Recargo de equivalencia"]).replace(".", ",");
    factura_sheet.getRange("F15").setValue("=(E15*" + ivaForFormula + ")") // Impuestos
    if (retForFormula !== "0" && retForFormula !== "0,0") {
      factura_sheet.getRange("G15").setValue("=(E15*" + retForFormula + ")") // Retención
    }
    if (recForFormula !== "0" && recForFormula !== "0,0") {
      factura_sheet.getRange("I15").setValue("=(E15*" + recForFormula + ")") // Recargo de equivalencia
    }
    // Importe línea = Subtotal + IVA + Recargo (sin retenciones ni descuentos)
    factura_sheet.getRange("J15").setValue("=(E15+F15+I15)")

  } else {
    hojaFactura.insertRowAfter(lastProductRow)
    rowParaTotalTaxes = taxSectionStartRow + 1
    rowParaDatos = lastProductRow + 1
    factura_sheet.getRange("A" + String(rowParaDatos)).setValue(dictInformacionProducto["codigo Producto"])
    factura_sheet.getRange("B" + String(rowParaDatos)).setValue(producto)
    factura_sheet.getRange("C" + String(rowParaDatos)).setValue(cantidad)
    factura_sheet.getRange("D" + String(rowParaDatos)).setValue(dictInformacionProducto["valor Unitario"])//valor unitario
    factura_sheet.getRange("E" + String(rowParaDatos)).setValue("=(D" + String(rowParaDatos) + "*C" + String(rowParaDatos) + ")")//subtotal : valor unitario * cantidad
    const ivaForFormula = String(dictInformacionProducto["IVA"]).replace(".", ",");
    const retForFormula = String(dictInformacionProducto["retencion"]).replace(".", ",");
    const recForFormula = String(dictInformacionProducto["Recargo de equivalencia"]).replace(".", ",");
    factura_sheet.getRange("F" + String(rowParaDatos)).setValue("=(E" + String(rowParaDatos) + "*" + ivaForFormula + ")")//Valor de los impuestos
    if (retForFormula !== "0" && retForFormula !== "0,0") {
      factura_sheet.getRange("G" + String(rowParaDatos)).setValue("=(E" + String(rowParaDatos) + "*" + retForFormula + ")")//Retencion
    }
    if (recForFormula !== "0" && recForFormula !== "0,0") {
      factura_sheet.getRange("I" + String(rowParaDatos)).setValue("=(E" + String(rowParaDatos) + "*" + recForFormula + ")")//Recargo de equivalencia
    }
    // Importe línea = Subtotal + IVA + Recargo (sin retenciones ni descuentos)
    factura_sheet.getRange("J" + String(rowParaDatos)).setValue("=(E" + String(rowParaDatos) + "+F" + String(rowParaDatos) + "+I" + String(rowParaDatos) + ")")//total linea

  }



  Logger.log("rowParaDatos " + rowParaDatos)
  Logger.log("Number(taxSectionStartRow-1) " + Number(taxSectionStartRow - 1))



  updateTotalProductCounter(rowParaDatos, productStartRow, hojaFactura, rowParaTotalTaxes);
  calcularImporteYTotal(rowParaDatos, productStartRow, rowParaTotalTaxes, hojaFactura)
}

function verificarEstadoConsecutivo(){
  const scriptProperties = PropertiesService.getDocumentProperties();
  numero = scriptProperties.getProperty('NumeroConescutivo');  // Ej: "123"
  letra  = scriptProperties.getProperty('LetraConescutivo');   // Ej: "abc"
  Logger.log("numero "+numero)
  Logger.log("letra "+letra)  
  if (numero==null || letra==null || numero=="" || letra==""){
    SpreadsheetApp.getUi().alert("Recuerda que antes de poder generar una factura es necesario guardado un nuevo consecutivo, dirígete a la hoja Datos de emisor y crea un nuevo consecutivo dándole click al botón crear consecutivo")
    return false
  }else{
    Logger.log("la consecutivo si existe");
    return true
  }
}

function onImageClick() {
  // Obtén el rango activo (última celda seleccionada)
  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();

  // Obtén la dirección de la celda
  var cellAddress = range.getA1Notation();

  // Muestra la celda en un diálogo
  SpreadsheetApp.getUi().alert('La celda activa es: ' + cellAddress);
}
function probarInsertarImagen() {
  insertarImagenBorrarFila(15)
}
function insertarImagenBorrarFila(fila) {
  var spreadsheet = SpreadsheetApp.getActive();
  let hojaFcatura = spreadsheet.getSheetByName('Factura');
  let imagenURL = "https://i.postimg.cc/RFZ45sgp/basura3.png"
  var cell = hojaFcatura.getRange('H' + fila);
  cell.setHorizontalAlignment('center');
  var imageBlob = UrlFetchApp.fetch(imagenURL).getBlob();
  var image = hojaFcatura.insertImage(imageBlob, cell.getColumn(), cell.getRow());
  var numFactura = hojaFcatura.getRange('A' + fila).getValue();
  image.assignScript("onImageClick");
  image.setHeight(20);
  image.setWidth(20);
  image.setAnchorCellXOffset(40);
}

function guardarFacturaHistorial() {
  
  var hojaFactura = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Factura');
  var hojaListado = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas Data');
  var numeroFactura = hojaFactura.getRange("G2").getValue();
  var cliente = hojaFactura.getRange("B2").getValue();
  var fechaEmision = hojaFactura.getRange("G4").getValue();
  var estado = "Creada";
  var informacionCliente = getCustomerInformation(cliente);
  var nif = informacionCliente.Identification;

  var lastRow = hojaListado.getLastRow();
  var newRow = lastRow + 1;
  var celdaNumFactura = hojaListado.getRange("A" + newRow);
  celdaNumFactura.setValue(numeroFactura);
  celdaNumFactura.setHorizontalAlignment('center');
  celdaNumFactura.setBorder(true, true, true, true, null, null, null, null);

  var celdaCliente = hojaListado.getRange("B" + newRow);
  celdaCliente.setValue(cliente);
  celdaCliente.setHorizontalAlignment('center');
  celdaCliente.setBorder(true, true, true, true, null, null, null, null);

  var celdaNIF = hojaListado.getRange("C" + newRow);
  celdaNIF.setValue(nif);
  celdaNIF.setHorizontalAlignment('center');
  celdaNIF.setBorder(true, true, true, true, null, null, null, null);

  var celdaFecha = hojaListado.getRange("D" + newRow);
  celdaFecha.setValue(fechaEmision);
  celdaFecha.setHorizontalAlignment('center');
  celdaFecha.setBorder(true, true, true, true, null, null, null, null);

  var celdaEstado = hojaListado.getRange("E" + newRow);
  celdaEstado.setValue(estado);
  celdaEstado.setHorizontalAlignment('center');
  celdaEstado.setBorder(true, true, true, true, null, null, null, null);

  var celdaImagen = hojaListado.getRange("F" + newRow);
  //insertarImagen(newRow);
  celdaImagen.setHorizontalAlignment('center');
  celdaImagen.setBorder(true, true, true, true, null, null, null, null);

  // var html = HtmlService.createHtmlOutputFromFile('postFactura')
  //   .setTitle('Menú');
  // SpreadsheetApp.getUi()
  //   .showSidebar(html);
  
  showCustomDialog()
}

function enviarFactura(){
  var spreadsheet = SpreadsheetApp.getActive();
  const scriptProps = PropertiesService.getDocumentProperties();
  let ambiente = scriptProps.getProperty('Ambiente')
  Logger.log("Ambiente: "+ambiente)
  
  // Nuevo endpoint para AddInvoice
  let url
  if (ambiente=="Pruebas"){
    url = "https://facturasapp-qa.cenet.ws/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/AddInvoice"
  }else{
    url = "https://www.facturasapp.com/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/AddInvoice";
  }
  
  // Obtener el JSON del nuevo formato desde ListadoEstado
  let listadoEstado = spreadsheet.getSheetByName('ListadoEstado');
  let lastRow = listadoEstado.getLastRow();
  let jsonFieldInvoice = listadoEstado.getRange(lastRow, 13).getValue(); // Columna M donde se guarda el nuevo JSON
  
  if (!jsonFieldInvoice) {
    SpreadsheetApp.getUi().alert("Error: No se encontró el JSON de la factura. Asegúrese de haber generado la factura primero.");
    return;
  }
  
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let APIkey = hojaDatos.getRange("I21").getValue()
  
  if (!APIkey) {
    SpreadsheetApp.getUi().alert("Error: No se encontró la API Key. Asegúrese de haber vinculado su cuenta de FacturasApp.");
    return;
  }
  
  let opciones = {
    "method": "post",
    "contentType": "application/json",
    "payload": jsonFieldInvoice,
    "headers": {"X-API-KEY": APIkey},
    'muteHttpExceptions': true
  };

  try {
    Logger.log("URL: " + url);
    Logger.log("API Key: " + APIkey);
    Logger.log("Payload length: " + jsonFieldInvoice.length);
    Logger.log("jsonFieldInvoice"+jsonFieldInvoice)
    
    // Verificar que el JSON es válido
    try {
      let testJson = JSON.parse(jsonFieldInvoice);
      Logger.log("JSON válido. Productos: " + testJson.products.length);
    } catch (parseError) {
      Logger.log("ERROR: JSON inválido - " + parseError.message);
      SpreadsheetApp.getUi().alert("Error: El JSON generado no es válido. " + parseError.message);
      return;
    }
    
    var respuesta = UrlFetchApp.fetch(url, opciones);
    let responseText = respuesta.getContentText();
    let responseCode = respuesta.getResponseCode();
    
    Logger.log("Status: " + responseCode);
    Logger.log("Response: " + responseText);
    
    if (responseCode === 200) {
      let responseData = JSON.parse(responseText);
      if (responseData.isError) {
        Logger.log("Error de FacturasApp: " + responseData.messages);
        SpreadsheetApp.getUi().alert("Error de FacturasApp: " + responseData.messages);
      } else {
        SpreadsheetApp.getUi().alert("Factura enviada correctamente a FacturasApp. ID: " + responseData.id);
        if (responseData.id) {
          Logger.log("Factura creada con ID: " + responseData.id+" y enviada a FacturasApp");
        }
      }
    } else if (responseCode === 500) {
      Logger.log("Error HTTP 500: " + responseText);
      SpreadsheetApp.getUi().alert("Ocurrió un error interno del servidor (500).\nPor favor, intenta cerrar sesión y volver a iniciarla en FacturasApp.\nSi el problema persiste, contacta a soporte.");
    } else {
      Logger.log("Error HTTP " + responseCode + ": " + responseText);
      SpreadsheetApp.getUi().alert("Error HTTP " + responseCode + ": " + responseText);
    }
  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
    Logger.log("Stack trace: " + error.stack);
    SpreadsheetApp.getUi().alert("Error al enviar la factura a FacturasApp. Error: " + error.message);
  }
}

function guardarIdArchivo(idArchivo, numeroFactura) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  var lastRow = hoja.getLastRow();
  var newRow = lastRow + 1;
  hoja.getRange("A" + newRow).setValue(numeroFactura).setBorder(true, true, true, true, null, null, null, null);
  hoja.getRange("B" + newRow).setValue(idArchivo).setBorder(true, true, true, true, null, null, null, null);

}

function convertPdfToBase64Historial() {

}

function convertPdfToBase64(historial = false, row = null) {
  let hojaFacturasID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  let hojaListadoEstao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListadoEstado');
  let dataRange = hojaListadoEstao.getDataRange()
  let data = dataRange.getValues()
  let lastRowFacturasId;
  let lastRowListadoEstado;
  if (historial) {
    lastRowFacturasId = row
    lastRowListadoEstado = row
    lastRowListadoEstado = lastRowListadoEstado - 1
  } else {
    lastRowFacturasId = hojaFacturasID.getLastRow()
    lastRowListadoEstado = hojaListadoEstao.getLastRow()
    lastRowListadoEstado = lastRowListadoEstado - 1
  }


  Logger.log("data: " + data)
  let jsonNuevoCol = 13;
  let jsonData = data[lastRowListadoEstado][jsonNuevoCol]
  Logger.log("json" + jsonData)
  let invoiceData = JSON.parse(jsonData)
  let infoACambiar = invoiceData.file;
  Logger.log("infoACambiar " + infoACambiar)

  Logger.log("lastRowFacturasId: " + lastRowFacturasId)
  var idArchivo = hojaFacturasID.getRange("B" + lastRowFacturasId).getValue();
  // usa la API avanzada de Drive
  var file = Drive.Files.get(idArchivo);
  const url = `https://www.googleapis.com/drive/v3/files/${idArchivo}?alt=media`;
  var pdfBlob = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
    },
  }).getBlob();

  var base64String = Utilities.base64Encode(pdfBlob.getBytes());
  // Logger.log("base64String "+base64String)
  Logger.log("File titel " + file.name)
  invoiceData.Document.fileName = String(file.name);
  invoiceData.file = base64String;

  Logger.log("Nuevo valor de invoiceData.file: " + invoiceData.Document.fileName);
  let nuevoJsonData = JSON.stringify(invoiceData);

  return nuevoJsonData;

}
function enviarFactura(){
  var spreadsheet = SpreadsheetApp.getActive();
  const scriptProps = PropertiesService.getDocumentProperties();
  let ambiente = scriptProps.getProperty('Ambiente')
  Logger.log("Ambiente: "+ambiente)
  
  // Nuevo endpoint para AddInvoice
  let url
  if (ambiente=="Pruebas"){
    url = "https://facturasapp-qa.cenet.ws/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/AddInvoice"
  }else{
    url = "https://www.facturasapp.com/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/AddInvoice";
  }
  
  // Obtener el JSON del nuevo formato desde ListadoEstado
  let listadoEstado = spreadsheet.getSheetByName('ListadoEstado');
  let lastRow = listadoEstado.getLastRow();
  let jsonFieldInvoice = listadoEstado.getRange(lastRow, 13).getValue(); // Columna M donde se guarda el nuevo JSON
  
  if (!jsonFieldInvoice) {
    SpreadsheetApp.getUi().alert("Error: No se encontró el JSON de la factura. Asegúrese de haber generado la factura primero.");
    return;
  }
  
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let APIkey = hojaDatos.getRange("I21").getValue()
  
  if (!APIkey) {
    SpreadsheetApp.getUi().alert("Error: No se encontró la API Key. Asegúrese de haber vinculado su cuenta de FacturasApp.");
    return;
  }
  
  let opciones = {
    "method": "post",
    "contentType": "application/json",
    "payload": jsonFieldInvoice,
    "headers": {"X-API-KEY": APIkey},
    'muteHttpExceptions': true
  };

  try {
    Logger.log("URL: " + url);
    Logger.log("API Key: " + APIkey);
    Logger.log("Payload length: " + jsonFieldInvoice.length);
    Logger.log("jsonFieldInvoice"+jsonFieldInvoice)
    
    // Verificar que el JSON es válido
    try {
      let testJson = JSON.parse(jsonFieldInvoice);
      Logger.log("JSON válido. Productos: " + testJson.products.length);
    } catch (parseError) {
      Logger.log("ERROR: JSON inválido - " + parseError.message);
      SpreadsheetApp.getUi().alert("Error: El JSON generado no es válido. " + parseError.message);
      return;
    }
    
    var respuesta = UrlFetchApp.fetch(url, opciones);
    let responseText = respuesta.getContentText();
    let responseCode = respuesta.getResponseCode();
    
    Logger.log("Status: " + responseCode);
    Logger.log("Response: " + responseText);
    
    if (responseCode === 200) {
      let responseData = JSON.parse(responseText);
      if (responseData.isError) {
        Logger.log("Error de FacturasApp: " + responseData.messages);
        SpreadsheetApp.getUi().alert("Error de FacturasApp: " + responseData.messages);
      } else {
        SpreadsheetApp.getUi().alert("Factura enviada correctamente a FacturasApp. ID: " + responseData.id);
        if (responseData.id) {
          //Logger.log("Factura creada con ID: " + responseData.id+" y enviada a FacturasApp");
        }
      }
    } else if (responseCode === 500) {
      Logger.log("Error HTTP 500: " + responseText);
      SpreadsheetApp.getUi().alert("Ocurrió un error interno del servidor (500).\nPor favor, intenta cerrar sesión y volver a iniciarla en FacturasApp.\nSi el problema persiste, contacta a soporte.");
    } else {
      Logger.log("Error HTTP " + responseCode + ": " + responseText);
      SpreadsheetApp.getUi().alert("Error HTTP " + responseCode + ": " + responseText);
    }
  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
    Logger.log("Stack trace: " + error.stack);
    SpreadsheetApp.getUi().alert("Error al enviar la factura a FacturasApp. Error: " + error.message);
  }
}


function enviarFacturaHistorial(numeroFactura) {
  let spreadsheet = SpreadsheetApp.getActive()
  const scriptProps = PropertiesService.getDocumentProperties();
  let ambiente = scriptProps.getProperty('Ambiente');
  let url;
  if (ambiente == "Pruebas"){
    url = "https://facturasapp-qa.cenet.ws/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/AddInvoice";
  } else {
    url = "https://www.facturasapp.com/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/AddInvoice";
  }
  let hojafFacturasID = spreadsheet.getSheetByName('Facturas ID');
  let lastRow = hojafFacturasID.getLastRow()
  let rangeFacturasID = hojafFacturasID.getRange(2, 1, lastRow - 1)
  let facturasIDList = rangeFacturasID.getValues().map(row => row[0]);
  Logger.log(facturasIDList)

  Logger.log(numeroFactura)
  let resultadoBusqueda = busquedaLineal(facturasIDList, numeroFactura)
  resultadoBusqueda = resultadoBusqueda + 2 //se le suma 2 debido al desface de la hoja de calculo, ojo con el retorno de -1
  let json = convertPdfToBase64(true, resultadoBusqueda)
  //verificar si exite la el apikey
  Logger.log("resultadoBusqueda:" + resultadoBusqueda)
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let APIkey = hojaDatos.getRange("I21").getValue()
  let opciones = {
    "method": "post",
    "contentType": "application/json",
    "payload": json,
    "headers": { "X-API-KEY": APIkey },
    'muteHttpExceptions': true
  };


  try {
    var respuesta = UrlFetchApp.fetch(url, opciones);
    Logger.log(respuesta.status); // Muestra la respuesta de la API en los logs
    SpreadsheetApp.getUi().alert("Factura enviada correctamente a FacturasApp.");
  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
    SpreadsheetApp.getUi().alert("Error al enviar la factura a FacturasApp. Intente de nuevo si el error presiste comuniquese con soporte");
  }
}

function busquedaLineal(lista, objetivo) {
  for (let i = 0; i < lista.length; i++) {
    if (lista[i] == objetivo) {
      return i; // Índice encontrado
    }
  }
  return -1; // No encontrado
}


function jsonAPIkey(usuario, contra) {
  let json = {
    "user": usuario,
    "password": contra
  }

  return json
}
function obtenerAPIkey(usuario, contra) {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let hojaDatos = spreadsheet.getSheetByName("Datos")
  const scriptProps = PropertiesService.getDocumentProperties();
  let ambiente = scriptProps.getProperty('Ambiente');
  let url;
  if (ambiente == "Pruebas"){
    url = "https://facturasapp-qa.cenet.ws/ApiGateway/AppSecurity/ApiKey";
  } else {
    url = "https://www.facturasapp.com/ApiGateway/AppSecurity/ApiKey";
  }
  let json = jsonAPIkey(usuario, contra);
  let opciones = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(json),
    'muteHttpExceptions': true
  };

  try {
    let respuesta = UrlFetchApp.fetch(url, opciones);
    let contenidoRespuesta = respuesta.getContentText();

    // Intentamos parsear la respuesta como JSON
    let respuestaJson;
    try {
      respuestaJson = JSON.parse(contenidoRespuesta);
    } catch (e) {
      throw new Error("Respuesta inesperada de la API. No es JSON válido.");
    }

    // Verificar si la respuesta contiene un API Key en el formato esperado
    if (Array.isArray(respuestaJson) && respuestaJson.length > 0 && typeof respuestaJson[0] === 'string') {
      let apiKey = respuestaJson[0]; // Extrae el API Key
      Logger.log("API Key obtenida: " + apiKey);
      SpreadsheetApp.getUi().alert("Se ha vinculado tu cuenta exitosamente");
      hojaDatosEmisor.getRange("B16").setBackground('#ccffc7')
      hojaDatosEmisor.getRange("B16").setValue("Vinculado")
      hojaDatos.getRange("I21").setValue(apiKey)
      try { showNuevaFactura(); } catch (_) {}
    } else {
      hojaDatosEmisor.getRange("B16").setBackground('#FFC7C7')
      hojaDatosEmisor.getRange("B16").setValue("Desvinculado")
      throw new Error("Error de la API: " + contenidoRespuesta); // Muestra el error de la API

    }
  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
    hojaDatosEmisor.getRange("B16").setBackground('#FFC7C7')
    hojaDatosEmisor.getRange("B16").setValue("Desvinculado")
    SpreadsheetApp.getUi().alert("Error al vincular tu cuenta. Verifica que el usuario y la contraseña estén correctos e intenta de nuevo. Si el error persiste, comunícate con soporte.");
  }
}



function convertPdfToBase64Prueba() {
  let hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  let hojaListadoEstao = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListadoEstado');
  let dataRange = hojaListadoEstao.getDataRange();
  let data = dataRange.getValues();

  let jsonNuevoCol = 13;
  let lastRow = hojaListadoEstao.getLastRow();
  let jsonData = data[lastRow - 1][jsonNuevoCol];
  Logger.log("json" + jsonData);

  let invoiceData = JSON.parse(jsonData);
  let infoACambiar = invoiceData.file;
  Logger.log("infoACambiar " + infoACambiar);

  let lastRowFacturasId = hoja.getLastRow();
  let idArchivo = hoja.getRange("B" + lastRowFacturasId).getValue();
  const file = DriveApp.getFileById(idArchivo);
  const base64String = Utilities.base64Encode(file.getBlob().getBytes());

  invoiceData.file = base64String;
  Logger.log("Nuevo valor de invoiceData.file: " + invoiceData.file);

  let nuevoJsonData = JSON.stringify(invoiceData);
  Logger.log("Nuevo JSON Data: " + nuevoJsonData);

  // Crear o actualizar el archivo 'prueba.json' en Google Drive
  let folder = DriveApp.getRootFolder(); // Aquí puedes especificar una carpeta en particular
  let files = folder.getFilesByName('prueba.json');
  let jsonFile = folder.createFile('prueba.json', nuevoJsonData, "application/json");

  if (files.hasNext()) {
    jsonFile = files.next();
    jsonFile.setContent(nuevoJsonData);
    Logger.log('Archivo "prueba.json" actualizado.');
  } else {
    jsonFile = folder.createFile('prueba.json', nuevoJsonData, MimeType.JSON);
    Logger.log('Archivo "prueba.json" creado.');
  }
}



function linkDescargaFactura() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  var lastRow = hoja.getLastRow();
  var idArchivo = hoja.getRange("B" + lastRow).getValue();
  var numFactura = hoja.getRange("A" + lastRow).getValue();

  if (!idArchivo) {
    throw new Error("El ID del archivo está vacío o no es válido.");
  }

  // Verificar el archivo y asignar permisos públicos usando Advanced Drive Service
  var permisos = {
    role: "reader",
    type: "anyone"
  };

  try {
    Drive.Permissions.create(permisos, idArchivo, { sendNotificationEmails: false });
  } catch (e) {
    throw new Error("Error al configurar permisos públicos: " + e.message);
  }

  // Generar la URL de descarga
  var url = "https://drive.google.com/uc?export=download&id=" + idArchivo;

  return {
    numFactura: numFactura,
    url: url
  };
}



function getDownloadLink() {
  var data = linkDescargaFactura();
  Logger.log("sale de linkdescargar")
  return data;
}

function enviarEmailPostFactura(email, historial = false, numFacturaAbuscar = null) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  var lastRow = hoja.getLastRow();
  let idArchivo;
  let numFactura;
  let lastRowfacturasID;
  if (historial) {
    let rangeFacturasID = hoja.getRange(2, 1, lastRow - 1)
    let facturasIDList = rangeFacturasID.getValues().map(row => row[0]);

    lastRowfacturasID = busquedaLineal(facturasIDList, numFacturaAbuscar)//que pasa cuando retorne -1 ?
    lastRowfacturasID = lastRowfacturasID + 2
    idArchivo = hoja.getRange("B" + lastRowfacturasID).getValue();
    numFactura = hoja.getRange("A" + lastRowfacturasID).getValue();
  } else {

    idArchivo = hoja.getRange("B" + lastRow).getValue();
    numFactura = hoja.getRange("A" + lastRow).getValue();
  }
  Logger.log("lastRowfacturasID " + lastRowfacturasID)

  Logger.log("email " + email)
  Logger.log("idArchivo " + idArchivo)
  Logger.log("numFactura " + numFactura)


  var file = Drive.Files.get(idArchivo);
  Logger.log("file obtenido exitosamente." + file.name);
  const url = `https://www.googleapis.com/drive/v3/files/${idArchivo}?alt=media`;
  var pdfBlob = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`,
    },
  }).getBlob();

  pdfBlob.setName(file.name)
  var subject = `Factura ${numFactura}`;
  var body = 'Adjunto encontrará la factura en formato PDF.';

  if (!email) {
    return "Por favor ingrese una dirección de correo válida.";
  }

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
    attachments: [pdfBlob.setName(file.name)]  // Adjuntar el archivo PDF
  });

  return "PDF generado y enviado por correo electrónico a " + email;
}


function ProcesarFormularioFactura(data) {
  var numFactura = data.numFactura;
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');

  var range = hoja.getRange('A2:A');
  var textFinder = range.createTextFinder(numFactura);
  var cell = textFinder.findNext();
  Logger.log("cell " + cell)
  if (!cell) {
    return 'Factura no encontrada';
  }

  var fila = cell.getRow();
  var idAsociado = hoja.getRange('B' + fila).getValue();

  if (!idAsociado) {
    return 'ID de factura no encontrado';
  }

  try {
    // descargarPDF(idAsociado);
    // makeFilePublic(idAsociado)
    // var url = `https://www.googleapis.com/drive/v3/files/${idAsociado}/download`;
    // var file = Drive.Files.get(idAsociado);
    // const url = file.webContentLink;
    var url = "https://drive.google.com/uc?export=download&id=" + idAsociado;
    return url;
  } catch (e) {
    return 'Error al obtener el archivo: ' + e.message;
  }
}

function descargarPDF(id) {
  var fileId = id; // Reemplaza con el ID real del archivo PDF
  var file = Drive.Files.get(fileId);
  var url = file.webContentLink; // Obtiene el enlace de descarga directo

  Logger.log("Enlace de descarga: " + url);

  // Opcional: Si lo ejecutas desde un script de Google Sheets, puedes mostrarlo en un cuadro de diálogo
  var ui = SpreadsheetApp.getUi();
  ui.alert("Haz clic en el enlace para descargar:\n" + url);
}

function makeFilePublic(fileId) {
  try {
    var permission = {
      'role': 'reader',
      'type': 'anyone'
    };
    Drive.Permissions.insert(permission, fileId, { sendNotificationEmails: false });
    return "Permiso actualizado. Intenta descargar nuevamente.";
  } catch (e) {
    return "Error al actualizar permisos: " + e.message;
  }
}


function verificarCodigo(codigo, nombreHoja, inHoja, lineEditada = null, codigoV = "") {
  Logger.log("Verificar códigos");
  Logger.log("linea editada: " + lineEditada)
  // Obtener la hoja por nombre
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  let codigoNumero = String(codigo)


  try {
    let columna;
    let lastActiveRow = sheet.getLastRow();
    let rangeDatos;
    let pruebaPostRow = 0
    Logger.log(lastActiveRow + "last acitrive row")
    // Determinar la columna y el rango según el tipo de hoja
    if (nombreHoja === "Contactos" && codigoV !== "codigo") {
      columna = 6; // Columna para el identificador de clientes
      rangeDatos = sheet.getRange(2, columna, lastActiveRow - (inHoja ? 2 : 1));
    } else if (nombreHoja === "Contactos" && codigoV === "codigo") {
      columna = 7;//columna codigo
      rangeDatos = sheet.getRange(2, columna, lastActiveRow - (inHoja ? 2 : 1));
    } else if (nombreHoja === "Productos") {
      columna = 2; // Columna para el código de productos
      pruebaPostRow = lastActiveRow - (inHoja ? 2 : 1)
      rangeDatos = sheet.getRange(2, columna, lastActiveRow - (inHoja ? 2 : 1));
    } else if (nombreHoja === "Historial Facturas Data") {
      columna = 1; // Columna para el número de factura
      rangeDatos = sheet.getRange(2, columna, lastActiveRow - 1);
    } else {
      Logger.log("Nombre de hoja no válido.");
      return false;
    }
    Logger.log("last active ro post" + pruebaPostRow)
    // Obtener los valores del rango como una matriz de números
    let datos = rangeDatos.getValues().flat().map(String);
    Logger.log("Datos obtenidos como números:");
    Logger.log(datos);

    // Convertir el código a número

    Logger.log(codigoNumero)
    // Verificar si algún valor en datos es exactamente igual al código
    for (let i = 0; i < datos.length; i++) {
      if (datos[i] === codigoNumero) {
        if (i === lineEditada - 2) {
          Logger.log("dentro de continue")

        } else {

          Logger.log(`El código "${codigoNumero}" ya existe en la hoja "${nombreHoja}".`);
          return true;
        }
      }
    }

    Logger.log(`El código "${codigoNumero}" no existe en la hoja "${nombreHoja}".`);
    return false;
  } catch (error) {
    Logger.log("Error al verificar el código: " + error.message);
    return false;
  }
}




function insertarImagen(fila) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas Data');
  var imageUrl = 'https://cdn.icon-icons.com/icons2/1674/PNG/512/download_111133.png'; // Reemplaza con la URL de tu imagen
  var cell = sheet.getRange('F' + fila);
  cell.setHorizontalAlignment('center');
  var imageBlob = UrlFetchApp.fetch(imageUrl).getBlob();
  var image = sheet.insertImage(imageBlob, cell.getColumn(), cell.getRow());
  image.assignScript("descargarFactura");
  image.setHeight(20);
  image.setWidth(20);
  image.setAnchorCellXOffset(40);
}

function descargarFactura() {
  var html = HtmlService.createHtmlOutputFromFile('descargaFacturaHistorial')
    .setTitle('Historial Facturas');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function guardarFilaFactura() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas Data');
  var cell = sheet.getActiveCell();
  var fila = cell.getRow();
  sheet.getRange('Z1').setValue(fila); // Guardar la fila en una celda oculta (Z1)
  generarPDFfactura();
}


function generarPDFfactura() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas Data');
  var fila = sheet.getRange('Z1').getValue(); // Leer el número de fila de la celda oculta
  var numeroFactura = sheet.getRange('A' + fila).getValue(); // Obtener el número de factura

  var resultado = obtenerDatosFactura(numeroFactura);
  if (resultado) {
    var pdfBlob = generarPDF();
  } else {
    Utilities.sleep(5000);
    var pdfBlob = generarPDF();
  }
  var url = generarPdfUrl(pdfBlob);

  // Crear un archivo temporal en el Drive para proporcionar un enlace de descarga
  var tempFile = DriveApp.createFile(pdfBlob);
  var tempFileUrl = tempFile.getDownloadUrl();
  Logger.log("generar pdf despues de getlinkdownload")
  // Enviar un enlace de descarga al usuario
  var html = '<html><body><a href="' + tempFileUrl + '">Descargar PDF de la Factura ' + numeroFactura + '</a></body></html>';
  var ui = HtmlService.createHtmlOutput(html)
    .setWidth(300)
    .setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(ui, 'Descargar PDF');
}


function generarPDF() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Copia de Plantilla');

  if (!sheet) {
    throw new Error('La hoja Plantilla no existe.');
  }

  var sheetId = sheet.getSheetId();
  var url = ss.getUrl().replace(/edit$/, '') + 'export?exportFormat=pdf&format=pdf' +
    '&gid=' + sheetId +
    '&size=A4' +  // Tamaño del papel
    '&portrait=true' +  // Orientación vertical
    '&fitw=true' +  // Ajustar a ancho de la página
    '&sheetnames=false&printtitle=false' +  // Opciones de impresión
    '&pagenumbers=false&gridlines=false' +  // Más opciones de impresión
    '&fzr=false' +  // Aislar filas congeladas
    '&top_margin=0.8' +  // Margen superior
    '&bottom_margin=0.00' +  // Margen inferior
    '&left_margin=0.50' +  // Margen izquierdo
    '&right_margin=0.50' +  // Margen derecho
    '&horizontal_alignment=CENTER' +  // Alineación horizontal
    '&vertical_alignment=TOP';  // Alineación vertical

  var token = ScriptApp.getOAuthToken();

  try {
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      var pdfBlob = response.getBlob().setName('Factura.pdf');
      return pdfBlob;
    } else {
      Logger.log('Error ' + response.getResponseCode() + ': ' + response.getContentText());
      throw new Error('Error ' + response.getResponseCode() + ': ' + response.getContentText());
    }
  } catch (e) {
    Logger.log('Exception: ' + e.message);
    throw new Error('Exception: ' + e.message);
  }
}

function generarPdfUrl(pdfBlob) {
  var base64Data = Utilities.base64Encode(pdfBlob.getBytes());
  var contentType = pdfBlob.getContentType();
  var name = pdfBlob.getName();
  return `data:${contentType};base64,${base64Data}`;
}


function limpiarHojaFactura() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const hojaFactura = spreadsheet.getSheetByName('Factura');
  const copiaFactura = spreadsheet.getSheetByName('Copia de Factura');
  const hojaInicio=spreadsheet.getSheetByName('Inicio');

  if (!copiaFactura) {
    Logger.log("No se encontró la hoja 'Copia facturas'.");
    return;
  }
  spreadsheet.setActiveSheet(hojaInicio)
  // Si existe la hoja Factura, elimínala
  if (hojaFactura) {
    spreadsheet.deleteSheet(hojaFactura);
  }
  
  // Copiar la hoja "Copia facturas" como nueva hoja llamada "Factura"
  const nuevaHojaFactura = copiaFactura.copyTo(spreadsheet);
  nuevaHojaFactura.setName('Factura');
  const hojaFacturaPost = spreadsheet.getSheetByName('Factura');
  spreadsheet.setActiveSheet(hojaFacturaPost)
  Logger.log("La hoja 'Factura' ha sido reemplazada correctamente.");
  generarNumeroFactura()
}



function inicarFacturaNueva() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let hojaInfoUsuario = spreadsheet.getSheetByName('Datos de emisor');
  let IABN = hojaInfoUsuario.getRange("B9").getValue()

  hojaFactura.getRange("B11").setValue(IABN)
  generarNumeroFactura();
  obtenerFechaYHoraActual();
}

function limpiarYEliminarFila(numeroFila, hoja, hojaTax) {
  //funcion para el boton que se va a agregar al final del producto
  if (numeroFila > 20 && numeroFila < hojaTax) {
    hoja.deleteRow(numeroFila)
  } else {
    hoja.getRange("A" + String(numeroFila)).setValue("");//producto
    hoja.getRange("B" + String(numeroFila)).setValue("");//ref
    hoja.getRange("C" + String(numeroFila)).setValue("");//cantidad
    hoja.getRange("D" + String(numeroFila)).setValue(0);//CON IVa
    hoja.getRange("E" + String(numeroFila)).setValue(0);//sin iva
    //sheet.getRange("C"+String(posicionTaxInfo)).setValue(valorEnPorcentaje);
  }
}

function verificarYCopiarContacto(e) {
  let hojaFacturas = e.source.getSheetByName('Factura');
  let hojaContactos = e.source.getSheetByName('Contactos');
  let celdaEditada = e.range;



  let nombreContacto = celdaEditada.getValue();
  let ultimaColumnaPermitida = 20; // Columna del estado en la hoja de contactos
  let datosARetornar = ["B", "O", "M", "L", "N", "Q"]; // Columnas que quiero de la hoja de contactos


  if (nombreContacto === "Cliente") {
    Logger.log("Estado default")
  } else {
    let listaConInformacion = obtenerInformacionCliente(nombreContacto);
    if (listaConInformacion["Estado"] === "No Valido") {
      SpreadsheetApp.getUi().alert("Error: El contacto seleccionado no es válido.");
    } else {
      //asigna el valor del coldigo solamente porque ese fue lo que me pidieron no mas
      hojaFacturas.getRange("B3").setValue(listaConInformacion["Código cliente"]);
    }
  }


}


function generarNumeroFactura() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('Factura');
  let sheetHistorial = spreadsheet.getSheetByName("Historial Facturas Data");
  let columnaNumeroFactura = 1;
  let lastActiveRow = sheetHistorial.getLastRow();

  if (lastActiveRow <= 2) {
    lastActiveRow = 2;
  }

  let rangeNumeroFactura = sheetHistorial.getRange(2, columnaNumeroFactura, lastActiveRow - 1);
  let numeroFacturas = rangeNumeroFactura.getValues();

  let numeroMayor = -Infinity;
  let ultimoConsecutivo = "";

  // Iterar sobre la columna para encontrar el mayor número
  for (let i = 0; i < numeroFacturas.length; i++) {
    let consecutivo = numeroFacturas[i][0];
    let cumple = cumpleEstructura(consecutivo)
    if (!cumple) {
      Logger.log("No cumple con la estructura")
    } else {
      let numero = obtenerParteNumerica(consecutivo);

      if (numero > numeroMayor) {
        numeroMayor = numero;
        ultimoConsecutivo = consecutivo; // Guardamos el último número en formato original
      }
    }
  }

  // Calcular el siguiente consecutivo DESPUÉS de recorrer el historial
  let nuevoConsecutivo = "";
  if (numeroMayor === -Infinity) {
    // No hay facturas previas en el historial. Usar el prefijo y longitud
    // configurados en las propiedades del script para construir el valor base.
    const scriptProperties = PropertiesService.getDocumentProperties();
    let numeroBase = String(scriptProperties.getProperty('NumeroConescutivo') || "1");
    let letraBase = String(scriptProperties.getProperty('LetraConescutivo') || "FCT");
    nuevoConsecutivo = letraBase + numeroBase; // Primer número según configuración
  } else {
    // Existe historial: incrementar manteniendo formato
    nuevoConsecutivo = generarNuevoConsecutivo(ultimoConsecutivo, numeroMayor + 1);
  }

  // Escribir una sola vez en la celda de número de factura
  sheet.getRange("G2").setValue(nuevoConsecutivo);
}

// Extrae la parte numérica de una cadena
function obtenerParteNumerica(str) {
  str = String(str);
  const match = str.match(/\d+$/);
  return match ? parseInt(match[0], 10) : 0;
}

// Genera el nuevo número con el mismo formato del original
function generarNuevoConsecutivo(original, nuevoNumero) {
  let match = original.match(/^(\D*)(\d+)$/); // Captura el prefijo y la parte numérica

  if (!match) {
    return String(nuevoNumero); // Si no hay formato reconocible, devuelve solo el número
  }

  let prefijo = match[1]; // Parte no numérica (ejemplo: "uuu", "xyz-")
  let parteNumerica = match[2]; // Parte numérica original (ejemplo: "000001")

  let nuevoNumeroStr = String(nuevoNumero).padStart(parteNumerica.length, '0'); // Mantiene los ceros iniciales

  return prefijo + nuevoNumeroStr;
}

function obtenerFechaYHoraActual() {
  let spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName('Factura');
  let zonaHorariaEspaña = "Europe/Madrid"
  let fecha = Utilities.formatDate(new Date(), zonaHorariaEspaña, "dd/MM/yyyy");
  let hora = Utilities.formatDate(new Date(), zonaHorariaEspaña, "HH:mm:ss");

  sheet.getRange("G4").setNumberFormat("dd/MM/yyyy");
  sheet.getRange("G4").setValue(String(fecha))
  sheet.getRange("G3").setValue(String(fecha))
  sheet.getRange("G7").setValue(hora)


  let valorFecha = sheet.getRange("G4").getValue();

  let fechaFormateada = Utilities.formatDate(new Date(valorFecha), zonaHorariaEspaña, "dd/MM/yyyy");
  Logger.log("valorFecha " + valorFecha)
  Logger.log("fecha " + fecha)
  Logger.log("fechaFormateada " + fechaFormateada)

}

function ObtenerFecha(opcion = null) {
  let spreadsheet = SpreadsheetApp.getActive();
  let fechaFormateada
  let valorFecha
  let zonaHorariaEspaña = "Europe/Madrid"
  if (opcion == "pago") {
    let sheet = spreadsheet.getSheetByName('Factura');
    valorFecha = sheet.getRange("G3").getValue();
    Logger.log("valorFecha 1" + String(valorFecha))
    fechaFormateada = Utilities.formatDate(new Date(valorFecha), zonaHorariaEspaña, "dd/MM/yyyy");
  } else {
    let sheet = spreadsheet.getSheetByName('Factura');
    valorFecha = sheet.getRange("G4").getValue();
    Logger.log("valorFecha " + String(valorFecha))
    Logger.log("valorFecha 2" + valorFecha)
    fechaFormateada = Utilities.formatDate(new Date(valorFecha), zonaHorariaEspaña, "dd/MM/yyyy");
  }
  Logger.log("fecha formateada" + fechaFormateada)
  Logger.log("valorFecha " + String(valorFecha))
  return fechaFormateada
}

function obtenerDatosProductos(sheet, range, e) {
  if (range.getA1Notation() === "A14" || range.getA1Notation() === "A15" || range.getA1Notation() === "A16" || range.getA1Notation() === "A17" || range.getA1Notation() === "A18") {
    Logger.log("entro a obtenerdatos")
    var selectedProduct = range.getValue();

    // Referencia a la hoja de productos
    var productSheet = e.source.getSheetByName("Productos");
    var data = productSheet.getDataRange().getValues();

    // Encuentra el producto en la hoja de productos
    for (var i = 1; i < data.length; i++) {
      Logger.log(data[i][1])
      Logger.log(selectedProduct)
      if (data[i][1] == selectedProduct) {
        sheet.getRange("B14").setValue(data[i][0]);  // Código de referencia
        sheet.getRange("D14").setValue(data[i][2]);  // Valor unitario
        sheet.getRange("E14").setValue(data[i][4]);  // Otros datos,  segun sea necesario
        break;
      }
    }
  }

}

function obtenerIdRateWithHoldings(valorRetencion, tipoWithHolding) {
  // Mapeo según la tabla de configuración
  // tipoWithHolding: 10 = Retención, 11 = Recargo de equivalencia
  
  // Convertir a número y redondear para evitar problemas de precisión
  let valor = Math.round(Number(valorRetencion) * 10) / 10;
  
  if (tipoWithHolding === 10) { // Retención
    switch (valor) {
      case 7:
        return "20";
      case 15:
        return "21";
      case 19:
        return "22";
      default:
        Logger.log("Valor de retención no reconocido: " + valor + ". Usando código por defecto 20");
        return "20"; // Valor por defecto
    }
  } else if (tipoWithHolding === 11) { // Recargo de equivalencia
    switch (valor) {
      case 5.2:
        return "23";
      case 1.4:
        return "24";
      case 0.5:
        return "25";
      case 1.75:
        return "26";
      default:
        Logger.log("Valor de recargo de equivalencia no reconocido: " + valor + ". Usando código por defecto 23");
        return "23"; // Valor por defecto
    }
  }
  
  Logger.log("Tipo de withholding no reconocido: " + tipoWithHolding + ". Usando código por defecto 20");
  return "20"; // Valor por defecto si no coincide
}

function mapIdPaymentCode(medioPagoTxt){
  if(!medioPagoTxt) return "ND"; // No definido
  const normalizado = String(medioPagoTxt).toLowerCase().trim();
  switch(normalizado){
    case 'Efectivo':
      return 'EF';
    case 'Transferencia bancaria':
      return 'TF';
    case 'Tarjeta bancaria':
    case 'tarjeta':
      return 'TB';
    case 'Domiciliación bancaria':
    case 'Domiciliacion bancaria':
      return 'DB';
    case 'PayPal':
      return 'PP';
    case 'Talón':
    case 'Talon':
      return 'TL';
    case 'Factoring':
      return 'FR';
    case 'Confirming':
      return 'CF';
    case 'No definido':
    default:
      return 'ND';
  }
}

// Normaliza valores tipo porcentaje a fracción (0.xx). Acepta 21, '21%', '0,21', 0.21
function parsePercentValue(value){
  if (typeof value === 'number') {
    if (isNaN(value)) return 0;
    return value > 1 ? value / 100 : value;
  }
  if (typeof value === 'string') {
    var v = value.replace('%','').replace(',', '.').trim();
    var n = Number(v);
    if (isNaN(n)) return 0;
    return n > 1 ? n / 100 : n;
  }
  return 0;
}

function parseNumberCell(value){
  if (typeof value === 'number') {
    return Number(isNaN(value) ? 0 : value);
  }
  if (typeof value === 'string') {
    var cleaned = value.replace(/[^0-9,.-]/g, '').trim();
    if (cleaned.indexOf(',') !== -1 && cleaned.indexOf('.') === -1) {
      cleaned = cleaned.replace(',', '.');
    } else if (cleaned.indexOf(',') !== -1 && cleaned.indexOf('.') !== -1) {
      cleaned = cleaned.replace(/\./g, '').replace(',', '.');
    }
    var n = Number(cleaned);
    return Number(isNaN(n) ? 0 : n);
  }
  return 0;
}

// Mapeos a códigos de 2 caracteres esperados por la API
function mapPersonType(tipoPersonaTxt){
  var t = String(tipoPersonaTxt || '').toLowerCase().trim();
  if (t === 'autonomo' || t === 'autónomo' || t === 'persona fisica' || t === 'persona física') return '01';
  if (t === 'empresa' || t === 'persona juridica' || t === 'persona jurídica') return '02';
  return '01';
}

function mapIdentificationTypeES(tipoDocTxt){
  var t = String(tipoDocTxt || '').toLowerCase().trim();
  if (t.indexOf('dni') !== -1) return '01';
  if (t.indexOf('nie') !== -1) return '02';
  if (t.indexOf('cif') !== -1 || t.indexOf('nif-iva') !== -1 || t.indexOf('nif') !== -1) return '03';
  if (t.indexOf('pasaporte') !== -1) return '04';
  return '03';
}

function mapContactTypeES(tipoContactoTxt){
  var t = String(tipoContactoTxt || '').toLowerCase().trim();
  if (t.indexOf('cliente') !== -1) return '01';
  if (t.indexOf('proveedor') !== -1) return '02';
  return '01';
}

function mapRegimeES(regimenTxt){
  var t = String(regimenTxt || '').toLowerCase().trim();
  if (!t) return '03'; // General por defecto
  if (t.indexOf('general') !== -1) return '03';
  if (t.indexOf('exento') !== -1) return '01';
  if (t.indexOf('simplificado') !== -1) return '02';
  return '03';
}

function getprefacturaValueA1(column, row) {
  return getsheetValueA1(prefactura_sheet, column, row);
}

function getprefacturaValue(prefactura_sheet, column, row) {

  return getsheetValue(prefactura_sheet, column, row);
}

function updateprefacturaValue(column, row, value) {
  updatesheetValue(prefactura_sheet, column, row, value);
  return;
}

function getInvoiceGeneralInformation() {
  //Browser.msgBox('getInvoiceGeneralInformation()');
  let spreadsheet = SpreadsheetApp.getActive();
  let prefactura_sheet = spreadsheet.getSheetByName('Factura');
  var InvoiceAuthorizationNumber = "nulo"//Resolución Autorización
  //
  range = prefactura_sheet.getRange("G6");//dias de vencimiento
  var DaysOff = range.getValue();

  var invoice_number = getprefacturaValue(prefactura_sheet, 2, 7);//cambiamos los valores para llamar el numero de factura
  var InvoiceGeneralInformation = {
    "InvoiceAuthorizationNumber": InvoiceAuthorizationNumber,
    "PreinvoiceNumber": invoice_number,
    "InvoiceNumber": invoice_number,
    "DaysOff": DaysOff,
    "Currency": "EUR",
    "ExchangeRate": "",
    "ExchangeRateDate": "",
    "SalesPerson": "",
    //"InvoiceDueDate": null,
    "Note": getprefacturaValue(prefactura_sheet, 10, 2), //cambia los valores para llamar la nota de la factura
    "ExternalGR": false
    //"AdditionalProperty": AdditionalProperty
  }


  return InvoiceGeneralInformation;
}
function getPaymentSummary(startingRowTaxation) {
  let spreadsheet = SpreadsheetApp.getActive();
  let prefactura_sheet = spreadsheet.getSheetByName('Factura');
  let posTotalFactura = startingRowTaxation + 7
  let posMontoNeto = startingRowTaxation + 12
  var total_factura = prefactura_sheet.getRange("A" + String(posTotalFactura)).getValue();// por ahora esto no lo utilizamos ya que no hay descuentos
  var monto_neto = prefactura_sheet.getRange("B" + String(posMontoNeto)).getValue();
  //var numeros = new Intl.NumberFormat('es-CO', {maximumFractionDigits:0, style: 'currency', currency: 'COP' }).format(monto);
  var numeros_total = new Intl.NumberFormat().format(total_factura);
  var numeros_neto = new Intl.NumberFormat().format(monto_neto);
  //Browser.msgBox(`${numeros_total}  ${numeros_neto}`);

  Logger.log("total_factura" + total_factura)
  Logger.log("monto_neto" + monto_neto)
  var PaymentTypeTxt = prefactura_sheet.getRange("G5").getValue();
  var PaymentMeansTxt = prefactura_sheet.getRange("E4").getValue();
  let PaymentNote = prefactura_sheet.getRange("D11").getValue();
  var PaymentSummary = {
    "PaymentType": PaymentTypeTxt,
    "PaymentMeans": "PaymentMeansTxt: No hay medio de pago",//a qui habia getPaymentMeans(PaymentMeansTxt)
    "PaymentNote": PaymentNote
  }
  return PaymentSummary;
}

function guardarYGenerarInvoice(){
  let spreadsheet = SpreadsheetApp.getActive();
  let listadoestado_sheet = spreadsheet.getSheetByName('ListadoEstado');
  let prefactura_sheet = spreadsheet.getSheetByName('Factura');
  let FacturaDatos = spreadsheet.getSheetByName('Datos de emisor');

  // Obtener el total de productos
  let posicionTotalProductos = prefactura_sheet.getRange("A16").getValue();
  let cantidadProductos;
  if (posicionTotalProductos === "Total filas"){
    cantidadProductos = prefactura_sheet.getRange("B16").getValue();
  } else {
    let startingRowTax = getTaxSectionStartRow(prefactura_sheet);
    let posicionTotalProductos = startingRowTax - 3;
    cantidadProductos = prefactura_sheet.getRange("B" + String(posicionTotalProductos)).getValue();
  }

  Logger.log("cantidadProductos: " + cantidadProductos);

  // Obtener información básica de la factura
  let cliente = prefactura_sheet.getRange("B2").getValue();
  let InvoiceGeneralInformation = getInvoiceGeneralInformation();
  let CustomerInformation = getCustomerInformation(cliente);
  let startingRowTaxation = getTaxSectionStartRow(prefactura_sheet);
  let usuario = FacturaDatos.getRange("B11").getValue()
  // Obtener fechas
  let fechaFactura = new Date(prefactura_sheet.getRange("G4").getValue());
  let fechaVencimiento = new Date(prefactura_sheet.getRange("G3").getValue());
  let horaFactura = new Date().toTimeString().split(' ')[0] + ".0000000";
  
  // Validar fechas
  if (isNaN(fechaFactura.getTime())) {
    fechaFactura = new Date();
  }
  if (isNaN(fechaVencimiento.getTime())) {
    fechaVencimiento = new Date();
  }

  // Recalcular invoiceExpiration coherente con días de vencimiento (G6)
  let diasVencimiento = Number(prefactura_sheet.getRange("G6").getValue() || 0);
  if (!isNaN(diasVencimiento) && diasVencimiento >= 0) {
    let fv = new Date(fechaFactura);
    fv.setDate(fv.getDate() + diasVencimiento);
    fechaVencimiento = fv;
  }

  // Procesar productos con estructura completa
  let products = [];
  let totalTaxBase = 0;
  let totalTax = 0;
  let totalSubTotal = 0;
  // Acumuladores explícitos para validaciones
  let sumIvaAmount = 0;            // Suma de IVA (CuotaRepercutida)
  let sumRecargoAmount = 0;        // Suma de Recargo Equivalencia (no forma parte de sumTotalTax)
  let hasAnyTaxOrSurcharge = false;
  let totalWithHoldings = 0;
  let totalSurCharges = 0;
  let totalDiscounts = 0;
  // Utilidad de redondeo a 2 decimales disponible en todo el bloque
  const round2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
  
  // Crear array para fieldTaxations (resumen de impuestos)
  let fieldTaxations = [];
  let taxGroups = {};            // IVA agrupado por porcentaje
  let recargoTaxGroups = {};     // Recargo equivalencia agrupado por porcentaje
  
  for (let i = 15; i < 15 + cantidadProductos; i++) {
    let filaActual = "A" + String(i) + ":K" + String(i);
    let rangoProducto = prefactura_sheet.getRange(filaActual);
    let productoData = rangoProducto.getValues()[0];
    
    let referencia = String(productoData[0] || "");
    let descripcion = String(productoData[1] || "");
    let cantidad = Number(productoData[2]) || 1;
    let precioUnitario = Number(productoData[3]) || 0;
    // subtotal en hoja (col F) YA tiene el descuento aplicado.
    // Para evitar doble descuento en el PDF/servicio, calculamos explícitamente
    // el bruto (antes de descuento) y la base neta (después de descuento).
    let subtotalHoja = Number(productoData[4]) || 0; // E: Subtotal (con descuento)
    // F: Impuestos (importe IVA), G: Retención (importe), H: % Descuento, I: Recargo (importe), J: Importe total línea
    let impuestosImporte = parseNumberCell(productoData[5]);
    let retencionImporte = parseNumberCell(productoData[6]);
    let descuentoRate = parsePercentValue(productoData[7]);
    let recargoImporte = parseNumberCell(productoData[8]);
    let totalLinea = Number(productoData[9]) || 0;

    // Calcular valores base
    let baseBruta = round2(precioUnitario * cantidad); // antes de descuento
    let discountAmount = round2(baseBruta * descuentoRate);
    let baseNeta = round2(baseBruta - discountAmount); // después de descuento
    
    // Calcular importes y derivar tasas reales
    let taxAmount = round2(impuestosImporte);
    let withHoldingsAmount = round2(retencionImporte);
    let surChargesAmount = round2(recargoImporte);
    let ivaRate = baseNeta > 0 ? (taxAmount / baseNeta) : 0;
    let retencionRate = baseNeta > 0 ? (withHoldingsAmount / baseNeta) : 0;
    let recargoEquivalenciaRate = baseNeta > 0 ? (surChargesAmount / baseNeta) : 0;
    
    // Validar campos obligatorios
    if (!descripcion || descripcion.trim() === "") {
      descripcion = "Producto sin descripción";
    }
    if (!referencia || referencia.trim() === "") {
      referencia = "REF-" + i;
    }
    if (cantidad <= 0) {
      cantidad = 1;
    }
    
    // Crear arrays de taxes, withHoldings y discounts según factura.json
    let taxes = [];
    if (ivaRate > 0) {
      taxes.push({
        taxName: "IVA",
        rate: ivaRate * 100, // Convertir a porcentaje
        taxBase: baseNeta,
        valueTax: taxAmount
      });
      
      // Agrupar para fieldTaxations
      let rateKey = ivaRate * 100;
      if (!taxGroups[rateKey]) {
        taxGroups[rateKey] = {
          taxName: "IVA",
          rate: rateKey,
          taxBase: 0,
          valueTax: 0
        };
      }
      taxGroups[rateKey].taxBase = round2(taxGroups[rateKey].taxBase + baseNeta);
      taxGroups[rateKey].valueTax = round2(taxGroups[rateKey].valueTax + taxAmount);
    }
    // Acumular para totales y regla de validación
    sumIvaAmount = round2(sumIvaAmount + taxAmount);
    sumRecargoAmount = round2(sumRecargoAmount + surChargesAmount);
    if (taxAmount > 0 || surChargesAmount > 0) hasAnyTaxOrSurcharge = true;
    
    let withHoldingsSurChargesDto = [];
    if (withHoldingsAmount > 0 && retencionRate > 0) {
      let codigoRetencion = obtenerIdRateWithHoldings(Math.round(retencionRate * 1000) / 10, 10);
      Logger.log("Producto: " + descripcion + " - Retención: " + (retencionRate * 100) + "% - Código: " + codigoRetencion);
      withHoldingsSurChargesDto.push({
        idRateWithHoldings: codigoRetencion, // Retención
        subTotalWithHoldings: baseNeta,
        cuotaWithHoldings: withHoldingsAmount
      });
    }
    
    // Agregar recargo de equivalencia si existe
    if (surChargesAmount > 0 && recargoEquivalenciaRate > 0) {
      let codigoRecargo = obtenerIdRateWithHoldings(Math.round(recargoEquivalenciaRate * 1000) / 10, 11);
      Logger.log("Producto: " + descripcion + " - Recargo: " + (recargoEquivalenciaRate * 100) + "% - Código: " + codigoRecargo);
      withHoldingsSurChargesDto.push({
        idRateWithHoldings: codigoRecargo, // Recargo de equivalencia
        subTotalWithHoldings: baseNeta,
        cuotaWithHoldings: surChargesAmount
      });
    }

    // Agrupar recargo de equivalencia para fieldTaxations (redondear tasa a 1 decimal)
    let recargoRateKey = recargoEquivalenciaRate * 100;
    if (recargoEquivalenciaRate > 0) {
      let recargoRateKeyRounded = Math.round(recargoRateKey * 10) / 10;
      if (!recargoTaxGroups[recargoRateKeyRounded]) {
        recargoTaxGroups[recargoRateKeyRounded] = {
          taxName: "RecargoEquivalencia",
          rate: recargoRateKeyRounded,
          taxBase: 0,
          valueTax: 0
        };
      }
      recargoTaxGroups[recargoRateKeyRounded].taxBase = round2(recargoTaxGroups[recargoRateKeyRounded].taxBase + baseNeta);
      recargoTaxGroups[recargoRateKeyRounded].valueTax = round2(recargoTaxGroups[recargoRateKeyRounded].valueTax + surChargesAmount);
    }
    
    let discountDtoModules = [];
    if (descuentoRate > 0) {
      discountDtoModules.push({
        discountName: "Descuento aplicado",
        discountRate: descuentoRate * 100, // Convertir a porcentaje
        discountBase: baseBruta,
        valueDiscount: discountAmount
      });
    }
    
    // Crear producto con estructura completa
    let producto = {
      typeUse: "VEN",
      reference: String(referencia).substring(0, 50),
      description: String(descripcion).substring(0, 100),
      unitPrice: Number(precioUnitario),
      quantity: Number(cantidad),
      // IMPORTANTE: Enviar subTotal BRUTO (antes de descuento) para que el servicio
      // aplique los módulos de descuento y no descuente doble.
      subTotal: baseBruta,
      // totalTax DEBE representar únicamente la cuota de IVA. El recargo
      // se informa en totalSurCharges y en withHoldingsSurChargesDto, pero
      // no debe sumarse aquí para evitar inflar sumTotalTax.
      totalTax: round2(taxAmount),
      totalwithHoldings: withHoldingsAmount,
      totalSurCharges: surChargesAmount,
      totaldiscount: discountAmount,
      taxes: taxes.length > 0 ? taxes : [],
      withHoldingsSurChargesDto: withHoldingsSurChargesDto.length > 0 ? withHoldingsSurChargesDto : [],
      discountDtoModules: discountDtoModules.length > 0 ? discountDtoModules : []
    };
    
    // Asignar null si los arrays están vacíos para mantener estructura
    if (!producto.taxes || producto.taxes.length === 0) producto.taxes = [];
    if (!producto.withHoldingsSurChargesDto || producto.withHoldingsSurChargesDto.length === 0) producto.withHoldingsSurChargesDto = [];
    if (!producto.discountDtoModules || producto.discountDtoModules.length === 0) producto.discountDtoModules = [];
    
    products.push(producto);
    
    // Acumular totales
    // totalSubTotal: suma de subTotal BRUTO (igual a hoja: Valor bruto sin impuestos)
    totalSubTotal = round2(totalSubTotal + baseBruta);
    // totalTaxBase: solicitado por el usuario como "Valor bruto" -> usar base BRUTA
    totalTaxBase = round2(totalTaxBase + baseBruta);
    // totalTax se calculará después del bucle con sumIvaAmount + sumRecargoAmount
    totalWithHoldings = round2(totalWithHoldings + withHoldingsAmount);
    totalSurCharges = round2(totalSurCharges + surChargesAmount);
    totalDiscounts = round2(totalDiscounts + discountAmount);
  }
  // sumTotalTax debe reflejar únicamente el total de IVA (sin recargo)
  totalTax = round2(sumIvaAmount);
  
  // Crear fieldTaxations desde grupos (IVA + Recargo Equivalencia)
  for (let rate in taxGroups) {
    fieldTaxations.push(taxGroups[rate]);
  }
  for (let rate in recargoTaxGroups) {
    fieldTaxations.push(recargoTaxGroups[rate]);
  }

  // Obtener totales de la factura
  let cargoTotal = 0;
  let totalFactura = 0;
  
  if (prefactura_sheet.getRange("A31").getValue() === "Total factura") {
    totalFactura = prefactura_sheet.getRange("B31").getValue();
    cargoTotal = prefactura_sheet.getRange("B17").getValue() || 0;
  } else {
    let rowTotalFactura = startingRowTaxation + 12;
    let rowCargoFactura = startingRowTaxation - 2;
    totalFactura = prefactura_sheet.getRange(rowTotalFactura, 2).getValue();
    cargoTotal = prefactura_sheet.getRange("B" + String(rowCargoFactura)).getValue() || 0;
  }

  // Validar datos del cliente
  usuario = String(usuario || "");
  if (!usuario || usuario.trim() === "") {
    usuario = "Cliente sin nombre";
  }
  
  // Obtener información completa del cliente
  let codigoCliente = String(prefactura_sheet.getRange("B3").getValue() || "CLIENTE001");
  // 'cliente' viene de B2 y puede estar como "Nombre - Código". Extraemos el nombre limpio.
  let nombreClienteArr = dividirString(cliente);
  let nombreClienteLimpio = nombreClienteArr[0] || String(cliente);
  
  // Mapear códigos exigidos por API desde hoja Datos
  let datos_sheet_local = spreadsheet.getSheetByName('Datos');
  let tipoPersonaTxt = datos_sheet_local ? datos_sheet_local.getRange('L2').getValue() : '';
  let tipoDocTxt = datos_sheet_local ? datos_sheet_local.getRange('J2').getValue() : '';
  let tipoContactoTxt = datos_sheet_local ? datos_sheet_local.getRange('AB2').getValue() : '';
  let regimenTxt = datos_sheet_local ? datos_sheet_local.getRange('M2').getValue() : '';

  // Crear contactos con estructura completa
  let contacts = [{
    contactType: mapContactTypeES(tipoContactoTxt || 'Cliente'),
    personType: mapPersonType(tipoPersonaTxt), 
    companyName: String(nombreClienteLimpio).substring(0, 450),
    customerCode: String(codigoCliente).substring(0,20),
    identificationType: mapIdentificationTypeES(tipoDocTxt),
    identification: String(CustomerInformation.Identification || "12345678A").substring(0, 20),
    tradeName: String(cliente).substring(0, 450),
    regime: mapRegimeES(regimenTxt), // Según factura.json
    country: "207", // Código España según factura.json
    province: "5102", // Código provincia según factura.json
    population: "32653", // Código población según factura.json
    addressCustomer: (function(){
      var ad = String(CustomerInformation.AddressLine || '').trim();
      if (!ad || ad.toLowerCase() === 'null') return null;
      return ad.substring(0, 200);
    })(),
    postalCodeCustomer: (function(){
      var cp = String(CustomerInformation.PostalZone || '').trim();
      if (!cp || cp.toLowerCase() === 'null') return null;
      return cp.substring(0, 10);
    })(),
    phoneCustomer: String(CustomerInformation.Telephone || "").substring(0, 20),
    webSite: String(CustomerInformation.WebSiteURI || "").substring(0, 100) || null,
    emailCustomer: String(CustomerInformation.Email || "").substring(0, 100) || null
  }];
  
  // Mantener todos los campos, asignar null si están vacíos
  if (!contacts[0].webSite || contacts[0].webSite.trim() === "") contacts[0].webSite = null;
  if (!contacts[0].emailCustomer || contacts[0].emailCustomer.trim() === "") contacts[0].emailCustomer = null;
  if (!contacts[0].phoneCustomer || contacts[0].phoneCustomer.trim() === "") contacts[0].phoneCustomer = null;
  
  // Validar número de factura
  let numeroFacturaValidado = String(InvoiceGeneralInformation.InvoiceNumber);
  if (!numeroFacturaValidado || numeroFacturaValidado.trim() === "") {
    numeroFacturaValidado = "FACT-" + Date.now();
  }
  
  // Extraer número actual y validar que sea mayor que 0
  let currentNumber = Number(numeroFacturaValidado.replace(/[^0-9]/g, ''));
  if (currentNumber <= 0) {
    currentNumber = Math.floor(Date.now() / 1000);
  }
  
  // Crear chargeAndDiscount - siempre incluir al menos un elemento
  let chargeAndDiscount = [];
  
  // Siempre agregar al menos un elemento base según la estructura requerida
  let baseFeeDiscountValue = totalTaxBase || 0;
  let totalFeeDiscountValue = cargoTotal > 0 ? cargoTotal : (baseFeeDiscountValue * 0.01); // 1% por defecto si no hay cargo específico
  
  chargeAndDiscount.push({
    idtypeFeeDiscount: "CG", // Según factura.json
    idTypeValueFeeDiscount: "PJ", // Según factura.json  
    baseFeeDiscount: baseFeeDiscountValue,
    valueFeeDiscount: 1,
    totalFeeDiscount: totalFeeDiscountValue
  });
  
  // Calcular totales finales (totalTax ya considera recargo)
  let sumTotalSubTotalAndTax = totalSubTotal + totalTax;
  // Neto a pagar = Total factura - Retenciones (nunca negativo)
  let sumTotalNetPayable = round2(Math.max(0, totalFactura - totalWithHoldings));
  
  // Crear el JSON con estructura EXACTA de factura.json
  // Fechas coherentes con hoja: invoiceDate = G4, invoiceExpiration = días de G6
  let diasExpiracion = Number(prefactura_sheet.getRange("G6").getValue() || 0);
  if (isNaN(diasExpiracion) || diasExpiracion < 0) diasExpiracion = 0;
  // idPayment dinámico según medio de pago (E4)
  const medioPagoTxt = String(prefactura_sheet.getRange("E4").getValue() || "");
  const idPaymentCode = mapIdPaymentCode(medioPagoTxt);
  let fieldInvoice = {
    textCustomerObservations: String(prefactura_sheet.getRange("D11").getValue() || "").substring(0, 350) || null,
    invoiceNumber: numeroFacturaValidado.substring(0, 50),
    currentNumber: currentNumber,
    invoiceDate: fechaFactura.toISOString(),
    invoiceTime: horaFactura,
    invoiceExpiration: String(diasExpiracion),
    invoiceIdTypeRegAEAT: "AI",// null
    invoiceIdTypeRegSIF: null,//null
    contactName: String(prefactura_sheet.getRange("G8").getValue()|| "").substring(0, 30) || "",
    contacts: contacts,
    products: products,
    idPayment: idPaymentCode,
    paymentNote: String(prefactura_sheet.getRange("D11").getValue() || "").substring(0, 300) || null,
    textObservations: String(prefactura_sheet.getRange("B10").getValue() || "").substring(0, 500) || null,
    idOperations: "N1", // Según factura.json
    // Si hay impuestos (IVA o recargo) no es exenta: usar E0. Si no hay impuestos, E3
    idOperationsExenta: (hasAnyTaxOrSurcharge) ? "E0" : "E3",
    valueExemptBase: 0,
    chargeAndDiscount: chargeAndDiscount, // Siempre incluir - nunca null
    fieldTaxations: fieldTaxations.length > 0 ? fieldTaxations : [],
    sumTotalSubTotal: totalSubTotal,
    sumTotalTaxBase: totalTaxBase,
    sumTotalTax: totalTax,
    sumTotalSubTotalAndTax: sumTotalSubTotalAndTax,
    sumTotalExemptBase: 0,
    sumTotalDiscount: totalDiscounts,
    sumTotalCharge: cargoTotal,
    // Total de la factura
    sumTotalTotal: round2(totalFactura),
    // Neto a pagar segun esquema: Total - Retenciones
    sumTotalNetPayable: sumTotalNetPayable,
    invoiceTypeId: 0, // Según factura.json
    invoiceRectificativeTypeId: 0,
    typeRectificativeId: 0,
    aditionalData: {
      invoiceId: 0, // Según factura.json
      startInvoiceId: 0 // Según factura.json
    }
  };
  
  // Mantener estructura completa - asegurar arreglo vacío cuando no hay impuestos
  if (!fieldInvoice.fieldTaxations || fieldInvoice.fieldTaxations.length === 0) {
    fieldInvoice.fieldTaxations = [];
  }

  // Guardar en el listado de estado
  let fecha = ObtenerFecha();
  let numeroFactura = InvoiceGeneralInformation.InvoiceNumber;
  let nameString = cliente;
  
  // Guardar el JSON corregido
  listadoestado_sheet.appendRow([
    "vacio", "vacio", "vacio", fecha, "vacio", numeroFactura, 
    nameString, codigoCliente, "vacio", "vacio", "representacion", 
    "Vacio", JSON.stringify(fieldInvoice), ""
  ]);
  
  Logger.log("JSON FieldInvoice generado con estructura completa:");
  Logger.log(JSON.stringify(fieldInvoice, null, 2));
  
  // Validación de estructura completa
  Logger.log("=== VALIDACIÓN ESTRUCTURA COMPLETA ===");
  Logger.log("✓ textCustomerObservations: " + (fieldInvoice.textCustomerObservations !== undefined));
  Logger.log("✓ invoiceExpiration: " + (fieldInvoice.invoiceExpiration !== undefined));  
  Logger.log("✓ invoiceIdTypeRegSIF: " + (fieldInvoice.invoiceIdTypeRegSIF !== undefined));
  Logger.log("✓ paymentNote: " + (fieldInvoice.paymentNote !== undefined));
  Logger.log("✓ textObservations: " + (fieldInvoice.textObservations !== undefined));
  Logger.log("✓ valueExemptBase: " + (fieldInvoice.valueExemptBase !== undefined));
  Logger.log("✓ chargeAndDiscount: " + (fieldInvoice.chargeAndDiscount !== undefined && fieldInvoice.chargeAndDiscount.length > 0));
  Logger.log("✓ fieldTaxations: " + (fieldInvoice.fieldTaxations !== undefined));
  Logger.log("✓ Contacto con todos los campos: " + (fieldInvoice.contacts[0].postalCodeCustomer !== undefined));
  
  if (fieldInvoice.products && fieldInvoice.products.length > 0) {
    Logger.log("✓ Producto con taxes: " + (fieldInvoice.products[0].taxes !== undefined));
    Logger.log("✓ Producto con withHoldingsSurChargesDto: " + (fieldInvoice.products[0].withHoldingsSurChargesDto !== undefined));
    Logger.log("✓ Producto con discountDtoModules: " + (fieldInvoice.products[0].discountDtoModules !== undefined));
  }
  
  SpreadsheetApp.getUi().alert("Factura generada con estructura JSON COMPLETA según factura.json");
}

function showMensajeRespuesta() {
  SpreadsheetApp.getUi().alert("Vinculacion exitosa");
}

function calcularPorcentaje(valor, total) {
  return (valor / total) * 100;
}

function showCustomDialog() {
  var html = HtmlService.createHtmlOutputFromFile('postFactura')
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Elige una opción');
}


function ConvertirFecha(opcion) {

  // Llama a la función ObtenerFecha para obtener la fecha formateada
  let fechaFormateada = ObtenerFecha(opcion);

  // Divide la fecha en día, mes y año
  let [dia, mes, año] = fechaFormateada.split("/");

  // Reorganiza la fecha en formato YYYY-MM-DD
  let fechaConvertida = `${año}-${mes}-${dia}`;

  return fechaConvertida;
}

function SumarDiasAFecha(dias) {
  // Obtiene la fecha en formato yyyy-MM-dd
  let fechaConvertida = ConvertirFecha();

  // Descompone la fecha en año, mes y día
  let [año, mes, dia] = fechaConvertida.split("-").map(Number);

  // Crea un objeto Date con los valores de año, mes y día
  let fecha = new Date(año, mes - 1, dia); // mes - 1 porque los meses en Date son indexados desde 0

  // Suma el número de días a la fecha
  fecha.setDate(fecha.getDate() + dias);

  // Formatea la nueva fecha en formato yyyy-MM-dd
  let nuevoAño = fecha.getFullYear();
  let nuevoMes = ("0" + (fecha.getMonth() + 1)).slice(-2); // Asegura dos dígitos para el mes
  let nuevoDia = ("0" + fecha.getDate()).slice(-2); // Asegura dos dígitos para el día

  let nuevaFecha = `${nuevoAño}-${nuevoMes}-${nuevoDia}`;

  return nuevaFecha;
}





//--------------------------------------------------------------------------------------------//
function obtenerDatosFactura(factura) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListadoEstado');
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copia de Plantilla');
  var invoiceColIndex = 5; // Columna F (indexada desde 0)
  var jsonColIndex = 12; // Columna M (indexada desde 0)
  if (!targetSheet) throw new Error("La hoja 'Copia de Plantilla' no existe.");
  var wasHidden = targetSheet.isSheetHidden();

  if (wasHidden) {
    targetSheet.showSheet(); // Mostrar la hoja temporalmente
  }

  Logger.log("factura " + factura)
  Logger.log("data length " + data.length)
  Logger.log(typeof (factura))
  //Logger.log("data +"+data)
  Logger.log(wasHidden)



  for (var i = 1; i < data.length; i++) { // Comienza en 1 para saltar la fila de encabezado
    //Logger.log(data[i][invoiceColIndex])
    //Logger.log(typeof(data[i][invoiceColIndex]))
    Logger.log("error " + data[i][invoiceColIndex])
    if (data[i][invoiceColIndex] == factura) {
      var jsonData = data[i][jsonColIndex];
      Logger.log("jsondata " + jsonData)
      if (jsonData) {
        try {
          var invoiceData = JSON.parse(jsonData);
          let Asesor = invoiceData.Delivery
          var facturaNumero = invoiceData.InvoiceGeneralInformation.InvoiceNumber;
          var cliente = invoiceData.CustomerInformation.RegistrationName;
          var nif = invoiceData.CustomerInformation.Identification;
          var codigo = invoiceData.CustomerInformation.CustomerCode;
          var direccion = invoiceData.CustomerInformation.AddressLine;
          var telefono = invoiceData.CustomerInformation.Telephone;
          var poblacion = invoiceData.CustomerInformation.CityName;
          var provincia = invoiceData.CustomerInformation.SubdivisionName;
          var pais = invoiceData.CustomerInformation.CountryName;
          var fechaEmision = invoiceData.CustomerInformation.DV;
          var formaPago = invoiceData.PaymentSummary.PaymentType;
          var listaProductos = invoiceData.ItemInformation;
          var numeroProductos = 0;
          var descuentosFactura = parseFloat(invoiceData.InvoiceTotal.PrePaidAmount);
          let descuentoGeneralesFactura = parseFloat(invoiceData.InvoiceTotal.GeneralPrePaidAmount);
          var cargosFactura = parseFloat(invoiceData.InvoiceTotal.ChargeTotalAmount);
          var totalFacturaJSON = parseFloat(invoiceData.InvoiceTotal.PayableAmount);
          let totalFacturaLetra = int2word(totalFacturaJSON)
          totalFacturaLetra = capitalizarPrimeraPalabra(totalFacturaLetra)
          Logger.log("totalFacturaLetra " + totalFacturaLetra)
          var valorPagar = totalFacturaLetra //arreglar
          var notaPago = invoiceData.PaymentSummary.PaymentNote;
          var observaciones = invoiceData.InvoiceGeneralInformation.Note;

          let ReqEquivalencia = parseFloat(invoiceData.InvoiceTotal.totalCargoEqui)
          let retenciones = parseFloat(invoiceData.InvoiceTotal.totalRet)
          let totalLinea = totalFacturaJSON


          var filasInsertadas = 0;
          var filasInsertadasPorProductos = 0;
          var grupoIva = {};

          var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copia de Plantilla'); // Hoja donde quieres insertar el NIF
          if (!targetSheet) {
            targetSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Copia de Plantilla');
          }

          var hojaCeldas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Celdas Plantilla');

          for (var j = 0; j < listaProductos.length; j++) {
            numeroProductos += 1;
            var numeroCelda = 19 + j;
            if (numeroProductos > 1) {
              targetSheet.insertRowAfter(numeroCelda);
              targetSheet.getRange('C' + (numeroCelda + 1) + ':E' + (numeroCelda + 1)).merge();
              filasInsertadas += 1;
              filasInsertadasPorProductos += 1;
            }
            // var celdaItem = targetSheet.getRange('A'+numeroCelda);
            // celdaItem.setBorder(true,true,true,true,null,null,null,null);
            // celdaItem.setValue(numeroProductos);
            // celdaItem.setHorizontalAlignment('center');

            var celdaReferencia = targetSheet.getRange('A' + numeroCelda);
            celdaReferencia.setBorder(true, true, true, true, null, null, null, null);
            celdaReferencia.setValue(listaProductos[j].ItemReference);
            celdaReferencia.setHorizontalAlignment('center');

            var celdaDespricion = targetSheet.getRange('C' + numeroCelda);
            celdaDespricion.setBorder(true, true, true, true, null, null, null, null);
            celdaDespricion.setValue(listaProductos[j].Name);
            celdaDespricion.setHorizontalAlignment('center');

            var celdaCantidad = targetSheet.getRange('F' + numeroCelda);
            celdaCantidad.setBorder(true, true, true, true, null, null, null, null);
            celdaCantidad.setValue(listaProductos[j].Quatity);
            celdaCantidad.setHorizontalAlignment('center');

            var celdaPrecioUnitario = targetSheet.getRange('G' + numeroCelda);
            celdaPrecioUnitario.setBorder(true, true, true, true, null, null, null, null);
            celdaPrecioUnitario.setValue(listaProductos[j].Price);
            celdaPrecioUnitario.setHorizontalAlignment('normal');
            celdaPrecioUnitario.setNumberFormat('€#,##0.00')

            var celdaSubtotal = targetSheet.getRange('H' + numeroCelda);
            celdaSubtotal.setBorder(true, true, true, true, null, null, null, null);
            celdaSubtotal.setValue(listaProductos[j].LineExtensionAmount);
            celdaSubtotal.setHorizontalAlignment('normal');
            celdaSubtotal.setNumberFormat('€#,##0.00')

            var celdaIva = targetSheet.getRange('I' + numeroCelda);
            celdaIva.setBorder(true, true, true, true, null, null, null, null);
            var percent = listaProductos[j].TaxesInformation[0].Percent;
            percent = percent.slice(0, -1);
            percent = parseFloat(percent);
            celdaIva.setValue(percent / 100);
            celdaIva.setNumberFormat('0%');
            celdaIva.setHorizontalAlignment('center');

            var celdaDescuento = targetSheet.getRange('J' + numeroCelda);
            celdaDescuento.setBorder(true, true, true, true, null, null, null, null);
            celdaDescuento.setValue(parseFloat(listaProductos[j].TaxesInformation[0].Descuento));
            celdaDescuento.setNumberFormat('0.00%')
            celdaDescuento.setHorizontalAlignment('center');

            var celdaRetencion = targetSheet.getRange('K' + numeroCelda);
            celdaRetencion.setBorder(true, true, true, true, null, null, null, null);
            celdaRetencion.setValue(parseFloat(listaProductos[j].TaxesInformation[0].Retencion));
            celdaRetencion.setNumberFormat('0%')
            celdaRetencion.setHorizontalAlignment('center');

            var celdaRecargoEquivalencia = targetSheet.getRange('L' + numeroCelda);
            celdaRecargoEquivalencia.setBorder(true, true, true, true, null, null, null, null);
            celdaRecargoEquivalencia.setValue(parseFloat(listaProductos[j].TaxesInformation[0].RecgEquivalencia));
            celdaRecargoEquivalencia.setNumberFormat('0.00%')
            celdaRecargoEquivalencia.setHorizontalAlignment('center');


            var celdaTotalLinea = targetSheet.getRange('M' + numeroCelda);
            celdaTotalLinea.setBorder(true, true, true, true, null, null, null, null);
            //subtotal+(subtotal*iva)+(subtotal*recargo)-(subtotal*retencion)
            Logger.log("LineTotal " + listaProductos[j].LineTotal)
            celdaTotalLinea.setValue(listaProductos[j].LineTotal);
            celdaTotalLinea.setNumberFormat('€#,##0.00');
            celdaTotalLinea.setHorizontalAlignment('normal');


            var producto = listaProductos[j]
            //crea un diccionario que la llave sea el % de iva y el valor sea el total de la linea
            Logger.log(grupoIva + "before")
            if (grupoIva.hasOwnProperty(percent)) {
              grupoIva[percent] += producto.TaxesInformation[0].TaxableAmount;
            } else {
              grupoIva[percent] = producto.TaxesInformation[0].TaxableAmount;
            }
            Logger.log("grupoIva after" + grupoIva)
          }
          var contador = 0;
          var auxiliarFilasInsertadas = filasInsertadas;
          for (var key in grupoIva) {
            Logger.log("grupo iva")
            if (grupoIva.hasOwnProperty(key)) {
              Logger.log("dentro del primer if grupo iva")
              var numeroCelda = 30 + auxiliarFilasInsertadas;
              if (contador > 0) {
                Logger.log("dentro del segundo if grupo iva")
                targetSheet.insertRowAfter(numeroCelda);
                targetSheet.getRange('A' + (numeroCelda + 1) + ':D' + (numeroCelda + 1)).merge();
                targetSheet.getRange('F' + (numeroCelda + 1) + ':H' + (numeroCelda + 1)).merge();
                targetSheet.getRange('I' + (numeroCelda + 1) + ':M' + (numeroCelda + 1)).merge();
                filasInsertadas += 1;
                auxiliarFilasInsertadas += 1;
              } else {
                auxiliarFilasInsertadas += 1;
              }
              Logger.log("auxiliarfilasinseretadas after: " + auxiliarFilasInsertadas)
              Logger.log("pasando el segundo if")
              var celdaBaseImponible = targetSheet.getRange('A' + numeroCelda);
              celdaBaseImponible.setBorder(true, true, true, true, null, null, null, null);
              celdaBaseImponible.setValue(grupoIva[key]);
              celdaBaseImponible.setNumberFormat('€#,##0.00');
              celdaBaseImponible.setHorizontalAlignment('normal');

              var celdaPorcentajeIva = targetSheet.getRange('E' + numeroCelda);
              celdaPorcentajeIva.setBorder(true, true, true, true, null, null, null, null);
              celdaPorcentajeIva.setValue(key / 100);
              celdaPorcentajeIva.setNumberFormat('0%');
              celdaPorcentajeIva.setHorizontalAlignment('center');

              var celdaIVA = targetSheet.getRange('F' + numeroCelda);
              celdaIVA.setBorder(true, true, true, true, null, null, null, null);
              celdaIVA.setFormula('=A' + numeroCelda + '*E' + numeroCelda);
              celdaIVA.setNumberFormat('€#,##0.00');
              celdaIVA.setHorizontalAlignment('normal');

              var celdaTotal = targetSheet.getRange('I' + numeroCelda);
              celdaTotal.setBorder(true, true, true, true, null, null, null, null);
              celdaTotal.setFormula('=A' + numeroCelda + '+F' + numeroCelda);
              celdaTotal.setNumberFormat('€#,##0.00');
              celdaTotal.setHorizontalAlignment('normal');

              contador += 1;
              Logger.log('IVA: ' + key + '%');

            }
          }

          //Extaccion celdas de datos cliente
          var clienteCeldaHoja = hojaCeldas.getRange('E3').getValue();
          var nifCeldaHoja = hojaCeldas.getRange('E4').getValue();
          var codigoCeldaHoja = hojaCeldas.getRange('E8').getValue();
          var direccionCeldaHoja = hojaCeldas.getRange('E5').getValue();
          var telefonoCeldaHoja = hojaCeldas.getRange('E7').getValue();
          var poblacionCeldaHoja = hojaCeldas.getRange('E6').getValue();
          var fechaEmisionCeldaHoja = hojaCeldas.getRange('E9').getValue();
          var formaPagoCeldaHoja = hojaCeldas.getRange('E10').getValue();
          let contactoCeldaHoja = hojaCeldas.getRange("E11").getValue();

          //factura
          var celdaNumFactura = targetSheet.getRange('A9');
          //Datos Cliente
          var clienteCell = targetSheet.getRange(clienteCeldaHoja);
          var nifCell = targetSheet.getRange(nifCeldaHoja);
          var codigoCell = targetSheet.getRange(codigoCeldaHoja);
          var direccionCell = targetSheet.getRange(direccionCeldaHoja);
          var telefonoCell = targetSheet.getRange(telefonoCeldaHoja);
          var poblacionCell = targetSheet.getRange(poblacionCeldaHoja);
          var fechaEmisionCell = targetSheet.getRange(fechaEmisionCeldaHoja);
          var formaPagoCell = targetSheet.getRange(formaPagoCeldaHoja);
          let contactoCell = targetSheet.getRange(contactoCeldaHoja);
          var valorPagarCell = targetSheet.getRange('B' + (41 + filasInsertadas));
          var notaPagoCell = targetSheet.getRange('A' + (45 + filasInsertadas));
          var observacionesCell = targetSheet.getRange('A' + (50 + filasInsertadas));
          var totalItemsCell = targetSheet.getRange('B' + (21 + filasInsertadasPorProductos));
          var descuentosCell = targetSheet.getRange('A' + (24 + filasInsertadasPorProductos));
          var cargosCell = targetSheet.getRange('D' + (24 + filasInsertadasPorProductos));
          var sumaBaseImponible = targetSheet.getRange('A' + (32 + filasInsertadas));
          var sumaImpIva = targetSheet.getRange('F' + (32 + filasInsertadas));
          var sumaTotal = targetSheet.getRange('I' + (32 + filasInsertadas));

          var totalRetenciones = targetSheet.getRange('A' + (36 + filasInsertadas));
          var totalCrgEquivalencia = targetSheet.getRange('D' + (36 + filasInsertadas));
          var totalCargos = targetSheet.getRange('G' + (36 + filasInsertadas));
          var totalDescuentos = targetSheet.getRange('K' + (36 + filasInsertadas));

          var totalDeFactura = targetSheet.getRange('H' + (38 + filasInsertadas));


          const resultado = dividirString(cliente)
          celdaNumFactura.setValue("FACTURA DE VENTA NO. " + facturaNumero);
          clienteCell.setValue(resultado[0]);
          nifCell.setValue(nif);
          // codigoCell.setValue(codigo);
          direccionCell.setValue(direccion);
          telefonoCell.setValue(telefono);
          contactoCell.setValue(Asesor)
          // Ajustar la forma en que se ve el pais - IMPORTANTE
          if (poblacion == "" || provincia == "" || pais == "") {
            var columnaPoblacion = poblacionCell.getColumn();
            var filaPoblacion = poblacionCell.getRow();
            targetSheet.getRange(filaPoblacion, columnaPoblacion - 1).clearContent();
          } else {
            poblacionCell.setValue(poblacion + ', ' + provincia + ', ' + pais);
          }

          totalRetenciones.setNumberFormat('€#,##0.00');
          totalRetenciones.setHorizontalAlignment('normal');

          totalCrgEquivalencia.setNumberFormat('€#,##0.00');
          totalCrgEquivalencia.setHorizontalAlignment('normal');

          totalDeFactura.setNumberFormat('€#,##0.00');
          totalDeFactura.setHorizontalAlignment('normal');

          cargosCell.setNumberFormat('€#,##0.00');

          totalDescuentos.setNumberFormat('€#,##0.00');

          descuentosCell.setNumberFormat('€#,##0.00');

          totalCargos.setNumberFormat('€#,##0.00')

          fechaEmisionCell.setValue(fechaEmision);
          formaPagoCell.setValue(formaPago);
          valorPagarCell.setValue(valorPagar);
          notaPagoCell.setValue(notaPago);
          observacionesCell.setValue(observaciones);
          // totalItemsCell.setValue(numeroProductos);
          Logger.log("descuentoGeneralesFactura: " + descuentoGeneralesFactura)
          descuentosCell.setValue(descuentoGeneralesFactura);
          cargosCell.setValue(cargosFactura);
          sumaBaseImponible.setFormula('=SUM(A' + (30 + numeroProductos - 1) + ':A' + (31 + filasInsertadas - 1) + ')');
          sumaImpIva.setFormula('=SUM(F' + (30 + numeroProductos - 1) + ':F' + (31 + filasInsertadas - 1) + ')');
          sumaTotal.setFormula('=SUM(I' + (30 + numeroProductos - 1) + ':I' + (31 + filasInsertadas - 1) + ')');
          totalRetenciones.setValue(retenciones);
          totalCrgEquivalencia.setValue(ReqEquivalencia);
          totalCargos.setValue(cargosFactura);
          Logger.log("descuentosFactura: " + descuentosFactura)
          totalDescuentos.setValue(descuentosFactura);

          totalDeFactura.setValue(totalLinea);






          var itemCellPrueba = targetSheet.getRange('A19')
          var hojaEnBlanco = clienteCell.isBlank() || formaPagoCell.isBlank() || celdaBaseImponible.isBlank();
          while (hojaEnBlanco) {
            Utilities.sleep(2000);
            Logger.log("dentro de while")
            hojaEnBlanco = clienteCell.isBlank() || formaPagoCell.isBlank() || itemCellPrueba.isBlank() || celdaBaseImponible.isBlank();
          }

          if (!hojaEnBlanco) {
            Logger.log("entrar hoja en blanco")
            var pdfFactura = generatePdfFromPlantilla();
            resetPlantilla();
            var id = subirFactura2(facturaNumero, pdfFactura);

            if (wasHidden) {
              targetSheet.hideSheet();
            }

            return id;
          }


        } catch (e) {
          Logger.log('Error parsing JSON for row ' + (i + 1) + ': ' + e.message);
        }
      }
      break//ojo esto debo de quitarlo
    }
  }



  Logger.log('Invoice number ' + factura + ' not found.');
}

function capitalizarPrimeraPalabra(cadena) {
  if (typeof cadena !== 'string') {
    throw new Error('El argumento debe ser una cadena.');
  }

  // Convertimos toda la cadena a minúsculas para garantizar consistencia
  cadena = cadena.toLowerCase();

  // Dividimos la cadena en palabras
  const palabras = cadena.split(' ');

  // Capitalizamos la primera palabra y unimos el resto sin modificar
  palabras[0] = palabras[0].charAt(0).toUpperCase() + palabras[0].slice(1);

  // Unimos las palabras de nuevo en una cadena
  return palabras.join(' ');
}

function dividirString(string) {
  if (!string || typeof string !== "string") return ["", ""];
  const partes = string.match(/^(.*?)([-+]?\d+.*)$/); // Divide texto y número
  if (!partes) return [string, ""];
  return [partes[1].trim(), partes[2].trim()];
}

function testWriteNIFToPlantilla() {
  var invoiceNumber = '192'; // Reemplaza con el número de factura deseado
  Logger.log(obtenerDatosFactura(invoiceNumber));
}

function resetPlantilla() {
  Logger.log("entro a reset")
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copia de Plantilla');

  // Borrar información de productos
  var colProductos = "A";
  var lineaProductos = 19;
  limpiarTablas(colProductos, lineaProductos);

  var colBases = "E";
  var lineaBases = 30;
  limpiarTablas(colBases, lineaBases);

  // Borrar información del cliente
  targetSheet.getRange('B12').clearContent();
  targetSheet.getRange('B13').clearContent();
  targetSheet.getRange('B14').clearContent();
  targetSheet.getRange('B15').clearContent();
  targetSheet.getRange('B16').clearContent();
  targetSheet.getRange('K12').clearContent();
  targetSheet.getRange('K13').clearContent();
  targetSheet.getRange('K14').clearContent();
  targetSheet.getRange('K15').clearContent();

  // Borrar valor a pagar, nota de pago y observaciones
  targetSheet.getRange('B41').clearContent();
  targetSheet.getRange('A45').clearContent();
  targetSheet.getRange('A50').clearContent();

  // Borrar total de items, descuentos y cargos
  targetSheet.getRange('B21').clearContent();
  targetSheet.getRange('A24').clearContent();
  targetSheet.getRange('D24').clearContent();

  // Borrar totales
  targetSheet.getRange('A32').clearContent();
  targetSheet.getRange('F32').clearContent();
  targetSheet.getRange('I32').clearContent();
  targetSheet.getRange('A36').clearContent();
  targetSheet.getRange('D36').clearContent();
  targetSheet.getRange('G36').clearContent();
  targetSheet.getRange('K36').clearContent();

}

function limpiarTablas(columna, linea) {
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copia de Plantilla');
  var primeraFila = targetSheet.getRange(linea + ":" + linea);
  primeraFila.clearContent();
  linea++;
  while (!targetSheet.getRange(columna + linea).isBlank()) {
    targetSheet.deleteRow(linea);
  }
}

function sacarColumnaFila(celda) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Celdas Plantilla');
  var celdaDestino = hoja.getRange(celda).getValue();
  var match = celdaDestino.match(/([A-Z]+)(\d+)/);
  if (match) {
    var columna = match[1];  // 'B'
    var fila = parseInt(match[2], 10);  // 21

    return [columna, fila];
  } else {
    Logger.log('No se pudo dividir la referencia de celda.');
  }
}

function pruebaSacar() {
  var lista = sacarColumnaFila("E18")
  Logger.log(lista)
}

function subirFactura(nombre, pdfBlob) {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  var folderId = hojaDatosEmisor.getRange("B14").getValue();
  var folder = DriveApp.getFolderById(folderId);
  var file = folder.createFile(pdfBlob.setName(`Factura ${nombre}.pdf`));
  var id = file.getId();
  return id;
}

function crearCarpeta() {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let folder = DriveApp.createFolder("FacturasApp");
  Logger.log('Folder created: ' + folder.getName() + ' (ID: ' + folder.getId() + ')');
  let id = folder.getId();
  hojaDatosEmisor.getRange("B14").setValue(id);
}


function crearCarpetaConDriveAPI() {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  var nombreCarpeta = "FacturasApp";

  var folderMetadata = {
    'name': nombreCarpeta,
    'mimeType': 'application/vnd.google-apps.folder'
  };

  var folder = Drive.Files.create(folderMetadata);  // Usamos el servicio avanzado de Drive

  var id = folder.id;  // Obtenemos el ID de la nueva carpeta
  hojaDatosEmisor.getRange("B14").setValue(id);
  Logger.log("Carpeta creada")
}

function eliminarCarpetaConDriveAPI() {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let idCarpeta = hojaDatosEmisor.getRange("B14").getValue();  // Obtenemos el ID de la carpeta desde una celda

  try {
    Drive.Files.remove(idCarpeta);  // Elimina la carpeta usando el servicio avanzado de Drive
    Logger.log("Carpeta eliminada exitosamente.");
    hojaDatosEmisor.getRange("B14").setValue("");
  } catch (e) {
    Logger.log("Error al eliminar la carpeta: " + e.message);
    hojaDatosEmisor.getRange("B14").setValue("");
  }
}


function subirFactura2(nombre, pdfBlob) {
  Logger.log("subir factura 2")
  let hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos de emisor');
  let IdCarpeta = hoja.getRange("B14").getValue()
  let fileMetadata = {
    'name': 'Factura ' + nombre + '.pdf',
    'mimeType': 'application/pdf',
    'parents': [IdCarpeta]  // Opcional: si quieres especificar una carpeta
  };
  let file = Drive.Files.create(fileMetadata, pdfBlob);


  Logger.log('PDF creado y subido: ' + file.id);
  return file.id;
}


function filtroHistorialFacturas(tipoFiltro) {

  Logger.log("debtro de fitrlo historial")
  Logger.log("tipoFiltro " + tipoFiltro)
  let Formula = ''
  let hojahistorial = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas');
  if (tipoFiltro == "Numero factura") {
    Formula = "=FILTER('Historial Facturas Data'!A2:E1000;ISNUMBER(SEARCH(C5;'Historial Facturas Data'!A2:A1000)))"
  } else if (tipoFiltro == "NIF") {
    Formula = "=FILTER('Historial Facturas Data'!A2:E1000;ISNUMBER(SEARCH(C5;'Historial Facturas Data'!C2:C1000)))"
  } else if (tipoFiltro == "Cliente") {
    Formula = "=FILTER('Historial Facturas Data'!A2:E1000;ISNUMBER(SEARCH(C5;'Historial Facturas Data'!B2:B1000)))"
  } else if (tipoFiltro == "Estado") {
    Formula = "=FILTER('Historial Facturas Data'!A2:E1000;ISNUMBER(SEARCH(C5;'Historial Facturas Data'!E2:E1000)))"
  }

  hojahistorial.getRange("B8").setValue(Formula)


}