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

var diccionarioCaluclarIva={
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



function verificarEstadoCarpeta(){
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let idCarpeta = hojaDatosEmisor.getRange("B15").getValue();  // Obtenemos el ID de la carpeta desde una celda
  Logger.log("idCarpeta "+idCarpeta)
  if (idCarpeta==""){
    //hoja toca ver si es la misma que esta en la google drive 
    SpreadsheetApp.getUi().alert("Necesitas primero crear la carpeta en donde se guardaran las facturas, dirigete a la hoja Datos de emisor y crea la carpeta dandole click al boton crear carpeta")
    return false
  }else{
    Logger.log("la carpeta si existe");
    return true
  }

}

function guardarFactura(){
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let estadoVinculacion=hojaDatosEmisor.getRange("B16").getValue();
  let estadoFactura=verificarEstadoValidoFactura();
  if(estadoVinculacion=="Desvinculado"){
    SpreadsheetApp.getUi().alert("Recuerda que antes de poder generar una factura es necesario haber vinculado tu cuenta de FacturasApp")
  }
  else if(estadoFactura.success){
    //factura valida
    // generar json
    
    let respuesta = verificarEstadoCarpeta()
    let respuestaEstadoConsecutivo=verificarEstadoConsecutivo()
    if(respuesta && respuestaEstadoConsecutivo){
      guardarYGenerarInvoice()
      guardarFacturaHistorial()
      Logger.log("guardar factura")
      limpiarHojaFactura()
      
    }else{
      return
    }

    
  }else{
    SpreadsheetApp.getUi().alert("Error al generar factura. "+estadoFactura.message)
  }
  

}
function agregarFilaNueva(){
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

function agregarProductoDesdeFactura(cantidad,producto){
  var spreadsheet = SpreadsheetApp.getActive();
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let factura_sheet = spreadsheet.getSheetByName('Factura');
  let taxSectionStartRow = getTaxSectionStartRow(hojaFactura);//recordar este devuelve el lugar en donde deberian estar base imponible, toca restar -1
  const productStartRow = 15;
  const lastProductRow = getLastProductRow(hojaFactura, productStartRow, taxSectionStartRow);

  let dictInformacionProducto ={}
  if(producto==="" || cantidad==="" || cantidad===0){
    throw new Error('Porfavor elige un producto y un cantidad adecuado');
  }else{
    Logger.log("entra a dictInformacionProducto")
    dictInformacionProducto = obtenerInformacionProducto(producto);
  }
  Logger.log("dictInformacionProducto "+dictInformacionProducto["codigo Producto"])
  let rowParaDatos=lastProductRow
  let rowParaTotalTaxes=taxSectionStartRow
  let cantidadProductos=hojaFactura.getRange("B16").getValue()//estado defaul de total productos
  if(cantidadProductos===0 || cantidadProductos===""){
    factura_sheet.getRange("A15").setValue(dictInformacionProducto["codigo Producto"])
    factura_sheet.getRange("B15").setValue(producto)
    factura_sheet.getRange("C15").setValue(cantidad)
    factura_sheet.getRange("D15").setValue(dictInformacionProducto["valor Unitario"])
    factura_sheet.getRange("G15").setValue(dictInformacionProducto["IVA"])
    factura_sheet.getRange("I15").setValue(dictInformacionProducto["retencion"])
    factura_sheet.getRange("J15").setValue(dictInformacionProducto["Recargo de equivalencia"])

  }else{
    hojaFactura.insertRowAfter(lastProductRow)
    rowParaTotalTaxes=taxSectionStartRow+1
    rowParaDatos=lastProductRow+1
    factura_sheet.getRange("A"+String(rowParaDatos)).setValue(dictInformacionProducto["codigo Producto"])
    factura_sheet.getRange("B"+String(rowParaDatos)).setValue(producto)
    factura_sheet.getRange("C"+String(rowParaDatos)).setValue(cantidad)
    factura_sheet.getRange("E"+String(rowParaDatos)).setValue("=D"+String(rowParaDatos)+"+(D"+String(rowParaDatos)+"*G"+String(rowParaDatos)+")")//AGG COSA DE CON IVA
    factura_sheet.getRange("F"+String(rowParaDatos)).setValue("=(D"+String(rowParaDatos)+"-(D"+String(rowParaDatos)+"*H"+String(rowParaDatos)+"))*C"+String(rowParaDatos))//subtotal
    factura_sheet.getRange("D"+String(rowParaDatos)).setValue(dictInformacionProducto["valor Unitario"])//valor unitario
    factura_sheet.getRange("G"+String(rowParaDatos)).setValue(dictInformacionProducto["IVA"])//IVA
    
    factura_sheet.getRange("I"+String(rowParaDatos)).setValue(dictInformacionProducto["retencion"])//Retencion
    factura_sheet.getRange("J"+String(rowParaDatos)).setValue(dictInformacionProducto["Recargo de equivalencia"])//Recargo de equivalencia
    factura_sheet.getRange("K"+String(rowParaDatos)).setValue("=F"+String(rowParaDatos)+"+(F"+String(rowParaDatos)+"*G"+String(rowParaDatos)+")-(F"+String(rowParaDatos)+"*I"+String(rowParaDatos)+")+(F"+String(rowParaDatos)+"*J"+String(rowParaDatos)+")")//total linea
  }

  

  Logger.log("rowParaDatos "+rowParaDatos)
  Logger.log("Number(taxSectionStartRow-1) "+Number(taxSectionStartRow-1))



  updateTotalProductCounter(rowParaDatos, productStartRow,hojaFactura, rowParaTotalTaxes);
  calcularImporteYTotal(rowParaDatos,productStartRow,rowParaTotalTaxes,hojaFactura)
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
function probarInsertarImagen(){
  insertarImagenBorrarFila(15)
}
function insertarImagenBorrarFila(fila){
  var spreadsheet = SpreadsheetApp.getActive();
  let hojaFcatura=spreadsheet.getSheetByName('Factura');
  let imagenURL="https://i.postimg.cc/RFZ45sgp/basura3.png"
  var cell = hojaFcatura.getRange('H'+fila);
  cell.setHorizontalAlignment('center');
  var imageBlob = UrlFetchApp.fetch(imagenURL).getBlob();
  var image = hojaFcatura.insertImage(imageBlob, cell.getColumn(), cell.getRow());
  var numFactura = hojaFcatura.getRange('A'+fila).getValue();
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

  var idArchivo = obtenerDatosFactura(numeroFactura);
  Logger.log("idarchivo"+idArchivo)
  guardarIdArchivo(idArchivo, numeroFactura);

  // var html = HtmlService.createHtmlOutputFromFile('postFactura')
  //   .setTitle('Menú');
  // SpreadsheetApp.getUi()
  //   .showSidebar(html);
  
  showCustomDialog()
}

function guardarIdArchivo(idArchivo, numeroFactura) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  var lastRow = hoja.getLastRow();
  var newRow = lastRow + 1;
  hoja.getRange("A" + newRow).setValue(numeroFactura).setBorder(true, true, true, true, null, null, null, null);
  hoja.getRange("B" + newRow).setValue(idArchivo).setBorder(true, true, true, true, null, null, null, null);

}

// Funciones obsoletas eliminadas - ahora se usa el nuevo API de FacturasApp
// - convertPdfToBase64Historial()
// - convertPdfToBase64()
// Reemplazadas por obtenerPDFFacturaBase64() que usa el endpoint PDFInvoice
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
        
        // Actualizar el estado en el historial si es necesario
        if (responseData.id) {
          Logger.log("Factura creada con ID: " + responseData.id);
        }
      }
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

function enviarFacturaHistorial(numeroFactura){
  let spreadsheet = SpreadsheetApp.getActive()
  const scriptProps = PropertiesService.getDocumentProperties();
  let ambiente = scriptProps.getProperty('Ambiente')
  
  // Nuevo endpoint para AddInvoice
  let url
  if (ambiente=="Pruebas"){
    url = "https://facturasapp-qa.cenet.ws/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/AddInvoice"
  }else{
    url = "https://www.facturasapp.com/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/AddInvoice";
  }
  
  // Buscar la factura en ListadoEstado por número de factura
  let listadoEstado = spreadsheet.getSheetByName('ListadoEstado');
  let dataRange = listadoEstado.getDataRange();
  let data = dataRange.getValues();
  let invoiceColIndex = 5; // Columna F (número de factura)
  let jsonColIndex = 12; // Columna M (JSON del nuevo formato)
  
  let jsonFieldInvoice = null;
  
  for (let i = 1; i < data.length; i++) { // Comienza en 1 para saltar la fila de encabezado
    if (data[i][invoiceColIndex] == numeroFactura) {
      jsonFieldInvoice = data[i][jsonColIndex];
      break;
    }
  }
  
  if (!jsonFieldInvoice) {
    SpreadsheetApp.getUi().alert("Error: No se encontró el JSON de la factura " + numeroFactura + ". Asegúrese de que la factura haya sido generada con el nuevo formato.");
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
    var respuesta = UrlFetchApp.fetch(url, opciones);
    let responseText = respuesta.getContentText();
    Logger.log("Status: " + respuesta.getResponseCode());
    Logger.log("Response: " + responseText);
    
    if (respuesta.getResponseCode() === 200) {
      let responseData = JSON.parse(responseText);
      if (responseData.isError) {
        SpreadsheetApp.getUi().alert("Error de FacturasApp: " + responseData.messages);
      } else {
        SpreadsheetApp.getUi().alert("Factura " + numeroFactura + " enviada correctamente a FacturasApp. ID: " + responseData.id);
        
        // Actualizar el estado en el historial si es necesario
        if (responseData.id) {
          Logger.log("Factura " + numeroFactura + " creada con ID: " + responseData.id);
        }
      }
    } else {
      SpreadsheetApp.getUi().alert("Error al enviar la factura: " + responseText);
    }
  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
    SpreadsheetApp.getUi().alert("Error al enviar la factura a FacturasApp. Intente de nuevo. Si el error persiste comuníquese con soporte. Error: " + error.message);
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


function jsonAPIkey(usuario,contra){
  let json={
    "user": usuario,
    "password": contra
  }

  return json
}
function obtenerAPIkey(usuario, contra) {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let hojaDatos=spreadsheet.getSheetByName("Datos")
  const scriptProps = PropertiesService.getDocumentProperties();
  let ambiente = scriptProps.getProperty('Ambiente')
  Logger.log("apikeuy")
  Logger.log("Ambiente: "+ambiente)
  let url
  if (ambiente=="Pruebas"){
    url = "https://facturasapp-qa.cenet.ws/ApiGateway/AppSecurity/ApiKey";
  }else{
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
      hojaDatosEmisor.getRange("B16").setBackground('#ccffc7')  // Almacena el API Key en la celda
      hojaDatosEmisor.getRange("B16").setValue("Vinculado")
      hojaDatos.getRange("I21").setValue(apiKey)
    } else {
      hojaDatosEmisor.getRange("B16").setBackground('#FFC7C7')
      hojaDatosEmisor.getRange("B16").setValue("Desvinculado")
      throw new Error("Error de la API: " + contenidoRespuesta); // Muestra el error de la API
      
    }
  } catch (error) {
    Logger.log("Error al enviar el JSON a la API: " + error.message);
    hojaDatosEmisor.getRange("B16").setBackground('#FFC7C7')
    hojaDatosEmisor.getRange("B16").setValue("Desvinculado")
    hojaDatos.getRange("I21").setValue(0)
    SpreadsheetApp.getUi().alert("Error al vincular tu cuenta. Verifica que el usuario y la contraseña estén correctos e intenta de nuevo. Si el error persiste, comunícate con soporte.");
  }
}



function obtenerPDFFactura(numeroFactura) {
  let spreadsheet = SpreadsheetApp.getActive();
  const scriptProps = PropertiesService.getDocumentProperties();
  let ambiente = scriptProps.getProperty('Ambiente')
  
  // Nuevo endpoint para PDFInvoice
  let url
  if (ambiente=="Pruebas"){
    url = "https://facturasapp-qa.cenet.ws/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/PDFInvoice?invoiceNumber=" + encodeURIComponent(numeroFactura)
  }else{
    url = "https://www.facturasapp.com/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/PDFInvoice?invoiceNumber=" + encodeURIComponent(numeroFactura);
  }
  
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let APIkey = hojaDatos.getRange("I21").getValue()
  
  if (!APIkey) {
    SpreadsheetApp.getUi().alert("Error: No se encontró la API Key. Asegúrese de haber vinculado su cuenta de FacturasApp.");
    return null;
  }
  
  let opciones = {
    "method": "post",
    "headers": {"X-API-KEY": APIkey},
    'muteHttpExceptions': true
  };

  try {
    var respuesta = UrlFetchApp.fetch(url, opciones);
    let responseCode = respuesta.getResponseCode();
    
    if (responseCode === 200) {
      // La respuesta debería ser el PDF como bytes
      let pdfBlob = respuesta.getBlob();
      
      // Guardar el PDF en Google Drive
      let fileName = "Factura_" + numeroFactura + "_" + new Date().getTime() + ".pdf";
      let file = DriveApp.createFile(pdfBlob.setName(fileName));
      
      SpreadsheetApp.getUi().alert("PDF generado correctamente. Archivo guardado como: " + fileName);
      
      // Retornar el ID del archivo para referencia
      return file.getId();
    } else {
      let responseText = respuesta.getContentText();
      SpreadsheetApp.getUi().alert("Error al obtener el PDF: " + responseText);
      return null;
    }
  } catch (error) {
    Logger.log("Error al obtener el PDF: " + error.message);
    SpreadsheetApp.getUi().alert("Error al obtener el PDF. Error: " + error.message);
    return null;
  }
}

function obtenerPDFFacturaBase64(numeroFactura) {
  let spreadsheet = SpreadsheetApp.getActive();
  const scriptProps = PropertiesService.getDocumentProperties();
  let ambiente = scriptProps.getProperty('Ambiente')
  
  // Nuevo endpoint para PDFInvoice
  let url
  if (ambiente=="Pruebas"){
    url = "https://facturasapp-qa.cenet.ws/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/PDFInvoice?invoiceNumber=" + encodeURIComponent(numeroFactura)
  }else{
    url = "https://www.facturasapp.com/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/PDFInvoice?invoiceNumber=" + encodeURIComponent(numeroFactura);
  }
  
  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let APIkey = hojaDatos.getRange("I21").getValue()
  
  if (!APIkey) {
    Logger.log("Error: No se encontró la API Key");
    return null;
  }
  
  let opciones = {
    "method": "post",
    "headers": {"X-API-KEY": APIkey},
    'muteHttpExceptions': true
  };

  try {
    var respuesta = UrlFetchApp.fetch(url, opciones);
    let responseCode = respuesta.getResponseCode();
    
    if (responseCode === 200) {
      // La respuesta debería ser el PDF como bytes
      let pdfBlob = respuesta.getBlob();
      let base64String = Utilities.base64Encode(pdfBlob.getBytes());
      
      return base64String;
    } else {
      let responseText = respuesta.getContentText();
      Logger.log("Error al obtener el PDF: " + responseText);
      return null;
    }
  } catch (error) {
    Logger.log("Error al obtener el PDF: " + error.message);
    return null;
  }
}

// Función obsoleta convertPdfToBase64Prueba() eliminada
function convertPdfToBase64Prueba_OBSOLETA() {
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
    Drive.Permissions.create(permisos, idArchivo, {sendNotificationEmails: false});
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
  const usuario =obtenerUsuario()
  const propietario= obtenerPropietario()
  Logger.log("usuario "+usuario)
  Logger.log("propietario "+propietario)
  Logger.log("sale de linkdescargar")
  return data;
}

function enviarEmailPostFactura(email,historial=false,numFacturaAbuscar=null) {
  let hojaListadoEstado = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListadoEstado');
  let lastRowListado = hojaListadoEstado.getLastRow()
  let numFactura;
  let invoiceTotal;
  let fieldInvoiceJson;
  
  if(historial){
    // Buscar la factura por número en el historial
    let dataRange = hojaListadoEstado.getDataRange();
    let data = dataRange.getValues();
    let invoiceColIndex = 5; // Columna F (número de factura)
    let jsonColIndex = 12; // Columna M (JSON del nuevo formato)
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][invoiceColIndex] == numFacturaAbuscar) {
        numFactura = data[i][invoiceColIndex];
        fieldInvoiceJson = data[i][jsonColIndex];
        break;
      }
    }
    
    if (!fieldInvoiceJson) {
      return "Error: No se encontró la factura " + numFacturaAbuscar + " en el historial.";
    }
  } else {
    // Obtener la última factura
    numFactura = hojaListadoEstado.getRange(lastRowListado, 6).getValue(); // Columna F
    fieldInvoiceJson = hojaListadoEstado.getRange(lastRowListado, 13).getValue(); // Columna M
  }
  
  if (!fieldInvoiceJson) {
    return "Error: No se encontró el JSON de la factura.";
  }
  
  // Parsear el JSON para obtener el total
  let fieldInvoiceData = JSON.parse(fieldInvoiceJson);
  invoiceTotal = fieldInvoiceData.sumTotalTotal;
  
  Logger.log("email " + email)
  Logger.log("numFactura " + numFactura)
  Logger.log("invoiceTotal " + invoiceTotal)

  if (!email) {
    return "Por favor ingrese una dirección de correo válida.";
  }

  // Obtener el PDF usando el nuevo API
  let base64PDF = obtenerPDFFacturaBase64(numFactura);
  
  if (!base64PDF) {
    return "Error: No se pudo obtener el PDF de la factura desde FacturasApp.";
  }
  
  // Convertir base64 a blob
  let pdfBytes = Utilities.base64Decode(base64PDF);
  let pdfBlob = Utilities.newBlob(pdfBytes, 'application/pdf', 'Factura_' + numFactura + '.pdf');

  let hojaDatosEmisor = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos de emisor');
  let nombreCliente = hojaDatosEmisor.getRange("B1").getValue();

  var subject = `📄 Nueva factura de ${nombreCliente}`;
  var body = `¡Hola!\n` +
           `${nombreCliente} te ha enviado la siguiente factura:\n` +
           `🔹 Número de factura: ${numFactura}\n` +
           `💰 Valor: ${invoiceTotal} €\n` +
           `Si tienes alguna duda, contacta directamente con ${nombreCliente}.\n` +
           `Saludos,\n` +
           `${nombreCliente}\n\n`+
           `📌 ¿Necesitas facturación electrónica? Ahorra tiempo y factura fácilmente con FacturasApp\n` +
           `👉 Ver más: https://www.facturasapp.com/Publico/`;

  try {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      attachments: [pdfBlob]
    });

    return "PDF generado desde FacturasApp y enviado por correo electrónico a " + email;
  } catch (error) {
    Logger.log("Error al enviar email: " + error.message);
    return "Error al enviar el email: " + error.message;
  }
}


function ProcesarFormularioFactura(data) {
  var numFactura = data.numFactura;
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');

  var range = hoja.getRange('A2:A');
  var textFinder = range.createTextFinder(numFactura);
  var cell = textFinder.findNext();
  Logger.log("cell "+cell)
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
    Drive.Permissions.insert(permission, fileId, {sendNotificationEmails: false});
    return "Permiso actualizado. Intenta descargar nuevamente.";
  } catch (e) {
    return "Error al actualizar permisos: " + e.message;
  }
}


function verificarCodigo(codigo, nombreHoja, inHoja,lineEditada=null,codigoV="") {
  Logger.log("Verificar códigos");
  Logger.log("linea editada: "+lineEditada)
  // Obtener la hoja por nombre
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  let codigoNumero = String(codigo)


  try {
    let columna;
    let lastActiveRow = sheet.getLastRow();
    let rangeDatos;
    let pruebaPostRow=0
    Logger.log(lastActiveRow+"last acitrive row")
    // Determinar la columna y el rango según el tipo de hoja
    if (nombreHoja === "Clientes" && codigoV!=="codigo") {
      columna = 6; // Columna para el identificador de clientes
      rangeDatos = sheet.getRange(2, columna, lastActiveRow - (inHoja ? 2 : 1));
    }else if(nombreHoja==="Clientes" && codigoV==="codigo"){
      columna = 7;//columna codigo
      rangeDatos = sheet.getRange(2, columna, lastActiveRow - (inHoja ? 2 : 1));
    } else if (nombreHoja === "Productos") {
      columna = 2; // Columna para el código de productos
      pruebaPostRow=lastActiveRow - (inHoja? 2: 1)
      rangeDatos = sheet.getRange(2, columna, lastActiveRow - (inHoja? 2: 1));
    } else if (nombreHoja === "Historial Facturas Data") {
      columna = 1; // Columna para el número de factura
      rangeDatos = sheet.getRange(2, columna, lastActiveRow - 1);
    } else {
      Logger.log("Nombre de hoja no válido.");
      return false;
    }
    Logger.log("last active ro post"+pruebaPostRow)
    // Obtener los valores del rango como una matriz de números
    let datos = rangeDatos.getValues().flat().map(String);
    Logger.log("Datos obtenidos como números:");
    Logger.log(datos);

    // Convertir el código a número
    
    Logger.log(codigoNumero)
    // Verificar si algún valor en datos es exactamente igual al código
    for (let i = 0; i < datos.length; i++) {
      Logger.log("Datos i; "+"i:"+i+"datos: "+datos[i])
      
      if (datos[i] === codigoNumero) {
        if(i===lineEditada-2){
          Logger.log("dentro de continue")
          
        }else{

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




function  insertarImagen(fila) {
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



function inicarFacturaNueva(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let hojaFactura = spreadsheet.getSheetByName('Factura');
  let hojaInfoUsuario= spreadsheet.getSheetByName('Datos de emisor');
  let IABN=hojaInfoUsuario.getRange("B10").getValue()
  
  hojaFactura.getRange("B11").setValue(IABN)
  generarNumeroFactura(); 
  obtenerFechaYHoraActual();
}

function limpiarYEliminarFila(numeroFila,hoja,hojaTax){
  //funcion para el boton que se va a agregar al final del producto
  if (numeroFila>20 && numeroFila<hojaTax){
    hoja.deleteRow(numeroFila)
  }else{
    hoja.getRange("A"+String(numeroFila)).setValue("");//producto
    hoja.getRange("B"+String(numeroFila)).setValue("");//ref
    hoja.getRange("C"+String(numeroFila)).setValue("");//cantidad
    hoja.getRange("D"+String(numeroFila)).setValue(0);//CON IVa
    hoja.getRange("E"+String(numeroFila)).setValue(0);//sin iva
    //sheet.getRange("C"+String(posicionTaxInfo)).setValue(valorEnPorcentaje);
  }
}

function verificarYCopiarContacto(e) {
  let hojaFacturas = e.source.getSheetByName('Factura');
  let hojaContactos = e.source.getSheetByName('Clientes');
  let celdaEditada = e.range;



  let nombreContacto = celdaEditada.getValue();
  let ultimaColumnaPermitida = 20; // Columna del estado en la hoja de contactos
  let datosARetornar = ["B", "O","M","L","N","Q"]; // Columnas que quiero de la hoja de contactos


  if (nombreContacto==="Cliente"){
    Logger.log("Estado default")
  }else{
    let listaConInformacion = obtenerInformacionCliente(nombreContacto);
    if (listaConInformacion["Estado"]==="No Valido"){
      SpreadsheetApp.getUi().alert("Error: El contacto seleccionado no es válido.");
    }else{
      //asigna el valor del coldigo solamente porque ese fue lo que me pidieron no mas
      hojaFacturas.getRange("B3").setValue(listaConInformacion["Código cliente"]);
    }
  }


}


function generarNumeroFactura(){
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
    if(!cumple){
      Logger.log("No cumple con la estructura")
    }else{
      let numero = obtenerParteNumerica(consecutivo);
      
      if (numero > numeroMayor) {
        numeroMayor = numero;
        ultimoConsecutivo = consecutivo; // Guardamos el último número en formato original
      }
    }
    let numeroActual=0
    if (numeroMayor==-Infinity){
      const scriptProperties = PropertiesService.getDocumentProperties();
      numero = scriptProperties.getProperty('NumeroConescutivo');  // Ej: "123"
      letra  = scriptProperties.getProperty('LetraConescutivo');   // Ej: "abc"
      let consecutivo = letra+numero
      numeroActual =consecutivo
    }else{
      numeroActual = numeroMayor + 1;
    }
    let nuevoConsecutivo = generarNuevoConsecutivo(ultimoConsecutivo, numeroActual);

    sheet.getRange("G2").setValue(nuevoConsecutivo);
  }
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

function obtenerFechaYHoraActual(){ 
  let spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName('Factura');
  let zonaHorariaEspaña = "Europe/Madrid"
  let fecha = Utilities.formatDate(new Date(), zonaHorariaEspaña, "dd/MM/yyyy");
  let hora= Utilities.formatDate(new Date(), zonaHorariaEspaña, "HH:mm:ss");

  sheet.getRange("G4").setNumberFormat("dd/MM/yyyy");
  sheet.getRange("G4").setValue(String(fecha))
  sheet.getRange("G3").setValue(String(fecha))
  sheet.getRange("G7").setValue(hora)

  
  let valorFecha=sheet.getRange("G4").getValue();

  let fechaFormateada = Utilities.formatDate(new Date(valorFecha), zonaHorariaEspaña, "dd/MM/yyyy");
  Logger.log("valorFecha "+valorFecha)
  Logger.log("fecha "+fecha)
  Logger.log("fechaFormateada "+fechaFormateada)

}

function ObtenerFecha(opcion=null){
  let spreadsheet = SpreadsheetApp.getActive();
  let fechaFormateada
  let valorFecha
  let zonaHorariaEspaña = "Europe/Madrid"
  if(opcion=="pago"){
    let sheet = spreadsheet.getSheetByName('Factura');
    valorFecha=sheet.getRange("G3").getValue();
    Logger.log("valorFecha 1"+String(valorFecha))
    fechaFormateada = Utilities.formatDate(new Date(valorFecha), zonaHorariaEspaña, "dd/MM/yyyy");
  }else{
    let sheet = spreadsheet.getSheetByName('Factura');
    valorFecha=sheet.getRange("G4").getValue();
    Logger.log("valorFecha "+String(valorFecha))
    Logger.log("valorFecha 2"+valorFecha)
    fechaFormateada = Utilities.formatDate(new Date(valorFecha), zonaHorariaEspaña, "dd/MM/yyyy");
  }
  Logger.log("fecha formateada"+fechaFormateada)
  Logger.log("valorFecha "+String(valorFecha))
  return fechaFormateada
}



function obtenerDatosProductos(sheet,range,e){
    if ( range.getA1Notation() === "A14" || range.getA1Notation()=== "A15" || range.getA1Notation() === "A16" || range.getA1Notation()=== "A17" || range.getA1Notation()=== "A18") {
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

function getprefacturaValueA1(column, row) {
  return getsheetValueA1(prefactura_sheet, column, row);
}

function getprefacturaValue(prefactura_sheet,column, row) {

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
  
  var invoice_number = getprefacturaValue(prefactura_sheet,2, 7);//cambiamos los valores para llamar el numero de factura
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
    "Note": getprefacturaValue(prefactura_sheet,10, 2), //cambia los valores para llamar la nota de la factura
    "ExternalGR": false
    //"AdditionalProperty": AdditionalProperty
  }


  return InvoiceGeneralInformation;
}
function getPaymentSummary(startingRowTaxation) {
  let spreadsheet = SpreadsheetApp.getActive();
  let prefactura_sheet = spreadsheet.getSheetByName('Factura');
  let posTotalFactura=startingRowTaxation+7
  let posMontoNeto=startingRowTaxation+12
  var total_factura = prefactura_sheet.getRange("A"+String(posTotalFactura)).getValue();// por ahora esto no lo utilizamos ya que no hay descuentos
  var monto_neto = prefactura_sheet.getRange("B"+String(posMontoNeto)).getValue();
  //var numeros = new Intl.NumberFormat('es-CO', {maximumFractionDigits:0, style: 'currency', currency: 'COP' }).format(monto);
  var numeros_total = new Intl.NumberFormat().format(total_factura);
  var numeros_neto = new Intl.NumberFormat().format(monto_neto);
  //Browser.msgBox(`${numeros_total}  ${numeros_neto}`);

  Logger.log("total_factura"+total_factura)
  Logger.log("monto_neto"+monto_neto)
  var PaymentTypeTxt = prefactura_sheet.getRange("G5").getValue();
  var PaymentMeansTxt = prefactura_sheet.getRange("E4").getValue();
  let PaymentNote=prefactura_sheet.getRange("D11").getValue();
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
  
  // Obtener fechas
  let fechaFactura = new Date(prefactura_sheet.getRange("G4").getValue());
  let fechaVencimiento = new Date(prefactura_sheet.getRange("G3").getValue());
  let horaFactura = new Date().toTimeString().split(' ')[0];
  
  // Validar fechas
  if (isNaN(fechaFactura.getTime())) {
    fechaFactura = new Date();
  }
  if (isNaN(fechaVencimiento.getTime())) {
    fechaVencimiento = new Date();
  }

  // Procesar productos
  let products = [];
  let totalTaxBase = 0;
  let totalTax = 0;
  let totalSubTotal = 0;
  
  for (let i = 15; i < 15 + cantidadProductos; i++) {
    let filaActual = "A" + String(i) + ":K" + String(i);
    let rangoProducto = prefactura_sheet.getRange(filaActual);
    let productoData = rangoProducto.getValues()[0];
    
    let referencia = String(productoData[0] || "");
    let descripcion = String(productoData[1] || "");
    let cantidad = Number(productoData[2]) || 1;
    let precioUnitario = Number(productoData[3]) || 0;
    let subtotal = Number(productoData[5]) || 0;
    let ivaRate = Number(productoData[6]) || 0;
    let descuento = Number(productoData[7]) || 0;
    let retencion = Number(productoData[8]) || 0;
    let recargoEquivalencia = Number(productoData[9]) || 0;
    let totalLinea = Number(productoData[10]) || 0;
    
    // Calcular impuestos
    let taxAmount = subtotal * ivaRate;
    let withHoldings = subtotal * retencion;
    let surCharges = subtotal * recargoEquivalencia;
    let totalDiscount = subtotal * descuento;
    
    // Simplificando - eliminando arrays de taxes, withHoldings y discounts
    
    // Validar campos obligatorios de productos
    if (!descripcion || descripcion.trim() === "") {
      descripcion = "Producto sin descripción";
    }
    if (!referencia || referencia.trim() === "") {
      referencia = "REF-" + i;
    }
    if (cantidad <= 0) {
      cantidad = 1;
    }
    if (precioUnitario < 0) {
      precioUnitario = 0;
    }
    if (subtotal < 0) {
      subtotal = 0;
    }
    
    // Producto según el esquema FieldInvoice - MINIMO
    let producto = {
      typeUse: "001",
      reference: String(referencia).substring(0, 50),
      description: String(descripcion).substring(0, 100),
      unitPrice: Number(precioUnitario) || 0,
      quantity: Math.min(Number(cantidad) || 1, 999),
      subTotal: Number(subtotal) || 0,
      totalTax: Number(taxAmount) || 0,
      totalwithHoldings: Number(withHoldings) || 0,
      totalSurCharges: Number(surCharges) || 0,
      totaldiscount: Number(totalDiscount) || 0
    };
    
    products.push(producto);
    
    // Acumular totales
    totalSubTotal += subtotal;
    totalTaxBase += subtotal;
    totalTax += taxAmount;
  }

  // Obtener totales de la factura
  let totalFactura = 0;
  let cargoTotal = 0;
  let descuentoTotal = 0;
  
  if (prefactura_sheet.getRange("A31").getValue() === "Total factura") {
    totalFactura = prefactura_sheet.getRange("B31").getValue();
    cargoTotal = prefactura_sheet.getRange("B17").getValue() || 0;
    descuentoTotal = prefactura_sheet.getRange("B18").getValue() || 0;
  } else {
    let rowTotalFactura = startingRowTaxation + 12;
    let rowCargoFactura = startingRowTaxation - 2;
    let rowDescuentoFactura = startingRowTaxation - 1;
    totalFactura = prefactura_sheet.getRange(rowTotalFactura, 2).getValue();
    cargoTotal = prefactura_sheet.getRange("B" + String(rowCargoFactura)).getValue() || 0;
    descuentoTotal = prefactura_sheet.getRange("B" + String(rowDescuentoFactura)).getValue() || 0;
  }

  // Validar datos del cliente
  cliente = String(cliente || "");
  if (!cliente || cliente.trim() === "") {
    cliente = "Cliente sin nombre";
  }
  
  // Contactos según el esquema - MINIMO
  let contacts = [{
    contactType: "01",
    personType: "01", 
    companyName: String(cliente).substring(0, 450),
    customerCode: String(prefactura_sheet.getRange("B3").getValue() || "CLIENTE001").substring(0, 150),
    identification: String(CustomerInformation.Identification || "12345678A").substring(0, 20),
    tradeName: String(cliente).substring(0, 450),
    regime: "01",
    country: "ESP",
    province: "28",
    population: "28079",
    addressCustomer: String(CustomerInformation.Address || "Calle Principal 123").substring(0, 200)
  }];

  // Eliminando arrays problemáticos para simplificar

  // Validar totales
  if (isNaN(totalSubTotal) || totalSubTotal < 0) totalSubTotal = 0;
  if (isNaN(totalTaxBase) || totalTaxBase < 0) totalTaxBase = 0;
  if (isNaN(totalTax) || totalTax < 0) totalTax = 0;
  if (isNaN(totalFactura) || totalFactura < 0) totalFactura = 0;
  if (isNaN(cargoTotal) || cargoTotal < 0) cargoTotal = 0;
  if (isNaN(descuentoTotal) || descuentoTotal < 0) descuentoTotal = 0;
  
  // Validar número de factura
  let numeroFacturaValidado = String(InvoiceGeneralInformation.InvoiceNumber);
  if (!numeroFacturaValidado || numeroFacturaValidado.trim() === "") {
    numeroFacturaValidado = "FACT-" + Date.now();
  }
  
  // Extraer número actual y validar que sea mayor que 0
  let currentNumber = Number(numeroFacturaValidado.replace(/[^0-9]/g, ''));
  if (currentNumber <= 0) {
    currentNumber = Math.floor(Date.now() / 1000); // Usar timestamp como número
  }
  
  // Crear el JSON según el esquema FieldInvoice - VERSION MINIMA
  let fieldInvoice = {
    // SOLO CAMPOS OBLIGATORIOS
    contactName: String(cliente).substring(0, 30),
    invoiceNumber: numeroFacturaValidado.substring(0, 50),
    currentNumber: currentNumber,
    invoiceDate: fechaFactura.toISOString(),
    invoiceTime: horaFactura,
    invoiceIdTypeRegAEAT: "01",
    idPayment: "01",
    idOperations: "01",
    idOperationsExenta: "01",
    
    // Contactos y productos - OBLIGATORIOS
    contacts: contacts,
    products: products,
    
    // Totales mínimos - OBLIGATORIOS
    sumTotalSubTotal: Number(totalSubTotal) || 0,
    sumTotalTaxBase: Number(totalTaxBase) || 0,
    sumTotalTax: Number(totalTax) || 0,
    sumTotalSubTotalAndTax: Number(totalSubTotal + totalTax) || 0,
    sumTotalExemptBase: 0,
    sumTotalDiscount: Number(descuentoTotal) || 0,
    sumTotalCharge: Number(cargoTotal) || 0,
    sumTotalTotal: Number(totalFactura) || 0,
    sumTotalNetPayable: Number(totalFactura) || 0,
    
    // Valores base
    valueExemptBase: 0,
    invoiceTypeId: 1,
    invoiceRectificativeTypeId: 0,
    typeRectificativeId: 0,
    
    // Datos adicionales
    aditionalData: {
      invoiceId: currentNumber,
      startInvoiceId: 1
    }
  };

  // Guardar en el listado de estado
  let fecha = ObtenerFecha();
  let numeroFactura = InvoiceGeneralInformation.InvoiceNumber;
  let nameString = cliente;
  let codigoCliente = prefactura_sheet.getRange("B3").getValue();
  
  // Guardar el nuevo JSON en lugar del antiguo
  listadoestado_sheet.appendRow([
    "vacio", "vacio", "vacio", fecha, "vacio", numeroFactura, 
    nameString, codigoCliente, "vacio", "vacio", "representacion", 
    "Vacio", JSON.stringify(fieldInvoice), ""
  ]);
  
  Logger.log("Nuevo JSON FieldInvoice generado:");
  Logger.log(JSON.stringify(fieldInvoice, null, 2));
  
  // Validación adicional del JSON
  if (!fieldInvoice.products || fieldInvoice.products.length === 0) {
    Logger.log("ADVERTENCIA: No hay productos en la factura");
  }
  if (!fieldInvoice.contacts || fieldInvoice.contacts.length === 0) {
    Logger.log("ADVERTENCIA: No hay contactos en la factura");
  }
  
  // Validar campos críticos que pueden causar el error 500
  Logger.log("=== VALIDACIÓN CRÍTICA ===");
  Logger.log("invoiceNumber: " + fieldInvoice.invoiceNumber);
  Logger.log("currentNumber: " + fieldInvoice.currentNumber);
  Logger.log("contactName: " + fieldInvoice.contactName);
  Logger.log("invoiceDate: " + fieldInvoice.invoiceDate);
  Logger.log("invoiceTime: " + fieldInvoice.invoiceTime);
  Logger.log("idPayment: " + fieldInvoice.idPayment);
  Logger.log("idOperations: " + fieldInvoice.idOperations);
  Logger.log("idOperationsExenta: " + fieldInvoice.idOperationsExenta);
  Logger.log("invoiceIdTypeRegAEAT: " + fieldInvoice.invoiceIdTypeRegAEAT);
  Logger.log("Productos count: " + fieldInvoice.products.length);
  Logger.log("Contactos count: " + fieldInvoice.contacts.length);
  
  // Validar cada producto
  fieldInvoice.products.forEach((product, index) => {
    Logger.log("Producto " + index + " - typeUse: " + product.typeUse);
    Logger.log("Producto " + index + " - reference: " + product.reference);
    Logger.log("Producto " + index + " - description: " + product.description);
    Logger.log("Producto " + index + " - quantity: " + product.quantity);
    Logger.log("Producto " + index + " - unitPrice: " + product.unitPrice);
    Logger.log("Producto " + index + " - subTotal: " + product.subTotal);
  });
  
  // Validar cada contacto
  fieldInvoice.contacts.forEach((contact, index) => {
    Logger.log("Contacto " + index + " - contactType: " + contact.contactType);
    Logger.log("Contacto " + index + " - personType: " + contact.personType);
    Logger.log("Contacto " + index + " - companyName: " + contact.companyName);
    Logger.log("Contacto " + index + " - identification: " + contact.identification);
    Logger.log("Contacto " + index + " - country: " + contact.country);
    Logger.log("Contacto " + index + " - province: " + contact.province);
    Logger.log("Contacto " + index + " - population: " + contact.population);
    Logger.log("Contacto " + index + " - addressCustomer: " + contact.addressCustomer);
  });
  
  SpreadsheetApp.getUi().alert("Factura generada y guardada satisfactoriamente con el nuevo formato API");
}

function showMensajeRespuesta(){
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

function CalcularDiasOFecha(opcion) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Factura');
  const fechaEmision = sheet.getRange("G4").getValue();
  const fechaVencimiento = sheet.getRange("G3").getValue();
  const diasVencimiento = sheet.getRange("G6").getValue();
  Logger.log(opcion+"opcion")

  // Verifica que haya fecha de emisión
  if (!fechaEmision) return;

  // Si hay días de vencimiento (incluso 0), calcula la fecha de vencimiento
  if (opcion === "Dias")  {
    const nuevaFecha = new Date(fechaEmision);
    nuevaFecha.setDate(nuevaFecha.getDate() + Number(diasVencimiento));
    sheet.getRange("G3").setValue(nuevaFecha);
  }
  
  // Si no hay días pero sí fecha de vencimiento, calcula los días
  else if (opcion === "Fecha") {
    const dias = Math.ceil((fechaVencimiento - fechaEmision) / (1000 * 60 * 60 * 24));
    sheet.getRange("G6").setValue(dias);
  }
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
function obtenerDatosFactura(factura){
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
  
  Logger.log("factura "+factura)
  Logger.log("data length "+data.length)
  Logger.log(typeof(factura))
  //Logger.log("data +"+data)
  Logger.log(wasHidden)

  

  for (var i = 1; i < data.length; i++) { // Comienza en 1 para saltar la fila de encabezado
    //Logger.log(data[i][invoiceColIndex])
    //Logger.log(typeof(data[i][invoiceColIndex]))
    Logger.log("error "+data[i][invoiceColIndex])
    if (data[i][invoiceColIndex] == factura) {
      var jsonData = data[i][jsonColIndex];
      Logger.log("jsondata "+jsonData)
      if (jsonData) {
        try {
          var invoiceData = JSON.parse(jsonData);
          let Asesor=invoiceData.Delivery
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
          let codigoPostal=invoiceData.CustomerInformation.CityCode;
          let fechaPago = invoiceData.CustomerInformation.FechaPago
          var formaPago = invoiceData.PaymentSummary.PaymentType;
          var listaProductos = invoiceData.ItemInformation;
          var numeroProductos = 0;
          var descuentosFactura = parseFloat(invoiceData.InvoiceTotal.PrePaidAmount);
          let descuentoGeneralesFactura=parseFloat(invoiceData.InvoiceTotal.GeneralPrePaidAmount);
          var cargosFactura = parseFloat(invoiceData.InvoiceTotal.ChargeTotalAmount);
          var totalFacturaJSON = parseFloat(invoiceData.InvoiceTotal.PayableAmount);
          let totalFacturaLetra=int2word(totalFacturaJSON)
          totalFacturaLetra=capitalizarPrimeraPalabra(totalFacturaLetra)
          Logger.log("totalFacturaLetra "+totalFacturaLetra)
          var valorPagar = totalFacturaLetra //arreglar
          var notaPago = invoiceData.PaymentSummary.PaymentNote;
          var observaciones = invoiceData.InvoiceGeneralInformation.Note;

          let ReqEquivalencia=parseFloat(invoiceData.InvoiceTotal.totalCargoEqui)
          let retenciones=parseFloat(invoiceData.InvoiceTotal.totalRet)
          let totalLinea=totalFacturaJSON
          

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
              targetSheet.getRange('C'+(numeroCelda+1)+':E'+(numeroCelda+1)).merge();
              targetSheet.getRange('A'+(numeroCelda+1)+':B'+(numeroCelda+1)).merge();
              filasInsertadas += 1;
              filasInsertadasPorProductos += 1;
            }
            // var celdaItem = targetSheet.getRange('A'+numeroCelda);
            // celdaItem.setBorder(true,true,true,true,null,null,null,null);
            // celdaItem.setValue(numeroProductos);
            // celdaItem.setHorizontalAlignment('center');

            var celdaReferencia = targetSheet.getRange('A'+numeroCelda);
            celdaReferencia.setBorder(true,true,true,true,null,null,null,null);
            celdaReferencia.setValue(listaProductos[j].ItemReference);
            celdaReferencia.setHorizontalAlignment('center');

            var celdaDespricion = targetSheet.getRange('C'+numeroCelda);
            celdaDespricion.setBorder(true,true,true,true,null,null,null,null);
            let nombreProducto=obtenerTextoSinNumero(listaProductos[j].Name)
            celdaDespricion.setValue(nombreProducto);
            celdaDespricion.setHorizontalAlignment('center');
            
            var celdaCantidad = targetSheet.getRange('F'+numeroCelda);
            celdaCantidad.setBorder(true,true,true,true,null,null,null,null);
            celdaCantidad.setValue(listaProductos[j].Quatity);
            celdaCantidad.setHorizontalAlignment('center');
            
            var celdaPrecioUnitario = targetSheet.getRange('G'+numeroCelda);
            celdaPrecioUnitario.setBorder(true,true,true,true,null,null,null,null);
            celdaPrecioUnitario.setValue(listaProductos[j].Price);
            celdaPrecioUnitario.setHorizontalAlignment('normal');
            celdaPrecioUnitario.setNumberFormat('€#,##0.00')

            var celdaSubtotal = targetSheet.getRange('H'+numeroCelda);
            celdaSubtotal.setBorder(true,true,true,true,null,null,null,null);
            celdaSubtotal.setValue(listaProductos[j].LineExtensionAmount);
            celdaSubtotal.setHorizontalAlignment('normal');
            celdaSubtotal.setNumberFormat('€#,##0.00')
            
            var celdaIva = targetSheet.getRange('I'+numeroCelda);
            celdaIva.setBorder(true,true,true,true,null,null,null,null);
            var percent = listaProductos[j].TaxesInformation[0].Percent;
            percent = percent.slice(0, -1);
            percent = parseFloat(percent);
            celdaIva.setValue(percent/100);
            celdaIva.setNumberFormat('0%');
            celdaIva.setHorizontalAlignment('center');

            var celdaDescuento = targetSheet.getRange('J'+numeroCelda);
            celdaDescuento.setBorder(true,true,true,true,null,null,null,null);
            celdaDescuento.setValue(parseFloat(listaProductos[j].TaxesInformation[0].Descuento));
            celdaDescuento.setNumberFormat('0.00%')
            celdaDescuento.setHorizontalAlignment('center');

            var celdaRetencion = targetSheet.getRange('K'+numeroCelda);
            celdaRetencion.setBorder(true,true,true,true,null,null,null,null);
            celdaRetencion.setValue(parseFloat(listaProductos[j].TaxesInformation[0].Retencion));
            celdaRetencion.setNumberFormat('0%')
            celdaRetencion.setHorizontalAlignment('center');

            var celdaRecargoEquivalencia = targetSheet.getRange('L'+numeroCelda);
            celdaRecargoEquivalencia.setBorder(true,true,true,true,null,null,null,null);
            celdaRecargoEquivalencia.setValue(parseFloat(listaProductos[j].TaxesInformation[0].RecgEquivalencia));
            celdaRecargoEquivalencia.setNumberFormat('0.00%')
            celdaRecargoEquivalencia.setHorizontalAlignment('center');

            
            var celdaTotalLinea = targetSheet.getRange('M'+numeroCelda);
            celdaTotalLinea.setBorder(true,true,true,true,null,null,null,null);
            //subtotal+(subtotal*iva)+(subtotal*recargo)-(subtotal*retencion)
            Logger.log("LineTotal "+listaProductos[j].LineTotal)
            celdaTotalLinea.setValue(listaProductos[j].LineTotal);
            celdaTotalLinea.setNumberFormat('€#,##0.00');
            celdaTotalLinea.setHorizontalAlignment('normal');
            

            var producto = listaProductos[j]
            //crea un diccionario que la llave sea el % de iva y el valor sea el total de la linea
            Logger.log(grupoIva+"before")
            if (grupoIva.hasOwnProperty(percent)) {
              grupoIva[percent] += producto.TaxesInformation[0].TaxableAmount;
            } else {
              grupoIva[percent] = producto.TaxesInformation[0].TaxableAmount;
            }
            Logger.log("grupoIva after"+grupoIva)
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
                targetSheet.getRange('A'+(numeroCelda+1)+':D'+(numeroCelda+1)).merge();
                targetSheet.getRange('F'+(numeroCelda+1)+':H'+(numeroCelda+1)).merge();
                targetSheet.getRange('I'+(numeroCelda+1)+':M'+(numeroCelda+1)).merge();
                filasInsertadas += 1;
                auxiliarFilasInsertadas += 1;
              } else {
                auxiliarFilasInsertadas += 1;
              }
              Logger.log("auxiliarfilasinseretadas after: "+auxiliarFilasInsertadas)
              Logger.log("pasando el segundo if")
              var celdaBaseImponible = targetSheet.getRange('A'+numeroCelda);
              celdaBaseImponible.setBorder(true,true,true,true,null,null,null,null);
              celdaBaseImponible.setValue(grupoIva[key]);
              celdaBaseImponible.setNumberFormat('€#,##0.00');
              celdaBaseImponible.setHorizontalAlignment('normal');
              
              var celdaPorcentajeIva = targetSheet.getRange('E'+numeroCelda);
              celdaPorcentajeIva.setBorder(true,true,true,true,null,null,null,null);
              celdaPorcentajeIva.setValue(key/100);
              celdaPorcentajeIva.setNumberFormat('0%');
              celdaPorcentajeIva.setHorizontalAlignment('center');
              
              var celdaIVA = targetSheet.getRange('F'+numeroCelda);
              celdaIVA.setBorder(true,true,true,true,null,null,null,null);
              celdaIVA.setFormula('=A'+numeroCelda+'*E'+numeroCelda);
              celdaIVA.setNumberFormat('€#,##0.00');
              celdaIVA.setHorizontalAlignment('normal');
              
              var celdaTotal = targetSheet.getRange('I'+numeroCelda);
              celdaTotal.setBorder(true,true,true,true,null,null,null,null);
              celdaTotal.setFormula('=A'+numeroCelda+'+F'+numeroCelda);
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
          let fechaPagoCeldaHoja = hojaCeldas.getRange('E8').getValue();
          var formaPagoCeldaHoja = hojaCeldas.getRange('E10').getValue();
          let contactoCeldaHoja=hojaCeldas.getRange("E11").getValue();

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
          let contactoCell=targetSheet.getRange(contactoCeldaHoja);
          let fechaPagoCell=targetSheet.getRange(fechaPagoCeldaHoja);
          var valorPagarCell = targetSheet.getRange('B'+(41+filasInsertadas));
          var notaPagoCell = targetSheet.getRange('A'+(45+filasInsertadas));
          var observacionesCell = targetSheet.getRange('A'+(50+filasInsertadas));
          var totalItemsCell = targetSheet.getRange('B'+(21+filasInsertadasPorProductos));
          var descuentosCell = targetSheet.getRange('A'+(24+filasInsertadasPorProductos));
          var cargosCell = targetSheet.getRange('D'+(24+filasInsertadasPorProductos));
          var sumaBaseImponible = targetSheet.getRange('A'+(32+filasInsertadas));
          var sumaImpIva = targetSheet.getRange('F'+(32+filasInsertadas));
          var sumaTotal = targetSheet.getRange('I'+(32+filasInsertadas));

          var totalRetenciones = targetSheet.getRange('A'+(36+filasInsertadas));
          var totalCrgEquivalencia = targetSheet.getRange('D'+(36+filasInsertadas));
          var totalCargos = targetSheet.getRange('G'+(36+filasInsertadas));
          var totalDescuentos = targetSheet.getRange('K'+(36+filasInsertadas));

          var totalDeFactura = targetSheet.getRange('H'+(38+filasInsertadas));


          const resultado = dividirString(cliente)
          celdaNumFactura.setValue("FACTURA DE VENTA NO. "+facturaNumero);
          clienteCell.setValue(resultado[0]);
          nifCell.setValue(nif);
          // codigoCell.setValue(codigo);
          direccionCell.setValue(direccion+" "+codigoPostal);
          telefonoCell.setValue(telefono);
          contactoCell.setValue(Asesor)
          // Ajustar la forma en que se ve el pais - IMPORTANTE
          if (poblacion == "" || provincia == "" || pais == "") {
            var columnaPoblacion = poblacionCell.getColumn();
            var filaPoblacion = poblacionCell.getRow();
            targetSheet.getRange(filaPoblacion, columnaPoblacion-1).clearContent();
          } else {
            poblacionCell.setValue(poblacion+', '+provincia+', '+pais);
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
          fechaPagoCell.setValue(fechaPago);
          formaPagoCell.setValue(formaPago);
          valorPagarCell.setValue(valorPagar);
          notaPagoCell.setValue(notaPago);
          observacionesCell.setValue(observaciones);
          // totalItemsCell.setValue(numeroProductos);
          Logger.log("descuentoGeneralesFactura: "+descuentoGeneralesFactura)
          descuentosCell.setValue(descuentoGeneralesFactura);
          cargosCell.setValue(cargosFactura);
          sumaBaseImponible.setFormula('=SUM(A'+(30+numeroProductos-1)+':A'+(31+filasInsertadas-1)+')');
          sumaImpIva.setFormula('=SUM(F'+(30+numeroProductos-1)+':F'+(31+filasInsertadas-1)+')');
          sumaTotal.setFormula('=SUM(I'+(30+numeroProductos-1)+':I'+(31+filasInsertadas-1)+')');
          totalRetenciones.setValue(retenciones);
          totalCrgEquivalencia.setValue(ReqEquivalencia);
          totalCargos.setValue(cargosFactura);
          Logger.log("descuentosFactura: "+descuentosFactura)
          totalDescuentos.setValue(descuentosFactura);
  
          totalDeFactura.setValue(totalLinea);




          
          
          var itemCellPrueba = targetSheet.getRange('A19')
          var hojaEnBlanco = clienteCell.isBlank() || formaPagoCell.isBlank() || celdaBaseImponible.isBlank();
          while (hojaEnBlanco) {
            Utilities.sleep(2000);
            Logger.log("dentro de while")
            hojaEnBlanco = clienteCell.isBlank() || formaPagoCell.isBlank() || itemCellPrueba.isBlank() || celdaBaseImponible.isBlank();
          }

          if (!hojaEnBlanco){
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

function obtenerTextoSinNumero(str) {
  // Elimina cualquier espacio alrededor del guion y separa el texto del número
  const partes = str.split('-');
  // Retorna solo la parte de texto
  return partes[0].trim();
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

function limpiarTablas(columna, linea){
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Copia de Plantilla');
  var primeraFila = targetSheet.getRange(linea+":"+linea);
  primeraFila.clearContent();
  linea++;
  while (!targetSheet.getRange(columna+linea).isBlank()) {
    targetSheet.deleteRow(linea);
  }
}

function sacarColumnaFila(celda){
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

function pruebaSacar(){
  var lista = sacarColumnaFila("E18")
  Logger.log(lista)
}

function subirFactura(nombre, pdfBlob) {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  var folderId = hojaDatosEmisor.getRange("B15").getValue();
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
  hojaDatosEmisor.getRange("B15").setValue(id);
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
  hojaDatosEmisor.getRange("B15").setValue(id);
  Logger.log("Carpeta creada")
}

function eliminarCarpetaConDriveAPI() {
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let idCarpeta = hojaDatosEmisor.getRange("B15").getValue();  // Obtenemos el ID de la carpeta desde una celda

  try {
    Drive.Files.remove(idCarpeta);  // Elimina la carpeta usando el servicio avanzado de Drive
    Logger.log("Carpeta eliminada exitosamente.");
    hojaDatosEmisor.getRange("B15").setValue("");
  } catch (e) {
    Logger.log("Error al eliminar la carpeta: " + e.message);
    hojaDatosEmisor.getRange("B15").setValue("");
  }
}


function subirFactura2(nombre, pdfBlob) {
  Logger.log("subir factura 2")
  let hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos de emisor');
  let IdCarpeta=hoja.getRange("B15").getValue()
  let fileMetadata = {
    'name': 'Factura ' + nombre + '.pdf',
    'mimeType': 'application/pdf',
    'parents': [IdCarpeta]  // Opcional: si quieres especificar una carpeta
  };
  let file = Drive.Files.create(fileMetadata, pdfBlob);


  Logger.log('PDF creado y subido: ' + file.id);
  return file.id; 
}


function filtroHistorialFacturas(tipoFiltro){
  
  Logger.log("debtro de fitrlo historial")
  Logger.log("tipoFiltro "+tipoFiltro)
  let Formula=''
  let hojahistorial=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Historial Facturas');
  if(tipoFiltro=="Numero factura"){
    Formula="=FILTER('Historial Facturas Data'!A2:E1000;ISNUMBER(SEARCH(C5;'Historial Facturas Data'!A2:A1000)))"
  }else if(tipoFiltro=="NIF"){
    Formula="=FILTER('Historial Facturas Data'!A2:E1000;ISNUMBER(SEARCH(C5;'Historial Facturas Data'!C2:C1000)))"
  }else if(tipoFiltro=="Cliente"){
    Formula="=FILTER('Historial Facturas Data'!A2:E1000;ISNUMBER(SEARCH(C5;'Historial Facturas Data'!B2:B1000)))"
  }else if(tipoFiltro=="Estado"){
    Formula="=FILTER('Historial Facturas Data'!A2:E1000;ISNUMBER(SEARCH(C5;'Historial Facturas Data'!E2:E1000)))"
  }

  hojahistorial.getRange("B8").setValue(Formula)
  

}