PREFACTURA_ROW = 3;
PREFACTURA_COLUMN = 2;
COL_TOTALES_PREFACTURA = 11;// K
FILA_INICIAL_PREFACTURA = 8;
COLUMNA_FINAL = 50;
ADDITIONAL_ROWS = 3 + 3; //(Personalizacion)

// Bandera para mostrar/ocultar el resumen de la factura al finalizar
const SHOW_SUMMARY_MODAL = false;


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
  // 5% ya no aplica; mantenemos 4%
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

  // El asesor comercial ya no es obligatorio



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
      return;
    }
    if (estadoFactura.success) {
      // Aviso al usuario en la interfaz de Google Sheets (tipo popup de la captura)
      SpreadsheetApp.getUi().alert("Tu factura se está generando, espera un momento por favor.");
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
      mensaje = "Error en los datos del cliente: Tipo de persona inválido. Debe ser 'Autónomo' o 'Empresa'. Verifica la columna 'Tipo de persona' en la hoja Clientes.";
    }
    SpreadsheetApp.getUi().alert("No se pudo guardar/enviar la factura. " + mensaje);
    Logger.log("guardarFactura error: " + mensaje);
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
    
    // J es Retención y I es Recargo según nuevo orden visual
    factura_sheet.getRange("I"+String(rowParaDatos)).setValue(dictInformacionProducto["Recargo de equivalencia"])//Tarifa recargo
    factura_sheet.getRange("J"+String(rowParaDatos)).setValue(dictInformacionProducto["retencion"])//Tarifa retención
    // Total de línea: Base gravable (F) + IVA + recargo - retenciones
    factura_sheet.getRange("K"+String(rowParaDatos)).setValue("=IF(F"+String(rowParaDatos)+"=\"\";0;F"+String(rowParaDatos)+"*(1+G"+String(rowParaDatos)+"+I"+String(rowParaDatos)+"-J"+String(rowParaDatos)+"))")
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

  // var html = HtmlService.createHtmlOutputFromFile('postFactura')
  //   .setTitle('Menú');
  // SpreadsheetApp.getUi()
  //   .showSidebar(html);
  
  // Mostrar el resumen sólo si la bandera lo permite
  if (SHOW_SUMMARY_MODAL) {
    showCustomDialog()
  }
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
    "User": usuario,
    "Password": contra
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
  Logger.log("opciones"+opciones["payload"])
  Logger.log("usuario "+usuario)
  Logger.log("contra "+contra)

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
    Logger.log("Error al vincular cuenta: " + error.message);
    hojaDatosEmisor.getRange("B16").setBackground('#FFC7C7')
    hojaDatosEmisor.getRange("B16").setValue("Desvinculado")
    hojaDatos.getRange("I21").setValue(0)
    // Intentar extraer mensaje de la API si viene en JSON
    let apiMsg = '';
    try {
      const errObj = JSON.parse((error.message || '').replace('Error de la API: ',''));
      apiMsg = (errObj && (errObj.exception || errObj.message)) ? String(errObj.exception || errObj.message) : '';
    } catch (e) {}
    const mensajeFinal = apiMsg ? apiMsg : 'Error al vincular tu cuenta. Verifica usuario y contraseña e inténtalo de nuevo.';
    SpreadsheetApp.getUi().alert(mensajeFinal);
  }
}



function obtenerPDFFacturaBase64(numeroFactura) {
  let spreadsheet = SpreadsheetApp.getActive();
  const scriptProps = PropertiesService.getDocumentProperties();
  let ambiente = scriptProps.getProperty('Ambiente');

  // Endpoint para PDFInvoice (puede devolver JSON con base64 ahora)
  let Burl;
  if (ambiente == "Pruebas") {
    Burl = "https://facturasapp-qa.cenet.ws/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/PDFInvoice";
  } else {
    Burl = "https://www.facturasapp.com/ApiGateway/ApiExternal/Invoice/api/InvoiceServices/PDFInvoice";
  }

  let params = {
    invoiceNumber: String(numeroFactura)
  };

  let url = buildUrlWithParams(Burl, params);

  let hojaDatos = spreadsheet.getSheetByName('Datos');
  let APIkey = hojaDatos.getRange("I21").getValue();

  if (!APIkey) {
    Logger.log("Error: No se encontró la API Key");
    return null;
  }

  let opciones = {
    "method": "post",
    "headers": { "X-API-KEY": APIkey },
    'muteHttpExceptions': true
  };

  try {
    var respuesta = UrlFetchApp.fetch(url, opciones);
    let responseCode = respuesta.getResponseCode();

    if (responseCode !== 200) {
      let responseText = respuesta.getContentText();
      Logger.log("Error al obtener el PDF: " + responseText);
      return null;
    }

    // Intentar detectar JSON y extraer 'toolObject'
    let contentText = '';
    try { contentText = respuesta.getContentText(); } catch (e) { contentText = ''; }

    if (contentText && contentText.trim().startsWith('{')) {
      try {
        let json = JSON.parse(contentText);
        // Estructura esperada: { id, messages, isError, toolObject }
        if (json && json.isError === false && json.toolObject) {
          return String(json.toolObject);
        }
        // Si viene error o falta toolObject, registrar mensaje
        Logger.log("Respuesta JSON sin toolObject o con error: " + contentText);
        return null;
      } catch (parseErr) {
        Logger.log("No fue posible parsear JSON de respuesta. Se intentará como binario. Error: " + parseErr);
      }
    }

    // Compatibilidad: si no es JSON, asumir binario PDF y convertir a base64
    let pdfBlob = respuesta.getBlob();
    let base64String = Utilities.base64Encode(pdfBlob.getBytes());
    return base64String;

  } catch (error) {
    Logger.log("Error al obtener el PDF: " + error.message);
    return null;
  }
}

function buildUrlWithParams(baseUrl, params) {
  const query = Object.entries(params)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join('&');
  return `${baseUrl}?${query}`;
}

function linkDescargaFactura() {
  let hojaID = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Facturas ID');
  var lastRow = hojaID.getLastRow();
  // var idArchivo = hoja.getRange("B" + lastRow).getValue();
  var numFactura = hojaID.getRange("A" + lastRow).getValue();

  // if (!idArchivo) {
  //   throw new Error("El ID del archivo está vacío o no es válido.");
  // }

  // // Verificar el archivo y asignar permisos públicos usando Advanced Drive Service
  // var permisos = {
  //   role: "reader",
  //   type: "anyone"
  // };

  // try {
  //   Drive.Permissions.create(permisos, idArchivo, {sendNotificationEmails: false});
  // } catch (e) {
  //   throw new Error("Error al configurar permisos públicos: " + e.message);
  // }

  // Generar la URL de descarga
  // var url = "https://drive.google.com/uc?export=download&id=" + idArchivo;
  
  return {
    numFactura: numFactura,
    // url: url
  };
}



function getDownloadLink() {
  var data = linkDescargaFactura();
  // const usuario =obtenerUsuario()
  // const propietario= obtenerPropietario()
  // Logger.log("usuario "+usuario)
  // Logger.log("propietario "+propietario)
  Logger.log("sale de linkdescargar")
  Logger.log("dataa"+data)
  Logger.log("dataa"+data.numFactura)
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
  
  // Parsear el JSON para obtener el total.
  // Usamos sumTotalNetPayable (Neto a pagar) para que coincida con el valor de la factura,
  // y si no existe, caemos a sumTotalTotal.
  let fieldInvoiceData = JSON.parse(fieldInvoiceJson);
  let totalNumber = Number(
    fieldInvoiceData.sumTotalNetPayable != null
      ? fieldInvoiceData.sumTotalNetPayable
      : fieldInvoiceData.sumTotalTotal || 0
  );
  // Formatear con estilo español (ej: 553,32)
  invoiceTotal = totalNumber.toLocaleString('es-ES', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  
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
  Logger.log(data)
  Logger.log("out loop")
  for (let key in data) {
    Logger.log("in loop")
    Logger.log(key + ': ' + data[key]);
  }

  var numFactura = data.numFactura;
  
  // Verificar que la factura existe en el historial
  var hojaListadoEstado = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListadoEstado');
  var dataRange = hojaListadoEstado.getDataRange();
  var dataValues = dataRange.getValues();
  var invoiceColIndex = 5; // Columna F (número de factura)
  
  var facturaEncontrada = false;
  for (let i = 1; i < dataValues.length; i++) {
    if (dataValues[i][invoiceColIndex] == numFactura) {
      facturaEncontrada = true;
      break;
    }
  }
  
  if (!facturaEncontrada) {
    return 'Factura no encontrada en el historial';
  }

  try {
    Logger.log("Descargando factura " + numFactura + " usando nuevo API PDFInvoice");
    
    // Usar la función directa que obtiene el PDF como base64
    var downloadUrl = descargarPDFDirecto(numFactura);
    
    if (!downloadUrl) {
      return 'Error al obtener el PDF desde FacturasApp';
    }

   
    
    return downloadUrl;
  } catch (e) {
    Logger.log("Error al descargar factura: " + e.message);
    return 'Error al obtener el archivo: ' + e.message;
  }
}

// Función alternativa para descargar PDF directamente desde FacturasApp
function descargarPDFDirecto(numeroFactura) {
  try {
    Logger.log("Descargando PDF directamente para factura: " + numeroFactura);
    
    // Obtener PDF como base64 desde FacturasApp
    var base64PDF = obtenerPDFFacturaBase64(numeroFactura);
    
    if (!base64PDF) {
      SpreadsheetApp.getUi().alert("Error al obtener el PDF desde FacturasApp");
      return null;
    }
    
    // Convertir base64 a blob
    var pdfBytes = Utilities.base64Decode(base64PDF);
    var pdfBlob = Utilities.newBlob(pdfBytes, 'application/pdf', 'Factura_' + numeroFactura + '.pdf');
    
    // Crear data URL directamente sin subir a Google Drive
    var downloadUrl = generarPdfUrl(pdfBlob);
    
    Logger.log("PDF preparado correctamente para descarga directa");
    
    return downloadUrl;
    
  } catch (error) {
    Logger.log("Error al descargar PDF: " + error.message);
    SpreadsheetApp.getUi().alert("Error al descargar PDF: " + error.message);
    return null;
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

  Logger.log("Generando PDF para factura: " + numeroFactura);

  try {
    var downloadUrl = descargarPDFDirecto(numeroFactura);

    if (!downloadUrl) {
      SpreadsheetApp.getUi().alert("Error al obtener el PDF desde FacturasApp para la factura " + numeroFactura);
      return;
    }

    Logger.log("PDF preparado correctamente desde FacturasApp");

    var html = '<html><body><a href="' + downloadUrl + '">Descargar PDF de la Factura ' + numeroFactura + '</a></body></html>';
    var ui = HtmlService.createHtmlOutput(html)
      .setWidth(300)
      .setHeight(100);
    SpreadsheetApp.getUi().showModalDialog(ui, 'Descargar PDF');
    
  } catch (error) {
    Logger.log("Error al generar PDF: " + error.message);
    SpreadsheetApp.getUi().alert("Error al generar PDF: " + error.message);
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
      const nuevoPrefijo = scriptProperties.getProperty('ConsecutivoPlantillaPrefijo');
      const nuevoDigitos = scriptProperties.getProperty('ConsecutivoPlantillaDigitos');
      let prefijo = nuevoPrefijo || scriptProperties.getProperty('LetraConescutivo') || "";
      let numeroPlantilla = scriptProperties.getProperty('NumeroConescutivo') || "0";
      let digitos = nuevoDigitos ? Number(nuevoDigitos) : String(numeroPlantilla).length;
      numeroActual = parseInt(String(numeroPlantilla),10);
      if(!isFinite(numeroActual)){
        numeroActual = 0;
      }
      // Construimos un original coherente para mantener el padding
      ultimoConsecutivo = String(prefijo) + String(numeroActual).padStart(digitos,'0');
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
  // Aceptar cualquier prefijo (incluye letras, números y símbolos) y separar el sufijo numérico final
  const match = String(original).match(/^(.*?)(\d+)$/);
  if (!match) {
    return String(nuevoNumero);
  }

  const prefijo = match[1];
  const parteNumerica = match[2];

  const nuevoNumeroStr = String(nuevoNumero).padStart(parteNumerica.length, '0');
  return prefijo + nuevoNumeroStr;
}

function obtenerFechaYHoraActual(){ 
  let spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName('Factura');
  let zonaHorariaEspaña = "Europe/Madrid"
  let hoy = new Date();
  let fechaHoy = Utilities.formatDate(hoy, zonaHorariaEspaña, "dd/MM/yyyy");
  let hora = Utilities.formatDate(hoy, zonaHorariaEspaña, "HH:mm:ss");

  // Establecer fecha de emisión = hoy
  sheet.getRange("G4").setNumberFormat("dd/MM/yyyy");
  sheet.getRange("G4").setValue(String(fechaHoy))

  // Calcular fecha de pago sumando los días de vencimiento actuales (G6)
  let diasVencimiento = Number(sheet.getRange("G6").getValue() || 0);
  let fechaPagoDate = new Date(hoy);
  if (!isNaN(diasVencimiento) && diasVencimiento >= 0) {
    fechaPagoDate.setDate(fechaPagoDate.getDate() + diasVencimiento);
  }
  let fechaPago = Utilities.formatDate(fechaPagoDate, zonaHorariaEspaña, "dd/MM/yyyy");
  sheet.getRange("G3").setNumberFormat("dd/MM/yyyy");
  sheet.getRange("G3").setValue(String(fechaPago))

  // Hora de emisión
  sheet.getRange("G7").setValue(hora)

  let valorFecha = sheet.getRange("G4").getValue();
  let fechaFormateada = Utilities.formatDate(new Date(valorFecha), zonaHorariaEspaña, "dd/MM/yyyy");
  Logger.log("valorFecha "+valorFecha)
  Logger.log("fecha "+fechaHoy)
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

// Mapea el medio de pago textual (E4) al código idPayment requerido por RG
function mapIdPaymentCode(medioPagoTxt){
  Logger.log("medioPagoTxt"+medioPagoTxt)
  if(!medioPagoTxt) return "ND"; // No definido
  // const normalizado = String(medioPagoTxt).toLowerCase().trim();
  switch(medioPagoTxt){
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

function validarValorWithHolding(valor, tipoWithHolding) {
  // Validar que el valor esté en la lista de valores permitidos
  let valoresPermitidos = [];
  
  if (tipoWithHolding === 10) { // Retención
    valoresPermitidos = [7, 15, 19];
  } else if (tipoWithHolding === 11) { // Recargo de equivalencia
    valoresPermitidos = [5.2, 1.4, 0.5, 1.75];
  }
  
  let valorRedondeado = Math.round(Number(valor) * 10) / 10;
  return valoresPermitidos.includes(valorRedondeado);
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
  let sumRecargoAmount = 0;        // Suma de Recargo Equivalencia
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
  
  // IRPF general (totales): si está configurado y no hay retenciones por producto
  const rowIrpf = startingRowTaxation - 2; // Fila donde está el IRPF en totales
  const irpfSelDisplay = String(prefactura_sheet.getRange(rowIrpf, 6).getDisplayValue() || '').trim().toLowerCase(); // Columna F
  let generalIrpfRate = 0; // fracción (0.07, 0.15, 0.19 o libre)
  if (irpfSelDisplay) {
    if (irpfSelDisplay === 'valor libre') {
      const libreVal = Number(prefactura_sheet.getRange(rowIrpf, 8).getValue() || 0); // Columna H
      generalIrpfRate = Number(libreVal) || 0;
      if (generalIrpfRate < 0) generalIrpfRate = 0;
    } else {
      const rateNum = parsePercentToNumberES(prefactura_sheet.getRange(rowIrpf, 6).getDisplayValue());
      generalIrpfRate = rateNum ? Number(rateNum) / 100 : 0;
    }
  }
  let hasPerProductRetention = false;
  const productBaseNetList = [];
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
    let subtotalHoja = Number(productoData[5]) || 0; // con descuento
    let ivaRate = Number(productoData[6]) || 0;
    let descuentoRate = Number(productoData[7]) || 0;
    // En la hoja: I = Tarifa recargo, J = Tarifa retención. Respetar ese orden.
    let recargoEquivalenciaRate = Number(productoData[8]) || 0; // Tarifa recargo
    let retencionRate = Number(productoData[9]) || 0; // Tarifa retención
    let totalLinea = Number(productoData[10]) || 0;

    // Calcular valores base
    let baseBruta = round2(precioUnitario * cantidad); // antes de descuento
    let discountAmount = round2(baseBruta * descuentoRate);
    let baseNeta = round2(baseBruta - discountAmount); // después de descuento
    
    // Calcular valores (igual que en la hoja): sobre la base neta
    let taxAmount = round2(baseNeta * ivaRate);
    let withHoldingsAmount = round2(baseNeta * retencionRate);
    let surChargesAmount = round2(baseNeta * recargoEquivalenciaRate);
    
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
    if (retencionRate > 0) {
      let codigoRetencion = obtenerIdRateWithHoldings(retencionRate * 100, 10);
      Logger.log("Producto: " + descripcion + " - Retención: " + (retencionRate * 100) + "% - Código: " + codigoRetencion);
      withHoldingsSurChargesDto.push({
        idRateWithHoldings: codigoRetencion, // Retención
        subTotalWithHoldings: baseNeta,
        cuotaWithHoldings: withHoldingsAmount
      });
      hasPerProductRetention = true;
    }
    
    // Agregar recargo de equivalencia si existe
    if (recargoEquivalenciaRate > 0) {
      let codigoRecargo = obtenerIdRateWithHoldings(recargoEquivalenciaRate * 100, 11);
      Logger.log("Producto: " + descripcion + " - Recargo: " + (recargoEquivalenciaRate * 100) + "% - Código: " + codigoRecargo);
      withHoldingsSurChargesDto.push({
        idRateWithHoldings: codigoRecargo, // Recargo de equivalencia
        subTotalWithHoldings: baseNeta,
        cuotaWithHoldings: surChargesAmount
      });

      // Agrupar recargo de equivalencia para fieldTaxations
      let recargoRateKey = recargoEquivalenciaRate * 100;
      if (!recargoTaxGroups[recargoRateKey]) {
        recargoTaxGroups[recargoRateKey] = {
          taxName: "RecargoEquivalencia",
          rate: recargoRateKey,
          taxBase: 0,
          valueTax: 0
        };
      }
      recargoTaxGroups[recargoRateKey].taxBase = round2(recargoTaxGroups[recargoRateKey].taxBase + baseNeta);
      recargoTaxGroups[recargoRateKey].valueTax = round2(recargoTaxGroups[recargoRateKey].valueTax + surChargesAmount);
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
    // Asegurar cantidad entera y mínima de 1 para cumplir con el esquema (int32)
    const quantityInt = Math.max(1, Math.trunc(Number(cantidad)));
    let producto = {
      typeUse: "VEN",
      reference: String(referencia).substring(0, 50),
      description: String(descripcion).substring(0, 100),
      unitPrice: Number(precioUnitario),
      quantity: quantityInt,
      // IMPORTANTE: Enviar subTotal BRUTO (antes de descuento) para que el servicio
      // aplique los módulos de descuento y no descuente doble.
      subTotal: baseBruta,
      // totalTax solo debe reflejar IVA. El recargo se reporta aparte.
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
    // Guardar base neta para posible IRPF general
    productBaseNetList.push(baseNeta);
  }
  // totalTax a nivel factura debe incluir únicamente IVA (no recargo)
  totalTax = round2(sumIvaAmount);
  
  // Aplicar IRPF general si NO hubo retenciones por producto.
  // En lugar de crear retenciones por producto, solo acumulamos el total
  // para reportarlo en sumTotalRetentionIRPF a nivel de factura.
  if (generalIrpfRate > 0 && !hasPerProductRetention && products.length === productBaseNetList.length) {
    for (let idx = 0; idx < products.length; idx++) {
      const baseNeta = productBaseNetList[idx];
      const cuota = round2(baseNeta * generalIrpfRate);
      if (cuota <= 0) continue;
      totalWithHoldings = round2(totalWithHoldings + cuota);
    }
  }
  
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
  let netoPagar = 0;
  
  if (prefactura_sheet.getRange("A31").getValue() === "Total factura") {
    totalFactura = prefactura_sheet.getRange("B31").getValue();
    cargoTotal = prefactura_sheet.getRange("B17").getValue() || 0;
    netoPagar = prefactura_sheet.getRange("B32").getValue()
  } else {
    let rowTotalFactura = startingRowTaxation + 12;
    let rowCargoFactura = startingRowTaxation - 2;
    let rowNetoPagar = startingRowTaxation +13;
    totalFactura = prefactura_sheet.getRange(rowTotalFactura, 2).getValue();
    cargoTotal = prefactura_sheet.getRange("B" + String(rowCargoFactura)).getValue() || 0;
    netoPagar = prefactura_sheet.getRange("B"+String(rowNetoPagar)).getValue()
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
  // Crear contactos con estructura completa
  let contacts = [{
    contactType: CustomerInformation.IdentificationType,
    personType: CustomerInformation.TypePerson, 
    companyName: String(nombreClienteLimpio).substring(0, 450),
    customerCode: String(codigoCliente).substring(0,20),
    identificationType:CustomerInformation.DocumentIdentificationType,
    identification: String(CustomerInformation.Identification || "12345678A").substring(0, 20),
    tradeName: String(cliente).substring(0, 450),
    regime: CustomerInformation.Regimen, // Según factura.json
    // Códigos oficiales de país / provincia / población tomados del catálogo externo
    country: String(CustomerInformation.CountryCode || "").substring(0, 10),
    province: String(CustomerInformation.ProvinceCode || "").substring(0, 10),
    population: String(CustomerInformation.PopulationCode || "").substring(0, 10),
    addressCustomer: String(CustomerInformation.AddressLine || null).substring(0, 200), //AddressLine
    postalCodeCustomer: String(CustomerInformation.CityCode || null).substring(0, 10), //CityCode
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
  
  // Crear chargeAndDiscount: incluir solo valores reales, sin cargos por defecto
  let chargeAndDiscount = [];
  if (cargoTotal > 0) {
    chargeAndDiscount.push({
      idtypeFeeDiscount: "CG",
      idTypeValueFeeDiscount: "PJ",
      baseFeeDiscount: totalTaxBase,
      valueFeeDiscount: 1,
      totalFeeDiscount: cargoTotal
    });
  } else {
    // Mantener estructura compatible con valores en cero
    chargeAndDiscount.push({
      idtypeFeeDiscount: "CG",
      idTypeValueFeeDiscount: "PJ",
      baseFeeDiscount: 0,
      valueFeeDiscount: 0,
      totalFeeDiscount: 0
    });
  }
  
  // Calcular totales finales
  // Para evitar doble conteo en el portal:
  // - sumTotalSubTotalAndTax: SubTotal + IVA (sin recargo)
  // - sumTotalTotal: igual que sumTotalSubTotalAndTax
  // - sumTotalNetPayable: (SubTotal + IVA) + Recargo - Retenciones
  let sumTotalSubTotalAndTax = round2(totalSubTotal + totalTax);
  let sumTotalTotalCalc = sumTotalSubTotalAndTax;
  let sumTotalNetPayable = round2(sumTotalSubTotalAndTax + totalSurCharges - totalWithHoldings);
  
  // Crear el JSON con estructura EXACTA de factura.json
  // Fechas coherentes con hoja: invoiceDate = G4, invoiceExpiration = días de G6
  let diasExpiracion = Number(prefactura_sheet.getRange("G6").getValue() || 0);
  if (isNaN(diasExpiracion) || diasExpiracion < 0) diasExpiracion = 0;
  // idPayment dinámico según medio de pago (E4)
  const medioPagoTxt = String(prefactura_sheet.getRange("G5").getValue() || "");
  const idPaymentCode = mapIdPaymentCode(medioPagoTxt);
  // Base neta (exenta) para facturas sin impuestos
  const baseNetaTotal = round2(totalSubTotal - totalDiscounts);
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
    idOperations: "S1", // Según factura.json
    // Si hay impuestos (IVA o recargo) no es exenta: usar E0. Si no hay impuestos, E3
    idOperationsExenta: "E0",
    valueExemptBase: (hasAnyTaxOrSurcharge) ? 0 : baseNetaTotal,
    chargeAndDiscount: chargeAndDiscount, // Siempre incluir - nunca null
    fieldTaxations: fieldTaxations.length > 0 ? fieldTaxations : [],
    sumTotalSubTotal: totalSubTotal,
    sumTotalTaxBase: totalTaxBase,
    sumTotalTax: totalTax,
    sumTotalSubTotalAndTax: sumTotalSubTotalAndTax,
    sumTotalExemptBase: 0,
    sumTotalDiscount: totalDiscounts,
    sumTotalCharge: cargoTotal,
    // Nuevo campo: total de retenciones IRPF (por producto + general)
    sumTotalRetentionIRPF: totalWithHoldings,
    // Total de la factura (sin recargo; el portal suma el recargo por separado)
    sumTotalTotal: sumTotalTotalCalc,
    // Neto a pagar: (SubTotal + IVA) + Recargo - Retenciones
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
  
  // Guardar JSON en propiedades del documento para mostrar resumen inmediatamente
  try {
    PropertiesService.getDocumentProperties().setProperty('lastInvoiceJson', JSON.stringify(fieldInvoice));
  } catch (e) {
    Logger.log('No se pudo guardar lastInvoiceJson: ' + e);
  }
  // Mostrar resumen de la factura sólo si está habilitado
  if (SHOW_SUMMARY_MODAL) {
    mostrarResumenFactura();
  }
}

function showMensajeRespuesta(){
  SpreadsheetApp.getUi().alert("Vinculacion exitosa");
}

function calcularPorcentaje(valor, total) {
  return (valor / total) * 100;
}

function recuperarJson(){
  const jsonStr = PropertiesService.getDocumentProperties().getProperty('lastInvoiceJson');
  return jsonStr || '{}';
}

function formatearPesos(valor){
  try {
    return new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR' }).format(Number(valor)||0);
  } catch(e){
    return String(valor);
  }
}

function plantillaResumenFactura(nombreCliente, numeroFactura, impuestos, invoiceTotal) {
  return `
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Roboto:ital,wght@0,100..900;1,100..900&display=swap');
      body { font-family: 'Roboto', sans-serif; font-size: 16px; margin: 0; padding: 0; }
      .container { padding: 10px; display: flex; flex-direction: column; height: auto; }
      .title { font-size: 18px; font-weight: bold; margin-bottom: 10px; }
      .info { margin-bottom: 10px; }
      .info span { font-weight: bold; }
      .columns { display: flex; justify-content: space-between; gap: 20px; }
      .column { flex: 1; padding: 10px; margin-right: 10px; box-sizing: border-box; }
      .column:last-child { margin-right: 0; }
      .button-container { display: none; }
      .red-text { color: rgb(231, 112, 14); font-size: 16px; font-family: 'Roboto', sans-serif; font-weight: 600; }
      .grey-text { color: grey; font-size: 16px; font-family: 'Roboto', sans-serif; font-weight: 600; }
      button { padding: 3px 12px; font-family: 'Roboto', sans-serif; background-color: rgba(255, 255, 255, 0); border: none; border-radius: 30px; cursor: pointer; }
      button:hover { background-color:rgba(255, 218, 187, 0.32); }
      ul { padding: 0; list-style-type: none; }
      li { display: flex; justify-content: space-between; padding: 5px 0; }
      .btn-orange { display:none; }
      .btn-grey-outline { display:none; }
    </style>
    <div class="container">
      <div class="columns">
        <div class="column">
          <div class="info"><span>Nombre del Cliente:</span> ${nombreCliente}</div>
          <div class="info"><span>Número de la Factura:</span> ${numeroFactura}</div>
          <div class="info"><span>Impuestos:</span></div>
          <ul>
            ${impuestos.map(function (impuesto) {
              return `<li><span>${impuesto.tipo} (${impuesto.percent}%):</span> <span>${formatearPesos(impuesto.amount)}</span></li>`;
            }).join('')}
          </ul>
        </div>
        <div class="column">
          <div class="info"><span>Totales de la Factura:</span></div>
          <ul>
            <li><span>Subtotal:</span> <span>${formatearPesos(invoiceTotal.LineExtensionAmount)}</span></li>
            <li><span>Impuestos Excluidos:</span> <span>${formatearPesos(invoiceTotal.TaxExclusiveAmount)}</span></li>
            <li><span>Impuestos Incluidos:</span> <span>${formatearPesos(invoiceTotal.TaxInclusiveAmount)}</span></li>
            <li><span>Descuentos:</span> <span>${formatearPesos(invoiceTotal.AllowanceTotalAmount)}</span></li>
            <li><span>Cargos:</span> <span>${formatearPesos(invoiceTotal.ChargeTotalAmount)}</span></li>
            <li><span>Pagos Anticipados:</span> <span>${formatearPesos(invoiceTotal.PrePaidAmount)}</span></li>
            <li><span>Total a Pagar:</span> <span>${formatearPesos(invoiceTotal.PayableAmount)}</span></li>
          </ul>
        </div>
      </div>
      <div class="button-container"></div>
    </div>
  `;
}

function mostrarResumenFactura() {
  var jsonString = recuperarJson();
  var json;
  try {
    json = JSON.parse(jsonString);
  } catch (e) {
    json = {};
  }

  var nombreCliente = (json.CustomerInformation && json.CustomerInformation.RegistrationName) ||
                      (json.contacts && json.contacts[0] && json.contacts[0].companyName) ||
                      json.contactName || '';
  var numeroFactura = (json.InvoiceGeneralInformation && json.InvoiceGeneralInformation.InvoiceNumber) ||
                      json.invoiceNumber || '';

  var impuestos = [];
  if (Array.isArray(json.InvoiceTaxTotal)) {
    impuestos = json.InvoiceTaxTotal.map(function (tax) {
      var tipoImpuesto = tax.Id === "01" ? "IVA" : tax.Id === "04" ? "INC" : "ReteRenta";
      return { tipo: tipoImpuesto, percent: tax.Percent, amount: tax.TaxAmount };
    });
  } else if (Array.isArray(json.fieldTaxations)) {
    impuestos = json.fieldTaxations.map(function (tax) {
      var tipo = tax.taxName || 'Impuesto';
      return { tipo: tipo, percent: tax.rate, amount: tax.valueTax };
    });
  }

  var invoiceTotal = json.InvoiceTotal || {
    LineExtensionAmount: Number(json.sumTotalSubTotal || 0),
    TaxExclusiveAmount: Number(json.sumTotalTaxBase || 0),
    TaxInclusiveAmount: Number(json.sumTotalSubTotalAndTax || 0),
    AllowanceTotalAmount: Number(json.sumTotalDiscount || 0),
    ChargeTotalAmount: Number(json.sumTotalCharge || 0),
    PrePaidAmount: Number(json.prePaidAmount || 0),
    PayableAmount: Number(json.sumTotalNetPayable || json.sumTotalTotal || 0)
  };

  var htmlContent = plantillaResumenFactura(nombreCliente, numeroFactura, impuestos, invoiceTotal);
  var htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(600).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Resumen de la Factura');
}

function showCustomDialog() {
  mostrarResumenFactura();
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

function dividirString(input) {
  if (!input || typeof input !== "string") return ["", ""];
  const s = String(input).trim();
  // Caso común: "Nombre - Código" donde el código puede iniciar con letras (p.ej. Q2811001C)
  let m = s.match(/^(.*?)-\s*([A-Za-z]*\d+[A-Za-z0-9]*)$/);
  if (m) return [m[1].trim(), m[2].trim()];
  // Respaldo: nombre seguido de una parte numérica sin guión
  m = s.match(/^(.*?)([-+]?\d+.*)$/);
  if (m) return [m[1].trim(), m[2].trim()];
  return [s, ""];
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
