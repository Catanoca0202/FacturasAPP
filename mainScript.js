// spreadsheet = SpreadsheetApp.getActive();
//let unidades_sheet = spreadsheet.getSheetByName('Unidades');
//let datos_sheet = spreadsheet.getSheetByName('Datos2');

// directorio alejandro C:\\Users\\catan\\OneDrive\\Documents\\Work\\Appsheets\\MisFacturasApp
// directorio sebastian C:\\Users\\elfue\\Documents\\MisFacturasApp
// directorio carlos /home/cley/src/MisFacturasApp

// function onInstall(e) {
//   onOpen(e); // Llama a onOpen durante la instalación
  //ups mal merge
// }

const PRODUCT_COLUMNS = {
  ESTADO: 1,
  CODIGO_REFERENCIA: 2,
  NOMBRE: 3,
  TIPO_PRODUCTO: 4,
  TIPO_USO: 5,
  VALOR_UNITARIO: 6,
  TIPO_IMPUESTO: 7,
  TARIFA_IMPUESTO: 8,
  PRECIO_CON_IMPUESTO: 9,
  CHECK_RECARGO: 10,
  TIPO_RETENCION: 11,
  TARIFA_RETENCION: 12,
  IDENTIFICADOR_UNICO: 13
};

const RETENCION_IRPF_TIPOS = ['IRPF'];
const RETENCION_IRPF_TARIFAS = [formatPercentES(7), formatPercentES(15), formatPercentES(19)];
const RETENCION_RECARGO_LABEL = 'Recargo de equivalencia';
const FACTURA_CHECKBOX_COL = 11; // Columna "Eliminar" en hoja Factura

function OnOpenVariablesGlobales(){
  var spreadsheet = SpreadsheetApp.getActive();
  var hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  var listadoestado_sheet = spreadsheet.getSheetByName('ListadoEstado');
  var prefactura_sheet = spreadsheet.getSheetByName('Factura');
  var unidades_sheet = spreadsheet.getSheetByName('Unidades');
  var listadoestado_sheet = spreadsheet.getSheetByName('ListadoEstado');
  var hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  var folderId = hojaDatosEmisor.getRange("B14").getValue();
  var datos_sheet = spreadsheet.getSheetByName('Datos');
}

function OnOpenSheetInicio(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = ss.getSheetByName("Inicio");
  SpreadsheetApp.setActiveSheet(sheet);
}

function iniciarHojasFactura() {
  Logger.log("Inicio instalación de hojas");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const plantillaID = "1qxbXlhH4RpCOsObk91wsuu4k8jarVK34XXRUlKaKS1U";
  const plantilla = SpreadsheetApp.openById(plantillaID);

  const nombresHojas = ["Inicio", "Productos", "Datos de emisor", "Clientes", "Factura", "Historial Facturas Data", "ClientesInvalidos", "Historial Facturas","Facturas ID", "Copia de Plantilla", "ListadoEstado", "Plantilla", "Celdas plantilla", "Copia de Plantilla", "Copia de Factura","Datos"];
  const hojasBloqueadasEInvisibles = ["ListadoEstado", "Plantilla", "Celdas plantilla", "Historial Facturas Data", "Facturas ID", "Datos", "ClientesInvalidos", "Copia de Plantilla", "Copia de Factura"];



  // Instalar hojas desde la plantilla si no existen
  nombresHojas.forEach(nombreHoja => {
    if (nombreHoja === "Datos") return; // Saltar "Datos" para instalarla al final

    let hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) {
      const hojaPlantilla = plantilla.getSheetByName(nombreHoja);
      if (hojaPlantilla) {
        // Copiar hoja y replicar protecciones
        const hojaCopia = hojaPlantilla.copyTo(ss).setName(nombreHoja);

        // Obtener las protecciones de la hoja original
        const protecciones = hojaPlantilla.getProtections(SpreadsheetApp.ProtectionType.RANGE);

        // Replicar las protecciones en la copia
        protecciones.forEach(proteccion => {
          const rango = proteccion.getRange();
          const rangoEnCopia = hojaCopia.getRange(rango.getA1Notation());

          const nuevaProteccion = rangoEnCopia.protect();
          nuevaProteccion.setDescription(proteccion.getDescription());
          nuevaProteccion.setWarningOnly(proteccion.isWarningOnly());

          // Transferir permisos de edición
          if (!proteccion.isWarningOnly()) {
            nuevaProteccion.addEditors(proteccion.getEditors());
            if (proteccion.canDomainEdit()) {
              nuevaProteccion.setDomainEdit(true);
            }
          }
        });

        // Bloquear la hoja completa si está en la lista de bloqueadas e invisibles
        if (hojasBloqueadasEInvisibles.includes(nombreHoja)) {
          hojaCopia.hideSheet(); // Hacer la hoja invisible
          const protection = hojaCopia.protect();
          protection.removeEditors(protection.getEditors()); // Bloquear completamente
          protection.addEditor(Session.getEffectiveUser()); // Solo el propietario tiene acceso
        }

      } else {
        SpreadsheetApp.getUi().alert('La hoja "' + nombreHoja + '" no existe en la plantilla.');
      }
    }
  });

  // Siempre instalar o reinstalar la hoja "Datos" al final
  reinstalarHojaDatos(ss, plantilla);

    // Eliminar hojas que no pertenezcan a la lista de hojas instaladas
    ss.getSheets().forEach(hoja => {
      const nombreHoja = hoja.getName();
      if (!nombresHojas.includes(nombreHoja)) {
        Logger.log(hoja.getName())
        Logger.log("hojaname")
        ss.deleteSheet(hoja);
      }
    });

  SpreadsheetApp.getUi().alert("Hojas instaladas satisfactoriamente.");
  //SpreadsheetApp.getUi().alert("Recuerda que antes de utilizar FacturasApp debes de crear la carpeta donde se guardarán las facturas. Dirígete a la hoja Datos de emisor y dale clic en el botón crear carpeta.");
}

function reinstalarHojaDatos(ss, plantilla) {
  Logger.log("Reinstalando hoja Datos...");

  const nombreHoja = "Datos";
  let hojaDatos = ss.getSheetByName(nombreHoja);
  Logger.log("After getting hojadatos")
  // Eliminar la hoja "Datos" si ya existe
  if (hojaDatos) {
    ss.deleteSheet(hojaDatos);
    Logger.log("AIFF ")
  }

  // Copiar la hoja "Datos" desde la plantilla
  const hojaPlantilla = plantilla.getSheetByName(nombreHoja);
  if (hojaPlantilla) {
    const hojaCopia = hojaPlantilla.copyTo(ss).setName(nombreHoja);
    Logger.log("dentro if")
    // Bloquear la hoja "Datos"
    const protection = hojaCopia.protect();
    protection.removeEditors(protection.getEditors());
    protection.addEditor(Session.getEffectiveUser());
    hojaCopia.hideSheet(); // Hacer la hoja invisible
    Logger.log("hoja aca")
  } else {
    SpreadsheetApp.getUi().alert('La hoja "Datos" no existe en la plantilla.');
  }
}

function cambiarConfiguracionRegional() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cambiar la configuración regional a España
  sheet.setSpreadsheetLocale("es_ES");
  
  Logger.log("Configuración regional cambiada a España (es_ES)");
}


function IniciarFacturasApp(){
  let ui = SpreadsheetApp.getUi();
  
  let hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos de emisor");

  if (hoja==null){
    iniciarHojasFactura()
    OnOpenSheetInicio()
    agregarDataValidations()
    cambiarConfiguracionRegional()

  }else{
    let respuesta = ui.alert('Si vuelves a instalar, solo se instalaran las hojas no existan o que hayan sido eliminadas?', ui.ButtonSet.YES_NO);
    if (respuesta == ui.Button.YES) {
      iniciarHojasFactura()
      OnOpenSheetInicio()
      agregarDataValidations()
      cambiarConfiguracionRegional()
    } else {
      return
    }
  }
  //OnOpenVariablesGlobales()
  
}
function onOpen(e) {
  //showSidebar()
  // let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi();
  // if(e && e.authMode ==ScriptApp.AuthMode.NONE){
  //   Logger.log("ScriptApp.AuthMode.NONE")
  //   ui.createAddonMenu()
  //   .addItem('Inicio', 'showSidebar2')
  //   .addToUi();
  // }else{

  // }
  Logger.log("ScriptApp.AuthMode.NONE")
  ui.createAddonMenu()
  .addItem('Inicio', 'showSidebar2')
  .addItem('Instalar', 'IniciarFacturasApp')
  .addItem("Desinstalar","eliminarHojasFactura").addToUi();

  // https://developers.google.com/apps-script/guides/menus



  //showSidebar()

  Logger.log("no entra a ")
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // var sheet = ss.getSheetByName("Inicio");
  // SpreadsheetApp.setActiveSheet(sheet);
  
  // OnOpenVariablesGlobales()
  // OnOpenSheetInicio()
  return;
}

// function installableOnOpen(e) {
//   // Esta función actúa como el disparador
//   var ui = SpreadsheetApp.getUi();
//   ui.createAddonMenu()
//     .addItem('Instalar', 'IniciarFacturasApp')
//     .addSeparator()
//     .addItem('Inicio', 'showSidebar2')
//     .addToUi();
// }

// function createInstallableTrigger() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   ScriptApp.newTrigger('installableOnOpen')  // Aquí ponemos la función con permisos adecuados
//     .forSpreadsheet(ss)
//     .onOpen()  // El evento que activará el trigger
//     .create();
// }

function pruebaLogo(){
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos de emisor");
  var celdaLogo = hoja.getRange("B13").getValue();
  hoja.getRange("B20").setValue(celdaLogo);
}

function showDesvincular(){
  var html = HtmlService.createHtmlOutputFromFile('menuDesvincular')
  .setTitle('Desvincular cuenta');
SpreadsheetApp.getUi()
  .showSidebar(html);
}

function showSidebar() {
  console.log("showSidebar Enters");
 
  console.log("setActiveSheet Inicio");
  var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inicio");
  SpreadsheetApp.setActiveSheet(sheet);


  var html = HtmlService.createHtmlOutputFromFile('main')
    .setTitle('Menú');
  SpreadsheetApp.getUi()
    .showSidebar(html);
  console.log("showSidebar Exits");    
}

function showSidebar2() {
  const usuario =obtenerUsuario()
  const propietario= obtenerPropietario()
  console.log("showSidebar2 Enters");
  let ui = SpreadsheetApp.getUi();
  console.log("setActiveSheet2 Inicio");
  const scriptProps = PropertiesService.getDocumentProperties();
  scriptProps.setProperties({
    'propietario': propietario,
  });
  // var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inicio");
  // SpreadsheetApp.setActiveSheet(sheet);
  let hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos de emisor");
  Logger.log("hoja "+hoja)
  Logger.log(typeof(hoja))
  if(hoja==null){
    let respuesta = ui.alert('Primero debes de instalar las hojas necesarias ¿Deseas instalarlas ya?', ui.ButtonSet.YES_NO);
    if (respuesta == ui.Button.YES) {
      iniciarHojasFactura()
      OnOpenSheetInicio()
      agregarDataValidations()
    } else {
      return
    }
  }else{
    var template = HtmlService.createTemplateFromFile('main');
    template.emailPropietario=propietario
    const html = template.evaluate().setTitle('Menú');
    SpreadsheetApp.getUi().showSidebar(html);
    console.log("showSidebar Exits"); 

  }
}

function showInstalarHojas() {
  var html = HtmlService.createHtmlOutputFromFile('instalarHojas')
      .setWidth(400)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Instala las hojas');
}

function showVincularCuenta() {
  var html = HtmlService.createHtmlOutputFromFile('menuVincular')
    .setTitle('Vincular cuenta');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showEliminarInfo(){
  var html = HtmlService.createHtmlOutputFromFile('menuEliminarInfo')
  .setTitle('Eliminar informacion');
SpreadsheetApp.getUi()
  .showSidebar(html);
}

function showPreProductos() {
  console.log("Attempting to show Productos");
  respuesta=verficiarPropietario()
  if(respuesta){
  var html = HtmlService.createHtmlOutputFromFile('preProductos')
    .setTitle('Productos');
  SpreadsheetApp.getUi()
    .showSidebar(html);
  }else{
    SpreadsheetApp.getUi().alert("No tienes permisos para acceder a esta función debido a que tienes activos dos correos en la hoja. Por favor desvincula uno de los correos y vuelve a intentar")
  }
}

function showAggProductos() {
  var html = HtmlService.createHtmlOutputFromFile('agregarProducto')
    .setTitle('Agregar Productos');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function openFacturaSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Factura");
  SpreadsheetApp.setActiveSheet(sheet);
}

function openHistorialSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Historial Facturas");
  SpreadsheetApp.setActiveSheet(sheet);
}

function showMenuFactura() {
  obtenerUsuario()
  obtenerPropietario()
  respuesta=verficiarPropietario()
  Logger.log(respuesta)
  if(respuesta){
    Logger.log("showMenuFactura adentro")
    var html = HtmlService.createHtmlOutputFromFile('menuFactura')
      .setTitle('Menú Factura');
    SpreadsheetApp.getUi()
      .showSidebar(html);
  }else{  
    Logger.log("showMenuFactura adentro2")
    SpreadsheetApp.getUi().alert("No tienes permisos para acceder a esta función debido a que tienes activos dos correos en la hoja. Por favor desvincula uno de los correos y vuelve a intentar")
  }

}

function obtenerUsuario() {
  var email = Session.getActiveUser().getEmail();
  Logger.log("Usuario activo: " + email);
  return(email)
}

function obtenerPropietario() {
  var email = Session.getEffectiveUser().getEmail();
  Logger.log("Propietario del script: " + email);
  return(email)
}

function verficiarPropietario() {
  try {
    Logger.log("verificar propietario");
    var email = Session.getEffectiveUser().getEmail();
    var scriptProps = PropertiesService.getDocumentProperties();
    var propietario = scriptProps.getProperty('propietario');
    
    Logger.log("propietario: " + propietario);
    Logger.log("email: " + email);
    
    if (email === propietario) {
      return true;
    } else {
      return false;
    }

  } catch (e) {
    Logger.log("Error en verificarPropietario: " + e);
    return false;
  }
}

function showNuevaFactura() {
  var html = HtmlService.createHtmlOutputFromFile('nuevaFactura').setTitle("Nueva factura")
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showAgregarProdcuto() {
  var html = HtmlService.createHtmlOutputFromFile('menuAgregarProducto').setTitle("Agregar Producto")
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function openClientesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Clientes");
  SpreadsheetApp.setActiveSheet(sheet);
}
function openDatosEmisorSheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Datos de emisor");
  SpreadsheetApp.setActiveSheet(sheet);
}

function showClientes() {
  respuesta=verficiarPropietario()
  if(respuesta){

  var html = HtmlService.createHtmlOutputFromFile('menuCliente')
    .setTitle('Menu cliente');
  SpreadsheetApp.getUi()
    .showSidebar(html);
  }else{
    SpreadsheetApp.getUi().alert("No tienes permisos para acceder a esta función debido a que tienes activos dos correos en la hoja. Por favor desvincula uno de los correos y vuelve a intentar")
  }
}

function showAjustes(){
  var html = HtmlService.createHtmlOutputFromFile('menuAjustes')
  .setTitle('Datos emisor');
SpreadsheetApp.getUi()
  .showSidebar(html);
}


function openProductosSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Productos");
  SpreadsheetApp.setActiveSheet(sheet);
}

function showEnviarEmail() {
  var html = HtmlService.createHtmlOutputFromFile('enviarEmail')
    .setTitle('Enviar Email');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function inicarFacturaNuevaMain() {
  inicarFacturaNueva();
}

function showPostFactura() {
  var html = HtmlService.createHtmlOutputFromFile('postFactura')
    .setTitle('Post Factura');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}
function showEnviarEmailHistorial(data){
  var html = HtmlService.createHtmlOutputFromFile('enviarEmailHistorial')
    .setTitle('Enviar Email Historial');
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function showEnviarEmailPost() {
  var html = HtmlService.createHtmlOutputFromFile('enviarEmailPost')
    .setWidth(400)
    .setHeight(400)
  SpreadsheetApp.getUi()
    .showModalDialog(html,"Digite el email a enviar");
}

function showEnviarEmailPostHistorial(){
  var html =HtmlService.createHtmlOutputFromFile('enviarEmailPostHistorial')
    .setWidth(400)
    .setHeight(400)
  SpreadsheetApp.getUi()
    .showModalDialog(html,"Digite el email a enviar")
}

function eliminarHojasFactura() {
  let ui = SpreadsheetApp.getUi();
  Logger.log("Inicio de eliminación de hojas");
  let respuesta = ui.alert('Recuerda que al desinstalar las hojas se eliminará toda la información de las mismas. Esta función solo debe ejecutarse si tienes un problema irreparable con las hojas. ¿Estás seguro de continuar?', ui.ButtonSet.YES_NO);
  if (respuesta == ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const nombresHojas = ["Inicio", "Productos", "Datos de emisor", "Historial Facturas", "Clientes", "Factura", "Historial Facturas Data", "Facturas ID", "Datos", "Copia de Plantilla", "ListadoEstado", "Plantilla", "Celdas plantilla", "ClientesInvalidos", "Copia de Plantilla", "Copia de Factura"];

    // Crear una hoja nueva en blanco
    let nuevaHoja = ss.getSheetByName("Hoja en blanco");
    if (!nuevaHoja) {
      nuevaHoja = ss.insertSheet("Hoja en blanco");
      Logger.log("Se creó una nueva hoja en blanco");
    }

    // Recorrer todas las hojas del archivo
    ss.getSheets().forEach(hoja => {
      const nombreHoja = hoja.getName();
      if (nombresHojas.includes(nombreHoja)) {
        ss.deleteSheet(hoja);
        Logger.log(`Hoja eliminada: ${nombreHoja}`);
      }
    });

    SpreadsheetApp.getUi().alert("Hojas eliminadas satisfactoriamente.");
  } else {
    return;
  }
}

function agregarDataValidations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDatos = ss.getSheetByName("Datos");
  const hojaFacturas = ss.getSheetByName("Factura");
  const hojaValoresC = ss.getSheetByName("Clientes");
  const hojaValoresP = ss.getSheetByName("Productos");
  const hojaValoresCInvalidos = ss.getSheetByName("ClientesInvalidos");
  const HojaValorescopiaFactura=ss.getSheetByName("Copia de Factura");

  // Rango donde aplicar los dropdowns
  const rangoDropdownCliente = hojaDatos.getRange("H2");
  const rangoDropdownClienteInvalido = hojaDatos.getRange("I6");
  const rangoDropdownProductos = hojaDatos.getRange("I11");
  const rangoDropdownClienteF = hojaFacturas.getRange("B2:C2");
  const rangoDropdownProductoF = hojaFacturas.getRange("B15");
  const rangoDropdownCopiaFacturaCliente=HojaValorescopiaFactura.getRange("B2:C2")
  const rangoDropdownCopiaFacturaProducto=HojaValorescopiaFactura.getRange("B15")

  // Rango de valores para los dropdowns
  const rangoValoresClienteInvalido = hojaValoresCInvalidos.getRange("V2:V1000");
  const rangoValoresClienteDatos = hojaValoresC.getRange("B2:B1000");
  const rangoValoresProductosDatos = hojaValoresP.getRange("M2:M1000");
  const rangoValoresClienteFactura = hojaValoresC.getRange("$B$2:$B$1000");
  const rangoValoresProductosFactura = hojaValoresP.getRange("$M$2:$M$1000");

  // Crear y aplicar validaciones
  const reglas = [
    {
      rango: rangoDropdownCliente,
      valores: rangoValoresClienteDatos
    },
    {
      rango: rangoDropdownClienteInvalido,
      valores: rangoValoresClienteInvalido
    },
    {
      rango: rangoDropdownProductos,
      valores: rangoValoresProductosDatos
    },
    {
      rango: rangoDropdownClienteF,
      valores: rangoValoresClienteFactura
    },
    {
      rango: rangoDropdownProductoF,
      valores: rangoValoresProductosFactura
    },
    {
      rango:rangoDropdownCopiaFacturaCliente,
      valores:rangoValoresClienteFactura
    },
    {
      rango:rangoDropdownCopiaFacturaProducto,
      valores:rangoValoresProductosFactura
    }
  ];

  reglas.forEach(({ rango, valores }) => {
    const regla = SpreadsheetApp.newDataValidation()
      .requireValueInRange(valores, true) // Usar valores del rango especificado
      .setAllowInvalid(false) // No permitir valores fuera del rango
      .build();
    rango.setDataValidation(regla); // Aplicar la regla
  });

  SpreadsheetApp.getUi().alert("Validaciones de datos aplicadas correctamente.");
}



function processForm(data) {
  let existe=verificarCodigo(data.codigoReferencia,"Productos",false)
  if(existe){
    SpreadsheetApp.getUi().alert("El codigo de referencia ya existe, por favor poner un codigo de referencia unico");
    throw new Error('por favor poner un Numero de Identificacion unico');
  }
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Productos");
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;

    const codigoReferencia = data.codigoReferencia;
    const nombre = data.nombre;
    const tipoProducto = data.tipoProducto || '';
    const tipoUso = data.tipoUso || '';
    const valorUnitario = parseFloat(data.valorUnitario);
    const tipoImpuesto = data.tipoImpuesto || 'IVA';
    const tarifaImpuestoRaw = data.tarifaImpuesto || data.iva || '';
    const tarifaImpuestoNum = tarifaImpuestoRaw === '' ? null : parsePercentToNumberES(tarifaImpuestoRaw);
    const tarifaImpuestoStr = tarifaImpuestoNum !== null ? formatPercentES(tarifaImpuestoNum) : '';

    const aplicarRecargoFormulario = String(data.aplicarRecargo || '').toLowerCase() === 'true';
    const recargoSeleccionado = data.recargo && String(data.recargo).toLowerCase() !== 'seleccione' ? parsePercentToNumberES(data.recargo) : null;
    const retencionSeleccionada = data.retenciones && String(data.retenciones).toLowerCase() !== 'seleccione' ? parsePercentToNumberES(data.retenciones) : null;

    let tipoRetencion = data.tipoRetencion || '';
    let tarifaRetencionStr = data.tarifaRetencion || '';
    let aplicarRecargo = aplicarRecargoFormulario || recargoSeleccionado !== null;

    if (!aplicarRecargo && retencionSeleccionada !== null) {
      tipoRetencion = 'IRPF';
      tarifaRetencionStr = formatPercentES(retencionSeleccionada);
    }

    if (aplicarRecargo) {
      const permitidoNum = recargoPermitidoParaIva(tarifaImpuestoNum);
      if (permitidoNum === null) {
        SpreadsheetApp.getUi().alert('La tarifa de IVA seleccionada no permite recargo de equivalencia.');
        aplicarRecargo = false;
        tipoRetencion = '';
        tarifaRetencionStr = '';
      } else {
        tipoRetencion = RETENCION_RECARGO_LABEL;
        tarifaRetencionStr = formatPercentES(permitidoNum);
      }
    } else if (tarifaRetencionStr !== '') {
      const tarifaIrpfNum = parsePercentToNumberES(tarifaRetencionStr);
      if (tarifaIrpfNum !== null && !isNaN(tarifaIrpfNum)) {
        tarifaRetencionStr = formatPercentES(tarifaIrpfNum);
      }
    }

    sheet.getRange(newRow, PRODUCT_COLUMNS.CODIGO_REFERENCIA).setValue(codigoReferencia);
    sheet.getRange(newRow, PRODUCT_COLUMNS.CODIGO_REFERENCIA).setHorizontalAlignment('center');

    sheet.getRange(newRow, PRODUCT_COLUMNS.NOMBRE).setValue(nombre);
    sheet.getRange(newRow, PRODUCT_COLUMNS.NOMBRE).setHorizontalAlignment('center');

    sheet.getRange(newRow, PRODUCT_COLUMNS.TIPO_PRODUCTO).setValue(tipoProducto);
    sheet.getRange(newRow, PRODUCT_COLUMNS.TIPO_USO).setValue(tipoUso);

    const valorUnitarioRange = sheet.getRange(newRow, PRODUCT_COLUMNS.VALOR_UNITARIO);
    if (isNaN(valorUnitario)) {
      valorUnitarioRange.clearContent();
    } else {
      valorUnitarioRange.setValue(valorUnitario);
      valorUnitarioRange.setNumberFormat('€#,##0.00');
    }

    sheet.getRange(newRow, PRODUCT_COLUMNS.TIPO_IMPUESTO).setValue(tipoImpuesto);
    if (tarifaImpuestoStr !== '') {
      const tarifaImpuestoRange = sheet.getRange(newRow, PRODUCT_COLUMNS.TARIFA_IMPUESTO);
      tarifaImpuestoRange.setValue(tarifaImpuestoStr);
      tarifaImpuestoRange.setNumberFormat('0.00%');
    }

    const precioConImpuestoFormula = `=IF(AND(F${newRow}<>"";H${newRow}<>"");F${newRow}*(1+H${newRow});"")`;
    sheet.getRange(newRow, PRODUCT_COLUMNS.PRECIO_CON_IMPUESTO).setFormula(precioConImpuestoFormula);
    sheet.getRange(newRow, PRODUCT_COLUMNS.PRECIO_CON_IMPUESTO).setNumberFormat('€#,##0.00');

    sheet.getRange(newRow, PRODUCT_COLUMNS.CHECK_RECARGO).setValue(aplicarRecargo);
    sheet.getRange(newRow, PRODUCT_COLUMNS.TIPO_RETENCION).setValue(tipoRetencion);
    sheet.getRange(newRow, PRODUCT_COLUMNS.TARIFA_RETENCION).setValue(tarifaRetencionStr);

    if (aplicarRecargo) {
      manejarCheckboxRecargo(sheet, newRow);
    } else {
      aplicarValidacionTipoRetencion(sheet, newRow, false);
      aplicarValidacionTarifaRetencion(sheet, newRow, false);
    }

    const camposRequeridos = [codigoReferencia, nombre, tipoProducto, tipoUso, tipoImpuesto, tarifaImpuestoStr];
    let estado = camposRequeridos.some(valor => valor === '' || valor === null || String(valor).toLowerCase() === 'seleccione') || isNaN(valorUnitario)
      ? 'No Valido'
      : 'Valido';

    if (estado === 'Valido') {
      if (aplicarRecargo && tarifaRetencionStr === '') {
        estado = 'No Valido';
      }
      if (!aplicarRecargo && tipoRetencion === 'IRPF' && tarifaRetencionStr === '') {
        estado = 'No Valido';
      }
    }

    sheet.getRange(newRow, PRODUCT_COLUMNS.ESTADO).setValue(estado);

    let referenciaUnica = '';
    if (estado === 'Valido') {
      referenciaUnica = nombre+"-"+codigoReferencia;
      sheet.getRange(newRow, PRODUCT_COLUMNS.IDENTIFICADOR_UNICO).setValue(referenciaUnica);
    }

    SpreadsheetApp.getUi().alert("Nuevo producto generado satisfactoriamente");
    return {message: "Datos guardados correctamente", refe: referenciaUnica || null};
  } catch (error) {
    return {message:"Error al guardar los datos: " + error.message,refe: null};
  }
}

function generatePdfFromPlantilla() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Copia de Plantilla');
  var celdaNumFactura = ss.getSheetByName('Factura').getRange('A9').getValue();
  var numFactura = celdaNumFactura.substring(20);

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
      var pdfBlob = response.getBlob().setName('Factura '&numFactura&'.pdf');
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



function getPdfUrl() {
  var pdfBlob = generatePdfFromPlantilla();
  var base64Data = Utilities.base64Encode(pdfBlob.getBytes());
  var contentType = pdfBlob.getContentType();
  var name = pdfBlob.getName();
  return `data:${contentType};base64,${base64Data}`;
}

function sendPdfByEmail(email) {
  let hojaDatosEmisor = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Datos de emisor');
  let hojaListadoEstado = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ListadoEstado');
  let lastRow = hojaListadoEstado.getLastRow();
  let numFactura;
  let lastRowfacturasID;

  let nombreCliente = hojaDatosEmisor.getRange("B2").getValue();
  var pdfFile = generatePdfFromFactura();

  busquedaLineal()
  var subject = `📄 Nueva factura de ${nombreCliente}`;
  var body = `¡Hola!\n` +
           `${nombreCliente} te ha enviado la siguiente factura:\n` +
           `🔹 Número de factura: ${numFactura}\n` +
           `💰 Valor: 0000 €\n` +
           `Si tienes alguna duda, contacta directamente con ${nombreCliente}.\n` +
           `Saludos,\n` +
           `${nombreCliente}\n\n`+
           `📌 ¿Necesitas facturación electrónica? Ahorra tiempo y factura fácilmente con FacturasApp\n` +
           `👉 Ver más: https://www.facturasapp.com/Publico/`;

  if (!email) {
    return "Por favor ingrese una dirección de correo válida.";
  }

  MailApp.sendEmail({
    to: email,
    subject: subject,
    body: body,
    attachments: [pdfFile.getAs(MimeType.PDF)]
  });

  return "PDF generado y enviado por correo electrónico a " + email;
}

function convertToPercentage(value) {
  return (value * 100).toFixed(2).replace('.', ',') + '%';
}

function mostrarAlertaDesdeServidor(mensaje) {
  SpreadsheetApp.getUi().alert(mensaje);
}


/** Utilidades de IVA y Recargo (2025) **/
function parsePercentToNumberES(value) {
  // Acepta números, strings con "," o "." y con/sin "%". Devuelve número (por ejemplo 5.2)
  if (value === null || typeof value === 'undefined') return null;
  if (typeof value === 'number') {
    // A veces se almacena 0.05 para 5%
    return value > 0 && value < 1 ? value * 100 : value;
  }
  const str = String(value).trim().replace('%', '').replace(',', '.');
  if (str === '') return null;
  const num = Number(str);
  if (isNaN(num)) return null;
  return num;
}

function formatPercentES(num) {
  if (num === null || typeof num === 'undefined') return '';
  if (num === 0) return '0%';
  const fixed = Number(num).toFixed(2).replace('.', ',');
  return `${fixed}%`;
}

// Mapa oficial 2025 de recargo permitido por IVA
// Nota: mantenemos 4% -> 0,50% por compatibilidad con productos existentes
const IVA_RECARGO_MAP_2025 = {
  '21': 5.2,
  '10': 1.4,
  '5': 0.5,
  '4': 0.5,
  '0': 0
};

function recargoPermitidoParaIva(ivaNum) {
  if (ivaNum === null) return null;
  const key = String(Math.round(ivaNum));
  return Object.prototype.hasOwnProperty.call(IVA_RECARGO_MAP_2025, key)
    ? IVA_RECARGO_MAP_2025[key]
    : null;
}

function aplicarValidacionTipoRetencion(hoja, fila, esRecargo) {
  const rango = hoja.getRange(fila, PRODUCT_COLUMNS.TIPO_RETENCION);
  if (esRecargo) {
    const regla = SpreadsheetApp.newDataValidation()
      .requireValueInList([RETENCION_RECARGO_LABEL], true)
      .setAllowInvalid(false)
      .build();
    rango.setDataValidation(regla);
    rango.setValue(RETENCION_RECARGO_LABEL);
  } else {
    const regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(RETENCION_IRPF_TIPOS, true)
      .setAllowInvalid(true)
      .build();
    rango.setDataValidation(regla);
  }
}

function aplicarValidacionTarifaRetencion(hoja, fila, esRecargo, ivaNum) {
  const rango = hoja.getRange(fila, PRODUCT_COLUMNS.TARIFA_RETENCION);
  if (esRecargo) {
    const permitido = recargoPermitidoParaIva(ivaNum);
    if (permitido === null) {
      rango.clearDataValidations();
      rango.clearContent();
      return;
    }
    const etiqueta = formatPercentES(permitido);
    const regla = SpreadsheetApp.newDataValidation()
      .requireValueInList([etiqueta], true)
      .setAllowInvalid(false)
      .build();
    rango.setDataValidation(regla);
    rango.setValue(etiqueta);
  } else {
    const regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(RETENCION_IRPF_TARIFAS, true)
      .setAllowInvalid(true)
      .build();
    rango.setDataValidation(regla);
  }
}

function manejarCheckboxRecargo(hoja, fila) {
  const esRecargo = hoja.getRange(fila, PRODUCT_COLUMNS.CHECK_RECARGO).getValue() === true;
  if (esRecargo) {
    const ivaDisplay = hoja.getRange(fila, PRODUCT_COLUMNS.TARIFA_IMPUESTO).getDisplayValue();
    const ivaNum = parsePercentToNumberES(ivaDisplay);
    aplicarValidacionTipoRetencion(hoja, fila, true);
    aplicarValidacionTarifaRetencion(hoja, fila, true, ivaNum);
  } else {
    const rangoTipo = hoja.getRange(fila, PRODUCT_COLUMNS.TIPO_RETENCION);
    const rangoTarifa = hoja.getRange(fila, PRODUCT_COLUMNS.TARIFA_RETENCION);
    aplicarValidacionTipoRetencion(hoja, fila, false);
    aplicarValidacionTarifaRetencion(hoja, fila, false);
    rangoTipo.clearContent();
    rangoTarifa.clearContent();
  }
}

function sincronizarRecargoSegunIva(hoja, fila) {
  const esRecargo = hoja.getRange(fila, PRODUCT_COLUMNS.CHECK_RECARGO).getValue() === true;
  if (!esRecargo) {
    return;
  }
  const ivaDisplay = hoja.getRange(fila, PRODUCT_COLUMNS.TARIFA_IMPUESTO).getDisplayValue();
  const ivaNum = parsePercentToNumberES(ivaDisplay);
  aplicarValidacionTarifaRetencion(hoja, fila, true, ivaNum);
}

function onEdit(e) {
  const lock = LockService.getScriptLock();
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaActual = e.source.getActiveSheet();
  let factura_sheet = spreadsheet.getSheetByName('Factura');

  //verificarTipoDeDatos(e);
  try{
    lock.waitLock(6000);
    if (hojaActual.getName() === "Factura") {
      
      let celdaEditada = e.range;
      let rowEditada = celdaEditada.getRow();
      let colEditada = celdaEditada.getColumn();
      let columnaContactos = 2; // Ajusta según sea necesario
      let rowContactos = 2;


      const productStartRow = 15; // prodcutos empeiza aca
      const productEndColumn = 8; //   procutos terminan en column H
      let taxSectionStartRow = getTaxSectionStartRow(hojaActual); // Assuming products end at column H
      let posRowTotalProductos=taxSectionStartRow-3//poscion (row) de Total productos
      Logger.log("taxSectionStartRow "+taxSectionStartRow)
      let VerificartotalFilas = hojaActual.getRange("A15").getValue()
      if(VerificartotalFilas==="Total filas"){
        SpreadsheetApp.getUi().alert("Intenta no eliminar todas las filas de productos, ya que se puede generar un error en la factura");
        limpiarHojaFactura()
        
      }
      else if (colEditada === columnaContactos && rowEditada === rowContactos) {
        //celda de elegir contacto en hoja factura
        Logger.log("No se editó un contacto válido");
        verificarYCopiarContacto(e);
        obtenerFechaYHoraActual()
        //generarNumeroFactura()
        let hojaInfoUsuario= spreadsheet.getSheetByName('Datos de emisor');
        let  iban= hojaInfoUsuario.getRange("B10").getValue();
        factura_sheet.getRange("B11").setValue(iban)
        generarNumeroFactura()

      }
      else if(rowEditada >= productStartRow && (colEditada == 2 || colEditada == 3) && rowEditada < posRowTotalProductos)  {//asegurar que si sea dentro del espacio permititdo(donde empieza el taxinfo)
        if (colEditada == 2){
          Logger.log("agg producto")
          let valorok=celdaEditada.getValue()
          Logger.log("valor+ "+valorok)
          let dictInformacionProducto = obtenerInformacionProducto(valorok);
          let Estado=dictInformacionProducto["Estado"]
          if(Estado==="No Valido"){
            SpreadsheetApp.getUi().alert("El prodcuto elegido tiene un estado invalido. Verifica que el prodcuto posea los datos minimos para ser valido y vuelve a elegir");
            celdaEditada.setValue("")
          }

        }
        const lastProductRow = getLastProductRow(hojaActual, productStartRow, taxSectionStartRow);//1 producto
        Logger.log("lastProductRow " + lastProductRow)
        Logger.log("taxSectionStartRow " + taxSectionStartRow)


        //proceso para agg el valor de %IVA y precio unitario
        for(let i=productStartRow;i <= lastProductRow;i++){
          Logger.log(rowEditada+" row editada")
          if(i!=rowEditada){
            Logger.log("no es la linea editada "+i)
          }else{
          //por aca seria el proceso de ver si el IVA del producto esta entre el rango de tiempo
          let productoFilaI = factura_sheet.getRange("B"+String(i)).getValue()
          if(productoFilaI===""){
            Logger.log("NO ha elegido producto")
            continue
          }
          let dictInformacionProducto = obtenerInformacionProducto(productoFilaI);
          let cantiadProducto= factura_sheet.getRange("C"+String(i)).getValue()

          let ivaProductoActual=dictInformacionProducto["IVA"]
          let valorFechaActual=ObtenerFecha()
      
          let verifcadorFecha=verificarDescuentoValido(valorFechaActual,ivaProductoActual)
          if (verifcadorFecha===false){
            SpreadsheetApp.getUi().alert("Alguno de tus productos posee un iva del 5%. La fecha de facturación debe estar comprendida entre el 1 de julio de 2022 y el 30 de septiembre de 2024")
            //poner rangos en 0
            continue
          }

          if(cantiadProducto===""){
            cantiadProducto=0
            //tal vez mirara si agrego el 0 de cantidad
            factura_sheet.getRange("A"+String(i)).setValue(dictInformacionProducto["codigo Producto"])
            factura_sheet.getRange("D"+String(i)).setValue(0)//unitario preciounitario
            factura_sheet.getRange("G"+String(i)).setValue(dictInformacionProducto["IVA"])//IVA
            const recargoLinea = dictInformacionProducto["Recargo de equivalencia"];
            factura_sheet.getRange("I"+String(i)).setValue(recargoLinea);//Tarifa recargo
            const totalFormula = `=IF(F${i}="";0;F${i}*(1+G${i}+I${i}))`;
            factura_sheet.getRange("J"+String(i)).setValue(totalFormula);
          }else{
            factura_sheet.getRange("A"+String(i)).setValue(dictInformacionProducto["codigo Producto"])
            factura_sheet.getRange("E"+String(i)).setValue("=D"+String(i)+"+(D"+String(i)+"*G"+String(i)+")")//AGG COSA DE CON IVA 
            factura_sheet.getRange("F"+String(i)).setValue("=(D"+String(i)+")*C"+String(i)+"-((D"+String(i)+")*C"+String(i)+")*H"+String(i))//subtotal
            factura_sheet.getRange("D"+String(i)).setValue(dictInformacionProducto["valor Unitario"])//valor unitario
            factura_sheet.getRange("G"+String(i)).setValue(dictInformacionProducto["IVA"])//IVA
            const recargoLinea = dictInformacionProducto["Recargo de equivalencia"];
            factura_sheet.getRange("I"+String(i)).setValue(recargoLinea);//Tarifa recargo
            const totalFormula = `=IF(F${i}="";0;F${i}*(1+G${i}+I${i}))`;
            factura_sheet.getRange("J"+String(i)).setValue(totalFormula);//total linea
          }
        }

        }

      }else if(colEditada==8 && rowEditada >= productStartRow && rowEditada < posRowTotalProductos) {
        //verificar descuentos
        let valorEditadoDescuneto = celdaEditada.getValue();
        Logger.log(typeof(valorEditadoDescuneto))
        Logger.log("valorEditadoDescuneto "+valorEditadoDescuneto)

        if(0.00 > valorEditadoDescuneto || valorEditadoDescuneto > 1.00){
          Logger.log("No se puede pasar de 100% el valor de descuento o menos de 0%")
          SpreadsheetApp.getUi().alert("No es valido un descuento mayor a 100% ni menor a 0%")
          celdaEditada.setValue("0%")
        }


      }else if (colEditada == 7 && rowEditada == 6) {
        // Entra a verificar días de vencimiento
        let valorDiasVencimiento = celdaEditada.getValue();
      
        // Verifica si es un entero positivo
        if (!Number.isInteger(valorDiasVencimiento) || valorDiasVencimiento <= 0) {
          // Muestra una alerta
          SpreadsheetApp.getUi().alert('El valor de días de vencimiento debe ser un entero positivo.');
      
          // Restablece el valor a 0
          celdaEditada.setValue(0);
        }
        Logger.log("dentro de fecha calcular")
        CalcularDiasOFecha("Dias")
        // Refrescar datos clave como cuando se elige el cliente
        obtenerFechaYHoraActual();
        let hojaInfoUsuario = spreadsheet.getSheetByName('Datos de emisor');
        let iban = hojaInfoUsuario.getRange("B10").getValue();
        factura_sheet.getRange("B11").setValue(iban);
        generarNumeroFactura();
      }else if (colEditada == 5 && rowEditada == 4) {
        // Medio de pago editado (E4): refrescar datos clave y fechas
        obtenerFechaYHoraActual();
        CalcularDiasOFecha("Dias");
        let hojaInfoUsuario = spreadsheet.getSheetByName('Datos de emisor');
        let iban = hojaInfoUsuario.getRange("B10").getValue();
        factura_sheet.getRange("B11").setValue(iban);
        generarNumeroFactura();
      }else if (colEditada == 7 && rowEditada == 5) {
        // Forma de pago editada (G5): refrescar datos clave
        obtenerFechaYHoraActual();
        let hojaInfoUsuario = spreadsheet.getSheetByName('Datos de emisor');
        let iban = hojaInfoUsuario.getRange("B10").getValue();
        factura_sheet.getRange("B11").setValue(iban);
        generarNumeroFactura();
      }else if(colEditada==7 && rowEditada==2){
        let valorFacturaNumero = celdaEditada.getValue();
        let coincideEstruct=cumpleEstructura(valorFacturaNumero)
        if(coincideEstruct==0){
          SpreadsheetApp.getUi().alert("No puedes editar el consecutivo de la factura sin primero guardaro una estructura del consecutivo");
          celdaEditada.setValue("");
        }
        else if(!coincideEstruct){
          SpreadsheetApp.getUi().alert("El consecutivo de la factura debe de coincidir con la estructura que tu elegiste");
          celdaEditada.setValue("");
          generarNumeroFactura()
        }else {
          Logger.log("coincideEstruct "+coincideEstruct)
          let existe=verificarCodigo(valorFacturaNumero,"Historial Facturas Data",false)
          if(existe){
            SpreadsheetApp.getUi().alert("El numero de factura ya existe, por favor poner un numero de factura unico");
            celdaEditada.setValue("");
            generarNumeroFactura()
            throw new Error('por favor poner un Numero de Identificacion unico');
          }
      }
      }else if(colEditada==7 && (rowEditada ==4 || rowEditada ==3) ){
        Logger.log("dentro de fecha calcular2")
        CalcularDiasOFecha("Fecha")
      }
      
      let lastRowProducto=getLastProductRow(hojaActual, productStartRow, taxSectionStartRow);
      if (lastRowProducto===productStartRow){
        Logger.log("dentro de agg info para TOTLA pero last y start son iguales")
        // //ESTADO DEAFULT no se hace nada
        hojaActual.getRange("B31").setValue("=B32+C29-A29")


      }else{
        Logger.log("dentro de agg info para totoal")
        Logger.log("lastRowProducto "+lastRowProducto)
        Logger.log("productStartRow"+productStartRow)
        calcularImporteYTotal(lastRowProducto,productStartRow,taxSectionStartRow,hojaActual)
      }

      // CalcularDiasOFecha()
      
      updateTotalProductCounter(lastRowProducto,productStartRow,hojaActual,taxSectionStartRow)

    } else if (hojaActual.getName() === "Clientes") {
      let celdaEditada = e.range;
      let hojaCliente=e.source.getActiveSheet();
      
      let rowEditada = celdaEditada.getRow();
      let colEditada = celdaEditada.getColumn();
      let colTipoDePersona=2
      let tipoPersona= obtenerTipoDePersona(e);

      if (colEditada ==6 && rowEditada>1){
        Logger.log("entro a ver si el edit es en numero")
        let numeroIdentificacion=hojaCliente.getRange(rowEditada,colEditada).getValue()
        Logger.log("num i"+numeroIdentificacion )
        let existe=verificarCodigo(numeroIdentificacion,"Clientes",true,rowEditada)
        if(existe){
          SpreadsheetApp.getUi().alert("El numero de identificacion ya existe, por favor elegir otro numero unico");
          celdaEditada.setValue("");
          verificarDatosObligatorios(e,tipoPersona)
          throw new Error('por favor poner un Numero de Identificacion unico');
        }
      }else if(colEditada ==7 && rowEditada>1){
        let numeroIdentificacion=hojaCliente.getRange(rowEditada,colEditada).getValue()
        let existe=verificarCodigo(numeroIdentificacion,"Clientes",true,rowEditada,"codigo")
        if(existe){
          SpreadsheetApp.getUi().alert("El codigo del cliente ya existe, por favor elegir otro numero unico");
          celdaEditada.setValue("");
          verificarDatosObligatorios(e,tipoPersona)
          throw new Error('por favor poner un Numero de Identificacion unico');
        }
      }

      verificarDatosObligatorios(e,tipoPersona)
      agregarCodigoIdentificador(e)
      validarEmailDeCelda(e)
    

    }else if (hojaActual.getName() === "Historial Facturas"){
      let celdaEditada = e.range;
      let rowEditada = celdaEditada.getRow();
      let colEditada = celdaEditada.getColumn();
      if(rowEditada==5 && colEditada==9){
        Logger.log("dentto de selccionar filtor")
        let valor = celdaEditada.getValue()
        filtroHistorialFacturas(valor)
      }
    }else if (hojaActual.getName() === "Productos"){
      const celdaEditada = e.range;
      const rowEditada = celdaEditada.getRow();
      const colEditada = celdaEditada.getColumn();
      verificarDatosObligatoriosProductos(e);
      agregarCodigoIdentificador(e);

      if (rowEditada <= 1) {
        return;
      }

      if (colEditada === PRODUCT_COLUMNS.CODIGO_REFERENCIA){
        const codigoReferencia = hojaActual.getRange(rowEditada, PRODUCT_COLUMNS.CODIGO_REFERENCIA).getValue();
        const existe = verificarCodigo(codigoReferencia, "Productos", true, rowEditada);
        if (existe){
          SpreadsheetApp.getUi().alert("El Codigo de referencia ya existe, por favor elegir otro numero unico");
          celdaEditada.setValue("");
          verificarDatosObligatoriosProductos(e);
          throw new Error('por favor poner un Numero de Identificacion unico');
        }
      }

      if (colEditada === PRODUCT_COLUMNS.TARIFA_IMPUESTO){
        sincronizarRecargoSegunIva(hojaActual, rowEditada);
        verificarDatosObligatoriosProductos(e);
        agregarCodigoIdentificador(e);
      }

      if (colEditada === PRODUCT_COLUMNS.CHECK_RECARGO){
        manejarCheckboxRecargo(hojaActual, rowEditada);
        verificarDatosObligatoriosProductos(e);
        agregarCodigoIdentificador(e);
      }

      if (colEditada === PRODUCT_COLUMNS.TIPO_RETENCION){
        const esRecargo = hojaActual.getRange(rowEditada, PRODUCT_COLUMNS.CHECK_RECARGO).getValue() === true;
        const valor = hojaActual.getRange(rowEditada, PRODUCT_COLUMNS.TIPO_RETENCION).getDisplayValue().trim();
        if (!esRecargo && valor === RETENCION_RECARGO_LABEL){
          SpreadsheetApp.getUi().alert('Activa la casilla de tarifa recargo para utilizar el tipo "Recargo de equivalencia".');
          celdaEditada.setValue("");
        }
        if (esRecargo && valor !== RETENCION_RECARGO_LABEL){
          hojaActual.getRange(rowEditada, PRODUCT_COLUMNS.TIPO_RETENCION).setValue(RETENCION_RECARGO_LABEL);
        }
        verificarDatosObligatoriosProductos(e);
        agregarCodigoIdentificador(e);
      }

      if (colEditada === PRODUCT_COLUMNS.TARIFA_RETENCION){
        const esRecargo = hojaActual.getRange(rowEditada, PRODUCT_COLUMNS.CHECK_RECARGO).getValue() === true;
        if (esRecargo){
          sincronizarRecargoSegunIva(hojaActual, rowEditada);
        }else{
          const tipoRet = hojaActual.getRange(rowEditada, PRODUCT_COLUMNS.TIPO_RETENCION).getDisplayValue().trim();
          const valor = hojaActual.getRange(rowEditada, PRODUCT_COLUMNS.TARIFA_RETENCION).getDisplayValue();
          const valorNormalizado = valor.replace(/\s/g, '');
          if (tipoRet === 'IRPF' && valor !== ""){
            const permitido = RETENCION_IRPF_TARIFAS
              .map(item => item.replace(/\s/g, ''))
              .includes(valorNormalizado);
            if (!permitido){
              SpreadsheetApp.getUi().alert('Selecciona una tarifa de retención válida (7%, 15% o 19%).');
              celdaEditada.setValue("");
            }
          }
          if (tipoRet === ''){
            celdaEditada.setValue("");
          }
        }
        verificarDatosObligatoriosProductos(e);
        agregarCodigoIdentificador(e);
      }
    }else if(hojaActual.getName() === "Datos de emisor"){
      Logger.log("datos emisor")

    }
  }catch(error){
    Logger.log("No se pudo obtener el lock o hubo error: " + error);
  }finally {
    lock.releaseLock();
  }
}
function validarEmailDeCelda(e) {
  sheet=e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();
  let colEditada = range.getColumn();
  // Nombre de la hoja y celda a leer
  const cellRange = "U"+String(rowEditada);

  // Obtiene la hoja y el valor de la celda

  const value = sheet.getRange(cellRange).getValue();

  // Verifica si el valor está vacío, si es número o si es un string
  if (!value) {
    // Si está vacío, no hace nada específico
    Logger.log("La celda está vacía. No se realizó ninguna acción.");
    return;
  }

  // Si es un número
  if (typeof value === 'number') {
    // Muestra un mensaje o realiza otra acción
    Logger.log("Valor inválido: Se esperaba un email y se encontró un número.");
    return;
  }

  // Si es un string, verificar la estructura de correo
  if (typeof value === 'string') {
    // Expresión regular básica para validar emails
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (emailRegex.test(value)) {
      Logger.log("Correo válido: " + value);
    } else {
      Logger.log("Formato de correo inválido: " + value);
      sheet.getRange(cellRange).setValue("");
      SpreadsheetApp.getUi().alert('Por favor ingresa un email valido');
    }
  }
}

function eliminarProductos() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Factura');

  const productStartRow = 15; 
  let taxSectionStartRow = getTaxSectionStartRow(sh);
  let lastProductRow = getLastProductRow(sh, productStartRow, taxSectionStartRow);

  let totalProductos = lastProductRow - productStartRow + 1;
  if (totalProductos <= 1) {
    SpreadsheetApp.getUi().alert('No puedes eliminar filas cuando solo hay un producto en la factura');
    return;
  }

  const range = sh.getRange(productStartRow, FACTURA_CHECKBOX_COL, totalProductos, 1);
  const values = range.getValues();
  let rowsSeleccionadas = [];

  for (let i = 0; i < values.length; i++) {
    let celda = values[i][0];
    if (celda && typeof celda === 'object' && 'value' in celda) {
      celda = celda.value;
    }
    const normalizado = celda === true || String(celda).toLowerCase() === 'true';
    if (normalizado) {
      rowsSeleccionadas.push(productStartRow + i);
    }
  }

  rowsSeleccionadas = [...new Set(rowsSeleccionadas)];

  const seleccionados = rowsSeleccionadas.length;

  if (seleccionados === 0) {
    SpreadsheetApp.getUi().alert('Selecciona al menos una fila a eliminar (columna Eliminar).');
    return;
  }

  if (seleccionados >= totalProductos) {
    SpreadsheetApp.getUi().alert('Debes dejar al menos un producto en la factura');
    return;
  }

  rowsSeleccionadas.sort((a,b) => b - a);
  rowsSeleccionadas.forEach(row => {
    sh.deleteRow(row);
  });

  taxSectionStartRow = getTaxSectionStartRow(sh);
  lastProductRow = getLastProductRow(sh, productStartRow, taxSectionStartRow);
  calcularImporteYTotal(lastProductRow, productStartRow, taxSectionStartRow, sh);
  updateTotalProductCounter(lastProductRow, productStartRow, sh, taxSectionStartRow);
}

function DesvincularFacturasApp(){
  Logger.log("Desvincular")
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let hojaDatos=spreadsheet.getSheetByName("Datos")
  let estadoVinculacion=hojaDatosEmisor.getRange("B16").getValue();
  if(estadoVinculacion=="Desvinculado"){
    SpreadsheetApp.getUi().alert('Tu estado ya es Desvinculado');
  }else{
  hojaDatosEmisor.getRange("B16").setBackground('#FFC7C7')
  hojaDatosEmisor.getRange("B16").setValue("Desvinculado")
  hojaDatos.getRange("I21").setValue(0)
  SpreadsheetApp.getUi().alert('Haz desvinculado exitosamente facturasApp ');
  }
}

function MensajeErrorDesvincularFacturasApp(){
  Logger.log("Error Desvincular")
  SpreadsheetApp.getUi().alert('Si deseas desvincular facturasApp asegurate de escribir DESVINCULAR en el campo');
  
}

function MensajeErrorEliminarFacturasApp(){
  Logger.log("eliminar mensaje error")
  SpreadsheetApp.getUi().alert('Si deseas eliminar toda la informacion de facturasApp asegurate de escribir ELIMINAR en el campo');
  
}


function eliminarTotalidadInformacion(){
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  let hojaHistorialFactura = spreadsheet.getSheetByName('Historial Facturas Data');
  let hojaProductos = spreadsheet.getSheetByName('Productos');
  let hojaCodigosFatura = spreadsheet.getSheetByName('Facturas ID');
  let hojaClientes=spreadsheet.getSheetByName("Clientes");
  let hojaListadoEstado=spreadsheet.getSheetByName('ListadoEstado');
  let ClientesInvalidos=spreadsheet.getSheetByName('ClientesInvalidos');
  
  limpiarHojaFactura()
  borrarInfoHoja(hojaHistorialFactura)
  borrarInfoHoja(hojaProductos)
  borrarInfoHoja(hojaCodigosFatura)
  borrarInfoHoja(hojaClientes)
  borrarInfoHoja(hojaListadoEstado)
  borrarInfoHoja(ClientesInvalidos)
  borrarInfoHoja(hojaDatosEmisor)
  eliminarCarpetaConDriveAPI()
  hojaDatosEmisor.getRange("B16").setBackground('#FFC7C7')
  hojaDatosEmisor.getRange("B16").setValue("Desvinculado")
  SpreadsheetApp.getUi().alert('Informacion eliminada correctamente');


  //falta borrar carpeta esa con pdfs
}

function borrarInfoHoja(hoja){
  let lastrow=Number(hoja.getLastRow())
  let nombreHoja=hoja.getSheetName()
  Logger.log("borrarInfoHoja")
  Logger.log("nombreHoja "+nombreHoja)
  Logger.log("lastrow "+lastrow)
  if (nombreHoja==="Datos de emisor" ){
    Logger.log("Hoja es datos emisor ")
    hoja.getRange(1,2,16).setValue("")
  }else{
    Logger.log("else")
    hoja.deleteRows(2,lastrow)
    let maxRows=hoja.getMaxRows()
    Logger.log("maxRows "+maxRows)
    let dif = 1000-maxRows
    Logger.log("dif "+dif)
    hoja.insertRows(maxRows,dif)

    // for(let j=2;j<=lastrow;j++){
    //   Logger.log("j "+j)


    // }
  }
}

function mensajeBorrarInfoError(){
  Logger.log("Error borrar info")
  SpreadsheetApp.getUi().alert('Si deseas eliminar toda la informacion de facturasApp asegurate de escribir ELIMINAR en el campo');
}


function agregarCodigoIdentificador(e){
  hoja=e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();
  let colEditada = range.getColumn();
  let estadoActual=hoja.getRange(rowEditada,1).getValue()
  Logger.log("entrado a codigo-identiicador")
  Logger.log("estado actual "+estadoActual)
  if(hoja.getName()=="Clientes"){
    let tipoPersona=obtenerTipoDePersona(e)
    if (estadoActual=="Valido"){
      let nombre=""
      // Normalizar: Persona Física se trata como Autonomo
      let tipoNorm = String(tipoPersona).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
      if(tipoNorm==="autonomo" || tipoNorm==="persona fisica"){
        let primerNombre=hoja.getRange(rowEditada,10).getValue()
        let apellido=hoja.getRange(rowEditada,12).getValue()
        nombre =primerNombre+" "+apellido
      }else{
        nombre=hoja.getRange(rowEditada,9).getValue()
      }
      
      let numeroIdentificacion=hoja.getRange(rowEditada,6).getValue()
      let identificadorUnico=nombre+"-"+numeroIdentificacion
      hoja.getRange(rowEditada,2).setValue(identificadorUnico)
    }
  }else if (hoja.getName()=="Productos"){
    if (estadoActual=="Valido"){
      let nombre=hoja.getRange(rowEditada,PRODUCT_COLUMNS.NOMBRE).getValue()
      let numeroIdentificacion=hoja.getRange(rowEditada,PRODUCT_COLUMNS.CODIGO_REFERENCIA).getValue()
      let identificadorUnico=nombre+"-"+numeroIdentificacion
      hoja.getRange(rowEditada,PRODUCT_COLUMNS.IDENTIFICADOR_UNICO).setValue(identificadorUnico)
    } else {
      hoja.getRange(rowEditada,PRODUCT_COLUMNS.IDENTIFICADOR_UNICO).clearContent();
    }
  }
}

function verificarDescuentoValido(valorFechaActual,ivaProductoActual){
  //1julio 2022 hasta 30 de junio 2024. 
  Logger.log("ivaProductoActual"+ivaProductoActual)
  Logger.log("valorFechaActual"+valorFechaActual)
  Logger.log("Entra a verificar fecha")
  var fechaInicio = new Date(2022, 6, 1);  // 1 de julio de 2022 (mes 6 porque enero es 0)
  var fechaFin = new Date(2024, 5, 30);    // 30 de junio de 2024 (mes 5 porque enero es 0)
  var partesFecha = valorFechaActual.split("/");
  var dia = parseInt(partesFecha[0]);
  var mes = parseInt(partesFecha[1]) - 1; // Restar 1 porque los meses en Date empiezan desde 0
  var anio = parseInt(partesFecha[2]);
  var fechaActual = new Date(anio, mes, dia);

  if(ivaProductoActual===0.05){
    Logger.log("fecha es igual a 5%")
    if (fechaActual >= fechaInicio && fechaActual <= fechaFin) {
      Logger.log("Fecha dentro del rango válido");
      return true;
    } else {
      Logger.log("Fecha fuera del rango válido");
      return false;
    }
  }else{
    Logger.log("No hay producto con 5% interes")
    return true
  }
}

function calcularImporteYTotal(lastRowProducto,productStartRow,taxSectionStartRow,hojaActual) {
  Logger.log("Entra a calcular importe")
  Logger.log("lastRowProducto "+lastRowProducto)
  Logger.log("productStartRow "+productStartRow )
  Logger.log("taxSectionStartRow"+taxSectionStartRow)

  //base Imponible
  let rowParaFormulaBaseImponible=taxSectionStartRow+1
  let rowEspacioIvasAgrupacion=taxSectionStartRow+5
  let rowTotalBaseImponibleEIvaGeneral=taxSectionStartRow+7
  hojaActual.getRange("A"+String(rowParaFormulaBaseImponible)).setValue("=ARRAYFORMULA(SUMIF(G15:G"+String(lastRowProducto)+"; B"+String(rowParaFormulaBaseImponible)+":B"+String(rowEspacioIvasAgrupacion)+"; F15:F"+String(lastRowProducto)+"))")
  //BASE imponilbre recargo
  hojaActual.getRange("E"+String(rowParaFormulaBaseImponible)).setValue("=ARRAYFORMULA(SUMIF(I15:I"+String(lastRowProducto)+"; F"+String(rowParaFormulaBaseImponible)+":F"+String(rowEspacioIvasAgrupacion)+"; F15:F"+String(lastRowProducto)+"))")

    //total base imponible e iva genberal
    hojaActual.getRange("A"+String(rowTotalBaseImponibleEIvaGeneral)).setValue("=SUM(A"+String(rowParaFormulaBaseImponible)+":A"+String(rowEspacioIvasAgrupacion)+")")
    hojaActual.getRange("C"+String(rowTotalBaseImponibleEIvaGeneral)).setValue("=SUM(C"+String(rowParaFormulaBaseImponible)+":C"+String(rowEspacioIvasAgrupacion)+")")
    
    //total base imponible y recargo
    hojaActual.getRange("E"+String(rowTotalBaseImponibleEIvaGeneral)).setValue("=SUM(E"+String(rowParaFormulaBaseImponible)+":E"+String(rowEspacioIvasAgrupacion)+")")
    hojaActual.getRange("G"+String(rowTotalBaseImponibleEIvaGeneral)).setValue("=SUM(G"+String(rowParaFormulaBaseImponible)+":G"+String(rowEspacioIvasAgrupacion)+")")

  //IVA%
  hojaActual.getRange("B"+String(rowParaFormulaBaseImponible)).setValue("=UNIQUE(G15:G"+String(lastRowProducto)+")")
  //RECARGO %
  hojaActual.getRange("F"+String(rowParaFormulaBaseImponible)).setValue("=UNIQUE(I15:I"+String(lastRowProducto)+")")


  let rowParaTotales=taxSectionStartRow+10
  //total retenciones
  const irpfCellAbs = "$F$" + String(taxSectionStartRow - 2);
  const totalBaseCellAbs = "$A$" + String(rowTotalBaseImponibleEIvaGeneral);
  hojaActual.getRange("A"+String(rowParaTotales)).setValue("=IF("+irpfCellAbs+"=\"\";0;"+totalBaseCellAbs+"*"+irpfCellAbs+")")

  //total cargo equivalencia
  hojaActual.getRange("B"+String(rowParaTotales)).setValue("=SUMPRODUCT(F15:F"+String(lastRowProducto)+";I15:I"+String(lastRowProducto)+")*10")

  //total descuentos FACTURA
  let rowDescuentos=taxSectionStartRow-1
  hojaActual.getRange("D"+String(rowParaTotales)).setValue("=B"+String(rowDescuentos)+"+(SUMPRODUCT(D15:D"+String(lastRowProducto)+";C15:C"+String(lastRowProducto)+";H15:H"+String(lastRowProducto)+"))")

  //netopagar
  let rowParaTotalFactura=taxSectionStartRow+12
  //hojaActual.getRange("B"+String(rowParaTotalFactura)).setValue("=SUM(J15:J"+String(lastRowProducto)+")+C"+String(rowParaTotales)+"-B"+String(rowDescuentos))

  //valorBruto
  hojaActual.getRange("E"+String(rowParaTotalFactura)).setValue("=SUMPRODUCT(C15:C"+String(lastRowProducto)+";D15:D"+String(lastRowProducto)+")")

}

function getLastProductRow(sheet, productStartRow, taxSectionStartRow) {

  Logger.log("funcion getLastProductRow")
  //retorna el numero de fila exacta donde esta el ulitmo producto agregado
  // si no encuntra producto agg si solo tiene un producto retorna el mismo productStartRow 
  let lastProductRow = productStartRow;
  
  for (let row = productStartRow; row < taxSectionStartRow; row++) {
    
    let valorCeldaActual=sheet.getRange(row, 1).getValue() 
    Logger.log("'Valor celda "+valorCeldaActual)

      if(valorCeldaActual==="Total filas"){
        Logger.log("dentro de if+ lastProductRow"+lastProductRow)
        return lastProductRow
      }else{
        lastProductRow = row;
      }
      Logger.log("lastProductRow "+lastProductRow)
      
    
  }
  Logger.log("No dentro del if lastProductRow"+lastProductRow)
  //aqui arrelgar error que se agrega una nueva linea cuando hay espacio arriba
  return lastProductRow;
}

function getTaxSectionStartRow(sheet) {
  //obtiene la row donde esta la seccion de taxinformation osea Base imponible
  const lastRow = sheet.getLastRow();
  let row = 14
  
  for (row; row < lastRow; row++) { // 14 por si esta vacio, pero deberia de dar igual si es desde la 15
    if (sheet.getRange(row, 1).getValue() === 'Base imponible') {

      Logger.log("dentro de getTax row " + row)
      return row;
    }
  }

  
  return row+1;// por si se borro todos los productos,creo que da igual 
}

function updateTotalProductCounter(lastRowProducto,productStartRow,hojaActual,taxSectionStartRow) {
  let totalProducts = 0;
  Logger.log(" dentro updateTotalProductCounter")

  for(let i=productStartRow;i<=lastRowProducto;i++){
    Logger.log("I"+i)
    Logger.log("lastRowProducto"+lastRowProducto)
    if(hojaActual.getRange("B"+String(i)).getValue()!=""){
      totalProducts++
    }
  }

  let rowTotalProductos=taxSectionStartRow-3
  hojaActual.getRange("B"+String(rowTotalProductos)).setValue(totalProducts)

}


function limpiarDict() {
  Logger.log("Limpiar el dict")
  diccionarioCaluclarIva = {
    "0.21": 0,
    "0.1": 0,
    "0.05": 0,
    "0.04": 0,
    "0": 0
  }
}

function slugifyF(str) {
  var map = {
    '-': ' ',
    '-': '_',
    'a': 'á|à|ã|â|À|Á|Ã|Â',
    'e': 'é|è|ê|É|È|Ê',
    'i': 'í|ì|î|Í|Ì|Î',
    'o': 'ó|ò|ô|õ|Ó|Ò|Ô|Õ',
    'u': 'ú|ù|û|ü|Ú|Ù|Û|Ü',
    'c': 'ç|Ç',
    'n': 'ñ|Ñ'
  };

  str = String(str)
  str = str.toLowerCase();

  for (var pattern in map) {
    str = str.replace(new RegExp(map[pattern], 'g'), pattern);
  };

  return str;
};
function getAdditionalDocuments() {
  //Browser.msgBox('getAddtionionalDocuments');
  var AdditionalDocuments = {
    "OrderReference": "",
    "DespatchDocumentReference": "",
    "ReceiptDocumentReference": "",
    "AdditionalDocument": []
  }
  return AdditionalDocuments;
}

var centenas = ['', 'Ciento ', 'Doscientos ', 'Trescientos ', 'Cuatrocientos ', 'Quinientos ', 'Seiscientos ',
  'Setecientos ', 'Ochocientos ', 'Novecientos ']

var decenas1 = ['Diez ', 'Once ', 'Doce ', 'Trece ', 'Catorce ', 'Quince ', 'Dieciseis ', 'Diecisiete ',
  'Dieciocho ', 'Diecinueve ']

var decenas2 = ['', 'Diez ', 'Veinte ', 'Treinta ', 'Cuarenta ', 'Cincuenta ', 'Sesenta ', 'Setenta ', 'Ochenta ', 'Noventa ']
var unidades = ['', 'Un ', 'Dos ', 'Tres ', 'Cuatro ', 'Cinco ', 'Seis ', 'Siete ', 'Ocho ', 'Nueve ']

function getPaymentMeans(PaymentMeansTxt) {
  switch (PaymentMeansTxt) {
    case 'Instrumento no definido':
      var PaymentMeans = 1;
      break;
    case 'Crédito ACH':
      var PaymentMeans = 2;
      break;
    case 'Débito ACH':
      var PaymentMeans = 3;
      break;
    case 'Reversión débito de demanda ACH':
      var PaymentMeans = 4;
      break;
    case 'Reversión crédito de demanda ACH':
      var PaymentMeans = 5;
      break;
    case 'Crédito de demanda ACH':
      var PaymentMeans = 6;
      break;
    case 'Débito de demanda ACH':
      var PaymentMeans = 7;
      break;
    case 'Mantener':
      var PaymentMeans = 8;
      break;
    case 'Clearing Nacional o Regional':
      var PaymentMeans = 9;
    case 'Efectivo':
      var PaymentMeans = 10;
      break;
    case 'Reversión Crédito Ahorro':
      var PaymentMeans = 11;
      break;
    case 'Reversión Débito Ahorro':
      var PaymentMeans = 12;
      break;
    case 'Crédito Ahorro':
      var PaymentMeans = 13;
      break;
    case 'Débito Ahorro':
      var PaymentMeans = 14;
      break;
    case 'Bookentry Crédito':
      var PaymentMeans = 15;
      break;
    case 'Bookentry Débito':
      var PaymentMeans = 16;
      break;
    case 'Concentración de la demanda en efectivo/Desembolso Crédito (CCD)':
      var PaymentMeans = 17;
      break;
    case 'Concentración de la demanda en efectivo/Desembolso (CCD) débito':
      var PaymentMeans = 18;
      break;
    case 'Crédito Pago negocio corporativo (CTP)':
      var PaymentMeans = 19;
      break;
    case 'Cheque':
      var PaymentMeans = 20;
      break;
    case 'Proyecto bancario':
      var PaymentMeans = 21;
      break;
    case 'Proyecto bancario certificado':
      var PaymentMeans = 22;
      break;
    case 'Cheque bancario':
      var PaymentMeans = 23;
      break;
    case 'Nota cambiaria esperando aceptación':
      var PaymentMeans = 24;
      break;
    case 'Cheque certificado':
      var PaymentMeans = 25;
      break;
    case 'Cheque Local':
      var PaymentMeans = 26;
      break;
    case 'Débito Pago Neogcio Corporativo (CTP)':
      var PaymentMeans = 27;
      break;
    case 'Crédito Negocio Intercambio Corporativo (CTX)':
      var PaymentMeans = 28;
      break;
    case 'Débito Negocio Intercambio Corporativo (CTX)':
      var PaymentMeans = 29;
      break;
    case 'Transferecia Crédito':
      var PaymentMeans = 30;
      break;
    case 'Transferencia Débito':
      var PaymentMeans = 31;
      break;
    case 'Concentración Efectivo/Desembolso Crédito plus (CCD+)':
      var PaymentMeans = 32;
      break;
    case 'Concentración Efectivo/Desembolso Débito plus (CCD+)':
      var PaymentMeans = 33;
      break;
    case 'Pago y depósito pre acordado (PPD)':
      var PaymentMeans = 34;
      break;
    case 'Concentración efectivo ahorros/Desembolso Crédito (CCD)':
      var PaymentMeans = 35;
      break;
    case 'Concentración efectivo ahorros / Desembolso Drédito (CCD)':
      var PaymentMeans = 36;
      break;
    case 'Pago Negocio Corporativo Ahorros Crédito (CTP)':
      var PaymentMeans = 37;
      break;
    case 'Pago Neogcio Corporativo Ahorros Débito (CTP)':
      var PaymentMeans = 38;
    default:
      Logger.log("Error: PaymentMeans");
      var PaymentMeans = 100
  }
  return PaymentMeans;

}



function getPaymentType(PaymentTypeTxt) {
  switch (PaymentTypeTxt) {
    case 'Contado':
      var PaymentType = 1;
      break;
    case 'Credito':
      var PaymentType = 2;
      break;
    default:
      Logger.log('Error: getPaymentType');
      Browser.msgBox("Oops! PaymentType");
  }
  return PaymentType;
}

function unos(n) {
  if (n == 0) {
    return '';
  }
  else {
    return unidades[n];
  }
}

function dieces(n) {
  var decena = Math.floor(n / 10);
  var unidad = n % 10;
  switch (true) {
    case ((n % 10) == 0):
      return (decenas2[n / 10]);
    case ((11 <= n) && (n <= 19)):
      return (decenas1[(n % 10)]);
    case (Math.floor(n / 10) == 2):
      return `Veinti${unos(unidad).toLowerCase()}`;
    case (0 <= n && n < 10):
      return (unos(n % 10));
    default:
      var letras = `${decenas2[decena]} y ${unos(unidad)}`;
      return (letras);
  }
}

function cienes(n) {
  if (n == 100) {
    return 'Cien ';
  }
  if (n < 100) {
    return dieces(n);
  }
  else {
    return (centenas[Math.floor(n / 100)] + dieces(n % 100));
  }
}

function int2word(n) {
  var euros = Math.floor(n);
  var centimos = Math.round((n - euros) * 100);

  var megas = Math.floor(euros / 1000 / 1000);
  var kilos = Math.floor((euros - megas * 1000000) / 1000);
  var ones = euros - megas * 1000000 - kilos * 1000;

  var letras = '';
  if (megas >= 1) {
    if (megas == 1) {
      letras = letras + 'Un Millón ';
    } else {
      letras = letras + cienes(megas) + ' Millones ';
    }
  }
  if (kilos >= 1) {
    if (kilos == 1) {
      letras = letras + 'Mil ';
    } else {
      letras = letras + cienes(kilos) + 'Mil ';
    }
  }

  if (ones >= 1) {
    if (ones == 1) {
      letras = letras + 'Un ';
    } else {
      letras = letras + cienes(ones);
    }
  }

  if (centimos > 0) {
    letras = letras + 'Euros' + ` Con ${cienes(centimos)}Céntimos`;
  }else{
    letras = letras + 'Euros';
  }

  return letras;
}

function getAdditionalProperty() {
  var AdditionalProperty = [];
  return AdditionalProperty;
}

function getdatosValueA1(datos_sheet,range) {
  var range = datos_sheet.getRange(range);
  return range.getValue();
}

function getDelivery() {
  var row = getdatosValueA1("C50");

  var Delivery = {
    "AddressLine": "",//getdatosValueA1("B61"),
    "CountryCode": "",//"CO",
    "CountryName": "",//"Colombia",
    "SubdivisionCode": "",//getdatosValueA1("D61"),//Departamento Codigo
    "SubdivisionName": "",//getdatosValueA1("G61"),///Departamento Nombre
    "CityCode": "",//getdatosValueA1("E61"),//Codigo Municipio
    "CityName": "",//getdatosValueA1("F61"),//Nombre Municip
    "ContactPerson": "",
    "DeliveryDate": "",
    "DeliveryCompany": ""
  };
  return Delivery;

}

function getMeasureUnitCode(measureName) {
  var range = unidades_sheet.getRange("E1");

  var formula = `=DGET($A$1:$B$1104,A$1,{"Descripcion";"=${measureName}"})`;
  range.setValue(formula);

  return range.getValue();
}


function verificarTipoDeDatos(e) {
  /*Funcion que verificar que celda o grupo de celdas editada
y verifica su valor para saber si es valido 
Input: e objeto que actua como una instancia del sheet editado 
Output: no tiene output pero regresa un mensaje en caso de que sea erroneo el tipo de dato*/

  let sheet = e.range.getSheet();

  if (sheet.getName() === "Clientes") {//aca filtro de hoja, por cada hoja verifica cosas distintas
    let numIdentificacion = sheet.getRange("D2:D1000");
    let codigoContacto = sheet.getRange("E2:E1000");
    let nomberComercial=sheet.getRange("G2:G1000");
    let primerNombre = sheet.getRange("H2:H1000");
    let segundoNombre = sheet.getRange("I2:I1000");
    let primeraApellido = sheet.getRange("J2:J1000");
    let segundoApellido = sheet.getRange("K2:K1000");
    let pais = sheet.getRange("l2:l1000");
    let provincia = sheet.getRange("M2:M1000");
    let poblacion = sheet.getRange("N2:N1000");
    let direccion = sheet.getRange("O2:O1000");
    let codigoPostal = sheet.getRange("P2:P1000");
    let telefono = sheet.getRange("Q2:Q1000");
    let sitioWeb = sheet.getRange("R2:R1000");
    let email = sheet.getRange("S2:S1000");
    let editedCell = e.range;

    esCeldaEnRango(numIdentificacion, editedCell, undefined, e);
    esCeldaEnRango(nomberComercial,editedCell,"string",e)
    esCeldaEnRango(codigoContacto, editedCell, undefined, e);
    esCeldaEnRango(primerNombre, editedCell, "string", e);
    esCeldaEnRango(segundoNombre, editedCell, "string", e);
    esCeldaEnRango(primeraApellido, editedCell, "string", e);
    esCeldaEnRango(segundoApellido, editedCell, "string", e);
    esCeldaEnRango(pais, editedCell, "string", e)
    esCeldaEnRango(provincia, editedCell, "string", e)
    esCeldaEnRango(poblacion, editedCell, "string", e)
    esCeldaEnRango(direccion, editedCell, "string", e)
    esCeldaEnRango(codigoPostal, editedCell, undefined, e);
    esCeldaEnRango(telefono, editedCell, undefined, e);
    esCeldaEnRango(sitioWeb, editedCell, "string", e)
    esCeldaEnRango(email, editedCell, "string", e)
  }
}

function esCeldaEnRango(range, editedCell, tipoDato = 'number', e) {
  if (editedCell.getRow() >= range.getRow() &&
    editedCell.getRow() <= range.getLastRow() &&
    editedCell.getColumn() >= range.getColumn() &&
    editedCell.getColumn() <= range.getLastColumn()) {
    let value = e.value;
    if (typeof value === "undefined") {// no funciona value===null || value ==="null" || value ===''
      Logger.log("Ingreso algo vacio")
    } else {
      let newValue = convertirANumero(value);
      if (typeof newValue !== tipoDato) {
        SpreadsheetApp.getUi().alert("Error: Solo se permite " + tipoDato + " en este rango");
        e.range.setValue("");
      } else {
        Logger.log("Ingreso el tipo de valor correcto")
      }
    }
  }
}

function convertirANumero(value) {

  let number = Number(value);
  if (!isNaN(number)) {
    return number;
  } else {
    return value;
  }

}

function getsheetValueA1(sheet, column, row) {
  var cell = column + row;
  var range = sheet.getRange(cell);
  return range.getValue();
}


function getsheetValue(sheet, column, row) {
  var range = sheet.getRange(column, row);
  return range.getValue();
}

function updatesheetValueA1(sheet, column, row, value) {
  var cell = column + row;
  var range = sheet.getRange(cell);
  range.setValue(value);
  return
}

function aumentarConsecutivo(Consecutivo){

}

function verificarConsecutivo(entrada, isNumero) {
  let regex;

  if (entrada === "" || entrada === "0") {
    return false;
  }

  if (isNumero) {
    // Uno o más dígitos
    regex = /^\d+$/;
  } else {
    // Cualquier cadena con al menos un caracter (permite letras, números y símbolos)
    // Evitamos que sea solo dígitos para distinguir de la parte numérica
    regex = /^(?=.*\D).+$/;
  }

  return regex.test(entrada);
}

function guardarConsecutivo(){
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');
  // Usar getDisplayValue para preservar ceros a la izquierda
  let letra = String(hojaDatosEmisor.getRange(24,1).getDisplayValue() || '').trim();
  let numero = String(hojaDatosEmisor.getRange(24,3).getDisplayValue() || '').trim();
  const scriptProperties = PropertiesService.getDocumentProperties();
  if(verificarConsecutivo(letra,false)){
    Logger.log("letra valida")
    if(verificarConsecutivo(numero,true)){
      Logger.log("numero valido")
      Logger.log("numero "+numero)
      Logger.log("letra "+letra)
      // Guardamos valores antiguos por compatibilidad
      scriptProperties.setProperties({
        'NumeroConescutivo': String(numero),
        'LetraConescutivo': String(letra)
      });
      // También guardamos una plantilla unificada y la longitud del sufijo numérico
      const plantilla = String(letra);
      const longitudNumerica = String(numero).length;
      scriptProperties.setProperties({
        'ConsecutivoPlantillaPrefijo': plantilla,
        'ConsecutivoPlantillaDigitos': String(longitudNumerica)
      });
      SpreadsheetApp.getUi().alert('Consecutivo válido y guardado');
    }else{
      SpreadsheetApp.getUi().alert('Por favor ingresa un consecutivo válido');
    }

  }else{
    SpreadsheetApp.getUi().alert('Por favor ingresa un consecutivo válido');
  }

  
}

function cumpleEstructura(str) {
  let numero;
  let letra;
  
  try {
    const scriptProperties = PropertiesService.getDocumentProperties();
    // Compatibilidad con propiedades antiguas
    numero = scriptProperties.getProperty('NumeroConescutivo');  // Ej: "123"
    letra  = scriptProperties.getProperty('LetraConescutivo');   // Ej: "abc"

    // Nuevas propiedades opcionales
    const nuevoPrefijo = scriptProperties.getProperty('ConsecutivoPlantillaPrefijo');
    const nuevoDigitos = scriptProperties.getProperty('ConsecutivoPlantillaDigitos');

    if (nuevoPrefijo && nuevoDigitos) {
      letra = nuevoPrefijo;
      numero = new Array(Number(nuevoDigitos)).fill('0').join('');
    }
  } catch (err) {
    Logger.log('Error leyendo propiedades: %s', err.message);
    return false;            // Maneja el error según tu caso
  }

  /* --- NUEVA COMPROBACIÓN --- */
  // Si número o letra son "0", vacíos o nulos → devolver 0
  if (!numero || numero === '0' || !letra || letra === '0') {
    return 0;
  }
  /* -------------------------- */

  // Calculamos la longitud de "numero", que será la cantidad de dígitos esperados
  const lengthNumeros = String(numero).length;

  // Escapamos el prefijo para expresión regular
  const prefijoEscapado = String(letra).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

  // Cualquier prefijo (con letras, números o símbolos) seguido de lengthNumeros dígitos
  const regex = new RegExp(`^${prefijoEscapado}\\d{${lengthNumeros}}$`);

  return regex.test(String(str));
}


function abrirLinkSoporte(){
  var url = 'https://soporte.facturasapp.com/hc/es-es/articles/34589820237979-Evite-problemas-de-cuenta-en-FacturasApp-for-Google-Sheets-creando-un-perfil-dedicado-en-su-navegador';
  var html = "<script>window.open('"+url+"');google.script.host.close();</script>";
  var userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Soporte');
}


function cambiarAmbienete(){
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaDatosEmisor = spreadsheet.getSheetByName('Datos de emisor');

  Logger.log("cambiar ambiente")
  let ui = SpreadsheetApp.getUi();
  let respuesta = ui.alert('Estas seguro de que quieres cambiar el ambiente?, tendras que volver a inicar sesion', ui.ButtonSet.YES_NO);
  if (respuesta == ui.Button.YES){
    //DesvincularFacturasApp()
    let AmbienteActual=hojaDatosEmisor.getRange("C1002").getValue()
    if(AmbienteActual=="Produccion"){
      AmbienteActual="Pruebas"
    }else{
      AmbienteActual="Produccion"
    }
    Logger.log("Ambiente actual "+AmbienteActual)
    const scriptProps = PropertiesService.getDocumentProperties();
    scriptProps.setProperties({
      'Ambiente': AmbienteActual
    });
    showVincularCuenta()
    DesvincularFacturasApp()
    hojaDatosEmisor.getRange("C1002").setValue(AmbienteActual)

  }else{
    ui.alert('No se ha cambiado el ambiente');
  }
}

function showModoFacturacion() {
  var html = HtmlService.createHtmlOutputFromFile('modoFacturacion')
    .setTitle('Modo de facturación')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}
