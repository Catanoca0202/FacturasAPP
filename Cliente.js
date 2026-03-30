// var datos_sheet = spreadsheet.getSheetByName('Datos');
// var spreadsheet = SpreadsheetApp.getActive();
// var factura_sheet= spreadsheet.getSheetByName("Factura")

function showNuevaClienteDesdeFactura() {
  renderSidebarFromFile('menuAgregarClienteDesdeF', 'Nuevo Cliente');
}

function showNuevaProductoDesdeFactura(){
  renderSidebarFromFile('agregarProductoDesdeF', 'Nuevo Producto');
}

function showNuevaClienteV2() {
  renderSidebarFromFile('menuAgregarCliente', 'Nuevo Cliente');
}

function showInactivarCliente() {
  renderSidebarFromFile('menuInactivarCliente', 'Inactivar Cliente');
}

function showActivarCliente() {
  renderSidebarFromFile('menuActivarCliente', 'Activar Cliente');
}

function inactivarCliente(cliente){
  let spreadsheet = SpreadsheetApp.getActive();
  let hojaClientesInactivos=spreadsheet.getSheetByName('ClientesInvalidos');
  let hojaClietnes=spreadsheet.getSheetByName("Clientes")
  let datos_sheet = spreadsheet.getSheetByName('Datos');
  Logger.log(cliente)
  datos_sheet.getRange("H2").setValue(cliente)
  

  let rowDelCliente=datos_sheet.getRange("G2").getValue();
  let rowMaximaClientesInactivos=hojaClientesInactivos.getLastRow()+1;
  let rowMaximaClientes=hojaClietnes.getLastRow()+1;

  let tipoContacto=datos_sheet.getRange("AB2").getValue();
  let tipoPersona=datos_sheet.getRange("L2").getValue();
  let tipoDoc=datos_sheet.getRange("J2").getValue();
  let numIdentificacion=datos_sheet.getRange("K2").getValue();
  let codigoContacto=datos_sheet.getRange("I2").getValue();
  let nomnbreComercial=datos_sheet.getRange("N2").getValue();
  let primerNombre=datos_sheet.getRange("O2").getValue();
  let segundoNombre=datos_sheet.getRange("P2").getValue();
  let primerApellido=datos_sheet.getRange("Q2").getValue();
  let segundoApellido=datos_sheet.getRange("R2").getValue();
  let pais=datos_sheet.getRange("S2").getValue();
  let provicnica=datos_sheet.getRange("AA2").getValue();
  let poblacion=datos_sheet.getRange("Z2").getValue();
  let direccion=datos_sheet.getRange("T2").getValue();
  let codigoPostal=datos_sheet.getRange("U2").getValue();
  let telefono=datos_sheet.getRange("V2").getValue();
  let sitioWeb=datos_sheet.getRange("W2").getValue();
  let email=datos_sheet.getRange("X2").getValue();
  let estado=datos_sheet.getRange("Y2").getValue();
  let nombreOriginal=datos_sheet.getRange("AC2").getValue();


  // Proceso para agregar a la hoja de clientes inactivos
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 1).setValue(estado);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 2).setValue(cliente);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 3).setValue(tipoContacto);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 4).setValue(tipoPersona);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 5).setValue(tipoDoc);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 6).setValue(numIdentificacion);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 7).setValue(codigoContacto);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 8).setValue(nomnbreComercial);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 9).setValue(primerNombre);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 10).setValue(segundoNombre);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 11).setValue(primerApellido);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 12).setValue(segundoApellido);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 13).setValue(pais);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 14).setValue(provicnica);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 15).setValue(poblacion);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 16).setValue(direccion);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 17).setValue(codigoPostal);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 18).setValue(telefono);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 19).setValue(sitioWeb);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 20).setValue(email);
  hojaClientesInactivos.getRange(rowMaximaClientesInactivos, 21).setValue(cliente);

  //eliminar cliente de la hoja clientes

  hojaClietnes.deleteRow(rowDelCliente)
  hojaClietnes.insertRowAfter(rowMaximaClientes)
}

function activarCliente(cliente) {
  let spreadsheet = SpreadsheetApp.getActive();
  let datos_sheet = spreadsheet.getSheetByName('Datos');
  let hojaClientesInactivos = spreadsheet.getSheetByName('ClientesInvalidos');
  let hojaClietnes = spreadsheet.getSheetByName("Clientes");
  Logger.log(cliente);

  datos_sheet.getRange("I6").setValue(cliente);
  let rowDelCliente = datos_sheet.getRange("G6").getValue();
  let rowMaximaClientesInactivos = hojaClientesInactivos.getLastRow() + 1;
  let rowMaximaClientes = hojaClietnes.getLastRow() + 1;

  // Obtener los valores necesarios desde la hoja 'Datos'
  let estado = datos_sheet.getRange('H6').getValue();
  let tipoPersona = datos_sheet.getRange('K6').getValue(); // Determina si es 'Autonomo' o 'Empresa'
  let values = [
    estado,
    cliente, // nombreOriginal
    datos_sheet.getRange('J6').getValue(), // tipoContacto
    tipoPersona,
    datos_sheet.getRange('L6').getValue(), // tipoDoc
    datos_sheet.getRange('M6').getValue(), // numIdentificacion
    datos_sheet.getRange('N6').getValue(), // codigoContacto
    datos_sheet.getRange('O6').getValue(), // nombreComercial
    datos_sheet.getRange('Q6').getValue(), // primerNombre
    datos_sheet.getRange('R6').getValue(), // segundoNombre
    datos_sheet.getRange('S6').getValue(), // primerApellido
    datos_sheet.getRange('T6').getValue(), // segundoApellido
    datos_sheet.getRange('U6').getValue(), // pais
    datos_sheet.getRange('V6').getValue(), // provincia
    datos_sheet.getRange('W6').getValue(), // poblacion
    datos_sheet.getRange('X6').getValue(), // direccion
    datos_sheet.getRange('Y6').getValue(), // codigoPostal
    datos_sheet.getRange('Z6').getValue(), // telefono
    datos_sheet.getRange('AA6').getValue(), // sitioWeb
    datos_sheet.getRange('AB6').getValue(), // email
    
  ];

  // Agregar cliente a la hoja 'Clientes'
  hojaClietnes.getRange(rowMaximaClientes, 1, 1, values.length).setValues([values]);

  // Verificar datos obligatorios después de agregar el cliente
  verificarDatosObligatoriosManual(hojaClietnes, rowMaximaClientes, tipoPersona);

  // Eliminar el cliente de la hoja 'ClientesInvalidos'
  hojaClientesInactivos.deleteRow(rowDelCliente);
  hojaClientesInactivos.insertRowAfter(rowMaximaClientesInactivos);
}

function verificarDatosObligatoriosManual(sheet, row, tipoPersona) {
  // Tratar "Persona Física" igual que "Autónomo" para obligaciones
  const esAutonomo =
    tipoPersona === "Autónomo" ||
    tipoPersona === "Persona Física";

  const columnasObligatorias = esAutonomo ? 
    [2, 3, 4, 5, 6, 7, 9, 11, 13, 14, 15, 17, 20] : // Para autónomos
    [2, 3, 4, 5, 6, 7, 8, 13, 14, 15, 17, 20]; // Para empresas

  const todasLasColumnas = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20];
  const estadosDefault = ["", "Tipo Documento", "Tipo de persona"];
  let estaCompleto = true;
  let estaVacioOPredeterminado = true;

  // Borrar colores de fondo antes de verificar
  todasLasColumnas.forEach(col => {
    sheet.getRange(row, col).setBackground(null);
  });

  // Verificar cada columna obligatoria
  columnasObligatorias.forEach(col => {
    const valorDeCelda = sheet.getRange(row, col).getValue();
    if (estadosDefault.includes(valorDeCelda)) {
      estaCompleto = false;
      sheet.getRange(row, col).setBackground('#FFC7C7'); // Resaltar en rojo claro
    } else {
      estaVacioOPredeterminado = false;
    }
  });

  // Actualizar estado en la primera columna
  if (estaVacioOPredeterminado) {
    sheet.getRange(row, 1).clearContent(); // Limpiar el estado
  } else {
    const status = estaCompleto ? "Valido" : "No Valido";
    sheet.getRange(row, 1).setValue(status);
  }
}


function buscarClientes(terminoBusqueda,hojaA) {
  let spreadsheet = SpreadsheetApp.getActive();
  var resultados = [];

  if(hojaA==="Inactivar"){
    var sheet = spreadsheet.getSheetByName('Clientes');
  }else{

    var sheet = spreadsheet.getSheetByName('ClientesInvalidos');
    var ultimaFila = sheet.getLastRow(); 
    var valores = sheet.getRange(2, 2, ultimaFila - 1, 1).getValues();

    for (var i = 0; i < valores.length; i++) {
      var valor = valores[i][0]; // Accede al primer (y único) valor de cada fila
      resultados.push(valor);}
      
    return resultados
}
  
  var ultimaFila = sheet.getLastRow(); 
  var valores = sheet.getRange(2, 2, ultimaFila - 1, 1).getValues(); // `ultimaFila - 1` porque empieza en la fila 2


  if(terminoBusqueda===""){
    return resultados
  }
  // Recorre los valores obtenidos
  for (var i = 0; i < valores.length; i++) {
    var valor = valores[i][0]; // Accede al primer (y único) valor de cada fila
    
    // Comprueba si el valor coincide con el término de búsqueda
    if (valor.toLowerCase().indexOf(terminoBusqueda.toLowerCase()) !== -1) {
      resultados.push(valor); // Añade el valor a la lista de resultados si coincide
    }
  }
  
  // Devuelve los resultados
  return resultados;
}
// ------------------------ CATALOGO PAISES / PROVINCIAS / POBLACIONES ------------------------ //

function quitarTildes(texto) {
  return String(texto || '').normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

/**
 * Lee y cachea en memoria el catálogo completo (paises, provincias, poblaciones).
 * Estructura:
 *  Country   : A=countryCode, B=Name
 *  Province  : A=countryCode, B=provinceCode, C=Name
 *  Population: A=countryCode, B=provinceCode, C=populationCode, D=Name
 */
function getLocationCatalog_() {
  // Usar SIEMPRE el spreadsheet activo: el catálogo vive en hojas locales
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const countrySheet = ss.getSheetByName('Country');
  const lastRowCountry = countrySheet.getLastRow();
  const countryValues = lastRowCountry > 1
    ? countrySheet.getRange(2, 1, lastRowCountry - 1, 2).getValues()
    : [];
  const countries = countryValues
    .map(r => ({ code: String(r[0]), name: String(r[1]) }))
    .filter(c => c.code && c.name);

  const provinceSheet = ss.getSheetByName('Province');
  const lastRowProv = provinceSheet.getLastRow();
  const provValues = lastRowProv > 1
    ? provinceSheet.getRange(2, 1, lastRowProv - 1, 3).getValues()
    : [];
  const provinces = provValues
    .map(r => ({
      countryCode: String(r[0]),
      code: String(r[1]),
      name: String(r[2])
    }))
    .filter(p => p.countryCode && p.code && p.name);

  const populationSheet = ss.getSheetByName('Population');
  const lastRowPop = populationSheet.getLastRow();
  const popValues = lastRowPop > 1
    ? populationSheet.getRange(2, 1, lastRowPop - 1, 4).getValues()
    : [];
  const populations = popValues
    .map(r => ({
      countryCode: String(r[0]),
      provinceCode: String(r[1]),
      code: String(r[2]),
      name: String(r[3])
    }))
    .filter(p => p.countryCode && p.provinceCode && p.code && p.name);

  return { countries, provinces, populations };
}

/** Devuelve sólo los nombres de país para usar en data validation. */
function getCountryNameList_() {
  const catalog = getLocationCatalog_();
  return catalog.countries.map(c => c.name).sort();
}

/** Devuelve nombres de provincias para un país (por nombre de país). */
function getProvinceNamesForCountry_(countryName) {
  if (!countryName) return [];
  const catalog = getLocationCatalog_();
  const normalized = String(countryName).trim();
  const country = catalog.countries.find(c => c.name === normalized);
  if (!country) return [];
  return catalog.provinces
    .filter(p => p.countryCode === country.code)
    .map(p => p.name)
    .sort();
}

/** Devuelve nombres de poblaciones para un país + provincia (por nombre). */
function getPopulationNames_(countryName, provinceName) {
  if (!countryName || !provinceName) return [];
  const catalog = getLocationCatalog_();
  const country = catalog.countries.find(c => c.name === String(countryName).trim());
  if (!country) return [];
  const province = catalog.provinces.find(p =>
    p.countryCode === country.code && p.name === String(provinceName).trim()
  );
  if (!province) return [];

  return catalog.populations
    .filter(pop => pop.countryCode === country.code && pop.provinceCode === province.code)
    .map(pop => pop.name)
    .sort();
}

/**
 * A partir de nombres (pais / provincia / poblacion) devuelve los códigos
 * definidos en el catálogo. Si algo no se encuentra, devuelve null en ese campo.
 */
function getLocationCodesFromNames(countryName, provinceName, populationName) {
  const catalog = getLocationCatalog_();
  let countryCode = null;
  let provinceCode = null;
  let populationCode = null;

  if (countryName) {
    const c = catalog.countries.find(cc => cc.name === String(countryName).trim());
    if (c) countryCode = c.code;

    if (provinceName) {
      const p = catalog.provinces.find(pp =>
        pp.countryCode === countryCode && pp.name === String(provinceName).trim()
      );
      if (p) provinceCode = p.code;

      if (populationName) {
        const pop = catalog.populations.find(po =>
          po.countryCode === countryCode &&
          po.provinceCode === provinceCode &&
          po.name === String(populationName).trim()
        );
        if (pop) populationCode = pop.code;
      }
    }
  }

  return {
    countryCode,
    provinceCode,
    populationCode
  };
}

// --- Wrappers expuestos al frontend (sidebar) ---

/** Devuelve lista de países para el sidebar de clientes. */
function apiGetCountries() {
  return getCountryNameList_();
}

/** Devuelve lista de provincias para un país (nombre) para el sidebar. */
function apiGetProvinces(countryName) {
  return getProvinceNamesForCountry_(countryName);
}

/** Devuelve lista de poblaciones para país + provincia (nombres) para el sidebar. */
function apiGetPopulations(countryName, provinceName) {
  return getPopulationNames_(countryName, provinceName);
}

function obtenerTipoDePersona(e){
  let sheet = e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();
  let colEditada = 4;

  let tipoPersona =sheet.getRange(rowEditada,colEditada).getValue()
  return tipoPersona
}

function saveClientData(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
  if (!sheet) {
    throw new Error('La hoja "Clientes" no existe.');
  }

  let existe = verificarCodigo(formData.numeroIdentificacion, "Clientes", false);
  let existeC=verificarCodigo(formData.codigoContacto, "Clientes", false,null,"codigo");
  if (existe) {
    return { success: false, message: 'El Número de Identificación ya existe. Por favor ingrese un número único.' };
  }else if(existeC){
    return { success: false, message: 'El Codigo ya existe. Por favor ingrese un número único.' };
  }

  // Buscar la primera fila disponible usando la columna A (Estado) vacía
  // Esto asegura escribir sobre filas de plantilla con validaciones ya configuradas
  const lastRow = sheet.getLastRow();
  let emptyRow = 0;
  for (let r = 2; r <= lastRow; r++) {
    const estadoCell = String(sheet.getRange(r, 1).getDisplayValue() || '').trim();
    const idUnicoCell = String(sheet.getRange(r, 2).getDisplayValue() || '').trim();
    if (estadoCell === '' && idUnicoCell === '') {
      emptyRow = r;
      break;
    }
  }
  if (emptyRow === 0) {
    emptyRow = lastRow + 1; // si no hay hueco, agregar al final
  }

  const values = [
    formData.tipoContacto,
    formData.tipoPersona,
    formData.tipoDocumento,
    formData.numeroIdentificacion,
    formData.codigoContacto,
    formData.nombreComercial,
    formData.primerNombre,
    formData.segundoNombre,
    formData.primerApellido,
    formData.segundoApellido,
    formData.pais,
    formData.provincia,
    formData.poblacion,
    formData.direccion,
    formData.codigoPostal,
    formData.telefono,
    formData.sitioWeb,
    formData.email,
  ];
  let nombre="";
  // Tratar "Persona Física" como autónomo para construir el identificador único
  // Normalizamos y comparamos en minúsculas sin tildes
  let tipoNormSave = String(formData.tipoPersona)
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().trim();
  if (tipoNormSave === "autonomo" || tipoNormSave === "persona fisica") {
    const primerNombre = formData.primerNombre || "";
    const apellido = formData.primerApellido || "";
    nombre = (primerNombre+" "+apellido).trim();
  }else{
    nombre = formData.nombreComercial || "";
  }

  sheet.getRange(emptyRow, 3, 1, values.length).setValues([values]);
  let referenciaUnica = nombre + "-" + formData.numeroIdentificacion;
  sheet.getRange(emptyRow, 2).setValue(referenciaUnica);
  Logger.log("dentro de ref unico "+referenciaUnica)
  sheet.getRange(emptyRow, 1).setValue("Valido");

  return { success: true, message: 'Cliente creado exitosamente.' , refe: referenciaUnica};
}
function agregarUltimoCliente(referenciaUnica){
  Logger.log("agregarUltimo")
  Logger.log("referenciaUnica "+referenciaUnica)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  let InfoCliente=obtenerInformacionCliente(referenciaUnica)
  var sheet = ss.getSheetByName("Factura");
  sheet.getRange("B2").setValue(referenciaUnica)
  sheet.getRange("B3").setValue(InfoCliente["Código cliente"])
  obtenerFechaYHoraActual()
}

function agregarUltimoProducto(refe){
  agregarProductoDesdeFactura(1,refe)
}


function verificarDatosObligatoriosProductos(e){
  const sheet = e.source.getActiveSheet();
  const rowEditada = e.range.getRow();
  const colEditada = e.range.getColumn();
  const columnasObligatorias = [
    PRODUCT_COLUMNS.CODIGO_REFERENCIA,
    PRODUCT_COLUMNS.NOMBRE,
    PRODUCT_COLUMNS.REGIMEN,
    PRODUCT_COLUMNS.TIPO_PRODUCTO,
    PRODUCT_COLUMNS.TIPO_USO,
    PRODUCT_COLUMNS.VALOR_UNITARIO,
    PRODUCT_COLUMNS.TIPO_IMPUESTO,
    PRODUCT_COLUMNS.TARIFA_IMPUESTO
  ];
  const columnasARevisar = columnasObligatorias.concat([
    PRODUCT_COLUMNS.PRECIO_CON_IMPUESTO,
    PRODUCT_COLUMNS.TARIFA_RECARGO,
    PRODUCT_COLUMNS.TARIFA_RETENCION
  ]);
  const estadosDefault = ["", "Seleccione", "Selecciona una opción"];
  const estadosDefaultLower = estadosDefault.map(item => item.toLowerCase());

  if (rowEditada <= 1) {
    return;
  }

  // Limpiar fondos previos
  columnasARevisar.forEach(columna => {
    sheet.getRange(rowEditada, columna).setBackground(null);
  });

  if (colEditada > PRODUCT_COLUMNS.IDENTIFICADOR_UNICO) {
    return;
  }

  let estaCompleto = true;
  let estaTotalmenteVacio = true;

  columnasObligatorias.forEach(columna => {
    const valor = sheet.getRange(rowEditada, columna).getDisplayValue().trim();
    const valorLower = valor.toLowerCase();
    if (valor !== "") {
      estaTotalmenteVacio = false;
    }
    if (estadosDefault.includes(valor) || estadosDefaultLower.includes(valorLower)) {
      estaCompleto = false;
      sheet.getRange(rowEditada, columna).setBackground('#FFC7C7');
    }
  });

  const esRecargo = sheet.getRange(rowEditada, PRODUCT_COLUMNS.CHECK_RECARGO).getValue() === true;
  const retencionActiva = sheet.getRange(rowEditada, PRODUCT_COLUMNS.CHECK_RETENCION).getValue() === true;
  const tarifaRetencionDisplay = sheet.getRange(rowEditada, PRODUCT_COLUMNS.TARIFA_RETENCION).getDisplayValue().trim();

  if (esRecargo) {
    const ivaNum = parsePercentToNumberES(sheet.getRange(rowEditada, PRODUCT_COLUMNS.TARIFA_IMPUESTO).getDisplayValue());
    const esperado = recargoPermitidoParaIva(ivaNum);
    const tarifaNum = parsePercentToNumberES(sheet.getRange(rowEditada, PRODUCT_COLUMNS.TARIFA_RECARGO).getDisplayValue());
    if (esperado === null || tarifaNum === null || Math.abs(tarifaNum - esperado) > 0.0001) {
      estaCompleto = false;
      sheet.getRange(rowEditada, PRODUCT_COLUMNS.TARIFA_RECARGO).setBackground('#FFC7C7');
    }
  }

  if (retencionActiva) {
    const tarifaNum = parsePercentToNumberES(tarifaRetencionDisplay);
    const permitido = RETENCION_IRPF_TARIFAS
      .map(valor => parsePercentToNumberES(valor))
      .some(valor => Math.abs(valor - tarifaNum) < 0.0001);
    if (!permitido) {
      estaCompleto = false;
      sheet.getRange(rowEditada, PRODUCT_COLUMNS.TARIFA_RETENCION).setBackground('#FFC7C7');
    }
  }

  if (estaTotalmenteVacio) {
    sheet.getRange(rowEditada, PRODUCT_COLUMNS.ESTADO).clearContent();
    sheet.getRange(rowEditada, PRODUCT_COLUMNS.IDENTIFICADOR_UNICO).clearContent();
    sheet.getRange(rowEditada, PRODUCT_COLUMNS.PRECIO_CON_IMPUESTO).clearContent();
    return;
  }

  const estado = estaCompleto ? 'Valido' : 'No Valido';
  sheet.getRange(rowEditada, PRODUCT_COLUMNS.ESTADO).setValue(estado);

  if (estaCompleto) {
    sheet.getRange(rowEditada, PRODUCT_COLUMNS.VALOR_UNITARIO).setNumberFormat('€#,##0.00');
    sheet.getRange(rowEditada, PRODUCT_COLUMNS.PRECIO_CON_IMPUESTO)
      .setFormula(`=IF(AND(G${rowEditada}<>"";I${rowEditada}<>"");G${rowEditada}*(1+I${rowEditada});"")`);
    sheet.getRange(rowEditada, PRODUCT_COLUMNS.PRECIO_CON_IMPUESTO).setNumberFormat('€#,##0.00');
  }
}

function verificarDatosObligatorios(e, tipoPersona) {
  let sheet = e.source.getActiveSheet();
  let range = e.range;
  let rowEditada = range.getRow();
  let colEditada = range.getColumn();
  let ultimaColumnaPermitida = 20;
  let columnasObligatorias = [];
  let todasLasColumnas = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20];

  if (tipoPersona === "") {
    Logger.log("Vacio hizo edicion no en tipoPersona, cogemos el viejo");
    tipoPersona = sheet.getRange("D" + String(rowEditada)).getValue();
  }

  const esAutonomo =
    tipoPersona === "Autónomo" ||
    tipoPersona === "Persona Física";

  if (esAutonomo) {
    // primerNombre(9), primerApellido(11), pais(13), provincia(14), poblacion(15), codigoPostal(17), email(20)
    columnasObligatorias = [3, 4, 5, 6, 7, 9, 11, 13, 14, 15, 17, 20];
  } else if (tipoPersona === "Empresa") {
    // nombreComercial(8), pais(13), provincia(14), poblacion(15), codigoPostal(17), email(20)
    columnasObligatorias = [3, 4, 5, 6, 7, 8, 13, 14, 15, 17, 20];
  } else {
    Logger.log("Vacio tipo de persona");
  }
  
  let estadosDefault = ["", "Tipo Documento", "Tipo de persona"]; // Aquí otros estados predeterminados si es necesario

  if (rowEditada > 1 && colEditada <= ultimaColumnaPermitida) {
    let estaCompleto = true;
    let estaVacioOPredeterminado = true;

    // Borrar el color de fondo de todas las celdas obligatorias antes de la verificación
    for (let i = 0; i < todasLasColumnas.length; i++) {
      sheet.getRange(rowEditada, todasLasColumnas[i]).setBackground(null);
    }

    // Verificar celdas obligatorias
    for (let i = 0; i < columnasObligatorias.length; i++) {
      let valorDeCelda = sheet.getRange(rowEditada, columnasObligatorias[i]).getValue();
      if (estadosDefault.includes(valorDeCelda)) {
        estaCompleto = false;
        sheet.getRange(rowEditada, columnasObligatorias[i]).setBackground('#FFC7C7'); // Resaltar en rojo claro
      } else {
        estaVacioOPredeterminado = false;
      }
    }

    // Actualizar el estado en la primera columna
    if (estaVacioOPredeterminado) {
      sheet.getRange(rowEditada, 1).clearContent(); // Limpiar contenido de "Estado"
    } else {
      let status = estaCompleto ? "Valido" : "No Valido";
      sheet.getRange(rowEditada, 1).setValue(status); // Establecer valor en "Estado"
    }
  }
}


function crearContacto(){
  Logger.log("imprima algo")
  showNuevaClienteDesdeFactura()

}

function crearProducto(){
  showNuevaProductoDesdeFactura()
}

function getIdentificationCode(IdentificationType) {


  if (IdentificationType === "Cliente") {
    return "01";
  } else if (IdentificationType === "Proveedor") {
    return "02";
  } else {
    throw new Error("Valor inválido en AB2. Debe ser 'Cliente' o 'Proveedor'.");
  }
}

function getIdentificationCodeDocument(IdentificationType) {
  switch (IdentificationType) {
    case "NIF-IVA":
      return "02";
    case "Pasaporte":
      return "03";
    case "Documento oficial de identificación expedido por":
      return "04";
    case "Certificado de residencia":
      return "05";
    case "Otro documento aprobado":
      return "06";
    case "No censado":
      return "07";
    default:
      throw new Error("Valor inválido en el tipo de identificación. Debe ser uno de los valores permitidos.");
  }
}

function getTypePersonCode(TypePerson) {
  Logger.log("TypePerson"+TypePerson)
  // Aceptar variaciones con/ sin tilde y espacios
  let tipoNorm = String(TypePerson)
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toLowerCase().trim();
  if (tipoNorm === "autonomo" || tipoNorm === "persona fisica") {
    return "01";
  } else if (tipoNorm === "empresa") {
    return "02";
  } else {
    throw new Error("Valor inválido para TypePerson. Debe ser 'Autónomo' o 'Empresa'.");
  }
}

function getRegimenCode(Regimen) {
  Logger.log("Regimen"+Regimen)
  const regimenMap = {
    "Operación de régimen general": "01",
    "Exportación": "02",
    "Operaciones a las que se aplique el régimen especial de bienes usados, objetos de arte, antigüedades y objetos de colección": "03",
    "Régimen especial del oro de inversión": "04",
    "Régimen especial de las agencias de viajes": "05",
    "Régimen especial grupo de entidades en IVA (Nivel Avanzado)": "06",
    "Régimen especial del criterio de caja": "07",
    "Operaciones sujetas al IPSI / IGIC (Impuesto sobre la Producción, los Servicios y la Importación / Impuesto General Indirecto Canario)": "08",
    "Facturación de las prestaciones de servicios de agencias de viaje que actúan como mediadoras en nombre y por cuenta ajena (D.A.4ª RD1619/2012)": "09",
    "Cobros por cuenta de terceros de honorarios profesionales o de derechos derivados de la propiedad industrial, de autor u otros por cuenta de sus socios, asociados o colegiados efectuados por sociedades, asociaciones, colegios profesionales u otras entidades que realicen estas funciones de cobro": "10",
    "Operaciones de arrendamiento de local de negocio": "11",
    "Factura con IVA pendiente de devengo en certificaciones de obra cuyo destinatario sea una Administración Pública": "14",
    "Factura con IVA pendiente de devengo en operaciones de tracto sucesivo": "15",
    "Operación acogida a alguno de los regímenes previstos en el Capítulo XI del Título IX (OSS e IOSS)": "17",
    "Recargo de equivalencia": "18",
    "Operaciones de actividades incluidas en el Régimen Especial de Agricultura, Ganadería y Pesca (REAGYP)": "19",
    "Régimen simplificado": "20"
  };

  const code = regimenMap[Regimen];
  if (!code) {
    throw new Error("Valor inválido para Regimen. Asegúrate de que el texto coincide exactamente con una de las opciones.");
  }
  return code;
}

function getCustomerInformation(customer) {
  /*esta funcion debe de cambiar para obtener son los datos directamente de la hoja cliente */
  // ojo de donde esta cogiendo el datosheet ?
  let spreadsheet = SpreadsheetApp.getActive();
  let datos_sheet = spreadsheet.getSheetByName('Datos');
  let celdaCliente = datos_sheet.getRange("H2");
  celdaCliente.setValue(customer);


  // var range = datos_sheet.getRange("D50");
  // var Customer = range.getValue();

  var range = datos_sheet.getRange("I2");
  var CustomerCode = range.getValue();

  //range = datos_sheet.getRange("C51");// aqui agarra es el numero mas no el tipo en si
  //var IdentificationType = range.getValue();
  let IdentificationType=datos_sheet.getRange("AB2").getValue();
  IdentificationType=getIdentificationCode(IdentificationType)

  let DocumentIdentificationType = datos_sheet.getRange("J2").getValue();
  DocumentIdentificationType= getIdentificationCodeDocument(DocumentIdentificationType)

  // Conservar el valor original para validaciones de negocio (p.ej. recargo de equivalencia)
  let TypePersonOriginal = datos_sheet.getRange("L2").getValue();
  let TypePerson = getTypePersonCode(TypePersonOriginal);

  range = datos_sheet.getRange("K2");
  var Identification = range.getValue();//numero de identificacion

  
  var DV = 0;//no existe en espana, predeterminado 0

  range = datos_sheet.getRange("T2");
  var Address = range.getValue();// aqui lo dividia entre 2 por el psotalcode
  
  

  range = datos_sheet.getRange("S2");//cambie en vez de ciudad pais, porque en espana no hay parametro ciudad
  var CityID = range.getValue();

  range = datos_sheet.getRange("V2");
  var Telephone = range.getValue();

  // switch (datos_sheet.getRange("C1").getValue()) {
  //   case "Pruebas":
  //     var range = datos_sheet.getRange("E1");
  //     break;
  //   case "Produccion":
  //     var range = datos_sheet.getRange("B63");
  //     break;
  //   default:
  //     Logger.log("Oops!...Error Ambiente")
  //     return;
  // }
  var range = datos_sheet.getRange("X2");
  var Email = range.getValue();
  //Browser.msgBox(Email);


  range = datos_sheet.getRange("W2");
  var WebSiteURI = range.getValue();

  var paisCliente= datos_sheet.getRange("S2").getValue();
  let codigoPostalCliente=datos_sheet.getRange("U2").getValue();
  let provinciaCliente = datos_sheet.getRange("AA2").getValue();
  let poblacionCliente = datos_sheet.getRange("Z2").getValue();

  // Obtener códigos oficiales desde el catálogo externo a partir de los nombres
  const locationCodes = getLocationCodesFromNames(paisCliente, provinciaCliente, poblacionCliente);

  if (IdentificationType == "#NUM!") {
    Browser.msgBox("ERROR: Seleccione Tipo de Identificacion en Clientes")
    return;
  }
  let valorFecha=ObtenerFecha()
  let valorFechaPago=ObtenerFecha("pago")
  var CustomerInformation = {
    "IdentificationType": IdentificationType,
    "Identification": Identification,//.toString(),
    "DocumentIdentificationType":DocumentIdentificationType,
    "DV": valorFecha,
    "RegistrationName": customer,
    // Código de país usado por el catálogo (no se muestra en hoja Clientes)
    "CountryCode": locationCodes.countryCode || "",
    "CountryName": paisCliente,
    "FechaPago": valorFechaPago,// 11, //Codigo de Municipio
    "SubdivisionName": provinciaCliente,// provicnica
    "CityCode": codigoPostalCliente,
    "CityName": poblacionCliente,//poblacion
    "AddressLine": String(Address),
    "PostalZone": datos_sheet.getRange("U2").getValue(),//Confundido con el codigo postal hay 2, de recepcion y de 
    "Email": Email,
    "CustomerCode": CustomerCode,
    "Telephone": Telephone,
    "WebSiteURI": WebSiteURI,
    "AdditionalAccountID": "Numero que representa el tipo de persona, en España no se sabe si se utiliza o no",//"1",//1, //1: Juridica, 2: Natural
    "TaxLevelCodeListName": "numero que representa unos impuestos, no se si en España exista",//"48" Impuesto sobre las ventas IVA 49 – No responsable de impuesto sobre las ventas IVA
    "TaxSchemeCode": "Numero que representa algo, no se si en España exista ",
    "TaxSchemeName": "",
    "FiscalResponsabilities": "Responsabiliades fiscales, no se si en España exista",

    "PartecipationPercent": 100,
    "AdditionalCustomer": [],
    "TypePerson":TypePerson,
    // Metadatos para reglas específicas (NO impactan el payload del API)
    "TypePersonName": TypePersonOriginal,
    "TypePersonNorm": String(TypePersonOriginal || "")
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .toLowerCase().trim(),
    "ProvinceCode": locationCodes.provinceCode || "",
    "PopulationCode": locationCodes.populationCode || ""


  }
  return CustomerInformation;
}

function obtenerInformacionCliente(cliente) {
  let spreadsheet = SpreadsheetApp.getActive();
  let datos_sheet = spreadsheet.getSheetByName('Datos');
  let celdaCliente = datos_sheet.getRange("H2");
  celdaCliente.setValue(cliente);



  let codigoContacto = datos_sheet.getRange("K2").getValue();
  let direccion = datos_sheet.getRange("T2").getValue();
  let pais = datos_sheet.getRange("S2").getValue();
  let provincia = datos_sheet.getRange("AA2").getValue();
  let poblacion = datos_sheet.getRange("Z2").getValue();
  let telefono = datos_sheet.getRange("V2").getValue();
  let estado = datos_sheet.getRange("Y2").getValue();

  let ubicacion = poblacion + ", " + provincia + ", " + pais;

  let informacionCliente = {
    "Código cliente": codigoContacto,
    "Dirección": direccion,
    "Ubicación": ubicacion,
    "Teléfono": telefono,
    "Estado": estado
  };

  return informacionCliente;
}
