function normalizarTexto(s) {
    return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

// Obtiene provincias según país con filtro de texto y límite
function getProvincias(pais, query, limit) {
    try {
        var ss = SpreadsheetApp.getActive();
        var hoja = ss.getSheetByName('Datos');
        if (!hoja) { return []; }
        // Paises: C (código), D (nombre) filas 26..275
        var paisesCD = hoja.getRange(26, 3, 250, 2).getValues();
        var mapPaisNombreACodigo = {};
        for (var p = 0; p < paisesCD.length; p++) {
            var codP = paisesCD[p][0];
            var nomP = paisesCD[p][1];
            if (!codP || !nomP) continue;
            mapPaisNombreACodigo[normalizarTexto(nomP)] = String(codP);
        }
        var normPais = normalizarTexto(pais);
        var codigoPais = mapPaisNombreACodigo[normPais] || '';
        if (!codigoPais) { return []; }

        // Provincias: F (código país), G (código provincia), H (nombre) filas 26..5110
        var values = hoja.getRange(26, 6, 5110 - 26 + 1, 3).getValues();
        var normQuery = normalizarTexto(query);
        var resultados = [];
        var max = Math.max(1, Number(limit || 30));
        for (var i = 0; i < values.length && resultados.length < max; i++) {
            var row = values[i];
            var codPais = row[0];
            var codProv = row[1];
            var nomProv = row[2];
            if (!codPais || !nomProv) continue;
            if (String(codPais) === String(codigoPais) && (!normQuery || normalizarTexto(nomProv).indexOf(normQuery) !== -1)) {
                resultados.push(String(nomProv));
            }
        }
        var uniq = Array.from(new Set(resultados));
        return uniq;
    } catch (e) {
        return [];
    }
}

// Obtiene poblaciones según país y provincia, con filtro incremental y límite
function getPoblaciones(pais, provincia, query, limit) {
    Logger.log("buscarPoblacion provincia = " + provincia);
    try {
        var ss = SpreadsheetApp.getActive();
        var hoja = ss.getSheetByName('Datos');
        if (!hoja) { return []; }

        // Mapa país nombre -> código (C:D)
        var paisesCD = hoja.getRange(26, 3, 250, 2).getValues();
        var mapPaisNombreACodigo = {};
        for (var p = 0; p < paisesCD.length; p++) {
            var codP = paisesCD[p][0];
            var nomP = paisesCD[p][1];
            if (!codP || !nomP) continue;
            mapPaisNombreACodigo[normalizarTexto(nomP)] = String(codP);
        }
        var normPais = normalizarTexto(pais);
        var codigoPais = mapPaisNombreACodigo[normPais] || '';
        if (!codigoPais) { return []; }

        // Obtener código de provincia a partir del nombre (F:G:H)
        var provs = hoja.getRange(26, 6, 5110 - 26 + 1, 3).getValues();
        var normProv = normalizarTexto(provincia);
        var codigoProv = '';
        for (var r = 0; r < provs.length; r++) {
            var codPaisR = provs[r][0]; // F
            var codProvR = provs[r][1]; // G
            var nomProvR = provs[r][2]; // H
            if (!codPaisR || !codProvR || !nomProvR) continue;
            if (String(codPaisR) === String(codigoPais) && normalizarTexto(nomProvR) === normProv) {
                codigoProv = String(codProvR);
                break;
            }
        }
        if (!codigoProv) { return []; }

        // Poblaciones: I (código país), L (código provincia), M (código población), N (nombre) filas 26..150659
        var startRow = 26;
        var lastRow = 150659; // límite declarado
        var numRows = lastRow - startRow + 1;

        // Cache por país+provincia para evitar recorrer 150k filas en cada búsqueda
        var cache = CacheService.getScriptCache();
        var cacheKey = 'pobl_' + String(codigoPais) + '_' + String(codigoProv);
        var cached = cache.get(cacheKey);
        var lista = null;
        if (cached) {
            try { lista = JSON.parse(cached); } catch (err) { lista = null; }
        }
        if (!lista) {
            // Buscar solo filas cuyo L (col 12) == codigoProv usando TextFinder en la columna L
            var colProvRange = hoja.getRange(startRow, 12, numRows, 1); // L26:L150659
            var matches = colProvRange.createTextFinder(String(codigoProv)).matchEntireCell(true).findAll();
            lista = [];
            for (var m = 0; m < matches.length; m++) {
                var rowIdx = matches[m].getRow();
                // Verificar país en I
                var codPaisR = hoja.getRange(rowIdx, 9).getValue();
                if (String(codPaisR) !== String(codigoPais)) continue;
                var nombre = hoja.getRange(rowIdx, 14).getValue(); // N
                if (nombre) lista.push(String(nombre));
            }
            try { cache.put(cacheKey, JSON.stringify(lista), 21600); } catch (err) {}
        }

        var normQuery = normalizarTexto(query);
        var resultados = [];
        var max = Math.max(1, Number(limit || 30));
        for (var j = 0; j < lista.length && resultados.length < max; j++) {
            var nombre = lista[j];
            if (!normQuery || normalizarTexto(nombre).indexOf(normQuery) !== -1) {
                resultados.push(nombre);
            }
        }
        var uniq = Array.from(new Set(resultados));
        return uniq;
    } catch (e) {
        return [];
    }
}

function existeProvincia(pais, provincia) {
    try {
        var lista = getProvincias(pais, provincia, 1000);
        var objetivo = normalizarTexto(provincia);
        for (var i = 0; i < lista.length; i++) {
            if (normalizarTexto(lista[i]) === objetivo) return true;
        }
        return false;
    } catch (e) {
        return true; // En caso de error, no bloquear
    }
}

function existePoblacion(pais, provincia, poblacion) {
    try {
        var props = PropertiesService.getDocumentProperties();
        var codPais = props.getProperty('paisCodigoSeleccionado');
        var codProv = props.getProperty('provCodigoSeleccionado');
        var lista = [];
        if (codPais && codProv) {
            lista = getPoblacionesPorCodigo(codPais, codProv, '', 1000);
        } else {
            lista = getPoblaciones(pais, provincia, '', 1000);
        }
        var objetivo = normalizarTexto(poblacion);
        for (var i = 0; i < lista.length; i++) {
            if (normalizarTexto(lista[i]) === objetivo) return true;
        }
        // Si no hay lista (sin catálogo cargado), no bloquear
        if (lista.length === 0) return true;
        return false;
    } catch (e) {
        return true;
    }
}

// NUEVO: obtener código de país por nombre (Datos!C:D filas 26..275)
function getCodigoPais(nombrePais){
    var ss = SpreadsheetApp.getActive();
    var hoja = ss.getSheetByName('Datos');
    if(!hoja){ return ''; }
    var paisesCD = hoja.getRange(26, 3, 250, 2).getValues();
    var objetivo = normalizarTexto(nombrePais);
    for (var i=0;i<paisesCD.length;i++){
        var cod = paisesCD[i][0];
        var nom = paisesCD[i][1];
        if(!cod || !nom) continue;
        if (normalizarTexto(nom) === objetivo) return String(cod);
    }
    return '';
}

// NUEVO: obtener código de provincia por nombre y país
function getCodigoProvincia(nombrePais, nombreProvincia){
    var ss = SpreadsheetApp.getActive();
    var hoja = ss.getSheetByName('Datos');
    if(!hoja){ return ''; }
    var codigoPais = getCodigoPais(nombrePais);
    if(!codigoPais){ return ''; }
    var provs = hoja.getRange(26, 6, 5110 - 26 + 1, 3).getValues(); // F:G:H
    var objetivo = normalizarTexto(nombreProvincia);
    for (var i=0;i<provs.length;i++){
        var codPaisR = provs[i][0];
        var codProvR = provs[i][1];
        var nomProvR = provs[i][2];
        if(!codPaisR || !codProvR || !nomProvR) continue;
        if(String(codPaisR)===String(codigoPais) && normalizarTexto(nomProvR)===objetivo){
            return String(codProvR);
        }
    }
    return '';
}

// NUEVO: obtener nombre de país por código
function getNombrePaisPorCodigo(codigoPais){
    var ss = SpreadsheetApp.getActive();
    var hoja = ss.getSheetByName('Datos');
    if(!hoja){ return ''; }
    var paisesCD = hoja.getRange(26, 3, 250, 2).getValues();
    var objetivo = String(codigoPais);
    for (var i=0;i<paisesCD.length;i++){
        var cod = String(paisesCD[i][0]);
        var nom = paisesCD[i][1];
        if(cod===objetivo){ return String(nom); }
    }
    return '';
}

// NUEVO: obtener nombre de provincia por códigos de país y provincia
function getNombreProvinciaPorCodigo(codigoPais, codigoProvincia){
    var ss = SpreadsheetApp.getActive();
    var hoja = ss.getSheetByName('Datos');
    if(!hoja){ return ''; }
    var provs = hoja.getRange(26, 6, 5110 - 26 + 1, 3).getValues(); // F:G:H
    var objetivoPais = String(codigoPais);
    var objetivoProv = String(codigoProvincia);
    for (var i=0;i<provs.length;i++){
        var codPaisR = String(provs[i][0]);
        var codProvR = String(provs[i][1]);
        var nomProvR = provs[i][2];
        if(codPaisR===objetivoPais && codProvR===objetivoProv){
            return String(nomProvR);
        }
    }
    return '';
}

// NUEVO: obtener poblaciones por código de país y provincia
function getPoblacionesPorCodigo(codigoPais, codigoProvincia, query, limit){
    try{
        var ss = SpreadsheetApp.getActive();
        var hoja = ss.getSheetByName('Datos');
        if(!hoja){ return []; }
        var startRow = 26;
        var lastRow = 150659;
        var numRows = lastRow - startRow + 1;
        var cache = CacheService.getScriptCache();
        var cacheKey = 'pobl_cod_' + String(codigoPais) + '_' + String(codigoProvincia);
        var cached = cache.get(cacheKey);
        var lista = null;
        if (cached){ try{ lista = JSON.parse(cached); }catch(e){ lista=null; } }
        if(!lista){
            var colProvRange = hoja.getRange(startRow, 12, numRows, 1); // L
            var matches = colProvRange.createTextFinder(String(codigoProvincia)).matchEntireCell(true).findAll();
            lista = [];
            for (var m=0;m<matches.length;m++){
                var rowIdx = matches[m].getRow();
                var codPaisR = hoja.getRange(rowIdx, 9).getValue(); // I
                if(String(codPaisR)!==String(codigoPais)) continue;
                var nombre = hoja.getRange(rowIdx, 14).getValue(); // N
                if(nombre) lista.push(String(nombre));
            }
            try{ cache.put(cacheKey, JSON.stringify(lista), 21600); }catch(e){}
        }
        var normQuery = normalizarTexto(query);
        var resultados = [];
        var max = Math.max(1, Number(limit||30));
        for (var j=0;j<lista.length && resultados.length<max;j++){
            var nombre = lista[j];
            if(!normQuery || normalizarTexto(nombre).indexOf(normQuery)!==-1){
                resultados.push(nombre);
            }
        }
        return Array.from(new Set(resultados));
    }catch(e){
        return [];
    }
}

// Helpers para trabajar con celdas de filtro en hoja Datos (P25: codigo pais, Q25: codigo provincia, R25: query prov, S25: query pob)
function setPaisFiltroPorNombre(nombrePais){
    var ss = SpreadsheetApp.getActive();
    var hoja = ss.getSheetByName('Datos');
    var codigo = getCodigoPais(nombrePais);
    hoja.getRange('P25').setValue(codigo || '');
    hoja.getRange('Q25').clearContent();
    hoja.getRange('R25').clearContent();
    hoja.getRange('S25').clearContent();
    // Guardar en propiedades para recuperar nombres si el front no los envía
    try{
        PropertiesService.getDocumentProperties().setProperty('paisCodigoSeleccionado', String(codigo||''));
    }catch(e){}
    return String(codigo||'');
}

function setProvinciaFiltroPorNombre(nombrePais, nombreProvincia){
    var ss = SpreadsheetApp.getActive();
    var hoja = ss.getSheetByName('Datos');
    var codPais = getCodigoPais(nombrePais);
    var codProv = getCodigoProvincia(nombrePais, nombreProvincia);
    hoja.getRange('P25').setValue(codPais || '');
    hoja.getRange('Q25').setValue(codProv || '');
    hoja.getRange('S25').clearContent();
    try{
        PropertiesService.getDocumentProperties().setProperty('paisCodigoSeleccionado', String(codPais||''));
        PropertiesService.getDocumentProperties().setProperty('provCodigoSeleccionado', String(codProv||''));
    }catch(e){}
    return String(codProv||'');
}

function setQueryProvincia(query){
    var ss = SpreadsheetApp.getActive();
    var hoja = ss.getSheetByName('Datos');
    hoja.getRange('R25').setValue(query||'');
    return true;
}

function setQueryPoblacion(query){
    var ss = SpreadsheetApp.getActive();
    var hoja = ss.getSheetByName('Datos');
    hoja.getRange('S25').setValue(query||'');
    return true;
}

function getSugerenciasProvincias(){
    var ss = SpreadsheetApp.getActive();
    var hoja = ss.getSheetByName('Datos');
    var last = hoja.getLastRow();
    if(last < 26){ return []; }
    var num = last - 25;
    var values = hoja.getRange(26, 16, num, 1).getValues(); // P26:P
    var res = [];
    for (var i=0;i<values.length && res.length<30;i++){
        var v = values[i][0];
        if(v && v !== '#N/A'){ res.push(String(v)); }
    }
    return res;
}

function getSugerenciasPoblaciones(){
    var ss = SpreadsheetApp.getActive();
    var hoja = ss.getSheetByName('Datos');
    var last = hoja.getLastRow();
    if(last < 26){ return []; }
    var num = last - 25;
    var values = hoja.getRange(26, 17, num, 1).getValues(); // Q26:Q
    var res = [];
    for (var i=0;i<values.length && res.length<30;i++){
        var v = values[i][0];
        if(v && v !== '#N/A'){ res.push(String(v)); }
    }
    return res;
}

// Combos: preparar y devolver lista ya recalculada
function prepararProvincias(pais){
    try{
        setPaisFiltroPorNombre(pais);
        setQueryProvincia('');
        SpreadsheetApp.flush();
        Utilities.sleep(100);
        return getSugerenciasProvincias();
    }catch(e){
        return [];
    }
}

function prepararPoblaciones(pais, provincia){
    try{
        setProvinciaFiltroPorNombre(pais, provincia);
        setQueryPoblacion('');
        SpreadsheetApp.flush();
        Utilities.sleep(100);
        return getSugerenciasPoblaciones();
    }catch(e){
        return [];
    }
}

