// ---------------------------------------------------------------------------
// Rectificativas (corrective invoices) through the FacturasApp API.
//
// UI: the sidebar (viewFacturaRectificar.html) only searches the original invoice
// and picks the kind; abrirDialogoRectificativa() then opens rectificativaDialog.html
// as a modal dialog that mirrors the two web pages:
//   - "completa":  AddRectificativeComplete ("rectificativa de anulación"). The
//                  backend duplicates the original with negative amounts; we only
//                  send number + reason.
//   - "articulos": AddRectificativeArticles ("rectificativa de sustitución"). We
//                  rebuild the invoice from the JSON stored in ListadoEstado col M
//                  when the original was sent, with the lines as edited in the
//                  dialog (price, quantity, taxes, retention, surcharge, discount).
//
// Verified against QA on 2026-10-28:
//   - Only typeRectificative "SU" (por sustitución) is accepted; "DI" is rejected.
//   - One complete rectificativa per original; after it, by-articles is blocked.
//   - The API validates references (must exist in the original) but not quantities
//     or totals, so we compute and validate everything here.
//   - A gateway 504 can still create the invoice; we confirm it by number.
//
// Numbering (v1.6.0.2, "mode B"): we always send invoiceNumber = nextInvoiceNumber
// of the default rectificativa series. Verified on QA 2026-11-04: letting the API
// assign the number is unreliable with several series (it ignores the default
// series' counter and collides with existing numbers), while sending the default
// series' next number works on both endpoints, also in Verifactu accounts, and
// advances that series. The number is checked with GetInvoiceStateByNumber first.
// ---------------------------------------------------------------------------

const RECTIFICATIVA_MOTIVOS_AEAT = [
  { codigo: 'R1', label: 'R1 - Factura rectificativa por error en los datos' },
  { codigo: 'R2', label: 'R2 - Modificación de la operación' },
  { codigo: 'R3', label: 'R3 - Acuerdo entre las partes' },
  { codigo: 'R4', label: 'R4 - Otras razones' }
];
// The API only accepts "por sustitución" today (see header). Kept in one place.
const RECTIFICATIVA_TIPO = 'SU';
const RECTIFICATIVA_MOTIVO_MAX = 255;
// Observations / payment note limit used by the web form (API allows up to 500)
const RECTIFICATIVA_OBS_MAX = 200;
// Product reference / description limits (same as recalcularLineaRectificativa_ truncation)
const RECT_REFERENCIA_MAX = 50;
const RECT_DESCRIPCION_MAX = 300;
const RECTIFICATIVA_DIALOG_WIDTH = 960;
const RECTIFICATIVA_DIALOG_HEIGHT = 700;
// Historial Facturas Data: extra columns that link a rectificativa to its original
const HISTORIAL_COL_ORIGINAL = 7;      // G
const HISTORIAL_COL_TIPO_RECT = 8;     // H
const RECTIFICATIVA_TIPO_LABEL = { completa: 'Completa', articulos: 'Por artículos' };
// ListadoEstado: invoice number and the JSON sent to AddInvoice
const LISTADO_COL_NUMERO = 6;          // F
const LISTADO_COL_JSON = 13;           // M
// TypeWithHoldings codes (see getIdTypeWithHoldings_ in Factura.js)
const RECT_ID_RETENCION = '10';
const RECT_ID_RETENCION_OTROS = '8';
const RECT_ID_RECARGO = '11';

function round2Rect_(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

/** Percent -> fraction without float noise (5.2 -> 0.052, not 0.052000000000000005). */
function fraccionRect_(pct) {
  return Math.round((Number(pct) || 0) * 10000) / 1000000;
}

/** Catalogue of AEAT reasons for the dialog combo. */
function getRectificativaMotivosAeat() {
  return RECTIFICATIVA_MOTIVOS_AEAT;
}

/** Payment methods for the "Forma de pago" select: first label per code wins. */
function getFormasPago_() {
  const vistos = {};
  const lista = [];
  Object.keys(ID_PAYMENT_LABELS).forEach(label => {
    const codigo = ID_PAYMENT_LABELS[label];
    if (vistos[codigo]) return;
    vistos[codigo] = true;
    lista.push({ codigo: codigo, label: label });
  });
  return lista;
}

/**
 * Company data shown read-only in the dialog header. Only B1 (name) is a fixed
 * cell of 'Datos de emisor'; the rest is picked by label from A1:B12 so a
 * different sheet layout degrades to "name only" instead of failing.
 */
function leerDatosEmisor_() {
  const emisor = { nombre: '', nif: '', direccion: '', poblacion: '' };
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName('Datos de emisor');
    if (!hoja) return emisor;
    const filas = hoja.getRange(1, 1, 12, 2).getDisplayValues();
    emisor.nombre = String(filas[0][1] || '').trim();
    filas.forEach(fila => {
      const label = String(fila[0] || '').trim();
      const valor = String(fila[1] || '').trim();
      if (!label || !valor) return;
      if (!emisor.nif && /\b(nif|cif|identificaci)/i.test(label)) emisor.nif = valor;
      else if (!emisor.direccion && /direcci/i.test(label)) emisor.direccion = valor;
      else if (!emisor.poblacion && /(poblaci|ciudad|localidad|municipio)/i.test(label)) emisor.poblacion = valor;
    });
  } catch (e) {
    Logger.log('leerDatosEmisor_: ' + e);
  }
  return emisor;
}

/**
 * Finds the JSON we sent for an invoice (ListadoEstado col M), newest first.
 * Returns the parsed object or null.
 */
function buscarJsonFacturaOriginal_(numeroFactura) {
  const hoja = SpreadsheetApp.getActive().getSheetByName('ListadoEstado');
  if (!hoja || hoja.getLastRow() < 2) return null;
  const buscado = normalizarNumeroFactura_(numeroFactura);
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, LISTADO_COL_JSON).getValues();
  for (let i = filas.length - 1; i >= 0; i--) {
    if (normalizarNumeroFactura_(filas[i][LISTADO_COL_NUMERO - 1]) !== buscado) continue;
    const texto = filas[i][LISTADO_COL_JSON - 1];
    if (!texto) continue;
    try {
      const json = JSON.parse(texto);
      // Older rows store a legacy format without products; keep looking
      if (json && Array.isArray(json.products)) return json;
    } catch (e) {
      Logger.log('ListadoEstado row ' + (i + 2) + ' has invalid JSON: ' + e);
    }
  }
  return null;
}

/** Row data of an invoice in Historial Facturas Data, or null. */
function buscarFilaHistorial_(numeroFactura) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HISTORIAL_DATA_SHEET);
  if (!hoja || hoja.getLastRow() < 2) return null;
  const buscado = normalizarNumeroFactura_(numeroFactura);
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, HISTORIAL_COL_TIPO_RECT).getValues();
  for (let i = filas.length - 1; i >= 0; i--) {
    if (normalizarNumeroFactura_(filas[i][HISTORIAL_COL_NUMERO - 1]) === buscado) {
      return {
        numero: filas[i][HISTORIAL_COL_NUMERO - 1],
        cliente: filas[i][1],
        nif: filas[i][2],
        fecha: filas[i][HISTORIAL_COL_FECHA - 1],
        estado: filas[i][HISTORIAL_COL_ESTADO - 1],
        original: filas[i][HISTORIAL_COL_ORIGINAL - 1],
        tipoRectificativa: filas[i][HISTORIAL_COL_TIPO_RECT - 1]
      };
    }
  }
  return null;
}

/** Rectificativas already registered locally for an original invoice. */
function rectificativasLocalesDe_(numeroOriginal) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HISTORIAL_DATA_SHEET);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const buscado = normalizarNumeroFactura_(numeroOriginal);
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, HISTORIAL_COL_TIPO_RECT).getValues();
  return filas
    .filter(fila => normalizarNumeroFactura_(fila[HISTORIAL_COL_ORIGINAL - 1]) === buscado)
    .map(fila => ({ numero: fila[HISTORIAL_COL_NUMERO - 1], tipo: fila[HISTORIAL_COL_TIPO_RECT - 1] }));
}

/**
 * JSON a new rectificativa should start from: the latest sustitución made from
 * this book, or the original invoice when there is none. A sustitución replaces
 * the whole invoice, so its values are the current ones (FacturasApp's web
 * prefills the original instead). Returns { json, desde } where desde is the
 * sustitución number used, or ''.
 */
function jsonVigenteDeFactura_(numeroOriginal, existentes) {
  existentes = existentes || rectificativasLocalesDe_(numeroOriginal);
  const sustituciones = existentes.filter(r => r.tipo === RECTIFICATIVA_TIPO_LABEL.articulos);
  for (let i = sustituciones.length - 1; i >= 0; i--) {
    const json = buscarJsonFacturaOriginal_(sustituciones[i].numero);
    if (json) return { json: json, desde: String(sustituciones[i].numero) };
  }
  return { json: buscarJsonFacturaOriginal_(numeroOriginal), desde: '' };
}

// States that only exist in Verifactu accounts and block a rectificativa (the API
// requires Correcto / Parcialmente correcto). "Creada" is excluded on purpose:
// accounts without Verifactu keep every invoice in "Creada" and can rectify it.
const ESTADOS_BLOQUEANTES_VERIFACTU = ['pendiente envio aeat', 'incorrecto', 'fallida'];

function normalizarEstado_(texto) {
  return String(texto || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();
}

function mensajeNumeroOcupado_(numero) {
  return 'El siguiente número de tu serie de rectificativas (' + numero + ') ya está usado en FacturasApp. ' +
    'Revisa la configuración de consecutivos de rectificativas en la web.';
}

/**
 * Looks an invoice number up in FacturasApp. Unlike facturaExisteEnApi_, a
 * network or server error is reported as 'error', never as "free".
 * Returns { estado: 'existe'|'libre'|'error', nombre, error }.
 */
function consultarNumeroFactura_(numero) {
  const res = invoiceServicesFetch_('/GetInvoiceStateByNumber/' + encodeURIComponent(String(numero)), 'get');
  if (res.ok && res.data && res.data.toolObject) {
    return { estado: 'existe', nombre: String(res.data.toolObject.stateName || '').trim(), error: '' };
  }
  if (res.status === 404) return { estado: 'libre', nombre: '', error: '' };
  return { estado: 'error', nombre: '', error: res.error || ('Error ' + res.status) };
}

/**
 * Number the next rectificativa will be sent with: nextInvoiceNumber of the
 * default series (or the first one), checked to be unused.
 * Returns { ok:true, numero, prefijo } or { ok:false, motivo, numero, prefijo, mensaje }
 * with motivo 'existe' | 'sin_serie' | 'error_consulta'.
 */
function prepararNumeroRectificativa_() {
  let data;
  try {
    data = fetchConsecutivesFromAPI_();
  } catch (e) {
    Logger.log('prepararNumeroRectificativa_: GetConsecutives falló: ' + e);
    data = null;
  }
  if (!data) {
    return { ok: false, motivo: 'error_consulta', numero: '', prefijo: '',
      mensaje: 'No se pudo consultar el consecutivo de rectificativas en FacturasApp. Inténtalo de nuevo.' };
  }
  const lista = Array.isArray(data.facturasRectificativas) ? data.facturasRectificativas : [];
  const entrada = lista.find(e => e && e.isDefault === true) || lista[0];
  const numero = entrada && entrada.nextInvoiceNumber ? String(entrada.nextInvoiceNumber).trim() : '';
  const prefijo = entrada && entrada.prefijo ? String(entrada.prefijo) : '';
  if (!numero) {
    return { ok: false, motivo: 'sin_serie', numero: '', prefijo: prefijo,
      mensaje: 'No tienes una serie de rectificativas activa en FacturasApp. Configúrala en la web antes de continuar.' };
  }
  const consulta = consultarNumeroFactura_(numero);
  if (consulta.estado === 'existe') {
    return { ok: false, motivo: 'existe', numero: numero, prefijo: prefijo, mensaje: mensajeNumeroOcupado_(numero) };
  }
  if (consulta.estado === 'error') {
    return { ok: false, motivo: 'error_consulta', numero: numero, prefijo: prefijo,
      mensaje: 'No se pudo comprobar si el número ' + numero + ' está libre en FacturasApp. Inténtalo de nuevo.' };
  }
  return { ok: true, numero: numero, prefijo: prefijo };
}

/**
 * Non-blocking warning, shown only for Verifactu-only states (see
 * ESTADOS_BLOQUEANTES_VERIFACTU). Returns '' when there is nothing to say.
 */
function avisoEstadoVerifactu_(numero, consulta) {
  if (!consulta || consulta.estado !== 'existe') return '';
  const estado = String(consulta.nombre || '').trim();
  if (ESTADOS_BLOQUEANTES_VERIFACTU.indexOf(normalizarEstado_(estado)) < 0) return '';
  return 'La factura ' + numero + ' está en estado "' + estado + '" en FacturasApp. ' +
    'Solo se pueden rectificar facturas en estado Correcto o Parcialmente correcto; ' +
    'espera a que la AEAT la procese antes de crear la rectificativa.';
}

/**
 * Editable view of a stored product line (rates in percent). The dialog edits
 * these fields and recalcularLineaRectificativa_ rebuilds the API object.
 */
function lineaDesdeProductoOriginal_(p, indice) {
  const tax = (p.taxes && p.taxes[0]) || {};
  const retenciones = p.withHoldingsSurChargesDto || [];
  const retencion = retenciones.find(w => w.isWithHolding === true) || null;
  const recargo = retenciones.find(w => String(w.idRateWithHoldings) === RECT_ID_RECARGO) || null;
  const descuento = (p.discountDtoModules && p.discountDtoModules[0]) || null;
  const porcentaje = w => {
    if (!w) return 0;
    if (w.rateValueWithHoldings != null) return round2Rect_(Number(w.rateValueWithHoldings) * 100);
    const base = Number(w.subTotalWithHoldings) || 0;
    return base > 0 ? round2Rect_((Number(w.cuotaWithHoldings) || 0) / base * 100) : 0;
  };
  return {
    origen: indice,
    reference: String(p.reference || ''),
    referenciaOriginal: String(p.reference || ''),
    description: String(p.description || ''),
    quantity: Number(p.quantity) || 0,
    unitPrice: Number(p.unitPrice) || 0,
    ivaRate: round2Rect_(tax.rate),
    taxCode: tax.taxCode || '01',
    taxName: tax.taxName || 'IVA',
    regime: tax.regime || p.regime || '01',
    isExempt: tax.isExemptOperation === true,
    idExento: tax.idExento || null,
    qualificationOperation: tax.qualificationOperation || null,
    productType: p.productType,
    typeUse: p.typeUse || 'VEN',
    descuentoRate: descuento ? round2Rect_(descuento.discountRate) : 0,
    retencionRate: porcentaje(retencion),
    retencionTipo: retencion ? (String(retencion.idRateWithHoldings) === RECT_ID_RETENCION_OTROS ? 'otros' : 'retencion') : null,
    retencionDescripcion: retencion && retencion.otherDescription ? String(retencion.otherDescription) : '',
    recargoRate: porcentaje(recargo)
  };
}

/** Read-only contact block for the dialog, from the stored invoice contact. */
function contactoDesdeOriginal_(original) {
  const c = (original && original.contacts && original.contacts[0]) || {};
  const tieneRecargo = (original.products || []).some(p =>
    (p.withHoldingsSurChargesDto || []).some(w => String(w.idRateWithHoldings) === RECT_ID_RECARGO));
  return {
    nombre: String(c.companyName || original.contactName || ''),
    identificacion: String(c.identification || ''),
    codigo: String(c.customerCode || ''),
    direccion: String(c.addressCustomer || ''),
    codigoPostal: String(c.postalCodeCustomer || ''),
    email: String(c.emailCustomer || ''),
    telefono: String(c.phoneCustomer || ''),
    pais: String(c.country || ''),
    provincia: String(c.province || ''),
    poblacion: String(c.population || ''),
    // Recargo de equivalencia only applies to autónomos (personType 01)
    aplicaRecargo: c.applySurchargeEquivalence === true || String(c.personType) === '01' || tieneRecargo
  };
}

/** Invoice-level charges/discounts/IRPF of the original, as the dialog edits them. */
function cargosDescuentosDesdeOriginal_(original) {
  return (original.chargeAndDiscount || []).map(c => ({
    tipo: c.idtypeFeeDiscount,
    tipoValor: c.idTypeValueFeeDiscount === 'PJ' ? 'PJ' : 'VR',
    valor: round2Rect_(c.valueFeeDiscount),
    total: round2Rect_(c.totalFeeDiscount)
  }));
}

/** Totals block from any FieldInvoice-shaped object (original or new payload). */
function totalesDesdePayload_(payload) {
  const retenciones = (payload.products || []).reduce((acc, p) => round2Rect_(acc + (Number(p.totalwithHoldings) || 0)), 0);
  return {
    subtotal: round2Rect_(payload.sumTotalSubTotal),
    descuentosLinea: round2Rect_((Number(payload.sumTotalSubTotal) || 0) - (Number(payload.sumTotalTaxBase) || 0)),
    baseImponible: round2Rect_(payload.sumTotalTaxBase),
    iva: round2Rect_(payload.sumTotalTax),
    recargo: round2Rect_((Number(payload.sumTotalTotal) || 0) - (Number(payload.sumTotalSubTotalAndTax) || 0)),
    retenciones: retenciones,
    cargos: round2Rect_(payload.sumTotalCharge),
    descuentoFactura: round2Rect_(payload.sumTotalDiscount),
    irpf: round2Rect_(payload.sumTotalRetentionIRPF),
    total: round2Rect_(payload.sumTotalTotal),
    neto: round2Rect_(payload.sumTotalNetPayable),
    baseExenta: round2Rect_(payload.sumTotalExemptBase)
  };
}

/** Display shape of a payload for the "Totales" section and the preview. */
function resumenDesdePayload_(payload) {
  const lineas = (payload.products || []).map(p => ({
    reference: String(p.reference || ''),
    importe: round2Rect_((Number(p.subTotal) || 0) - (Number(p.totaldiscount) || 0) +
      (Number(p.totalTax) || 0) + (Number(p.totalSurCharges) || 0) - (Number(p.totalwithHoldings) || 0))
  }));
  const impuestos = (payload.fieldTaxations || []).map(t => ({
    tipo: t.taxName === 'RecargoEquivalencia' ? 'Rec. equivalencia' : String(t.taxName || 'Impuesto'),
    tarifa: round2Rect_(t.rate),
    base: round2Rect_(t.taxBase),
    total: round2Rect_(t.valueTax),
    exento: !!t.idExento
  }));
  return {
    lineas: lineas,
    impuestos: impuestos,
    cargosDescuentos: cargosDescuentosDesdeOriginal_(payload),
    totales: totalesDesdePayload_(payload)
  };
}

/**
 * Everything the sidebar step and the dialog need to rectify an invoice. Never
 * throws: the UI shows `error` when ok is false.
 * opciones.sinConsecutivo skips the GetConsecutives call (preview recalculations).
 */
function cargarFacturaParaRectificar(numeroFactura, opciones) {
  opciones = opciones || {};
  const numero = String(numeroFactura || '').trim();
  if (!numero) return { ok: false, error: 'Ingresa el número de la factura a rectificar.' };
  if (!getApiKey_()) return { ok: false, error: 'La cuenta no está vinculada a FacturasApp.' };

  const historial = buscarFilaHistorial_(numero);
  if (!historial) {
    return { ok: false, error: 'La factura ' + numero + ' no está en el historial de este libro. Solo se pueden rectificar facturas enviadas desde aquí.' };
  }
  if (historial.tipoRectificativa) {
    return { ok: false, error: 'La factura ' + numero + ' ya es una rectificativa (de ' + historial.original + ') y no se puede rectificar de nuevo.' };
  }

  const tz = getSpreadsheetTimeZone_();
  const ahora = new Date();
  const existentes = rectificativasLocalesDe_(numero);
  const vigente = jsonVigenteDeFactura_(numero, existentes);
  const original = vigente.json;
  const tieneCompleta = existentes.some(r => r.tipo === RECTIFICATIVA_TIPO_LABEL.completa);

  // Number to send (mode B) and AEAT state of the original. Skipped for preview
  // recalculations, which run on every keystroke.
  let numeracion = null;
  let consecutivo = { prefijo: '', siguiente: '' };
  let avisoVerifactu = '';
  if (!opciones.sinConsecutivo) {
    numeracion = prepararNumeroRectificativa_();
    consecutivo = { prefijo: numeracion.prefijo || '', siguiente: numeracion.numero || '' };
    try {
      avisoVerifactu = avisoEstadoVerifactu_(numero, consultarNumeroFactura_(numero));
    } catch (e) {
      Logger.log('No se pudo consultar el estado de ' + numero + ': ' + e);
    }
  }

  const info = {
    ok: true,
    numero: String(historial.numero),
    cliente: String(historial.cliente || ''),
    nif: String(historial.nif || ''),
    fecha: historial.fecha instanceof Date
      ? Utilities.formatDate(historial.fecha, tz, 'dd/MM/yyyy')
      : String(historial.fecha || ''),
    estado: String(historial.estado || ''),
    tieneJson: !!original,
    tieneCompleta: tieneCompleta,
    rectificativas: existentes.map(r => String(r.numero)),
    datosDesde: vigente.desde,
    siguienteNumero: consecutivo.siguiente,
    consecutivo: consecutivo,
    numeracion: numeracion,
    avisoVerifactu: avisoVerifactu,
    motivosAeat: RECTIFICATIVA_MOTIVOS_AEAT,
    hoy: {
      fecha: Utilities.formatDate(ahora, tz, 'yyyy-MM-dd'),
      hora: Utilities.formatDate(ahora, tz, 'HH:mm')
    },
    emisor: leerDatosEmisor_(),
    formasPago: getFormasPago_(),
    obsMax: RECTIFICATIVA_OBS_MAX,
    referenciaMax: RECT_REFERENCIA_MAX,
    descripcionMax: RECT_DESCRIPCION_MAX,
    motivoMax: RECTIFICATIVA_MOTIVO_MAX,
    lineas: [],
    contacto: null,
    pago: null,
    observaciones: null,
    cargosDescuentos: [],
    totales: null,
    totalesAnulacion: null
  };

  if (original) {
    info.lineas = original.products.map(lineaDesdeProductoOriginal_);
    info.contacto = contactoDesdeOriginal_(original);
    info.pago = { idPayment: original.idPayment || 'ND', paymentNote: String(original.paymentNote || '') };
    info.observaciones = {
      cliente: String(original.textCustomerObservations || ''),
      factura: String(original.textObservations || '')
    };
    info.cargosDescuentos = cargosDescuentosDesdeOriginal_(original);
    info.totales = totalesDesdePayload_(original);
    // AddRectificativeComplete always duplicates the alta in negative, even after
    // a sustitución, so the anulación shows the alta amounts.
    const alta = vigente.desde ? buscarJsonFacturaOriginal_(numero) : original;
    info.totalesAnulacion = alta ? totalesDesdePayload_(alta) : null;
    if (!info.nif && original.contacts && original.contacts[0]) info.nif = String(original.contacts[0].identification || '');
  }
  return info;
}

/** Edited text when the dialog sent one, else the stored value. */
function textoEditado_(editado, guardado) {
  const texto = String(editado == null ? '' : editado).trim();
  return texto || String(guardado || '');
}

/**
 * Stored product an edited line comes from: by position (origen) since the
 * reference itself can be edited, by reference for callers without origen.
 */
function productoOrigenDeLinea_(productos, linea) {
  const origen = Number(linea.origen);
  if (linea.origen !== undefined && linea.origen !== null && linea.origen !== '' &&
      Number.isInteger(origen) && origen >= 0 && origen < productos.length) {
    return { producto: productos[origen], clave: 'i' + origen };
  }
  const ref = String(linea.reference || '').trim().toUpperCase();
  for (let i = 0; i < productos.length; i++) {
    const p = productos[i];
    const r = String((p && (p.referenciaOriginal || p.reference)) || '').trim().toUpperCase();
    if (r === ref) return { producto: p, clave: 'i' + i };
  }
  return null;
}

/**
 * Rebuilds one product object of the API from the stored line and the values
 * edited in the dialog (rates in percent). Pure.
 * Mirrors the per-line maths of guardarYGenerarInvoice (Factura.js, "Procesar
 * productos"): round2 at every step, subTotal is gross, taxes on the net base.
 */
function recalcularLineaRectificativa_(producto, edicion) {
  const tax = (producto.taxes && producto.taxes[0]) || {};
  const isExempt = tax.isExemptOperation === true;
  const cantidad = Math.max(1, Math.trunc(Number(edicion.quantity) || 0));
  const precio = round2Rect_(edicion.unitPrice);
  const ivaPct = isExempt ? 0 : round2Rect_(edicion.ivaRate);
  const dtoPct = round2Rect_(edicion.descuentoRate);
  const retPct = round2Rect_(edicion.retencionRate);
  const recPct = round2Rect_(edicion.recargoRate);

  const baseBruta = round2Rect_(precio * cantidad);
  const discountAmount = round2Rect_(baseBruta * dtoPct / 100);
  const baseNeta = round2Rect_(baseBruta - discountAmount);
  const taxAmount = round2Rect_(baseNeta * ivaPct / 100);
  const withHoldingsAmount = round2Rect_(baseNeta * retPct / 100);
  const surChargesAmount = round2Rect_(baseNeta * recPct / 100);

  // Regime "18" is the recargo override; drop it when the surcharge is removed
  let regime = String(tax.regime || producto.regime || '01');
  if (recPct > 0) regime = '18';
  else if (regime === '18') regime = '01';

  const taxEntry = {
    taxName: tax.taxName || 'IVA',
    rate: ivaPct,
    taxBase: baseNeta,
    valueTax: taxAmount,
    taxCode: tax.taxCode || '01',
    isExemptOperation: isExempt,
    qualificationOperation: isExempt ? null : (tax.qualificationOperation || 'S1'),
    regime: regime
  };
  if (isExempt && tax.idExento) taxEntry.idExento = tax.idExento;

  const withHoldingsSurChargesDto = [];
  if (retPct > 0) {
    const retOriginal = (producto.withHoldingsSurChargesDto || []).find(w => w.isWithHolding === true) || null;
    const retDto = {
      isWithHolding: true,
      idRateWithHoldings: retOriginal ? String(retOriginal.idRateWithHoldings) : RECT_ID_RETENCION,
      rateValueWithHoldings: fraccionRect_(retPct),
      subTotalWithHoldings: baseNeta,
      cuotaWithHoldings: withHoldingsAmount
    };
    if (retOriginal && retOriginal.otherDescription) retDto.otherDescription = String(retOriginal.otherDescription);
    withHoldingsSurChargesDto.push(retDto);
  }
  if (recPct > 0) {
    withHoldingsSurChargesDto.push({
      isWithHolding: false,
      idRateWithHoldings: RECT_ID_RECARGO,
      rateValueWithHoldings: fraccionRect_(recPct),
      subTotalWithHoldings: baseNeta,
      cuotaWithHoldings: surChargesAmount
    });
  }

  const discountDtoModules = [];
  if (dtoPct > 0) {
    discountDtoModules.push({
      discountName: 'Descuento aplicado',
      discountRate: dtoPct,
      discountBase: baseBruta,
      valueDiscount: discountAmount
    });
  }

  return {
    typeUse: producto.typeUse || 'VEN',
    reference: String(textoEditado_(edicion.reference, producto.reference)).substring(0, 50),
    description: String(textoEditado_(edicion.description, producto.description)).substring(0, 300),
    regime: regime,
    productType: producto.productType,
    unitPrice: precio,
    quantity: cantidad,
    subTotal: baseBruta,
    totalTax: taxAmount,
    totalwithHoldings: withHoldingsAmount,
    totalSurCharges: surChargesAmount,
    totaldiscount: discountAmount,
    taxes: [taxEntry],
    withHoldingsSurChargesDto: withHoldingsSurChargesDto,
    discountDtoModules: discountDtoModules
  };
}

/**
 * Builds the AddRectificativeArticles payload from the original invoice JSON and
 * the lines edited in the dialog. Pure (no Sheets access) so it can be unit tested.
 *
 * @param {Object} original        JSON stored when the original was sent (FieldInvoice)
 * @param {Array}  lineasEditadas  [{ origen, reference, description, quantity, unitPrice, ivaRate, descuentoRate, retencionRate, recargoRate }]
 *                                 origen = index in original.products; reference and description may be edited
 *                                 (like the web, they only travel on this invoice, no product is created)
 * @param {Object} opciones        { numeroOriginal, invoiceNumber, motivoAeat, motivoTexto, invoiceDate, invoiceTime,
 *                                   operationDate, observacionesCliente, idPayment, paymentNote,
 *                                   textObservations, cargosDescuentos: [{ tipo, tipoValor, valor }] }
 */
function construirPayloadRectificativaArticulos_(original, lineasEditadas, opciones) {
  if (!original || !Array.isArray(original.products) || original.products.length === 0) {
    throw new Error('La factura original no tiene líneas guardadas.');
  }
  if (!Array.isArray(lineasEditadas) || lineasEditadas.length === 0) {
    throw new Error('La rectificativa debe conservar al menos una línea.');
  }

  const products = [];
  const vistas = {};
  lineasEditadas.forEach(linea => {
    const encontrado = productoOrigenDeLinea_(original.products, linea);
    if (!encontrado) throw new Error('El producto ' + linea.reference + ' no corresponde a ninguna línea de la factura original.');
    if (vistas[encontrado.clave]) throw new Error('El producto ' + linea.reference + ' está repetido.');
    vistas[encontrado.clave] = true;
    products.push(recalcularLineaRectificativa_(encontrado.producto, linea));
  });

  // Totals mirror guardarYGenerarInvoice(): subTotal is gross, taxBase is net
  let totalSubTotal = 0, totalTaxBase = 0, totalTax = 0, totalWithHoldings = 0, totalSurCharges = 0, totalExemptBase = 0;
  let firstIdExento = null;
  const taxGroups = {};
  const recargoGroups = {};
  products.forEach(p => {
    const baseNeta = round2Rect_(p.subTotal - p.totaldiscount);
    totalSubTotal = round2Rect_(totalSubTotal + p.subTotal);
    totalTaxBase = round2Rect_(totalTaxBase + baseNeta);
    totalTax = round2Rect_(totalTax + p.totalTax);
    totalWithHoldings = round2Rect_(totalWithHoldings + p.totalwithHoldings);
    totalSurCharges = round2Rect_(totalSurCharges + p.totalSurCharges);
    (p.taxes || []).forEach(t => {
      if (t.isExemptOperation === true) {
        totalExemptBase = round2Rect_(totalExemptBase + t.taxBase);
        if (!firstIdExento && t.idExento) firstIdExento = t.idExento;
      }
      const key = [t.taxCode || '', t.rate, t.isExemptOperation ? 'ex' : ''].join('_');
      if (!taxGroups[key]) {
        taxGroups[key] = { taxName: t.taxName, rate: t.rate, taxBase: 0, valueTax: 0,
          taxCode: t.taxCode, qualificationOperation: t.qualificationOperation || null, regime: t.regime };
        if (t.idExento) taxGroups[key].idExento = t.idExento;
      }
      taxGroups[key].taxBase = round2Rect_(taxGroups[key].taxBase + t.taxBase);
      taxGroups[key].valueTax = round2Rect_(taxGroups[key].valueTax + t.valueTax);
    });
    (p.withHoldingsSurChargesDto || []).forEach(w => {
      if (String(w.idRateWithHoldings) !== RECT_ID_RECARGO || !(w.subTotalWithHoldings > 0)) return;
      const rate = round2Rect_((w.rateValueWithHoldings != null ? Number(w.rateValueWithHoldings) : w.cuotaWithHoldings / w.subTotalWithHoldings) * 100);
      const taxCode = (p.taxes && p.taxes[0] && p.taxes[0].taxCode) || '01';
      const key = taxCode + '_' + rate;
      if (!recargoGroups[key]) {
        recargoGroups[key] = { taxName: 'RecargoEquivalencia', rate: rate, taxBase: 0, valueTax: 0, taxCode: taxCode, regime: '18' };
      }
      recargoGroups[key].taxBase = round2Rect_(recargoGroups[key].taxBase + w.subTotalWithHoldings);
      recargoGroups[key].valueTax = round2Rect_(recargoGroups[key].valueTax + w.cuotaWithHoldings);
    });
  });

  // Invoice-level charges (CG), discounts (DT) and IRPF (RT) as edited in the
  // dialog. PJ = percentage of the new net base, VR = fixed amount.
  let cargoTotal = 0, descuentoFactura = 0, irpfGlobal = 0;
  const chargeAndDiscount = (opciones.cargosDescuentos || []).map(c => {
    const tipo = String(c.tipo || '').toUpperCase();
    const valor = round2Rect_(c.valor);
    const esPorcentaje = String(c.tipoValor || '').toUpperCase() === 'PJ';
    const entrada = {
      idtypeFeeDiscount: tipo,
      idTypeValueFeeDiscount: esPorcentaje ? 'PJ' : 'VR',
      baseFeeDiscount: esPorcentaje ? totalTaxBase : 0,
      valueFeeDiscount: valor,
      totalFeeDiscount: esPorcentaje ? round2Rect_(totalTaxBase * valor / 100) : valor
    };
    if (tipo === 'CG') cargoTotal = round2Rect_(cargoTotal + entrada.totalFeeDiscount);
    if (tipo === 'DT') descuentoFactura = round2Rect_(descuentoFactura + entrada.totalFeeDiscount);
    if (tipo === 'RT') irpfGlobal = round2Rect_(irpfGlobal + entrada.totalFeeDiscount);
    return entrada;
  }).filter(c => c.totalFeeDiscount > 0);

  const sumTotalSubTotalAndTax = round2Rect_(totalTaxBase + totalTax);
  const sumTotalTotal = round2Rect_(sumTotalSubTotalAndTax + totalSurCharges);
  const sumTotalNetPayable = round2Rect_(sumTotalTotal - totalWithHoldings - descuentoFactura + cargoTotal - irpfGlobal);

  const fieldTaxations = Object.keys(taxGroups).map(k => taxGroups[k])
    .concat(Object.keys(recargoGroups).map(k => recargoGroups[k]));

  const limitar = (texto, max) => {
    const t = String(texto || '').trim();
    return t ? t.substring(0, max) : null;
  };
  const motivoTexto = limitar(opciones.motivoTexto, RECTIFICATIVA_MOTIVO_MAX);
  return {
    originalInvoiceNumber: String(opciones.numeroOriginal),
    cancellationReason: motivoTexto,
    typeRectificative: RECTIFICATIVA_TIPO,
    invoiceRectificativeType: String(opciones.motivoAeat),
    // Sent explicitly (mode B); undefined keys are dropped by JSON.stringify
    invoiceNumber: opciones.invoiceNumber ? String(opciones.invoiceNumber) : undefined,
    textCustomerObservations: limitar(opciones.observacionesCliente, RECTIFICATIVA_OBS_MAX),
    invoiceDate: opciones.invoiceDate,
    invoiceTime: opciones.invoiceTime,
    operationDate: opciones.operationDate,
    invoiceExpiration: original.invoiceExpiration != null ? String(original.invoiceExpiration) : null,
    invoiceIdTypeRegAEAT: original.invoiceIdTypeRegAEAT || 'AI',
    invoiceIdTypeRegSIF: original.invoiceIdTypeRegSIF || null,
    contactName: original.contactName || '',
    contacts: JSON.parse(JSON.stringify(original.contacts || [])),
    products: products,
    idPayment: opciones.idPayment || original.idPayment || 'ND',
    paymentNote: limitar(opciones.paymentNote, RECTIFICATIVA_OBS_MAX),
    textObservations: limitar(opciones.textObservations, RECTIFICATIVA_OBS_MAX),
    idOperations: original.idOperations || 'S1',
    idOperationsExenta: totalExemptBase > 0 ? (firstIdExento || original.idOperationsExenta || 'E1') : 'E0',
    valueExemptBase: totalExemptBase,
    chargeAndDiscount: chargeAndDiscount,
    fieldTaxations: fieldTaxations,
    sumTotalSubTotal: totalSubTotal,
    sumTotalTaxBase: totalTaxBase,
    sumTotalTax: totalTax,
    sumTotalSubTotalAndTax: sumTotalSubTotalAndTax,
    sumTotalExemptBase: totalExemptBase,
    sumTotalDiscount: descuentoFactura,
    sumTotalCharge: cargoTotal,
    sumTotalRetentionIRPF: irpfGlobal,
    sumTotalTotal: sumTotalTotal,
    sumTotalNetPayable: sumTotalNetPayable
  };
}

/** Keeps the surcharge flag of the stored contact consistent with the new lines. */
function ajustarRecargoEnContactos_(payload) {
  const tieneRecargo = (payload.products || []).some(p => (p.totalSurCharges || 0) > 0);
  (payload.contacts || []).forEach(c => { c.applySurchargeEquivalence = tieneRecargo; });
}

function esFechaIso_(texto) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(texto || ''));
}

function esHora_(texto) {
  return /^\d{2}:\d{2}$/.test(String(texto || ''));
}

/**
 * Validates the dialog data against the loaded invoice. Returns '' when valid
 * or the message to show. Shared by the preview and the send path.
 */
function validarDatosRectificativa_(datos, info) {
  const tipo = datos.tipo === 'articulos' ? 'articulos' : 'completa';
  const motivoAeat = String(datos.motivoAeat || '').trim().toUpperCase();
  const motivoTexto = String(datos.motivoTexto || '').trim();
  if (!RECTIFICATIVA_MOTIVOS_AEAT.some(m => m.codigo === motivoAeat)) return 'Selecciona el tipo de factura rectificativa (R1 a R4).';
  if (motivoTexto.length > RECTIFICATIVA_MOTIVO_MAX) return 'El motivo no puede superar ' + RECTIFICATIVA_MOTIVO_MAX + ' caracteres.';
  if (tipo === 'completa' && !motivoTexto) return 'Describe el motivo de la anulación.';
  if (datos.enviarEmail) {
    const email = String(datos.email || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return 'Ingresa un correo electrónico válido para el envío.';
  }
  if (datos.fechaOperacion && !esFechaIso_(datos.fechaOperacion)) return 'La fecha de operación no es válida.';
  if (tipo === 'completa') return '';

  if (!info.tieneJson) return 'No se encontraron las líneas de la factura ' + info.numero + ' en este libro. Usa la rectificativa de anulación.';
  if (!esFechaIso_(datos.fechaEmision)) return 'La fecha de emisión no es válida.';
  if (!esHora_(datos.horaEmision)) return 'La hora de emisión no es válida.';
  const lineas = Array.isArray(datos.lineas) ? datos.lineas : [];
  if (lineas.length === 0) return 'La rectificativa debe conservar al menos una línea.';
  const vistas = {};
  for (let i = 0; i < lineas.length; i++) {
    const l = lineas[i];
    const encontrado = productoOrigenDeLinea_(info.lineas || [], l);
    if (!encontrado) return 'El producto ' + l.reference + ' no corresponde a ninguna línea de la factura original.';
    if (vistas[encontrado.clave]) return 'El producto ' + l.reference + ' está repetido.';
    vistas[encontrado.clave] = true;
    const referencia = String(l.reference || '').trim();
    if (!referencia) return 'La referencia del producto ' + (i + 1) + ' no puede estar vacía.';
    if (referencia.length > RECT_REFERENCIA_MAX) return 'La referencia ' + referencia + ' no puede superar ' + RECT_REFERENCIA_MAX + ' caracteres.';
    if (String(l.description || '').trim().length > RECT_DESCRIPCION_MAX) return 'La descripción de ' + referencia + ' no puede superar ' + RECT_DESCRIPCION_MAX + ' caracteres.';
    const cantidad = Number(l.quantity);
    if (!isFinite(cantidad) || cantidad < 1 || Math.trunc(cantidad) !== cantidad) return 'La cantidad de ' + l.reference + ' debe ser un entero mayor o igual a 1.';
    const precio = Number(l.unitPrice);
    if (!isFinite(precio) || precio < 0) return 'El precio de ' + l.reference + ' no es válido.';
    const pct = (v, nombre) => {
      const n = Number(v);
      return (!isFinite(n) || n < 0 || n > 100) ? 'El porcentaje de ' + nombre + ' de ' + l.reference + ' debe estar entre 0 y 100.' : '';
    };
    const errPct = pct(l.ivaRate, 'IVA') || pct(l.descuentoRate, 'descuento') || pct(l.retencionRate, 'retención') || pct(l.recargoRate, 'recargo');
    if (errPct) return errPct;
    if (Number(l.recargoRate) > 0 && info.contacto && !info.contacto.aplicaRecargo) {
      return 'El recargo de equivalencia solo aplica a clientes autónomos.';
    }
  }
  const cargos = Array.isArray(datos.cargosDescuentos) ? datos.cargosDescuentos : [];
  let irpfs = 0;
  for (let k = 0; k < cargos.length; k++) {
    const c = cargos[k];
    const tipoCargo = String(c.tipo || '').toUpperCase();
    if (['CG', 'DT', 'RT'].indexOf(tipoCargo) < 0) return 'Tipo de cargo o descuento no válido.';
    const valor = Number(c.valor);
    if (!isFinite(valor) || valor < 0) return 'El valor de los cargos y descuentos debe ser mayor o igual a 0.';
    if (String(c.tipoValor || '').toUpperCase() === 'PJ' && valor > 100) return 'Un porcentaje de cargo o descuento no puede superar 100.';
    if (tipoCargo === 'RT') irpfs++;
  }
  if (irpfs > 1) return 'Solo puede haber una retención IRPF a nivel de factura.';
  const limiteObs = v => String(v || '').trim().length > RECTIFICATIVA_OBS_MAX;
  if (limiteObs(datos.observacionesCliente) || limiteObs(datos.paymentNote) || limiteObs(datos.textObservations)) {
    return 'Las observaciones no pueden superar ' + RECTIFICATIVA_OBS_MAX + ' caracteres.';
  }
  return '';
}

/** Dialog form -> opciones for construirPayloadRectificativaArticulos_. */
function opcionesPayloadDesdeDatos_(datos, numeroOriginal) {
  const tz = getSpreadsheetTimeZone_();
  const emision = Utilities.parseDate(datos.fechaEmision + ' ' + datos.horaEmision, tz, 'yyyy-MM-dd HH:mm');
  const operacion = esFechaIso_(datos.fechaOperacion)
    ? Utilities.parseDate(datos.fechaOperacion, tz, 'yyyy-MM-dd')
    : emision;
  return {
    numeroOriginal: numeroOriginal,
    motivoAeat: String(datos.motivoAeat || '').trim().toUpperCase(),
    motivoTexto: String(datos.motivoTexto || '').trim(),
    invoiceDate: formatInvoiceDateForPortal_(emision, tz),
    invoiceTime: Utilities.formatDate(emision, tz, 'HH:mm') + ':00.0000000',
    operationDate: Utilities.formatDate(operacion, tz, 'dd-MM-yyyy'),
    observacionesCliente: datos.observacionesCliente,
    idPayment: datos.idPayment,
    paymentNote: datos.paymentNote,
    textObservations: datos.textObservations,
    cargosDescuentos: datos.cargosDescuentos || []
  };
}

/**
 * Preview for the dialog: same validation and payload builder as the send path,
 * without calling the API or writing anything. Returns { ok, resumen } or
 * { ok:false, error }.
 */
function calcularTotalesRectificativa(datos) {
  try {
    datos = datos || {};
    const numeroOriginal = String(datos.numeroOriginal || '').trim();
    const info = cargarFacturaParaRectificar(numeroOriginal, { sinConsecutivo: true });
    if (!info.ok) return { ok: false, error: info.error };
    const error = validarDatosRectificativa_(datos, info);
    if (error) return { ok: false, error: error };
    if (datos.tipo !== 'articulos') {
      const original = buscarJsonFacturaOriginal_(numeroOriginal);
      return { ok: true, resumen: original ? resumenDesdePayload_(original) : null };
    }
    const original = jsonVigenteDeFactura_(numeroOriginal).json;
    const payload = construirPayloadRectificativaArticulos_(original, datos.lineas, opcionesPayloadDesdeDatos_(datos, numeroOriginal));
    return { ok: true, resumen: resumenDesdePayload_(payload) };
  } catch (error) {
    const mensaje = String(error && error.message ? error.message : error);
    Logger.log('calcularTotalesRectificativa error: ' + mensaje);
    return { ok: false, error: mensaje };
  }
}

// Waits before each existence check after a gateway timeout
const RECTIFICATIVA_ESPERAS_TIMEOUT_MS = [3000, 6000];

/**
 * Sends the rectificativa with its explicit invoiceNumber and confirms it by that
 * exact number when the gateway times out. Never retries the POST.
 * Returns { ok, numero, mensaje, nuevoNumero, reconfirmar }.
 */
function enviarRectificativa_(path, payload) {
  const numero = String(payload.invoiceNumber || '');
  const res = invoiceServicesFetch_(path, 'post', payload);
  if (res.ok) {
    const asignado = (res.data && typeof res.data.toolObject === 'string' && res.data.toolObject) || numero;
    return { ok: true, numero: asignado, mensaje: '' };
  }

  const esTimeout = res.status === 0 || res.status === 502 || res.status === 503 || res.status === 504;
  if (esTimeout && numero) {
    // The API may have created the invoice before the gateway gave up (seen on QA).
    let huboError = false;
    for (let i = 0; i < RECTIFICATIVA_ESPERAS_TIMEOUT_MS.length; i++) {
      Utilities.sleep(RECTIFICATIVA_ESPERAS_TIMEOUT_MS[i]);
      const consulta = consultarNumeroFactura_(numero);
      if (consulta.estado === 'existe') {
        return { ok: true, numero: numero, mensaje: 'confirmada tras un tiempo de espera del servidor' };
      }
      if (consulta.estado === 'error') huboError = true;
    }
    return { ok: false, numero: '', mensaje: huboError
      ? 'El servidor de FacturasApp no respondió a tiempo y no se pudo confirmar si la rectificativa ' + numero + ' se creó. Revisa el listado de facturas antes de reintentar.'
      : 'El servidor de FacturasApp no respondió a tiempo y la rectificativa ' + numero + ' no aparece registrada. Revisa el listado de facturas antes de reintentar.' };
  }

  let mensaje = res.error || 'Error al crear la rectificativa.';
  if (/ya existe una factura rectificativa/i.test(mensaje)) {
    mensaje = 'La factura original ya tiene una rectificativa completa en FacturasApp. No se pueden crear más rectificativas sobre ella.';
  } else if (/ya existe una factura con el n[uú]mero/i.test(mensaje)) {
    // Someone took the number between our check and the POST. Do not retry:
    // re-read the series so the dialog can show the new number.
    const nueva = prepararNumeroRectificativa_();
    mensaje = 'FacturasApp indica que el número ' + (numero || '') + ' ya existe; puede que se haya creado otra rectificativa al mismo tiempo. No se creó nada. ' +
      (nueva.ok ? 'El siguiente número disponible es ' + nueva.numero + '. Revisa y confirma de nuevo.' : nueva.mensaje);
    return { ok: false, numero: '', mensaje: mensaje, nuevoNumero: nueva.ok ? nueva.numero : '', reconfirmar: nueva.ok };
  } else if (/solo se pueden incluir productos/i.test(mensaje)) {
    mensaje = 'FacturasApp no aceptó alguna referencia de producto (solo admite referencias de la factura original). Revisa las referencias editadas. No se creó nada. Detalle: ' + mensaje;
  } else if (/para verifactu/i.test(mensaje)) {
    // Keep FacturasApp's own reason, plus what the user can do about it
    mensaje = 'FacturasApp rechazó la rectificativa: ' + mensaje.trim() +
      ' La factura original aún no ha sido aceptada por la AEAT; espera a que la procese e inténtalo de nuevo. No se creó nada.';
  }
  return { ok: false, numero: '', mensaje: mensaje };
}

/** Appends the rectificativa to Historial Facturas Data, linked to its original. */
function registrarRectificativaEnHistorial_(numero, info, tipo) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HISTORIAL_DATA_SHEET);
  if (!hoja) return;
  if (!hoja.getRange(1, HISTORIAL_COL_ORIGINAL).getValue()) {
    hoja.getRange(1, HISTORIAL_COL_ORIGINAL, 1, 2).setValues([['Factura original', 'Tipo rectificativa']]);
  }
  const fila = hoja.getLastRow() + 1;
  const valores = [[numero, info.cliente || '', info.nif || '', new Date(), 'Creada', '', info.numero, RECTIFICATIVA_TIPO_LABEL[tipo]]];
  const rango = hoja.getRange(fila, 1, 1, HISTORIAL_COL_TIPO_RECT);
  rango.setValues(valores);
  hoja.getRange(fila, 1, 1, 6).setHorizontalAlignment('center').setBorder(true, true, true, true, null, null, null, null);
}

/**
 * Appends the rectificativa to ListadoEstado (same row shape as
 * guardarYGenerarInvoice) so "Listado de facturas" can download its PDF and
 * email it by number.
 */
function registrarRectificativaEnListadoEstado_(numero, info, payload) {
  try {
    const hoja = SpreadsheetApp.getActive().getSheetByName('ListadoEstado');
    if (!hoja) return;
    const fecha = Utilities.formatDate(new Date(), getSpreadsheetTimeZone_(), 'dd/MM/yyyy');
    const codigoCliente = (info.contacto && info.contacto.codigo) || '';
    hoja.appendRow([
      'vacio', 'vacio', 'vacio', fecha, 'vacio', numero,
      info.cliente || '', codigoCliente, 'vacio', 'vacio', 'representacion',
      'Vacio', JSON.stringify(payload), ''
    ]);
  } catch (e) {
    Logger.log('No se pudo registrar la rectificativa en ListadoEstado: ' + e);
  }
}

/**
 * Emails the PDF of a rectificativa to the customer. Returns { ok, mensaje }.
 * Sibling of enviarEmailPostFactura(): same PDF source, copy adapted to a
 * corrective invoice.
 */
function enviarEmailRectificativa(numero, email, numeroOriginal) {
  try {
    numero = String(numero || '').trim();
    email = String(email || '').trim();
    if (!numero) return { ok: false, mensaje: 'Falta el número de la rectificativa.' };
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, mensaje: 'Ingresa un correo electrónico válido.' };

    const base64PDF = obtenerPDFFacturaBase64(numero);
    if (!base64PDF) return { ok: false, mensaje: 'FacturasApp todavía está generando el PDF de ' + numero + '. Inténtalo de nuevo en unos segundos desde el listado de facturas.' };
    const pdfBlob = Utilities.newBlob(Utilities.base64Decode(base64PDF), 'application/pdf', 'Rectificativa_' + numero + '.pdf');

    const emisor = leerDatosEmisor_().nombre || 'tu proveedor';
    const original = numeroOriginal ? ' sobre la factura ' + numeroOriginal : '';
    MailApp.sendEmail({
      to: email,
      subject: 'Factura rectificativa ' + numero + ' de ' + emisor,
      body: '¡Hola!\n' +
        emisor + ' te ha enviado la factura rectificativa ' + numero + original + '.\n' +
        'Encontrarás el documento adjunto en PDF.\n' +
        'Si tienes alguna duda, contacta directamente con ' + emisor + '.\n' +
        'Saludos,\n' + emisor + '\n\n' +
        '📌 ¿Necesitas facturación electrónica? Ahorra tiempo y factura fácilmente con FacturasApp\n' +
        '👉 Ver más: https://www.facturasapp.com/Publico/',
      attachments: [pdfBlob]
    });
    return { ok: true, mensaje: 'PDF enviado por correo electrónico a ' + email + '.' };
  } catch (error) {
    const mensaje = String(error && error.message ? error.message : error);
    Logger.log('enviarEmailRectificativa error: ' + mensaje);
    return { ok: false, mensaje: 'No se pudo enviar el correo: ' + mensaje };
  }
}

/**
 * Dialog entry point. datos = {
 *   numeroOriginal, tipo: 'completa'|'articulos', motivoAeat, motivoTexto,
 *   fechaEmision, horaEmision, fechaOperacion, observacionesCliente,
 *   lineas: [{ reference, quantity, unitPrice, ivaRate, descuentoRate, retencionRate, recargoRate }],
 *   idPayment, paymentNote, textObservations, cargosDescuentos: [{ tipo, tipoValor, valor }],
 *   enviarEmail, email, numeroEsperado }
 * numeroEsperado is the rectificativa number the dialog showed; if the series moved
 * in the meantime nothing is sent and { reconfirmar, nuevoNumero } come back.
 * Always returns a crearMensaje() object, extended with { numero, neto, emailEnviado,
 * emailMensaje } on success.
 */
function crearRectificativa(datos) {
  try {
    datos = datos || {};
    const numeroOriginal = String(datos.numeroOriginal || '').trim();
    const tipo = datos.tipo === 'articulos' ? 'articulos' : 'completa';
    if (!numeroOriginal) return crearMensaje('Ingresa el número de la factura a rectificar.', 'warning', 'Rectificativa');

    const info = cargarFacturaParaRectificar(numeroOriginal, { sinConsecutivo: true });
    if (!info.ok) return crearMensaje(info.error, 'error', 'Rectificativa');
    if (info.tieneCompleta) {
      return crearMensaje('La factura ' + numeroOriginal + ' ya tiene una rectificativa completa (' + info.rectificativas.join(', ') + '). FacturasApp no permite más rectificativas sobre ella.', 'warning', 'Rectificativa');
    }
    const error = validarDatosRectificativa_(datos, info);
    if (error) return crearMensaje(error, 'warning', 'Rectificativa');

    // Mode B: reserve the number now and refuse if it is not the one the user saw
    const numeracion = prepararNumeroRectificativa_();
    if (!numeracion.ok) {
      const bloqueo = crearMensaje(numeracion.mensaje, numeracion.motivo === 'error_consulta' ? 'error' : 'warning', 'Rectificativa');
      bloqueo.numeroOcupado = numeracion.motivo === 'existe';
      return bloqueo;
    }
    const numeroEsperado = String(datos.numeroEsperado || '').trim();
    if (numeroEsperado && numeroEsperado !== numeracion.numero) {
      const cambio = crearMensaje('El número de la rectificativa cambió de ' + numeroEsperado + ' a ' + numeracion.numero +
        ' porque se usó el anterior mientras tanto. No se creó nada. Revisa y confirma de nuevo.', 'warning', 'Número actualizado');
      cambio.nuevoNumero = numeracion.numero;
      cambio.reconfirmar = true;
      return cambio;
    }
    const numeroRect = numeracion.numero;

    const motivoAeat = String(datos.motivoAeat || '').trim().toUpperCase();
    const motivoTexto = String(datos.motivoTexto || '').trim();
    const original = jsonVigenteDeFactura_(numeroOriginal).json;
    let envio;
    let payloadListado;
    if (tipo === 'completa') {
      const payload = {
        originalInvoiceNumber: numeroOriginal,
        invoiceNumber: numeroRect,
        cancellationReason: motivoTexto,
        typeRectificative: RECTIFICATIVA_TIPO,
        invoiceRectificativeType: motivoAeat
      };
      // Not in the verified contract; the gateway ignores unknown fields, so the
      // dialog's "Fecha de operación" is forwarded in case the API honours it.
      if (esFechaIso_(datos.fechaOperacion)) {
        const tz = getSpreadsheetTimeZone_();
        payload.operationDate = Utilities.formatDate(Utilities.parseDate(datos.fechaOperacion, tz, 'yyyy-MM-dd'), tz, 'dd-MM-yyyy');
      }
      envio = enviarRectificativa_('/AddRectificativeComplete', payload);
      // Minimal row for ListadoEstado: the backend builds the negative duplicate
      // of the alta (never of a later sustitución)
      const alta = buscarJsonFacturaOriginal_(numeroOriginal);
      payloadListado = Object.assign({}, payload, {
        products: [],
        contacts: alta ? alta.contacts : [],
        sumTotalTotal: alta ? -round2Rect_(alta.sumTotalTotal) : 0,
        sumTotalNetPayable: alta ? -round2Rect_(alta.sumTotalNetPayable) : 0
      });
    } else {
      const opcionesPayload = opcionesPayloadDesdeDatos_(datos, numeroOriginal);
      opcionesPayload.invoiceNumber = numeroRect;
      const payload = construirPayloadRectificativaArticulos_(original, datos.lineas || [], opcionesPayload);
      ajustarRecargoEnContactos_(payload);
      Logger.log('AddRectificativeArticles payload: ' + JSON.stringify(payload));
      envio = enviarRectificativa_('/AddRectificativeArticles', payload);
      payloadListado = payload;
    }

    if (!envio.ok) {
      const fallo = crearMensaje(envio.mensaje, envio.reconfirmar ? 'warning' : 'error', 'Error al crear la rectificativa');
      if (envio.nuevoNumero) fallo.nuevoNumero = envio.nuevoNumero;
      if (envio.reconfirmar) fallo.reconfirmar = true;
      return fallo;
    }

    registrarRectificativaEnHistorial_(envio.numero, info, tipo);
    payloadListado.invoiceNumber = envio.numero;
    registrarRectificativaEnListadoEstado_(envio.numero, info, payloadListado);

    let emailEnviado = false;
    let emailMensaje = '';
    if (datos.enviarEmail && envio.numero) {
      const resultadoEmail = enviarEmailRectificativa(envio.numero, datos.email, numeroOriginal);
      emailEnviado = resultadoEmail.ok;
      emailMensaje = resultadoEmail.mensaje;
    }

    const neto = tipo === 'completa'
      ? (info.totalesAnulacion ? -info.totalesAnulacion.neto : null)
      : round2Rect_(payloadListado.sumTotalNetPayable);
    const mensaje = crearMensaje(
      'Rectificativa ' + (envio.numero || '') + ' creada sobre la factura ' + numeroOriginal + '.' +
      (envio.mensaje ? ' (' + envio.mensaje + ')' : '') +
      (emailMensaje ? '\n' + emailMensaje : '') +
      '\nPuedes descargar su PDF desde el listado de facturas.',
      'success', 'Rectificativa creada');
    mensaje.numero = envio.numero || '';
    mensaje.neto = neto;
    mensaje.emailEnviado = emailEnviado;
    mensaje.emailMensaje = emailMensaje;
    return mensaje;
  } catch (error) {
    const mensaje = String(error && error.message ? error.message : error);
    Logger.log('crearRectificativa error: ' + mensaje);
    return crearMensaje(mensaje, 'error', 'Rectificativa');
  }
}

/**
 * Sidebar entry point: loads the invoice and opens the modal dialog in the
 * requested mode. Returns { ok: true } or a crearMensaje() object for the sidebar.
 */
function abrirDialogoRectificativa(numeroFactura, tipo) {
  try {
    tipo = tipo === 'articulos' ? 'articulos' : 'completa';
    const info = cargarFacturaParaRectificar(numeroFactura);
    if (!info.ok) return crearMensaje(info.error, 'error', 'Rectificativa');
    if (info.tieneCompleta) {
      return crearMensaje('La factura ' + info.numero + ' ya tiene una rectificativa completa (' + info.rectificativas.join(', ') + '). FacturasApp no permite más rectificativas sobre ella.', 'warning', 'Rectificativa');
    }
    if (tipo === 'articulos' && (!info.tieneJson || info.lineas.length === 0)) {
      return crearMensaje('No se encontraron las líneas de la factura ' + info.numero + ' en este libro. Usa la rectificativa de anulación.', 'error', 'Rectificativa');
    }
    if (info.numeracion && !info.numeracion.ok) {
      return crearMensaje(info.numeracion.mensaje, info.numeracion.motivo === 'error_consulta' ? 'error' : 'warning', 'Rectificativa');
    }
    info.tipo = tipo;
    const template = HtmlService.createTemplateFromFile('rectificativaDialog');
    // Force-printed inside a <script>: escape "<" so descriptions can't close the tag
    template.datosJson = JSON.stringify(info).replace(/</g, '\\u003c');
    const html = template.evaluate()
      .setWidth(RECTIFICATIVA_DIALOG_WIDTH)
      .setHeight(RECTIFICATIVA_DIALOG_HEIGHT);
    SpreadsheetApp.getUi().showModalDialog(html, tipo === 'completa'
      ? 'Factura rectificativa de anulación'
      : 'Factura rectificativa de sustitución');
    return { ok: true };
  } catch (error) {
    const mensaje = String(error && error.message ? error.message : error);
    Logger.log('abrirDialogoRectificativa error: ' + mensaje);
    return crearMensaje(mensaje, 'error', 'Rectificativa');
  }
}
