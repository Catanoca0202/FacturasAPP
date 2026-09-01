/**
 * Simulador fiel de guardarYGenerarInvoice() (Factura.js:1962)
 * Reproduce paso a paso la misma aritmética del add-in a partir de una
 * "hoja Factura" simulada, y emite el JSON exacto que se enviaría a
 * POST /ApiGateway/ApiExternal/Invoice/api/InvoiceServices/AddInvoice
 */

const round2 = (n) => Math.round((Number(n) || 0) * 100) / 100;

// --- Mapeos copiados del código real -----------------------------------
const getTaxCodeFromName_ = (name) => ({ IVA: '01', IPSI: '02', IGIC: '03', OTROS: '05' }[String(name || '').trim().toUpperCase()] || '01');
const getTaxNameForCode_ = (code) => ({ '01': 'IVA', '02': 'IPSI', '03': 'IGIC', '05': 'Otros' }[code] || 'IVA');
const getIdTypeWithHoldings_ = (t) => ({ retencion: '10', recargo: '11', irpf: '99' }[t] || null);

function generarFactura(hoja) {
  const productos = hoja.lineas;

  let products = [];
  let taxGroups = {}, recargoTaxGroups = {};
  let totalSubTotal = 0, totalTaxBase = 0;
  let sumIvaAmount = 0, totalWithHoldings = 0, totalSurCharges = 0, totalDiscounts = 0;
  let totalExemptBase = 0, firstIdExento = null;
  let hasRecargoEquivalencia = false;

  // ---------- BUCLE POR LÍNEA (Factura.js:2074-2312) ----------
  for (const L of productos) {
    const cantidad       = Number(L.C) || 1;   // col C
    const precioUnitario = Number(L.D) || 0;   // col D
    const ivaRate        = Number(L.G) || 0;   // col G
    const descuentoRate  = Number(L.H) || 0;   // col H
    const recargoRate    = Number(L.I) || 0;   // col I
    const retencionRate  = Number(L.J) || 0;   // col J

    const baseBruta      = round2(precioUnitario * cantidad);
    const discountAmount = round2(baseBruta * descuentoRate);
    const baseNeta       = round2(baseBruta - discountAmount);
    const taxAmount      = round2(baseNeta * ivaRate);
    const withHoldings   = round2(baseNeta * retencionRate);
    const surCharges     = round2(baseNeta * recargoRate);

    if (recargoRate > 0) hasRecargoEquivalencia = true;

    const taxCode = getTaxCodeFromName_(L.tipoImpuesto);
    const taxName = getTaxNameForCode_(taxCode);
    let regime = L.regimen;
    if (recargoRate > 0) regime = '18';           // override obligatorio

    // ----- taxes[] -----
    const taxes = [];
    if (ivaRate > 0) {
      taxes.push({
        taxName, rate: ivaRate * 100, taxBase: baseNeta, valueTax: taxAmount,
        taxCode, isExemptOperation: false,
        qualificationOperation: L.exento ? null : L.calificacion,
        regime
      });
      const groupKey = taxCode + '_' + (ivaRate * 100);
      if (!taxGroups[groupKey]) taxGroups[groupKey] = {
        taxName, rate: ivaRate * 100, taxBase: 0, valueTax: 0, taxCode,
        qualificationOperation: L.exento ? null : L.calificacion, regime
      };
      taxGroups[groupKey].taxBase  = round2(taxGroups[groupKey].taxBase + baseNeta);
      taxGroups[groupKey].valueTax = round2(taxGroups[groupKey].valueTax + taxAmount);
    } else {
      const idExento = L.exento ? (L.operacionExenta || 'E1') : null;
      const t = {
        taxName, rate: 0, taxBase: baseNeta, valueTax: 0, taxCode,
        isExemptOperation: !!L.exento,
        qualificationOperation: L.exento ? null : L.calificacion,
        regime
      };
      if (L.exento && idExento) t.idExento = idExento;
      taxes.push(t);
      if (!firstIdExento && idExento) firstIdExento = idExento;

      const groupKey = taxCode + '_0' + (L.exento ? '_ex' : '');
      if (!taxGroups[groupKey]) {
        taxGroups[groupKey] = {
          taxName, rate: 0, taxBase: 0, valueTax: 0, taxCode,
          qualificationOperation: L.exento ? null : L.calificacion, regime
        };
        if (L.exento && idExento) taxGroups[groupKey].idExento = idExento;
      }
      taxGroups[groupKey].taxBase = round2(taxGroups[groupKey].taxBase + baseNeta);
      totalExemptBase = round2(totalExemptBase + baseNeta);
    }

    sumIvaAmount = round2(sumIvaAmount + taxAmount);

    // ----- withHoldingsSurChargesDto[] -----
    const whs = [];
    if (retencionRate > 0) whs.push({
      isWithHolding: true, idRateWithHoldings: getIdTypeWithHoldings_('retencion'),
      rateValueWithHoldings: retencionRate, subTotalWithHoldings: baseNeta,
      cuotaWithHoldings: withHoldings
    });
    if (recargoRate > 0) {
      whs.push({
        isWithHolding: false, idRateWithHoldings: getIdTypeWithHoldings_('recargo'),
        rateValueWithHoldings: recargoRate, subTotalWithHoldings: baseNeta,
        cuotaWithHoldings: surCharges
      });
      const k = taxCode + '_' + (recargoRate * 100);
      if (!recargoTaxGroups[k]) recargoTaxGroups[k] = {
        taxName: 'RecargoEquivalencia', rate: recargoRate * 100,
        taxBase: 0, valueTax: 0, taxCode, regime: '18'
      };
      recargoTaxGroups[k].taxBase  = round2(recargoTaxGroups[k].taxBase + baseNeta);
      recargoTaxGroups[k].valueTax = round2(recargoTaxGroups[k].valueTax + surCharges);
    }

    // ----- discountDtoModules[] -----
    const discounts = [];
    if (descuentoRate > 0) discounts.push({
      discountName: 'Descuento aplicado',
      discountRate: descuentoRate * 100,
      discountBase: baseBruta,
      valueDiscount: discountAmount
    });

    products.push({
      typeUse: 'VEN',
      reference: String(L.A).substring(0, 50),
      description: String(L.descripcion).substring(0, 100),
      regime,
      productType: L.tipoProducto === 'Producto' ? 1 : L.tipoProducto === 'Servicio' ? 2 : 0,
      unitPrice: precioUnitario,
      quantity: Math.max(1, Math.trunc(cantidad)),
      subTotal: baseBruta,              // BRUTO, antes de descuento
      totalTax: round2(taxAmount),      // solo IVA
      totalwithHoldings: withHoldings,
      totalSurCharges: surCharges,
      totaldiscount: discountAmount,
      taxes,
      withHoldingsSurChargesDto: whs,
      discountDtoModules: discounts
    });

    totalSubTotal     = round2(totalSubTotal + baseBruta);
    totalTaxBase      = round2(totalTaxBase + baseNeta);
    totalWithHoldings = round2(totalWithHoldings + withHoldings);
    totalSurCharges   = round2(totalSurCharges + surCharges);
    totalDiscounts    = round2(totalDiscounts + discountAmount);
  }

  if (hasRecargoEquivalencia && hoja.tipoPersonaNombre.toLowerCase() !== 'autónomo') {
    throw new Error('Regla de validación: recargo de equivalencia con cliente no Autónomo.');
  }

  const totalTax = round2(sumIvaAmount);

  // ---------- fieldTaxations ----------
  const fieldTaxations = [...Object.values(taxGroups), ...Object.values(recargoTaxGroups)];

  // ---------- chargeAndDiscount ----------
  const cargoTotal        = round2(hoja.cargoFactura);        // B{T-2}
  const descuentoFactura  = round2(hoja.descuentoFactura);    // D{T-2}
  // E{T+10}: la plantilla calcula base imponible × tarifa cuando el selector es 7/15/19 %
  const irpfGlobalAmount  = hoja.irpfModo === 'percent'
    ? round2(totalTaxBase * hoja.irpfTarifa)
    : round2(hoja.irpfImporteFijo || 0);

  const chargeAndDiscount = [];
  if (cargoTotal > 0) chargeAndDiscount.push({
    idtypeFeeDiscount: 'CG', idTypeValueFeeDiscount: 'VR',
    baseFeeDiscount: 0, valueFeeDiscount: cargoTotal, totalFeeDiscount: cargoTotal
  });
  if (descuentoFactura > 0) chargeAndDiscount.push({
    idtypeFeeDiscount: 'DT', idTypeValueFeeDiscount: 'VR',
    baseFeeDiscount: 0, valueFeeDiscount: descuentoFactura, totalFeeDiscount: descuentoFactura
  });
  if (irpfGlobalAmount > 0) {
    if (hoja.irpfModo === 'percent') chargeAndDiscount.push({
      idtypeFeeDiscount: 'RT', idTypeValueFeeDiscount: 'PJ',
      baseFeeDiscount: totalTaxBase,
      valueFeeDiscount: round2(hoja.irpfTarifa * 100),
      totalFeeDiscount: irpfGlobalAmount
    });
    else chargeAndDiscount.push({
      idtypeFeeDiscount: 'RT', idTypeValueFeeDiscount: 'VR',
      baseFeeDiscount: 0, valueFeeDiscount: irpfGlobalAmount, totalFeeDiscount: irpfGlobalAmount
    });
  }

  // ---------- Totales ----------
  const sumTotalSubTotalAndTax = round2(totalTaxBase + totalTax);
  const sumTotalTotal          = round2(sumTotalSubTotalAndTax + totalSurCharges);
  const sumTotalNetPayable     = round2(sumTotalTotal - totalWithHoldings - descuentoFactura + cargoTotal - irpfGlobalAmount);

  return {
    textCustomerObservations: hoja.observacionesCliente,
    invoiceNumber: hoja.invoiceNumber,
    currentNumber: Number(String(hoja.invoiceNumber).replace(/[^0-9]/g, '')),
    invoiceDate: hoja.fechaFactura + 'T12:00:00.000Z',
    invoiceTime: hoja.horaFactura + '.0000000',
    invoiceExpiration: String(hoja.diasVencimiento),
    invoiceIdTypeRegAEAT: 'AI',
    invoiceIdTypeRegSIF: null,
    contactName: String(hoja.contactName).substring(0, 30),
    operationDate: hoja.fechaOperacion,
    contacts: [hoja.contacto(hasRecargoEquivalencia)],
    products,
    idPayment: hoja.idPayment,
    paymentNote: hoja.notaPago,
    textObservations: hoja.observaciones,
    idOperations: 'S1',
    idOperationsExenta: totalExemptBase > 0 ? (firstIdExento || 'E1') : 'E0',
    valueExemptBase: totalExemptBase,
    chargeAndDiscount,
    fieldTaxations,
    sumTotalSubTotal: totalSubTotal,
    sumTotalTaxBase: totalTaxBase,
    sumTotalTax: totalTax,
    sumTotalSubTotalAndTax,
    sumTotalExemptBase: totalExemptBase,
    sumTotalDiscount: descuentoFactura,
    sumTotalCharge: cargoTotal,
    sumTotalRetentionIRPF: irpfGlobalAmount,
    sumTotalTotal,
    sumTotalNetPayable,
    invoiceTypeId: 0,
    invoiceRectificativeTypeId: 0,
    typeRectificativeId: 0,
    aditionalData: { invoiceId: 0, startInvoiceId: 0 }
  };
}

// =======================================================================
// ESCENARIO A — Empresa · IRPF global 15 % · descuento de línea ·
//               cargo y descuento de factura · una línea exenta
// =======================================================================
const escenarioA = {
  invoiceNumber: 'FAC-2026-0042',
  fechaFactura: '2026-08-20',
  horaFactura: '10:24:07',
  fechaOperacion: '20-08-2026',
  diasVencimiento: 30,
  contactName: 'Alejandro Cataño',
  idPayment: 'TF',                       // "Transferencia bancaria"
  notaPago: 'Pago a 30 días a la cuenta ES12 3456 7890 1234 5678 9012',
  observaciones: 'Servicios prestados durante agosto de 2026.',
  observacionesCliente: 'Referencia de pedido PO-8842',
  tipoPersonaNombre: 'Empresa',
  cargoFactura: 25,
  descuentoFactura: 50,
  irpfModo: 'percent',
  irpfTarifa: 0.15,
  contacto: () => ({
    contactType: '01',                   // Cliente
    personType: '02',                    // Empresa
    companyName: 'Construcciones Delta S.L.',
    customerCode: 'CLI-000148',
    identificationType: '02',            // NIF-IVA
    identification: 'B87654321',
    tradeName: 'Construcciones Delta S.L. - CLI-000148',
    regime: '01',
    country: '724',
    province: '28',
    population: '28079',
    addressCustomer: 'Calle Mayor 14, 2º B',
    postalCodeCustomer: '28013',
    phoneCustomer: '+34 915 555 123',
    webSite: null,
    emailCustomer: 'administracion@delta.example',
    applySurchargeEquivalence: false
  }),
  lineas: [
    { A: 'SRV-CONS', descripcion: 'Consultoría técnica', tipoProducto: 'Servicio',
      C: 10, D: 60.00, G: 0.21, H: 0.10, I: 0, J: 0,
      tipoImpuesto: 'IVA', regimen: '01', calificacion: 'S1', exento: false },
    { A: 'LIC-SW01', descripcion: 'Licencia software anual', tipoProducto: 'Servicio',
      C: 2, D: 120.00, G: 0.21, H: 0, I: 0, J: 0,
      tipoImpuesto: 'IVA', regimen: '01', calificacion: 'S1', exento: false },
    { A: 'FORM-BON', descripcion: 'Formación bonificada', tipoProducto: 'Servicio',
      C: 1, D: 300.00, G: 0, H: 0, I: 0, J: 0,
      tipoImpuesto: 'IVA', regimen: '01', calificacion: 'S1', exento: true, operacionExenta: 'E1' }
  ]
};

// =======================================================================
// ESCENARIO B — Autónomo · recargo de equivalencia · retención por línea
// =======================================================================
const escenarioB = {
  invoiceNumber: 'FAC-2026-0043',
  fechaFactura: '2026-08-20',
  horaFactura: '11:05:33',
  fechaOperacion: '20-08-2026',
  diasVencimiento: 15,
  contactName: 'Alejandro Cataño',
  idPayment: 'EF',                       // "Efectivo"
  notaPago: 'Pagado en efectivo en el momento de la entrega',
  observaciones: null,
  observacionesCliente: null,
  tipoPersonaNombre: 'Autónomo',
  cargoFactura: 0,
  descuentoFactura: 0,
  irpfModo: 'none',
  irpfTarifa: 0,
  irpfImporteFijo: 0,
  contacto: (recargo) => ({
    contactType: '01',
    personType: '01',                    // Autónomo / Persona física
    companyName: 'Marta Ruiz Peña',
    customerCode: 'CLI-000212',
    identificationType: '02',
    identification: '12345678Z',
    tradeName: 'Marta Ruiz Peña - CLI-000212',
    regime: '01',
    country: '724',
    province: '46',
    population: '46250',
    addressCustomer: 'Avenida del Puerto 88',
    postalCodeCustomer: '46023',
    phoneCustomer: '+34 963 111 222',
    webSite: null,
    emailCustomer: 'marta.ruiz@example.com',
    applySurchargeEquivalence: recargo
  }),
  lineas: [
    { A: 'MON-27', descripcion: 'Monitor 27 pulgadas', tipoProducto: 'Producto',
      C: 3, D: 210.00, G: 0.21, H: 0, I: 0.052, J: 0,
      tipoImpuesto: 'IVA', regimen: '01', calificacion: 'S1', exento: false },
    { A: 'CBL-HDMI', descripcion: 'Cable HDMI 2 m', tipoProducto: 'Producto',
      C: 10, D: 8.50, G: 0.10, H: 0, I: 0.014, J: 0,
      tipoImpuesto: 'IVA', regimen: '01', calificacion: 'S1', exento: false },
    { A: 'SRV-INST', descripcion: 'Servicio de instalación', tipoProducto: 'Servicio',
      C: 1, D: 150.00, G: 0.21, H: 0, I: 0, J: 0.15,
      tipoImpuesto: 'IVA', regimen: '01', calificacion: 'S1', exento: false }
  ]
};

const fs = require('fs');
const a = generarFactura(escenarioA);
const b = generarFactura(escenarioB);
fs.writeFileSync('/home/claude/factura_ejemplo_A.json', JSON.stringify(a, null, 2));
fs.writeFileSync('/home/claude/factura_ejemplo_B.json', JSON.stringify(b, null, 2));

const chk = (n, o) => {
  console.log('\n===== ' + n + ' =====');
  console.log('sumTotalSubTotal        ', o.sumTotalSubTotal);
  console.log('sumTotalTaxBase         ', o.sumTotalTaxBase);
  console.log('sumTotalTax             ', o.sumTotalTax);
  console.log('sumTotalSubTotalAndTax  ', o.sumTotalSubTotalAndTax);
  console.log('recargo (Σ)             ', round2(o.products.reduce((s, p) => s + p.totalSurCharges, 0)));
  console.log('retenciones (Σ)         ', round2(o.products.reduce((s, p) => s + p.totalwithHoldings, 0)));
  console.log('sumTotalExemptBase      ', o.sumTotalExemptBase);
  console.log('sumTotalRetentionIRPF   ', o.sumTotalRetentionIRPF);
  console.log('sumTotalTotal           ', o.sumTotalTotal);
  console.log('sumTotalNetPayable      ', o.sumTotalNetPayable);
};
chk('ESCENARIO A', a);
chk('ESCENARIO B', b);
