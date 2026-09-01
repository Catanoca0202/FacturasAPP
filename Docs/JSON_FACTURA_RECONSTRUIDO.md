# JSON de factura — reconstruido por ingeniería inversa

**No** está basado en `factura.json` del repo (ese sí está desactualizado). Está reconstruido
ejecutando la lógica real de `guardarYGenerarInvoice()` (Factura.js:1962) sobre una hoja
simulada, y contrastado contra el OpenAPI `Api Applications 1.json` y contra el contrato
actual de `invoice_create` del MCP.

Archivos que acompañan a este documento:

| Archivo | Qué es |
|---|---|
| `simular_factura.js` | Reimplementación fiel del cálculo. `node simular_factura.js` regenera los JSON. |
| `factura_ejemplo_A.json` | Payload real — empresa, IRPF 15 %, descuento de línea, cargo y descuento de factura, línea exenta |
| `factura_ejemplo_B.json` | Payload real — autónomo, recargo de equivalencia, retención por línea |
| `factura_nuevo_contrato_PLANTILLA.json` | El escenario A traducido al contrato **nuevo** de `invoice_create` (con IDs de catálogo pendientes) |

---

## 1. Hay TRES contratos en juego, no uno

| # | Contrato | Dónde vive | Forma |
|---|---|---|---|
| 1 | `factura.json` del repo | carpeta del proyecto | **Obsoleto.** Le faltan `regime`, `taxCode`, `qualificationOperation`, `isExemptOperation`, `idExento`, `productType`, `operationDate`, `aditionalData`… |
| 2 | OpenAPI `Api Applications 1.json` → `FieldInvoice` | carpeta del proyecto | Describe `POST .../InvoiceServices/AddInvoice`, que es lo que el add-in llama hoy. **También va por detrás del código.** |
| 3 | `invoice_create` del MCP FacturasApp | API actual | Contrato **nuevo**: IDs enteros de catálogo en vez de códigos string. |

Lo que el add-in emite hoy = contrato 2 **+ extensiones Veri\*factu** que el OpenAPI no declara.

### Campos que el código envía y el OpenAPI del repo NO declara

| Objeto | Campos extra |
|---|---|
| `FieldInvoice` | `operationDate`, `sumTotalRetentionIRPF` |
| `Products` | `regime`, `productType` |
| `Taxes` | `taxCode`, `isExemptOperation`, `qualificationOperation`, `regime`, `idExento` |
| `FieldTaxations` | `taxCode`, `qualificationOperation`, `regime`, `idExento` |
| `WithHoldingsSurChargesDto` | `isWithHolding`, `rateValueWithHoldings` |
| `Contacts` | `identificationType`, `applySurchargeEquivalence` |

Restricciones del OpenAPI que el código sí respeta: `contactName` ≤ 30, `paymentNote` ≤ 300,
`textCustomerObservations` ≤ 350, `description` ≤ 100, `reference` ≤ 50, `identification` ≤ 20,
`quantity` **int32** (por eso el `Math.max(1, Math.trunc(cantidad))`).

---

## 2. Escenario A — empresa, IRPF 15 %

### Entrada (hoja `Factura`)

| Fila | A ref | B producto | C cant | D precio | G %IVA | H %dto | I recargo | J retención |
|---|---|---|---|---|---|---|---|---|
| 15 | SRV-CONS | Consultoría técnica | 10 | 60,00 | 21 % | 10 % | — | — |
| 16 | LIC-SW01 | Licencia software anual | 2 | 120,00 | 21 % | — | — | — |
| 17 | FORM-BON | Formación bonificada (exenta E1) | 1 | 300,00 | 0 % | — | — | — |

`B17` cargo = 25,00 · `D17` descuento factura = 50,00 · `F17` IRPF = 15 %

### Aritmética, línea a línea

| Línea | baseBruta | descuento | baseNeta | IVA | recargo | retención |
|---|---|---|---|---|---|---|
| 1 | 10 × 60 = **600,00** | 600 × 0,10 = **60,00** | **540,00** | 540 × 0,21 = **113,40** | 0 | 0 |
| 2 | 2 × 120 = **240,00** | 0 | **240,00** | 240 × 0,21 = **50,40** | 0 | 0 |
| 3 | 1 × 300 = **300,00** | 0 | **300,00** (exenta) | 0 | 0 | 0 |

### Totales

```
sumTotalSubTotal       = 600 + 240 + 300              = 1.140,00   (bruto)
sumTotalTaxBase        = 540 + 240 + 300              = 1.080,00   (base imponible)
sumTotalTax            = 113,40 + 50,40               =   163,80   (solo IVA)
sumTotalSubTotalAndTax = 1.080,00 + 163,80            = 1.243,80
sumTotalTotal          = 1.243,80 + 0 (recargo)       = 1.243,80
IRPF                   = 1.080,00 × 0,15              =   162,00
sumTotalNetPayable     = 1.243,80 − 0 − 50 + 25 − 162 = 1.056,80
```

`fieldTaxations` queda con **dos** entradas (21 % y 0 % exenta), y `idOperationsExenta` = `"E1"`
porque hay base exenta.

→ payload completo en **`factura_ejemplo_A.json`**

---

## 3. Escenario B — autónomo con recargo de equivalencia

### Entrada

| Fila | A ref | B producto | C cant | D precio | G %IVA | H %dto | I recargo | J retención |
|---|---|---|---|---|---|---|---|---|
| 15 | MON-27 | Monitor 27 pulgadas | 3 | 210,00 | 21 % | — | 5,2 % | — |
| 16 | CBL-HDMI | Cable HDMI 2 m | 10 | 8,50 | 10 % | — | 1,4 % | — |
| 17 | SRV-INST | Servicio de instalación | 1 | 150,00 | 21 % | — | — | 15 % |

Sin cargo, sin descuento de factura, sin IRPF global.

### Aritmética

| Línea | baseBruta = baseNeta | IVA | recargo | retención |
|---|---|---|---|---|
| 1 | **630,00** | **132,30** | 630 × 0,052 = **32,76** | 0 |
| 2 | **85,00** | **8,50** | 85 × 0,014 = **1,19** | 0 |
| 3 | **150,00** | **31,50** | 0 | 150 × 0,15 = **22,50** |

### Totales

```
sumTotalSubTotal       = 865,00
sumTotalTaxBase        = 865,00
sumTotalTax            = 172,30
sumTotalSubTotalAndTax = 1.037,30
Σ recargo              =    33,95
sumTotalTotal          = 1.037,30 + 33,95 = 1.071,25
Σ retenciones          =    22,50
sumTotalNetPayable     = 1.071,25 − 22,50 = 1.048,75
```

Efectos del recargo, todos automáticos:

- `regime` de esas líneas se fuerza a `"18"` (Recargo de equivalencia).
- Se añade una entrada extra en `fieldTaxations` con `taxName: "RecargoEquivalencia"`, una por cada % distinto.
- `contacts[0].applySurchargeEquivalence = true`.
- Si el cliente no es **Autónomo**, la generación aborta con error (`clienteEsSoloAutonomo_`).

→ payload completo en **`factura_ejemplo_B.json`**

---

## 4. Traducción al contrato NUEVO (`invoice_create`)

Este es el cambio de fondo: el contrato nuevo sustituye **todos** los códigos string por
**IDs enteros de catálogo**.

### Nivel factura

| Legacy (`FieldInvoice`) | Nuevo (`invoice`) | Nota |
|---|---|---|
| `invoiceNumber` | `invoiceNumber` | ahora sale de `invoice_list_consecutives` → `nextInvoiceNumber` |
| `invoiceIdTypeRegAEAT: "AI"` | `aeatTypeId` (int) | `catalog_list_aeat_types` |
| `idPayment: "TF"` | `paymentTypeId` (int) | `catalog_list_payment_types` |
| `invoiceIdTypeRegSIF` | `sifTypeId` (int/null) | `catalog_list_sif` |
| `operationDate: "20-08-2026"` | `operationDate: "2026-08-20"` | ⚠️ **cambia el formato** a `yyyy-MM-dd` |
| `contacts[]` | `customers[]` | |
| `currentNumber`, `invoiceDate`, `invoiceTime`, `idOperations`, `idOperationsExenta`, `textCustomerObservations`, `invoiceTypeId`, `invoiceRectificativeTypeId`, `typeRectificativeId`, `aditionalData`, `sumTotalRetentionIRPF` | — | ya no existen; fecha/hora las pone el servidor y el IRPF va solo en `chargeAndDiscount` |
| `sumTotal*` y `valueExemptBase` | idénticos | mismos nombres y misma aritmética |

### Cliente

| Legacy | Nuevo |
|---|---|
| `contactType` `"01"` | `typeContactId` |
| `personType` `"01"/"02"` | `typePersonId` — `2`/`4` = autónomo / persona física → usar `personName`; `3` = empresa → usar `tradeName` |
| `identificationType` `"02"` | `identificationTypeId` |
| `identification` | `document` |
| `emailCustomer` | `email` (obligatorio) |
| `customerCode` | `codeClient` |
| `country` / `province` / `population` | `countryId` / `stateId` / `cityId` |
| `addressCustomer` / `postalCodeCustomer` | `address` / `postalCode` |
| `phoneCustomer` / `webSite` | `phone` / `siteWeb` |
| `companyName` + `regime` | — |

### Línea de producto

| Legacy | Nuevo |
|---|---|
| `description` | `productName` (≤ 150) |
| `unitPrice` | `value` (+ `ivaIncluded`, `valueWithIva`) |
| `typeUse: "VEN"` | `useTypeId: 2` (Venta) |
| `productType: 1/2` | `productTypeId: 1/2` |
| `quantity` int32 | `quantity` **number ≥ 0,0001** — ya admite decimales |
| `taxes[].taxCode "01"` | `taxes[].taxId` |
| `taxes[].rate 21` | `taxes[].tariffId` si es IVA · `taxes[].manualRate` si no es IVA |
| `taxes[].regime "01"` | `taxes[].regimeId` |
| `taxes[].qualificationOperation "S1"` | `taxes[].qualificationOperationId` |
| `taxes[].idExento "E1"` | `taxes[].operationExentaId` |
| recargo en `withHoldingsSurChargesDto` | `taxes[].equivalenceSurchargeRate` + `applySurchargeEquivalence` (y elegir un `tariffId` con `surchargeEquivalence > 0`) |
| retención en `withHoldingsSurChargesDto` | `retentions[]` → `{ retentionSurchargeId, idTariffRetention, manualRate }` |
| `discountDtoModules[]` | `discounts[]` (mismos 4 campos) |
| `taxes[]` con varias entradas | **máximo 1 impuesto por línea** |

### Cargos y descuentos

| Legacy | Nuevo |
|---|---|
| `idtypeFeeDiscount` `"CG"/"DT"/"RT"` | `feeDiscountTypeId` (int) |
| `idTypeValueFeeDiscount` `"VR"/"PJ"` | `valueFeeDiscountTypeId` (int) |
| `baseFeeDiscount` / `valueFeeDiscount` / `totalFeeDiscount` | idénticos |

→ esqueleto en **`factura_nuevo_contrato_PLANTILLA.json`**

---

## 5. Trampas al portar (lo que rompería en el contrato nuevo)

1. **`quantity` truncado a entero.** `Math.max(1, Math.trunc(cantidad))` era una imposición del `int32` viejo. En el contrato nuevo destruye cualquier factura con horas u unidades fraccionadas (2,5 h → 2 h). Hay que quitarlo.
2. **Formato de `operationDate`.** `dd-MM-yyyy` → `yyyy-MM-dd`.
3. **Todos los mapeos de string a código están hardcodeados** (`getRegimenCode`, `mapIdPaymentCode`, `getTaxCodeFromName_`, `getIdTypeWithHoldings_`, `extraerCodigoExento`). En el contrato nuevo hay que resolver contra los catálogos vivos, no contra tablas fijas.
4. **`getIdentificationCodeDocument` no contempla "NIF"** — solo NIF-IVA, Pasaporte, etc. Un cliente con "NIF" a secas lanza excepción.
5. **`taxes[]` máximo 1 por línea** en el contrato nuevo. Hoy el código siempre mete exactamente 1, así que no hay problema, pero el recargo ya no puede ir como segunda entrada.
6. **El IRPF sale de una celda, no del código.** `E{T+10}` la calcula la plantilla. En el simulador asumí `base imponible × tarifa`; **conviene verificarlo abriendo la hoja**, porque de ahí sale `sumTotalRetentionIRPF` y el `totalFeeDiscount` del `RT`. Si la plantilla excluye la base exenta, el escenario A daría 117,00 € en vez de 162,00 €.
7. **Redondeo.** La hoja calcula con precisión completa; el JSON redondea a 2 decimales por línea. Con muchas líneas, el total de la hoja y el del payload pueden diferir en céntimos.

---

## 6. Lo que falta para tener un JSON 100 % listo para enviar

Los `<<...>>` de la plantilla son IDs de catálogo. Se resuelven con una `apiKey` de compañía:

```
catalog_list_taxes(apiKey)                              → taxId (IVA)
catalog_list_tax_tariffs(apiKey, date)                  → tariffId (21 / 10 / 0), surchargeEquivalence
catalog_list_regimen(apiKey, taxId)                     → regimeId (01 general, 18 recargo)
catalog_list_operations(apiKey, taxId)                  → qualificationOperationId (S1)
catalog_list_operation_exenta(apiKey, taxId)            → operationExentaId (E1…E8)
catalog_list_retention_surcharges(apiKey)               → retentionSurchargeId
catalog_list_retention_surcharge_tariffs(apiKey, id)    → idTariffRetention
catalog_list_fee_discount_types(apiKey)                 → feeDiscountTypeId (CG / DT / RT)
catalog_list_value_fee_discount_types(apiKey)           → valueFeeDiscountTypeId (VR / PJ)
catalog_list_aeat_types / payment_types / person_types / identification_types / contact_types
catalog_list_countries / provinces / cities
```
