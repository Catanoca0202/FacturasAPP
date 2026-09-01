# FacturasApp — Modelo de cálculo completo (hoja Factura → JSON API)

Referencia extraída de `mainScript.js`, `Factura.js` y `productos.js`.
Todo está en **orden de ejecución**: del dato origen hasta el payload que se envía a la API.

---

## 0. Convenciones

| Símbolo | Significado |
|---|---|
| `T` | Fila donde está la etiqueta **"Base imponible"** (`getTaxSectionStartRow()`). En la plantilla por defecto `T = 19`. |
| `n` | Última fila de producto (`getLastProductRow()`). |
| `15` | Primera fila de producto (constante `productStartRow`). |

Toda la sección de totales es **relativa a `T`** porque se desplaza cuando se insertan líneas de producto.
Los porcentajes (`G`, `H`, `I`, `J`) se guardan como **fracción** (0,21 = 21%).

---

## 1. Origen de los datos — hoja `Productos`

`obtenerInformacionProducto()` (productos.js:4) busca la fila por identificador único (col. R) y devuelve:

| Campo devuelto | Columna en `Productos` |
|---|---|
| `codigo Producto` | B — Código de referencia |
| `regimen` | D |
| `tipoProducto` | E (Producto / Servicio) |
| `valor Unitario` | G |
| `impuestos` | H (IVA / IGIC / IPSI / Otros) |
| `IVA` | I — Tarifa impuesto |
| `precio Con Iva` | J |
| `calificacionOperacion` | K |
| `exento` | L (checkbox) |
| `operacionExenta` | M |
| `Recargo de equivalencia` | O, **solo si** N (check recargo) = TRUE, si no `""` |
| `retencion` | Q, **solo si** P (check retención) = TRUE, si no `""` |
| `Estado` | A |

---

## 2. Layout de la hoja `Factura`

### 2.1 Línea de producto (filas `15` … `T-4`)

| Col | Contenido | Origen |
|---|---|---|
| A | Código de referencia | valor, desde `Productos` |
| B | Producto (identificador único) | lo elige el usuario (validación de datos) |
| C | Cantidad | lo escribe el usuario |
| D | Precio unitario | valor, desde `Productos` |
| E | Precio unitario con IVA | **fórmula** |
| F | Base gravable (ya con descuento) | **fórmula** |
| G | % IVA | valor, desde `Productos` |
| H | % descuento de línea | lo escribe el usuario (0 – 1) |
| I | Tarifa recargo de equivalencia | valor, desde `Productos` |
| J | Tarifa retención | valor, desde `Productos` |
| K | Total de línea | **fórmula** |

### 2.2 Bloque de totales

| Fila | Contenido |
|---|---|
| `T-3` | A = "Total filas", B = nº de líneas con producto |
| `T-2` | B = Cargos · D = Descuento de factura · F = selector IRPF (7%/15%/19%/"Valor libre") · H = valor IRPF cuando es "Valor libre" |
| `T` | Cabecera "Base imponible" |
| `T+1 … T+5` | Agrupación de impuestos (5 filas máx.) |
| `T+7` | Subtotales de la agrupación |
| `T+10` | Fila de totales (retenciones, recargo, descuentos, IRPF) |
| `T+12` | B = Neto a pagar · E = Valor bruto |
| `T+13` | B = Importe total |

---

## 3. PASO 1 — Cálculo de la línea de producto

Disparador: `onEdit` sobre la columna **B** (producto) o **C** (cantidad), dentro del rango `15 … T-4`.
Código: `mainScript.js` líneas ~1679-1748.

Se escriben **fórmulas** (no valores) en E, F y K:

```
E{i} = D{i} + (D{i} * G{i})

F{i} = (D{i} * C{i}) - ((D{i} * C{i}) * H{i})

K{i} = IF(F{i}=""; 0; F{i} * (1 + G{i} + I{i} - J{i}))
```

Equivalencias:

```
E  = precio unitario con IVA
F  = base gravable neta          = precio × cantidad × (1 − %descuento)
K  = total de línea              = base × (1 + %IVA + %recargo − %retención)
```

> **Clave:** `F` ya viene **neta de descuento**. Todo el resto del modelo cuelga de `F`.

Si la cantidad está vacía se escribe `D = 0` y solo se pone la fórmula de `K` (sin E ni F).

---

## 4. PASO 2 — Contador de líneas

`updateTotalProductCounter()` (mainScript.js:2548) cuenta las filas `15…n` con `B` no vacío y lo escribe en:

```
B{T-3} = nº de líneas con producto
```

Ese contador es el que después usa el generador de JSON para saber cuántas filas leer.

---

## 5. PASO 3 — Agrupación de impuestos (`calcularImporteYTotal`, mainScript.js:2433)

Filas `T+1` … `T+5` (hasta 5 tipos distintos).

**Bloque IVA (columnas A, B, C):**

```
B{T+1} = UNIQUE(G15:G{n})

A{T+1} = ARRAYFORMULA(SUMIF(G15:G{n}; B{T+1}:B{T+5}; F15:F{n}))

C{T+1} = A{T+1} * B{T+1}        ← viene de la plantilla, el código NO la escribe
```

**Bloque recargo de equivalencia (columnas E, F, G):**

```
F{T+1} = UNIQUE(I15:I{n})

E{T+1} = ARRAYFORMULA(SUMIF(I15:I{n}; F{T+1}:F{T+5}; F15:F{n}))

G{T+1} = E{T+1} * F{T+1}        ← viene de la plantilla, el código NO la escribe
```

Lectura: `B` = los % de IVA distintos que hay en la factura, `A` = la base imponible acumulada de cada uno,
`C` = la cuota. Igual para recargo con `F` / `E` / `G`.

---

## 6. PASO 4 — Subtotales de la agrupación (fila `T+7`)

```
A{T+7} = SUM(A{T+1}:A{T+5})     → Total base imponible
C{T+7} = SUM(C{T+1}:C{T+5})     → Total cuota IVA
E{T+7} = SUM(E{T+1}:E{T+5})     → Total base con recargo
G{T+7} = SUM(G{T+1}:G{T+5})     → Total cuota recargo
```

---

## 7. PASO 5 — Fila de totales (`T+10`)

```
A{T+10} = SUMPRODUCT(F15:F{n}; J15:J{n})     → Total retenciones

B{T+10} = SUMPRODUCT(F15:F{n}; I15:I{n})     → Total recargo de equivalencia

D{T+10} = D{T-2} + SUMPRODUCT(D15:D{n}; C15:C{n}; H15:H{n})
          → Total descuentos = descuento de factura + descuentos de línea

E{T+10} = Valor IRPF            ← lo calcula la plantilla a partir de F{T-2} / H{T-2}

F{T+10} = (se limpia con clearContent)
```

**Reglas del IRPF** (`onEdit`, mainScript.js:1822-1860):

- `F{T-2}` = "7%", "15%" o "19%" → `H{T-2}` se bloquea y se limpia.
- `F{T-2}` = "Valor libre" → `H{T-2}` queda editable (se le pone la nota `VALOR_LIBRE`) y se acepta un **importe fijo**, no un porcentaje.
- Editar `H{T-2}` sin estar en modo "Valor libre" se revierte con alerta.

---

## 8. PASO 6 — Importe total y neto a pagar

```
B{T+13} = A{T+7} + C{T+7} + G{T+7}
          → Importe total = base imponible + IVA + recargo

B{T+12} = B{T+13} - A{T+10} - D{T-2} + C{T+10} - E{T+10}
          → Neto a pagar = importe total − retenciones − descuento factura + cargos − IRPF

E{T+12} = SUMPRODUCT(C15:C{n}; D15:D{n})
          → Valor bruto (cantidad × precio, sin descuentos ni impuestos)
```

**Caso especial** (mainScript.js:1866-1873): cuando solo hay **una** línea de producto
(`lastRowProducto === 15`), no se llama a `calcularImporteYTotal()` y se escriben las fórmulas
hardcodeadas de la plantilla por defecto (`T = 19`):

```
B32 = A26 + C26 + G26
B31 = B32 - A29 - D17 + C29 - E29
```

---

## 9. PASO 7 — Recálculo por línea en JS (`guardarYGenerarInvoice`, Factura.js:1962)

Al generar la factura **no se confía en los totales de la hoja**: se lee `A{i}:K{i}` de cada línea
y se recalcula todo en JavaScript, redondeando a 2 decimales en **cada** paso.

```js
const round2 = (n) => Math.round((Number(n) || 0) * 100) / 100;

// Lectura de la fila (índices 0-based del rango A:K)
cantidad        = productoData[2]   // C
precioUnitario  = productoData[3]   // D
subtotalHoja    = productoData[5]   // F  (se lee pero NO se usa para calcular)
ivaRate         = productoData[6]   // G
descuentoRate   = productoData[7]   // H
recargoRate     = productoData[8]   // I
retencionRate   = productoData[9]   // J
totalLinea      = productoData[10]  // K  (se lee pero NO se usa)

// Recálculo
baseBruta      = round2(precioUnitario * cantidad);
discountAmount = round2(baseBruta * descuentoRate);
baseNeta       = round2(baseBruta - discountAmount);      // ≈ columna F
taxAmount      = round2(baseNeta * ivaRate);
withHoldings   = round2(baseNeta * retencionRate);
surCharges     = round2(baseNeta * recargoRate);
```

### Payload de cada producto

```js
{
  typeUse: "VEN",
  reference, description, regime, productType,
  unitPrice: precioUnitario,
  quantity: Math.max(1, Math.trunc(cantidad)),   // entero, mínimo 1
  subTotal: baseBruta,          // ← BRUTO, antes de descuento (evita doble descuento)
  totalTax: taxAmount,          // ← solo IVA, el recargo va aparte
  totalwithHoldings: withHoldings,
  totalSurCharges: surCharges,
  totaldiscount: discountAmount,
  taxes: [...],
  withHoldingsSurChargesDto: [...],
  discountDtoModules: [...]
}
```

Reglas asociadas:

- **`regime`** se fuerza a `"18"` si la línea tiene recargo de equivalencia (`recargoRate > 0`).
- **Retención** → `withHoldingsSurChargesDto` con `idRateWithHoldings: "10"`, `isWithHolding: true`, `subTotalWithHoldings: baseNeta`.
- **Recargo** → `withHoldingsSurChargesDto` con `idRateWithHoldings: "11"`, `isWithHolding: false`, `subTotalWithHoldings: baseNeta`.
- **Descuento de línea** → `discountDtoModules` con `discountBase: baseBruta`, `discountRate: %`, `valueDiscount: discountAmount`.
- **Validación:** si alguna línea lleva recargo y el cliente no es *Autónomo / Persona Física*, se lanza error y no se genera la factura.

---

## 10. PASO 8 — Agrupación de impuestos en el JSON (`fieldTaxations`)

Se construyen dos diccionarios y luego se concatenan.

**IVA > 0** — clave `taxCode + '_' + (rate*100)`:

```js
{ taxName, rate: ivaRate*100, taxBase: Σ baseNeta, valueTax: Σ taxAmount,
  taxCode, qualificationOperation, regime }
```

**IVA = 0** — clave `taxCode + '_0' (+ '_ex' si exento)`:

```js
{ taxName, rate: 0, taxBase: Σ baseNeta, valueTax: 0, taxCode,
  isExemptOperation, idExento }     // idExento solo si el producto es exento
```

y además acumula `totalExemptBase += baseNeta`.

**Recargo** — entrada separada, clave `taxCode + '_' + (recargo*100)`:

```js
{ taxName: "RecargoEquivalencia", rate: recargoRate*100,
  taxBase: Σ baseNeta, valueTax: Σ surCharges, taxCode, regime: "18" }
```

Códigos de impuesto: `01` = IVA · `02` = IPSI · `03` = IGIC · `05` = Otros.

---

## 11. PASO 9 — `chargeAndDiscount` (nivel factura)

Se leen **de la hoja**, no se recalculan:

| Concepto | Celda origen | Entrada generada |
|---|---|---|
| Cargos | `B{T-2}` | `{ idtypeFeeDiscount: "CG", idTypeValueFeeDiscount: "VR", baseFeeDiscount: 0, valueFeeDiscount, totalFeeDiscount }` |
| Descuento de factura | `D{T-2}` | `{ "DT", "VR", 0, valor, valor }` |
| IRPF global | `E{T+10}` | `"RT"` — modo **PJ** si el selector era 7/15/19% (`baseFeeDiscount = totalTaxBase`, `valueFeeDiscount = %`), modo **VR** si era "Valor libre" |

> El descuento de factura que se envía es **solo `D{T-2}`**, no el total `D{T+10}` (que ya incluye los descuentos de línea, y esos van en `discountDtoModules` de cada producto).

---

## 12. PASO 10 — Totales finales del JSON

```js
sumTotalSubTotal        = Σ baseBruta                      // bruto, antes de descuentos
sumTotalTaxBase         = Σ baseNeta                       // base imponible
sumTotalTax             = Σ taxAmount                      // solo IVA
sumTotalSubTotalAndTax  = sumTotalTaxBase + sumTotalTax
sumTotalTotal           = sumTotalSubTotalAndTax + Σ surCharges
sumTotalNetPayable      = sumTotalTotal
                          - Σ withHoldings      // retenciones por producto
                          - descuentoFactura    // D{T-2}
                          + cargoTotal          // B{T-2}
                          - irpfGlobalAmount    // E{T+10}

sumTotalExemptBase      = Σ baseNeta de líneas con IVA 0%
sumTotalDiscount        = descuentoFactura      // solo el de factura
sumTotalCharge          = cargoTotal
sumTotalRetentionIRPF   = irpfGlobalAmount
valueExemptBase         = totalExemptBase
idOperationsExenta      = totalExemptBase > 0 ? (primer idExento || "E1") : "E0"
```

Equivalencia con la hoja:

| JSON | Celda equivalente |
|---|---|
| `sumTotalTaxBase` | `A{T+7}` |
| `sumTotalTax` | `C{T+7}` |
| `Σ surCharges` | `G{T+7}` = `B{T+10}` |
| `sumTotalTotal` | `B{T+13}` |
| `sumTotalNetPayable` | `B{T+12}` |
| `Σ withHoldings` | `A{T+10}` |
| `sumTotalSubTotal` | `E{T+12}` (valor bruto) |

---

## 13. Referencia rápida de celdas (plantilla por defecto, `T = 19`)

| Celda | Contenido |
|---|---|
| B2 / C2 | Cliente |
| G2 | Número de factura |
| G3 | Fecha de vencimiento |
| G4 | Fecha de factura |
| G6 | Días de vencimiento |
| G7 | Fecha de operación |
| E4 / G5 | Medio de pago / Forma de pago |
| B11 | IBAN · D11 | Nota de pago |
| 15 … 15+k | Líneas de producto |
| A16 / B16 | "Total filas" / contador |
| B17 | Cargos |
| D17 | Descuento de factura |
| F17 / H17 | Selector IRPF / valor libre IRPF |
| 19 | "Base imponible" (cabecera) |
| 20 – 24 | Agrupación de impuestos |
| 26 | Subtotales de la agrupación |
| 29 | Totales (retenciones / recargo / descuentos / IRPF) |
| B31 / E31 | Neto a pagar / Valor bruto |
| B32 | Importe total |

---

## 14. Inconsistencias detectadas (a corregir al portar)

1. **Columnas invertidas en la primera línea** — `agregarProductoDesdeFactura` (Factura.js:508-513) escribe `I15 = retención` y `J15 = recargo`, al revés del resto del código (`onEdit` y el lector del JSON usan `I = recargo`, `J = retención`). Esa rama tampoco escribe las fórmulas de `E`, `F` ni `K`.
2. **`K` no es sumable** — el total de línea resta la retención (`−J`), pero el importe total de la factura no la resta (se descuenta después, en el neto). `SUM(K15:K{n}) ≠ B{T+13}`.
3. **`C{T+10}` en la fórmula del neto** — se suma como "cargos" pero el código nunca escribe esa celda; el JSON en cambio lee los cargos de `B{T-2}`. Si la plantilla no tiene `C29 = B17`, hoja y PDF discrepan.
4. **`SUMPRODUCT` de descuentos** — `D{T+10}` usa `D × C × H`, que es descuento sobre bruto; correcto, pero conviene verificar que coincida con `Σ discountAmount` del JSON (que redondea a 2 decimales por línea y la hoja no).
5. **Redondeo asimétrico** — la hoja trabaja con precisión completa y el JSON redondea a 2 decimales en cada paso intermedio. Con muchas líneas puede aparecer 1-2 céntimos de diferencia entre lo que ve el usuario y lo que recibe la API.
6. **Código muerto** — `totalFactura`, `netoPagar` (Factura.js:2371-2384) y `baseNetaTotal` (2514) se calculan y nunca se usan; además en esa rama se asignan cruzados (`totalFactura = B31`, cuando B31 es el neto).
7. **Límite de 5 grupos de impuesto** — la agrupación ocupa `T+1 … T+5`. Con más de 5 tipos de IVA (o de recargo) distintos, el `UNIQUE`/`ARRAYFORMULA` desborda sobre la fila de subtotales.
