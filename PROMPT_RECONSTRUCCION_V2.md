# Prompt para Reconstruir FacturasApp v2 - Google Apps Script Extension

## 🎯 CONTEXTO DEL PROYECTO

Eres un experto desarrollador especializado en Google Apps Script, complementos de Google Sheets, arquitectura de software y mejores prácticas de desarrollo. Tu misión es reconstruir desde cero **FacturasApp**, un complemento de Google Sheets para la generación y administración de facturas electrónicas en España.

La versión actual funciona pero tiene problemas de mantenibilidad, código duplicado, falta de modularización y testing. El objetivo es crear una versión 2.0 con arquitectura limpia, modular, testeable y extensible.

---

## 📋 DESCRIPCIÓN DEL PRODUCTO

### Propósito
FacturasApp es un complemento (add-on) de Google Sheets que permite a usuarios españoles:
- Gestionar clientes (contactos)
- Administrar catálogo de productos/servicios
- Crear facturas con cálculos automáticos de IVA, retenciones IRPF y recargo de equivalencia
- Enviar facturas a la API de FacturasApp para generación oficial
- Descargar PDFs y enviarlos por email

### Stack Tecnológico
- **Backend**: Google Apps Script (V8 runtime)
- **Frontend**: HTML/CSS/JavaScript en Sidebar (Bootstrap 5)
- **Servicios Google**: Sheets API v4, Drive API v3, MailApp
- **API Externa**: FacturasApp REST API (producción + QA)
- **Integraciones futuras**: n8n webhooks para MCP de clientes/productos

---

## 🏗️ ARQUITECTURA ACTUAL (A MEJORAR)

### Archivos JavaScript Principales
```
mainScript.js    - Menú, funciones de UI, instalación, validaciones de productos
Cliente.js       - CRUD clientes, catálogos de ubicación (País/Provincia/Población)
productos.js     - Búsqueda y obtención de información de productos
Factura.js       - Lógica de facturación, generación JSON, envío API, historial
cards.js         - Cards para el homepage trigger del add-on
```

### Hojas del Spreadsheet
| Hoja | Propósito | Visible |
|------|-----------|---------|
| Inicio | Landing page del usuario | ✅ |
| Clientes | Registro de clientes/contactos | ✅ |
| Productos | Catálogo de productos/servicios | ✅ |
| Factura | Interfaz principal para crear facturas | ✅ |
| Historial Facturas | Vista de facturas creadas | ✅ |
| Historial Facturas Data | Datos crudos del historial | ❌ |
| Datos de emisor | Configuración del emisor (nombre, IBAN, etc.) | ✅ |
| Datos | Hoja intermedia con fórmulas INDEX/MATCH | ❌ |
| ListadoEstado | Almacena JSONs de facturas enviadas | ❌ |
| Facturas ID | Registro de IDs de facturas | ❌ |
| Country/Province/Population | Catálogos de ubicaciones | ❌ |
| Copia de Factura | Plantilla para limpiar factura | ❌ |
| ClientesInvalidos | Clientes inactivados | ❌ |

### Problemas Identificados en la Versión Actual
1. **Código espagueti**: Funciones de 200+ líneas sin separación de responsabilidades
2. **Hardcoding de rangos**: `getRange("A15")`, `getRange("B2:C2")` por todo el código
3. **Sin tests**: Cero cobertura de pruebas
4. **Duplicación**: Misma lógica repetida en múltiples lugares
5. **Acoplamiento UI-Lógica**: La lógica de negocio está mezclada con manipulación de hojas
6. **Manejo de errores inconsistente**: Mix de `throw`, `alert()`, y returns silenciosos
7. **Sin documentación de código**: Pocos comentarios y sin JSDoc
8. **Configuración hardcodeada**: URLs de API, IDs de plantillas en el código
9. **Estado global**: Variables globales y `PropertiesService` usado sin estructura

---

## 📐 ARQUITECTURA PROPUESTA V2

### Principios de Diseño
1. **Separación de capas**: UI → Servicios → Repositorios → Hojas
2. **Configuración centralizada**: Constantes en un solo lugar
3. **Dependency Injection**: Facilitar testing con mocks
4. **Handlers de errores unificados**: Try-catch consistente
5. **Validación en capas**: DTOs validados antes de procesamiento
6. **Documentación JSDoc**: Tipos y documentación en todas las funciones públicas

### Estructura de Archivos Propuesta
```
/src
  /config
    Config.js           - URLs, IDs de plantilla, constantes
    SheetNames.js       - Nombres de hojas como constantes
    ColumnMaps.js       - Mapeo de columnas por hoja
  
  /models
    Cliente.js          - DTO/Modelo de Cliente
    Producto.js         - DTO/Modelo de Producto
    Factura.js          - DTO/Modelo de Factura
    LineaFactura.js     - DTO/Modelo de línea de producto en factura
    
  /repositories
    BaseRepository.js   - Clase base con operaciones CRUD genéricas
    ClienteRepository.js
    ProductoRepository.js
    FacturaRepository.js
    ConfigRepository.js  - Lee/escribe configuración del emisor
    
  /services
    ClienteService.js   - Lógica de negocio de clientes
    ProductoService.js  - Lógica de negocio de productos
    FacturaService.js   - Lógica de facturación, cálculos, JSON
    APIService.js       - Comunicación con API FacturasApp
    EmailService.js     - Envío de emails con PDFs
    CatalogoService.js  - País/Provincia/Población
    ValidationService.js - Validaciones centralizadas
    
  /controllers
    MenuController.js   - Manejo del menú del add-on
    SidebarController.js - Funciones llamadas desde el sidebar
    TriggerController.js - onOpen, onEdit, onInstall
    
  /utils
    DateUtils.js        - Formateo de fechas (zona horaria España)
    NumberUtils.js      - Formateo de moneda, porcentajes
    StringUtils.js      - Normalización, slugify
    SheetUtils.js       - Helpers para manipular hojas
    ErrorHandler.js     - Manejo centralizado de errores
    Logger.js           - Wrapper sobre console.log con niveles
    
  /ui
    main.html           - SPA principal del sidebar
    /views              - Vistas parciales
    /css                - Estilos (si se embeben)
    
  /tests
    /mocks
      MockSheet.js
      MockSpreadsheet.js
    ClienteService.test.js
    ProductoService.test.js
    FacturaService.test.js
    ValidationService.test.js

appsscript.json         - Manifest del add-on
```

---

## 📊 MODELO DE DATOS

### Hoja: Clientes
| Columna | Campo | Tipo | Obligatorio |
|---------|-------|------|-------------|
| A | Estado | "Valido"/"No Valido" | Auto |
| B | Identificador único | "Nombre-NIF" | Auto |
| C | Tipo contacto | "Cliente"/"Proveedor" | ✅ |
| D | Tipo persona | "Autónomo"/"Empresa"/"Persona Física" | ✅ |
| E | Tipo documento | NIF-IVA, Pasaporte, etc. | ✅ |
| F | Número identificación | String | ✅ (único) |
| G | Código contacto | String | ✅ (único) |
| H | Régimen fiscal | Dropdown con 20 opciones | ✅ |
| I | Nombre comercial | String | ✅ si Empresa |
| J | Primer nombre | String | ✅ si Autónomo |
| K | Segundo nombre | String | - |
| L | Primer apellido | String | ✅ si Autónomo |
| M | Segundo apellido | String | - |
| N | País | Dropdown | ✅ |
| O | Provincia | Dropdown dependiente | ✅ |
| P | Población | Dropdown dependiente | ✅ |
| Q | Dirección | String | - |
| R | Código postal | String | ✅ |
| S | Teléfono | String | - |
| T | Sitio web | URL | - |
| U | Email | Email válido | ✅ |

### Hoja: Productos
| Columna | Campo | Tipo | Obligatorio |
|---------|-------|------|-------------|
| A | Estado | "Valido"/"No Valido" | Auto |
| B | Código referencia | String | ✅ (único) |
| C | Nombre | String | ✅ |
| D | Tipo producto | "Producto"/"Servicio" | ✅ |
| E | Tipo uso | "Venta"/"Compra" | ✅ |
| F | Valor unitario | Número €#,##0.00 | ✅ |
| G | Tipo impuesto | "IVA" | ✅ |
| H | Tarifa impuesto | 0%/4%/10%/21% | ✅ |
| I | Precio con impuesto | Fórmula | Auto |
| J | Aplicar recargo | Checkbox | - |
| K | Tarifa recargo | 0.5%/1.4%/5.2% según IVA | Condicional |
| L | Aplicar retención | Checkbox | - |
| M | Tarifa retención | 7%/15%/19% IRPF | Condicional |
| N | Identificador único | "Nombre-Código" | Auto |

### Hoja: Factura (Interfaz)
| Celda | Campo |
|-------|-------|
| B2:C2 | Cliente (dropdown) |
| B3 | Código cliente |
| E4 | Medio de pago |
| G2 | Número factura (consecutivo) |
| G3 | Fecha vencimiento |
| G4 | Fecha emisión |
| G5 | Forma de pago |
| G6 | Días de vencimiento |
| G7 | Hora emisión |
| G8 | Asesor comercial |
| B10 | Observaciones |
| B11 | IBAN |
| D11 | Observaciones de pago |
| Filas 15+ | Líneas de productos |

### Línea de Producto en Factura
| Columna | Campo |
|---------|-------|
| A | Código referencia |
| B | Producto (dropdown) |
| C | Cantidad |
| D | Precio unitario |
| E | Precio con IVA (fórmula) |
| F | Subtotal (fórmula) |
| G | % IVA |
| H | % Descuento |
| I | % Recargo equivalencia |
| J | % Retención |
| K | Total línea (fórmula) |
| L | Checkbox eliminar |

---

## 🔄 FLUJOS PRINCIPALES

### 1. Instalación
```
Usuario instala add-on
  → onOpen crea menú "FacturasApp"
  → Usuario clica "Instalar"
  → IniciarFacturasApp():
      1. Copia hojas desde plantilla (ID: 1-ZkL7SKO8IqBwgfj9bta1ELuZoXcelp4K1a_Xd2FA0c)
      2. Aplica protecciones a hojas sensibles
      3. Oculta hojas internas (Datos, Copia de Factura, etc.)
      4. Aplica Data Validations (dropdowns)
      5. Configura locale a es_ES
```

### 2. Crear Cliente
```
Usuario abre sidebar → Contactos → Crear contacto
  → Formulario con campos dinámicos según tipo persona
  → País → carga Provincias → carga Poblaciones (cascada)
  → Validación frontend de campos obligatorios
  → saveClientData(formData):
      1. Verificar unicidad de NIF y Código
      2. Encontrar primera fila vacía
      3. Escribir datos
      4. Calcular identificador único
      5. Establecer estado "Valido"
```

### 3. Crear Producto
```
Usuario abre sidebar → Productos → Crear producto
  → processForm(data):
      1. Verificar unicidad de código referencia
      2. Normalizar IVA a valores permitidos (0,4,10,21)
      3. Calcular recargo según IVA si aplica
      4. Escribir en hoja Productos
      5. Establecer validaciones y fórmulas
```

### 4. Crear Factura
```
Usuario selecciona cliente en hoja Factura (B2)
  → onEdit detecta cambio:
      1. verificarYCopiarContacto() - valida estado cliente
      2. obtenerFechaYHoraActual() - fecha emisión
      3. generarNumeroFactura() - siguiente consecutivo
      
Usuario agrega productos (dropdown B15+)
  → onEdit detecta cambio:
      1. obtenerInformacionProducto() - lookup en hoja Datos
      2. Calcular IVA, recargo, retención
      3. Actualizar fórmulas de totales
      
Usuario clica "Guardar"
  → guardarFactura():
      1. verificarEstadoValidoFactura() - todas las validaciones
      2. verificarEstadoConsecutivo() - consecutivo configurado
      3. guardarYGenerarInvoice() - construir JSON completo
      4. guardarFacturaHistorial() - registro en hoja historial
      5. enviarFactura() - POST a API FacturasApp
      6. limpiarHojaFactura() - reiniciar para siguiente
```

### 5. Enviar Factura por Email
```
Usuario selecciona factura del historial
  → enviarEmailPostFactura(email, historial, numFactura):
      1. obtenerPDFFacturaBase64(numFactura) - GET de API
      2. Convertir base64 a Blob
      3. MailApp.sendEmail() con attachment
```

---

## 🧮 REGLAS DE NEGOCIO CRÍTICAS

### IVA (España 2025)
```javascript
const IVA_PERMITIDOS = [0, 4, 10, 21]; // porcentajes
```

### Recargo de Equivalencia (Solo para Autónomos)
```javascript
const IVA_RECARGO_MAP = {
  '21': 5.2,  // IVA 21% → Recargo 5.2%
  '10': 1.4,  // IVA 10% → Recargo 1.4%
  '5': 0.5,   // IVA 5% → Recargo 0.5%
  '4': 0.5,   // IVA 4% → Recargo 0.5%
  '0': 0
};
```
- **Regla**: Solo aplica a clientes tipo "Autónomo" (no "Persona Física" ni "Empresa")
- **Regla**: No aplica a productos tipo "Servicio"

### Retenciones IRPF
```javascript
const RETENCION_TARIFAS = ['7%', '15%', '19%'];
const RETENCION_CODIGOS = {
  7: "20",
  15: "21", 
  19: "22"
};
```

### Consecutivo de Factura
- Formato: `{PREFIJO}{NÚMERO con padding}`
- Ejemplo: `ABC001`, `FACT-0001`, `2025/00001`
- Se guarda en `DocumentProperties`: 
  - `LetraConescutivo`: prefijo
  - `NumeroConescutivo`: número actual
  - `ConsecutivoPlantillaDigitos`: cantidad de dígitos

### Tipos de Persona → Códigos API
```javascript
const TIPO_PERSONA_CODIGOS = {
  'autonomo': '01',
  'persona fisica': '01',
  'empresa': '02'
};
```

### Regímenes Fiscales
```javascript
const REGIMEN_CODIGOS = {
  "Operación de régimen general": "01",
  "Exportación": "02",
  // ... 20 opciones totales
  "Recargo de equivalencia": "18",
  "Régimen simplificado": "20"
};
```

---

## 🔗 API FacturasApp

### Autenticación
```
POST /ApiGateway/AppSecurity/ApiKey
Body: { "User": "email", "Password": "pass" }
Response: ["API_KEY_STRING"]
```
- La API Key se guarda en `Datos!I21`
- Estado de vinculación en `Datos de emisor!B16`

### Enviar Factura
```
POST /ApiGateway/ApiExternal/Invoice/api/InvoiceServices/AddInvoice
Headers: { "X-API-KEY": "..." }
Body: { invoiceNumber, contacts, products, fieldTaxations, ... }
```

### Obtener PDF
```
POST /ApiGateway/ApiExternal/Invoice/api/InvoiceServices/PDFInvoice?invoiceNumber={num}
Headers: { "X-API-KEY": "..." }
Response: { id, messages, isError, toolObject: "BASE64_PDF" }
```

### Ambientes
```javascript
const API_URLS = {
  produccion: 'https://www.facturasapp.com',
  qa: 'https://facturasapp-qa.cenet.ws'
};
```

---

## 📄 ESTRUCTURA JSON DE FACTURA

```json
{
  "textCustomerObservations": "string|null",
  "invoiceNumber": "ABC001",
  "currentNumber": 1,
  "invoiceDate": "2025-01-19T00:00:00.000Z",
  "invoiceTime": "10:30:00.0000000",
  "invoiceExpiration": "30",
  "invoiceIdTypeRegAEAT": "AI",
  "invoiceIdTypeRegSIF": null,
  "contactName": "Asesor comercial",
  
  "contacts": [{
    "contactType": "01|02",
    "personType": "01|02",
    "companyName": "Nombre",
    "customerCode": "CODIGO",
    "identificationType": "02|03|04|05|06|07",
    "identification": "NIF",
    "tradeName": "Nombre comercial",
    "regime": "01-20",
    "country": "207",
    "province": "5102",
    "population": "32653",
    "addressCustomer": "Dirección",
    "postalCodeCustomer": "28001",
    "phoneCustomer": "600000000",
    "webSite": "www.ejemplo.com",
    "emailCustomer": "email@ejemplo.com"
  }],
  
  "products": [{
    "typeUse": "VEN",
    "reference": "CODIGO",
    "description": "Nombre producto",
    "unitPrice": 100.00,
    "quantity": 2,
    "subTotal": 200.00,
    "totalTax": 42.00,
    "totalwithHoldings": 0,
    "totalSurCharges": 0,
    "totaldiscount": 0,
    "taxes": [{
      "taxName": "IVA",
      "rate": 21,
      "taxBase": 200.00,
      "valueTax": 42.00
    }],
    "withHoldingsSurChargesDto": [{
      "idRateWithHoldings": "20|21|22|23|24|25|26",
      "subTotalWithHoldings": 200.00,
      "cuotaWithHoldings": 14.00
    }],
    "discountDtoModules": [{
      "discountName": "Descuento",
      "discountRate": 5,
      "discountBase": 200.00,
      "valueDiscount": 10.00
    }]
  }],
  
  "idPayment": "ND|EF|TF|TB|DB|PP|FR|CF|TL|TC|TD|PA|CH|RB|CD|BZ|LC|CP",
  "paymentNote": "string|null",
  "textObservations": "string|null",
  "idOperations": "S1|N1",
  "idOperationsExenta": "E0|E3",
  "valueExemptBase": 0,
  
  "chargeAndDiscount": [{
    "idtypeFeeDiscount": "CG",
    "idTypeValueFeeDiscount": "PJ",
    "baseFeeDiscount": 200.00,
    "valueFeeDiscount": 1,
    "totalFeeDiscount": 2.00
  }],
  
  "fieldTaxations": [{
    "taxName": "IVA",
    "rate": 21,
    "taxBase": 200.00,
    "valueTax": 42.00
  }],
  
  "sumTotalSubTotal": 200.00,
  "sumTotalTaxBase": 200.00,
  "sumTotalTax": 42.00,
  "sumTotalSubTotalAndTax": 242.00,
  "sumTotalExemptBase": 0,
  "sumTotalDiscount": 0,
  "sumTotalCharge": 0,
  "sumTotalRetentionIRPF": 0,
  "sumTotalTotal": 242.00,
  "sumTotalNetPayable": 242.00,
  
  "invoiceTypeId": 0,
  "invoiceRectificativeTypeId": 0,
  "typeRectificativeId": 0,
  "aditionalData": {
    "invoiceId": 0,
    "startInvoiceId": 0
  }
}
```

---

## 🧪 ESTRATEGIA DE TESTING

### Niveles de Test
1. **Unit Tests**: Servicios y utilidades puras
2. **Integration Tests**: Repositorios con mocks de Sheets
3. **E2E Manual**: Flujos completos en spreadsheet de prueba

### Mocks Necesarios
```javascript
// MockSpreadsheetApp.js
const MockSpreadsheetApp = {
  getActive: () => MockSpreadsheet,
  getUi: () => MockUi
};

// MockSpreadsheet.js
const MockSpreadsheet = {
  getSheetByName: (name) => sheets[name] || null,
  // ...
};

// MockSheet.js
class MockSheet {
  constructor(data) {
    this.data = data;
  }
  getRange(a1) { /* ... */ }
  getLastRow() { /* ... */ }
  // ...
}
```

### Tests Prioritarios
1. `ValidationService.validarCliente()`
2. `FacturaService.calcularTotales()`
3. `FacturaService.construirJSON()`
4. `NumberUtils.parsePercentToNumberES()`
5. `DateUtils.formatearFechaEspana()`

---

## 🔮 INTEGRACIONES FUTURAS (n8n)

### MCP Disponible
El MCP de n8n ya permite:
- Crear/Editar/Eliminar/Consultar Clientes
- Crear/Editar/Eliminar/Consultar Productos

### Webhooks a Implementar
```javascript
// Ejemplo: Sincronizar cliente desde n8n
function webhookCrearCliente(payload) {
  // payload viene de n8n con datos del cliente
  const cliente = ClienteService.crearDesdeWebhook(payload);
  return { success: true, id: cliente.identificadorUnico };
}

// Registrar webhook en appsscript.json o mediante Web App
```

### Consideraciones
- Autenticación de webhooks (API Key o token)
- Rate limiting
- Logs de sincronización
- Manejo de conflictos (mismo cliente editado en ambos lados)

---

## 📝 CRITERIOS DE ACEPTACIÓN

### Funcionalidad
- [ ] Todas las funcionalidades actuales funcionan igual o mejor
- [ ] Los JSONs generados son idénticos a la versión actual
- [ ] La instalación copia todas las hojas correctamente
- [ ] Los dropdowns dependientes (País→Provincia→Población) funcionan
- [ ] El cálculo de IVA, retenciones y recargo es preciso

### Calidad de Código
- [ ] Cobertura de tests > 70% en servicios
- [ ] Sin funciones de más de 50 líneas
- [ ] JSDoc en todas las funciones públicas
- [ ] Constantes centralizadas (cero hardcoding)
- [ ] Manejo de errores consistente

### UX
- [ ] Sidebar SPA fluido sin recargas
- [ ] Mensajes de error claros y en español
- [ ] Validaciones en frontend antes de enviar a backend
- [ ] Indicadores de carga durante operaciones

### Performance
- [ ] Uso de `LockService` para operaciones concurrentes
- [ ] Batch operations donde sea posible
- [ ] Cacheo de catálogos en memoria durante sesión

---

## 🚀 PLAN DE MIGRACIÓN SUGERIDO

### Fase 1: Infraestructura (Semana 1)
1. Crear estructura de carpetas
2. Configurar `appsscript.json` con scopes
3. Implementar Config, Logger, ErrorHandler
4. Crear mocks básicos para testing

### Fase 2: Modelos y Repositorios (Semana 2)
1. Definir DTOs/modelos con validación
2. Implementar BaseRepository
3. Implementar repositorios específicos
4. Tests de repositorios

### Fase 3: Servicios (Semana 3-4)
1. ValidationService
2. CatalogoService
3. ClienteService + tests
4. ProductoService + tests
5. FacturaService + tests
6. APIService + tests

### Fase 4: UI y Controllers (Semana 5)
1. Migrar main.html a nuevo sistema
2. Refactorizar views
3. Implementar controllers
4. Tests E2E manuales

### Fase 5: Migración y QA (Semana 6)
1. Pruebas en spreadsheet de desarrollo
2. Pruebas con API QA
3. Documentación final
4. Deploy a producción

---

## ⚠️ RESTRICCIONES Y CONSIDERACIONES

1. **Compatibilidad**: Los usuarios existentes NO deben perder datos
2. **Plantilla**: El ID de la plantilla debe mantenerse o actualizarse coordinadamente
3. **API**: La estructura JSON no puede cambiar (rompe la API backend)
4. **Scopes OAuth**: Mantener los scopes actuales para no requerir re-autorización
5. **Zona horaria**: Siempre usar `Europe/Madrid`
6. **Locale**: Siempre `es_ES` para formato de números y fechas

---

## 📚 REFERENCIAS

- [Google Apps Script Best Practices](https://developers.google.com/apps-script/guides/support/best-practices)
- [Sheets API Reference](https://developers.google.com/sheets/api/reference/rest)
- [Apps Script Add-on Guidelines](https://developers.google.com/workspace/add-ons/guides/alternate-runtimes)
- ID Plantilla actual: `1-ZkL7SKO8IqBwgfj9bta1ELuZoXcelp4K1a_Xd2FA0c`
- ID Spreadsheet Desarrollo: `1dvaPCCRhKAS7xjs6WhjJcOdN5iJN5ekAY1vn2E_pYhalw6e3M8Q4ivJj`
- ID Spreadsheet QA: `186NgOivey1zIlfzL_1MNh5ubRLOUjrUjQMinKi41N0UmeV9dUZd-jbvI`

---

**NOTA FINAL**: Este prompt contiene toda la información necesaria para reconstruir FacturasApp v2. El agente IA debe:
1. Seguir la arquitectura propuesta
2. Implementar tests desde el inicio (TDD cuando sea posible)
3. Documentar cada decisión de diseño
4. Mantener compatibilidad con la API existente
5. Priorizar mantenibilidad sobre features nuevas

¡Éxito en la reconstrucción! 🚀
