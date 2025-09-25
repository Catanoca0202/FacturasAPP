# FacturasApp Project Context

## Product Overview
- FacturasApp es un producto interno de la empresa enfocado en la generación y administración de facturas (invoices) para clientes finales.
- Se expone como un complemento de Google Sheets (Apps Script) que guía al usuario para crear clientes, productos y facturas sin salir de la hoja de cálculo.
- La experiencia se apoya en barras laterales (sidebar) para navegar opciones de menú, crear registros y ejecutar acciones administrativas.

## Hojas Principales
- **Clientes**: mantiene registros completos de clientes. Columnas destacadas: identificador único (columna B), número de identificación (F), código de contacto (G), régimen (H), nombres y apellidos (J–M), país (N), provincia (O), población (P) y dirección (Q). Validaciones de datos controlan listas desplegables.
- **Productos**: administra catálogo de productos/servicios. Columnas clave: estado (A), código de referencia (B), nombre (C), tipo de producto (D), tipo de uso (E), valor unitario (F), tipo impuesto (G), tarifa impuesto (H), precio con impuesto (I), aplicar recargo (J, checkbox), tipo retención (K), tarifa retención (L) e identificador único (M).
- **Factura**: interfaz principal para armar facturas. Incluye selección de cliente (A2), NIF, datos de pago (F4–G7), observaciones/IBAN (fila 11), tabla de líneas (filas 13–33) con referencia, producto, cantidad y cálculos de IVA/retención/recargo, además de totales en la parte inferior.
- **Datos (hoja intermedia)**: hoja oculta que sincroniza información entre *Factura*, *Clientes* y *Productos*. Contiene:
  - Celdas para lookup de clientes/prod (p.ej. `I2` cliente seleccionado, `I11` producto) con fórmulas tipo `INDEX/MATCH` (ver ejemplos en capturas “HojaDatos2”).
  - Campos auxiliares como API key (`B23` en captura) y parámetros para validaciones.
  - Posiciones en hoja (columnas G–H) para saber en qué fila están los registros seleccionados.

## Flujo Operativo
1. El usuario crea/actualiza clientes y productos mediante formularios en el sidebar o editando directamente las hojas dedicadas.
2. En la hoja *Factura*, la validación de datos permite elegir cliente/producto; esa selección se escribe en la hoja *Datos*.
3. Fórmulas de *Datos* buscan en *Clientes* y *Productos* y devuelven los campos requeridos (identificación, régimen, IVA, descuentos, etc.).
4. Las celdas enlazadas en *Factura* consumen esos valores para completar la factura.
5. Antes: se generaba el PDF/archivo de factura en Google Drive (con posibilidad de crear carpeta en OneDrive). Ahora: se arma un JSON conforme a los requerimientos de la API de FacturasApp y se envía para que el backend genere y entregue la factura descargable.

## Sidebar y Menús
- `mainScript.js` define el menú **FacturasApp** con acciones: Inicio (`showSidebar2`), Instalar (copia hojas desde plantilla) y Desinstalar.
- El sidebar (`main.html` y asociados) habilita navegación y acciones como crear clientes (`showNuevaClienteDesdeFactura`, `showNuevaClienteV2`) o productos (`showAggProductos`).
- El flujo de instalación copia hojas desde la plantilla Drive `1qxbXlhH4RpCOsObk91wsuu4k8jarVK34XXRUlKaKS1U`, protege hojas sensibles y oculta *Datos* automáticamente.

## Integraciones y Configuración
- `appsscript.json` habilita servicios avanzados de Google Sheets v4 y Drive v3, define zona horaria `Europe/Madrid` y autoriza scopes de lectura/escritura en hojas, Drive, envío de correo y peticiones externas (`UrlFetchApp`).
- `Data.md` lista IDs de los libros de prueba (`prueba desarollo`) y producción (`produccion/version`) utilizados como contenedores de hojas.
- JSON de ejemplo (`factura*.json`, `1producto.json`, etc.) describen payloads de factura/producto alineados con la API.

## Consideraciones Técnicas
- La lógica de Apps Script depende críticamente de rangos específicos en *Datos*, *Clientes* y *Productos*. Cualquier cambio en encabezados o desplazamientos debe sincronizarse con las funciones que usan `getRange`.
- El proceso de instalación desinstala/reinstala la hoja *Datos* para asegurar que las protecciones y fórmulas estén actualizadas.
- Validaciones de datos se “queman” directamente en las hojas para evitar errores al copiar desde la plantilla.

## Referencias Visuales
- Capturas proporcionadas: “HojaDatos2”, “Clientes”, “Factura”, “Productos” ilustran la disposición exacta de columnas y celdas activas. Consultarlas para ubicar columnas críticas y validaciones.

Mantén este documento actualizado cuando cambien flujos, rangos clave o integraciones externas.
