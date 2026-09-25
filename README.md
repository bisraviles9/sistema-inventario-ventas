# Sistema de inventario y ventas

Aplicación local en Python con panel de control, base de datos SQLite y soporte para Excel. La interfaz está en español y maneja valores en USD.

## Funciones incluidas

- Dashboard: ventas del mes, unidades y valor de inventario, productos con stock bajo, tendencia de ventas y productos más vendidos.
- Productos: alta y actualización por SKU; control de costo, precio de venta, existencias y stock mínimo.
- Ventas: carrito con varios productos, cliente opcional, forma de pago e historial por fechas. Al confirmar una venta se valida y descuenta el inventario.
- Excel: importar y actualizar productos desde `.xlsx`; exportar cuatro hojas: Productos, Ventas, Detalle de ventas y Stock bajo.
- Base local SQLite: los datos quedan guardados en `inventario.db` junto a la aplicación.

## Cómo iniciar

1. Instala Python 3.10 o posterior.
2. Abre una terminal dentro de esta carpeta.
3. Instala las dependencias: `python -m pip install -r requirements.txt`
4. Ejecuta: `python -m streamlit run app.py`
5. Streamlit abrirá el sistema en el navegador.

## Importar productos desde Excel

Descarga la plantilla desde la pestaña Excel de la aplicación. La primera hoja incluye estas columnas: `SKU`, `Nombre`, `Categoría`, `Precio compra`, `Precio venta`, `Stock` y `Stock mínimo`. La plantilla tiene una fila de ejemplo; elimínala o reemplázala antes de importar.

Los SKU se usan como identificador único. Si importas un SKU que ya existe, se actualizan sus datos y stock; por eso conviene exportar o respaldar el inventario antes de importar cambios masivos.

## Respaldo y alcance

Cierra la aplicación y copia `inventario.db` para respaldar la información. Esta es una versión inicial para uso local y una persona/equipo pequeño; no incluye usuarios, permisos, sincronización en la nube ni control de múltiples cajas simultáneas. Antes de usarla para operación crítica, prueba el flujo con datos de muestra y conserva copias de seguridad.
