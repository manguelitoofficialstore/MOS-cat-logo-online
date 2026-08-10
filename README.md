# MOS Store

Tienda estática preparada para GitHub Pages. El catálogo se genera automáticamente desde `Productos.xlsx` cada vez que la página carga y vuelve a consultar el archivo cada 60 segundos.

## Estructura del Excel utilizada

La primera hoja contiene los productos. El sistema usa estas columnas:

- A `Referencia`
- B `Colección`
- C `categoria`
- D `subcategoría`
- E `Descripción`
- F `Talla`
- G `Género`
- H `Color`
- L `Precio de venta`
- M `Inventario`
- P `Foto`

La descripción visible se construye exactamente en este orden:

`E (Descripción) + F (Talla) + H (Color) + G (Género)`

La columna L es la única utilizada como precio mostrado al cliente.

## Orden del catálogo

1. Productos oficiales
   1. Ropa
   2. Accesorios
   3. Items
2. Colección Basic
   1. Ropa
   2. Accesorios
   3. Items

El código normaliza automáticamente `Accesorio`/`Accesorios` e `Item`/`Items`.

## Variantes e inventario

Cada combinación de referencia, talla, color y género se maneja como una variante independiente. Esto permite que una referencia repetida en el Excel con tallas diferentes se agregue al carrito sin mezclar inventarios.

Las filas con inventario `0` permanecen visibles como `Agotado`. Las filas sin referencia, colección, categoría o precio de venta no se publican.

## Fotografías

Coloca las imágenes en la carpeta `Imágenes_Fnales/`.

La tienda toma el nombre indicado en la columna `P (Foto)` y busca ese archivo dentro de `Imágenes_Fnales`. Si el nombre viene sin extensión, la tienda intenta automáticamente: nombre exacto, `.jpg`, `.jpeg`, `.png` y `.webp`.

## Interacción del producto

Cada producto tiene un botón **Ver información**. Ese botón abre un panel de detalle con imagen grande, referencia, precio de venta, descripción, color, talla, género, existencia, selector de cantidad, subtotal y botón **Agregar al carrito**.

## WhatsApp

La segunda hoja contiene los números de venta en la primera columna. Se leen todos los valores desde A1, sin exigir encabezado. Los celulares colombianos de 10 dígitos que empiezan por `3` reciben automáticamente el indicativo `57` para abrir WhatsApp.

La finalización de compra siempre se hace desde el carrito, tomando el número de ventas seleccionado desde la configuración del Excel y generando un informe completo del pedido para continuar la atención por WhatsApp.


## Logos oficiales integrados

Se integraron los logos entregados en `assets/logos/`.

- El encabezado principal de la tienda usa `mos-oficial.png`.
- La sección **Productos oficiales** muestra el logo oficial MOS.
- La sección **Colección Basic** muestra el logo `mb.png`.
- El hero principal utiliza el banner `mos-banner.png`.

Las demás variantes del logo quedaron incluidas dentro de `assets/logos/` como respaldo visual para futuros ajustes.
