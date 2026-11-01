# Gasto Rápido

Página para cargar gastos en segundos (desde el celu o la compu) que se guardan directo en tu planilla de Google Sheets.

```
Vos ──► Página (GitHub Pages) ──► Apps Script (tu API) ──► Google Sheets (tu planilla)
```

- `docs/` — la página. GitHub Pages la publica tal cual (HTML, CSS y JS, sin build).
- `apps-script/` — el script que va dentro de la planilla y recibe los movimientos.

**Este repo no tiene datos tuyos.** Tus movimientos viven solo en tu Google Sheets. La página guarda la dirección del script y la clave únicamente en el navegador de cada dispositivo.

## Puesta en marcha (una sola vez, ~10 minutos)

### 1. La planilla
1. Subí `Finanzas.xlsx` a Google Drive y abrila con **Hojas de cálculo de Google**.
2. Las fórmulas automáticas de Movimientos (columnas G, H e I) las instala el script en el paso 2 (también: menú Gasto Rápido → Preparar planilla).

### 2. El script (Apps Script)
1. En la planilla: **Extensiones → Apps Script**.
2. Borrá lo que haya en `Código.gs` y pegá el contenido de `apps-script/Codigo.gs`. Guardá.
3. ⚙️ **Configuración del proyecto** → zona horaria **(GMT-03:00) Buenos Aires**.
4. Volvé a la planilla y recargala: aparece el menú **Gasto Rápido**.
5. **Gasto Rápido → Generar clave nueva.** Google pide permisos: *Revisar permisos → tu cuenta → Configuración avanzada → Ir a … (no seguro) → Permitir*. Aparece "no seguro" porque el script es tuyo y no está verificado por Google; es normal.
6. Copiá la clave que aparece.
7. En Apps Script: **Implementar → Nueva implementación → ⚙️ Aplicación web**.
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
   - **Implementar** y copiá la **URL de la aplicación web** (termina en `/exec`).

> "Cualquier usuario" significa que la dirección responde sin pedir login de Google, para que la página pueda escribir. Sin la clave no hace nada: rechaza cualquier pedido.

### 3. La página (GitHub Pages)
1. Creá un repo en GitHub (por ejemplo `gasto-rapido`) y subí estos archivos.
2. **Settings → Pages → Build and deployment**: Source **Deploy from a branch**, Branch **main**, carpeta **/docs** → Save.
3. En un minuto queda en `https://TU-USUARIO.github.io/gasto-rapido/`.

### 4. Conectar
Abrí la página, pegá la URL del script y la clave → **Conectar**. Hacelo una vez en cada dispositivo (celu, compu).

En el celu: menú del navegador → **Agregar a pantalla de inicio**. Queda como una app.

## Uso
- **Escribir:** `cabi 3500`, `ayer kfc 18880`, `vea 143.860 santander`, `+ingreso 155000 icbc`, `pan 1500 #salidas`.
- **Extracto del banco:** subí el Excel/CSV del home banking o pegá la tabla.
- **Fijos del mes:** salen de la hoja Fijos; marcás cuáles y se agregan con la fecha de este mes.
- Revisás la lista y tocás **Guardar en la planilla**. Si no hay internet, queda pendiente en el dispositivo.

Las categorías, cuentas y reglas se editan en la planilla (hojas Config y Reglas); la página las lee de ahí. También podés agregar reglas desde la página.

## Si cambiás el script
Cada vez que edites `Codigo.gs`: **Implementar → Administrar implementaciones → ✏️ → Versión: Nueva versión → Implementar**. La URL no cambia.

## Seguridad
- La clave se puede regenerar desde el menú Gasto Rápido; la anterior deja de funcionar al instante.
- Si perdés un celular, regenerá la clave.
