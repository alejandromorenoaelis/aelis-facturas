# Aelis Facturas — add-in de Outlook

Botón independiente en Outlook que guarda los PDF adjuntos de un correo
en la carpeta de SharePoint del acreedor que elijas.

Es un add-in **separado** de "Aelis Procesos": GUID e IDs propios, se instala
y se desinstala por su cuenta. Los dos pueden convivir en el mismo Outlook.

## Contenido

    manifest.xml       ← el que se carga en Outlook
    facturas.html      ← panel
    facturas.js        ← lógica (aquí va la URL del flujo)
    assets/            ← iconos y logo

## Paso 1 — Publicar en GitHub Pages

1. Crea un repo nuevo llamado **`aelis-facturas`**.
2. Sube el contenido de esta carpeta a la raíz del repo.
3. Settings → Pages → Deploy from branch → `main` / raíz.
4. Comprueba que carga: `https://alejandromorenoaelis.github.io/aelis-facturas/facturas.html`

Si prefieres otro nombre de repo, busca y reemplaza `aelis-facturas` en
`manifest.xml` por el que uses. Aparece 7 veces.

## Paso 2 — Crear el flujo en Power Automate

Flujo nuevo: **GuardarFacturaAcreedor**

**Trigger — Cuando se recibe una solicitud HTTP**, con este esquema:

```json
{
  "type": "object",
  "properties": {
    "acreedor": { "type": "string" },
    "remitente": { "type": "string" },
    "asunto": { "type": "string" },
    "fecha": { "type": "string" },
    "adjuntos": {
      "type": "array",
      "items": {
        "type": "object",
        "properties": {
          "nombre": { "type": "string" },
          "tipo": { "type": "string" },
          "tamano": { "type": "integer" },
          "base64": { "type": "string" }
        }
      }
    }
  }
}
```

**Apply to each** sobre `triggerBody()?['adjuntos']`

**Create file** (dentro del bucle):

| Campo        | Valor |
|--------------|-------|
| Folder Path  | `/Flujos de Prueba/ExtraerPDFAcreedores/@{triggerBody()?['acreedor']}` |
| File Name    | `items('Apply_to_each')?['nombre']` |
| File Content | `base64ToBinary(items('Apply_to_each')?['base64'])` |

`base64ToBinary` es obligatorio. Sin él SharePoint guarda el texto base64
y el PDF no se abre.

**Respuesta** (fuera del bucle): Status 200 y cabecera
`Access-Control-Allow-Origin: *`

## Paso 3 — Conectar el panel con el flujo

En `facturas.js`, arriba del todo:

```js
const FLOW_URL = "…";                          // URL POST del trigger HTTP
const ACREEDORES = ["Adam", "SGS", "Sin clasificar"];
```

Los nombres de `ACREEDORES` deben coincidir exactamente con las carpetas
de SharePoint.

Vuelve a subir el archivo tras editarlo.

## Paso 4 — Cargar el add-in en Outlook

Outlook en la web / nuevo Outlook:

1. Abre un correo → **… (Más acciones)** → **Complementos**
2. **Mis complementos** → **Complementos personalizados** → **Agregar desde archivo**
3. Selecciona `manifest.xml` y acepta el aviso
4. Abre cualquier correo: verás el botón **Guardar factura** (grupo *Aelis Facturas*)

## Lista de acreedores dinámica (opcional)

Para no editar el JS cada vez que entre un acreedor nuevo, monta un segundo
flujo con trigger HTTP **GET** que lea las carpetas de
`/Flujos de Prueba/ExtraerPDFAcreedores` con *Get files (properties only)*
y devuelva un array de nombres:

```json
["Adam", "SGS", "Sin clasificar"]
```

Pega su URL en `ACREEDORES_URL`. El array `ACREEDORES` queda como respaldo
si el flujo no responde.

## Seguridad

La URL del flujo lleva su firma (`sig=`) incrustada en `facturas.js`. Si el
repo de GitHub es **público**, cualquiera que lo encuentre puede invocar el
flujo y escribir archivos en el SharePoint.

Mínimo recomendable: repo **privado** (Pages sigue funcionando en cuentas de
pago) o alojarlo en Azure Static Web Apps.
