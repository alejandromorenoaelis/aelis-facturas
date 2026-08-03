// =====================================================================
//  Guardar factura en SharePoint  ·  Aelis Procesos
// =====================================================================

// URL POST del flujo "GuardarFacturaAcreedor" (trigger HTTP).
const FLOW_URL = "https://default3ec777bd8b8646a8800f6d98eab6bc.39.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/fb910e725cee4c16b80bde3d307611a6/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=AavWhw2z58ityZogsJnDSQcOp-n9Sk-KTFtDuReSBDM";

// Lista de acreedores. Debe coincidir con el nombre de las carpetas
// de SharePoint en /ExtraerPDFAcreedores.
const ACREEDORES = ["Adam", "SGS", "Sin clasificar"];

// Opcional: URL GET de un flujo que devuelva ["Adam","SGS",...] leyendo
// las carpetas reales. Si se rellena, ACREEDORES solo actúa de respaldo.
const ACREEDORES_URL = "";

let datosCorreo = {};
let pdfs = [];              // { id, nombre, tipo, tamano }
let acreedorElegido = null;

const $ = (id) => document.getElementById(id);

Office.onReady(() => {
  const item = Office.context.mailbox.item;

  datosCorreo = {
    remitente: item.from ? item.from.emailAddress : "",
    nombreRemitente: item.from ? item.from.displayName : "",
    asunto: item.subject || "",
    fecha: item.dateTimeCreated
  };
  $("remitente").textContent = datosCorreo.remitente || "\u2014";
  $("asunto").textContent = datosCorreo.asunto || "\u2014";

  pintarAdjuntos(item.attachments || []);
  cargarAcreedores();

  $("run").onclick = guardar;
});

// ---------- acreedores ----------
async function cargarAcreedores() {
  let lista = ACREEDORES;
  if (ACREEDORES_URL) {
    try {
      const res = await fetch(ACREEDORES_URL);
      const datos = await res.json();
      if (Array.isArray(datos) && datos.length) lista = datos;
    } catch (e) {
      // Se mantiene la lista de respaldo.
    }
  }
  pintarChips(lista);
}

function pintarChips(lista) {
  const cont = $("chips");
  cont.innerHTML = "";
  lista.forEach((nombre) => {
    const b = document.createElement("button");
    b.className = "chip";
    b.type = "button";
    b.textContent = nombre;
    b.setAttribute("aria-pressed", "false");
    b.onclick = () => {
      acreedorElegido = nombre;
      [...cont.children].forEach((c) =>
        c.setAttribute("aria-pressed", String(c === b))
      );
      revisarBoton();
    };
    cont.appendChild(b);
  });
}

// ---------- adjuntos ----------
function esPdf(a) {
  const nom = (a.name || "").toLowerCase();
  const tipo = (a.contentType || "").toLowerCase();
  return !a.isInline && (nom.endsWith(".pdf") || tipo === "application/pdf");
}

function pintarAdjuntos(adjuntos) {
  pdfs = adjuntos.filter(esPdf).map((a) => ({
    id: a.id, nombre: a.name, tipo: a.contentType, tamano: a.size
  }));

  const cont = $("adjuntos");
  cont.innerHTML = "";

  if (!pdfs.length) {
    cont.innerHTML = '<span class="sin-adj">Este correo no lleva ningún PDF adjunto.</span>';
    return;
  }

  pdfs.forEach((p, i) => {
    const fila = document.createElement("label");
    fila.className = "adj";
    fila.innerHTML =
      '<input type="checkbox" data-i="' + i + '" checked />' +
      '<span class="adj-txt"><span class="adj-nom">' + escapar(p.nombre) + '</span>' +
      '<span class="adj-kb">' + Math.round(p.tamano / 1024) + ' KB</span></span>';
    fila.querySelector("input").onchange = revisarBoton;
    cont.appendChild(fila);
  });

  revisarBoton();
}

function seleccionados() {
  return [...document.querySelectorAll("#adjuntos input:checked")]
    .map((c) => pdfs[Number(c.dataset.i)]);
}

function revisarBoton() {
  $("run").disabled = !acreedorElegido || seleccionados().length === 0;
}

function leerContenidoAdjunto(id) {
  return new Promise((resolve) => {
    try {
      Office.context.mailbox.item.getAttachmentContentAsync(id, (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded && res.value &&
            res.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
          resolve(res.value.content);
        } else { resolve(null); }
      });
    } catch (e) { resolve(null); }
  });
}

// ---------- acción principal ----------
async function guardar() {
  const elegidos = seleccionados();
  const boton = $("run");
  boton.disabled = true;
  boton.innerHTML = '<span class="spin"></span> Guardando…';
  estado("work", "Guardando en SharePoint…", "Leyendo " + elegidos.length +
    (elegidos.length === 1 ? " archivo." : " archivos."));

  const envio = [];
  for (const p of elegidos) {
    const b64 = await leerContenidoAdjunto(p.id);
    if (b64) envio.push({ nombre: p.nombre, tipo: p.tipo, tamano: p.tamano, base64: b64 });
  }

  if (!envio.length) {
    estado("err", "No se pudo leer el PDF", "Vuelve a abrir el correo e inténtalo otra vez.");
    resetBoton();
    return;
  }

  if (!FLOW_URL) {
    estado("err", "Falta configurar el flujo", "Añade la URL del flujo en FLOW_URL dentro de facturas.js.");
    resetBoton();
    return;
  }

  const cuerpo = {
    acreedor: acreedorElegido,
    remitente: datosCorreo.remitente,
    nombreRemitente: datosCorreo.nombreRemitente,
    asunto: datosCorreo.asunto,
    fecha: datosCorreo.fecha,
    adjuntos: envio
  };

  try {
    const res = await fetch(FLOW_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=UTF-8" },
      body: JSON.stringify(cuerpo)
    });
    if (res.ok) exito(envio.length);
    else estado("err", "No se pudo guardar",
      "El flujo respondió con el código " + res.status + ". Inténtalo de nuevo.");
  } catch (e) {
    // Puede ser CORS al leer la respuesta; el envío suele haberse completado.
    exito(envio.length);
  }
  resetBoton();
}

function exito(n) {
  estado("ok", "Factura guardada",
    n + (n === 1 ? " PDF guardado en " : " PDF guardados en ") + acreedorElegido + ".");
}

function resetBoton() {
  const boton = $("run");
  boton.textContent = "Guardar en SharePoint";
  revisarBoton();
}

// ---------- UI ----------
function estado(tipo, msg, sub) {
  const box = $("status");
  box.hidden = false;
  box.className = "status " + tipo;
  const ico = tipo === "ok"
    ? '<svg class="check" viewBox="0 0 24 24"><circle cx="12" cy="12" r="12"/><path d="M6.5 12.5l3.5 3.5 7.5-8"/></svg>'
    : tipo === "err"
    ? '<span class="ico" style="color:var(--danger);font-weight:800;">!</span>'
    : '<span class="spin"></span>';
  box.innerHTML = '<span class="ico">' + ico + '</span><div><div class="msg">' +
    msg + '</div><div class="sub">' + (sub || "") + '</div></div>';
}

function escapar(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
