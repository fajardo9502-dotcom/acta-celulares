const form = document.getElementById("formActa");
const canvas = document.getElementById('pizarra');
const ctx = canvas.getContext('2d');
const inputFirma = document.getElementById('Firma_Digital');
const btnLimpiar = document.getElementById('btnLimpiar');
let dibujando = false;

// --- 1. CONFIGURACIÓN DE LA FIRMA ---
ctx.lineWidth = 2;
ctx.lineCap = 'round';
ctx.strokeStyle = '#000';

function obtenerCoordenadas(e) {
    const rect = canvas.getBoundingClientRect();
    const x = (e.clientX || e.touches[0].clientX) - rect.left;
    const y = (e.clientY || e.touches[0].clientY) - rect.top;
    return { x, y };
}

function empezarDibujo(e) {
    dibujando = true;
    const { x, y } = obtenerCoordenadas(e);
    ctx.beginPath();
    ctx.moveTo(x, y);
}

function dibujar(e) {
    if (!dibujando) return;
    e.preventDefault();
    const { x, y } = obtenerCoordenadas(e);
    ctx.lineTo(x, y);
    ctx.stroke();
}

function detenerDibujo() {
    if (dibujando) {
        dibujando = false;
        // Guardamos el dibujo en el input oculto para que Java lo reciba
        inputFirma.value = canvas.toDataURL();
    }
}

// Eventos de ratón y táctil
canvas.addEventListener('mousedown', empezarDibujo);
canvas.addEventListener('mousemove', dibujar);
window.addEventListener('mouseup', detenerDibujo);

canvas.addEventListener('touchstart', empezarDibujo);
canvas.addEventListener('touchmove', dibujar);
canvas.addEventListener('touchend', detenerDibujo);

btnLimpiar.addEventListener('click', () => {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    inputFirma.value = "";
});

// --- 2. ENVÍO DEL FORMULARIO ---
form.addEventListener("submit", async function(e) {
    e.preventDefault();

    try {
        // A. Generar PDF pagina por pagina
        const botonSubmit = form.querySelector("button[type='submit']");
        const botonLimpiar = document.getElementById("btnLimpiar");
        botonSubmit.style.display = "none";
        botonLimpiar.style.display = "none";

        const paginas = document.querySelectorAll(".pagina");
        const { jsPDF } = window.jspdf || {};

        console.log("Generando PDF...");

        // Capturar cada pagina como imagen y agregarla al PDF
        const pdf = new jsPDF({ unit: 'in', format: 'letter', orientation: 'portrait' });
        const margen = 0.5; // pulgadas
        const anchoUtil = 8.5 - (margen * 2);
        const altoUtil = 11 - (margen * 2);

        for (let i = 0; i < paginas.length; i++) {
            if (i > 0) pdf.addPage();

            const canvasPag = await html2canvas(paginas[i], {
                scale: 2,
                useCORS: true,
                backgroundColor: '#ffffff'
            });

            const imgData = canvasPag.toDataURL('image/jpeg', 0.98);
            const proporcion = canvasPag.height / canvasPag.width;
            let imgAncho = anchoUtil;
            let imgAlto = anchoUtil * proporcion;

            // Si la imagen es mas alta que la pagina, escalar para que quepa
            if (imgAlto > altoUtil) {
                imgAlto = altoUtil;
                imgAncho = altoUtil / proporcion;
            }

            pdf.addImage(imgData, 'JPEG', margen, margen, imgAncho, imgAlto);
        }

        pdf.save('acta_entrega.pdf');

        // Restaurar botones
        botonSubmit.style.display = "";
        botonLimpiar.style.display = "";

        // B. Preparar los datos para Java (incluir PDF en base64)
        const datos = new FormData(form);
        const objetoDatos = Object.fromEntries(datos.entries());

        // Obtener el PDF como base64 para enviarlo al servidor
        const pdfBlob = pdf.output('datauristring');
        objetoDatos.pdfBase64 = pdfBlob;

        // C. Enviar a la API de Java
        console.log("Enviando datos a Java...");
        
        const respuesta = await fetch("http://localhost:8080/api/acta", {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(objetoDatos)
        });

        if (!respuesta.ok) throw new Error("El servidor Java no respondió correctamente");

        const textoServidor = await respuesta.text();
        console.log("Respuesta Java:", textoServidor);

        // D. Éxito y Limpieza
        alert("¡Éxito! PDF guardado en carpeta PDFs y datos registrados en Excel ");
        form.reset();
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        inputFirma.value = "";

    } catch (error) {
        console.error("Error en el proceso:", error);
        alert("Hubo un error al procesar el acta ❌. Revisa la terminal de Java.");
    }
});