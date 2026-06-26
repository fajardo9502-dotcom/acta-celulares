// ================================================================
// 1. CONFIGURACIÓN DE LA PIZARRA DE FIRMA
// ================================================================
const canvas = document.getElementById('pizarra');
const ctx = canvas.getContext('2d');
const btnLimpiar = document.getElementById('btnLimpiar');
const inputFirma = document.getElementById('Firma_Digital');
let dibujando = false;

// Función para empezar a dibujar
canvas.addEventListener('mousedown', () => dibujando = true);
canvas.addEventListener('mouseup', () => {
    dibujando = false;
    ctx.beginPath();
    // Guardamos el dibujo como texto base64 para enviarlo a Python
    inputFirma.value = canvas.toDataURL(); 
});

// El movimiento del mouse sobre el canvas
canvas.addEventListener('mousemove', (event) => {
    if (!dibujando) return;
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.strokeStyle = 'black';

    ctx.lineTo(event.offsetX, event.offsetY);
    ctx.stroke();
});

// Botón para borrar el canvas
btnLimpiar.addEventListener('click', () => {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    inputFirma.value = ""; 
});

// ================================================================
// 2. GENERADOR DE PDF (MUTACIÓN TEMPORAL DEL DOM)
// ================================================================
async function generarPDF() {
    const { jsPDF } = window.jspdf;
    const paginas = document.getElementById('formActa').querySelectorAll('.pagina');
    
    // Buscamos todos los inputs para convertirlos en texto plano para la captura del PDF
    const inputs = document.querySelectorAll('.pagina input');

    inputs.forEach(input => {
        // Si el input es el de la firma, lo saltamos para que no se altere la imagen
        if (input.id === 'Firma_Digital') {
            return; 
        }

        // El resto de los inputs se convierten en spans visibles
        const span = document.createElement('span');
        span.textContent = input.value;
        span.style.borderBottom = '1px solid black';
        span.style.minWidth = input.offsetWidth + 'px';
        span.style.display = 'inline-block';
        span.style.textAlign = 'center';
        input.parentNode.replaceChild(span, input);
        input._span = span;
        span._input = input;
    });

    // Ocultar botón limpiar temporalmente
    const btnLimpiar = document.getElementById('btnLimpiar');
    if (btnLimpiar) btnLimpiar.style.display = 'none';

    await new Promise(resolve => setTimeout(resolve, 1000));

    const pdf = new jsPDF('p', 'mm', 'a4');
    let primera = true;

    for (const pagina of paginas) {
        const canvasPDF = await html2canvas(pagina, { 
            scale: 2,
            useCORS: true,
            logging: false,
            backgroundColor: '#ffffff',
            ignoreElements: (element) => element.id === 'btnLimpiar'
        });
        
        const imgData = canvasPDF.toDataURL('image/png');
        const ancho = pdf.internal.pageSize.getWidth();
        const alto = (canvasPDF.height * ancho) / canvasPDF.width;

        if (!primera) pdf.addPage();
        pdf.addImage(imgData, 'PNG', 0, 0, ancho, alto);
        primera = false;
    }

    // Restaurar los inputs originales del DOM
    document.querySelectorAll('.pagina span').forEach(span => {
        if (span._input) {
            span.parentNode.replaceChild(span._input, span);
        }
    });

    if (btnLimpiar) btnLimpiar.style.display = 'inline-block';

    return pdf.output('datauristring');
}

// ================================================================
// 3. ENVÍO DE DATOS AL BACKEND (FastAPI)
// ================================================================
const enviarDatos = async (datos) => {
    const respuesta = await fetch('/api/acta', { 
        method: 'PUT',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(datos)
    });
    const resultado = await respuesta.json();
    console.log(resultado.mensaje);
};

// ================================================================
// 4. PROCESO DE CONTROL Y ENVÍO DEL FORMULARIO (SUBMIT)
// ================================================================
document.getElementById('formActa').addEventListener('submit', async (event) => {
    event.preventDefault(); // Detiene la recarga nativa inicial

    // Validación preliminar de seguridad
    if (!inputFirma.value) {
        alert('Por favor firme el documento antes de guardar.');
        return;
    }

    // ------------------------------------------------------------
    // PASO 1 : Captura FormData antes de mutar el DOM
    // ------------------------------------------------------------
    const campos = new FormData(document.getElementById('formActa'));

    // Funciones helper de formateo local
    const limpiarTextoMayuscula = (valor) => {
        if (valor) {
            return valor.toString().trim().toUpperCase();
        }
        return "";
    };

    const capitalizarTexto = (valor) => {
        if (valor) {
            return valor.toString().trim().toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
        }
        return "";
    };

    // ------------------------------------------------------------
    // PASO 2 : Estructura el objeto datos en memoria
    // ------------------------------------------------------------
    const datos = {
        Funcionario:   capitalizarTexto(campos.get('Funcionario')),
        Cedula:        campos.get('Cedula') ? campos.get('Cedula').toString().trim() : "",
        MODELO:        limpiarTextoMayuscula(campos.get('MODELO')),
        TIPOEQUIPO:    limpiarTextoMayuscula(campos.get('TIPOEQUIPO')), // Extraído limpiamente del <select>
        marca:         limpiarTextoMayuscula(campos.get('marca')),
        IMEI1:         limpiarTextoMayuscula(campos.get('IMEI1')),
        IMEI2:         limpiarTextoMayuscula(campos.get('IMEI2')),
        Telefono:      limpiarTextoMayuscula(campos.get('Telefono')),
        C_Costos:      limpiarTextoMayuscula(campos.get('C_Costos') || campos.get('Centro de Costos')),
        Supervisor:    capitalizarTexto(campos.get('Supervisor')),
        Zona_o_Cargo:  limpiarTextoMayuscula(campos.get('Zona o Cargo')),
        Codigo:        limpiarTextoMayuscula(campos.get('Código')),
        firma_digital: inputFirma.value, // Datos binarios de firma resguardados
        Novedades:     capitalizarTexto(campos.get('novedades')),
    };

    // ------------------------------------------------------------
    // PASO 3 : Genera el render del PDF y despacha
    // ------------------------------------------------------------
    try {
        console.log("Estructura de datos asegurada. Renderizando documento...");
        
        // Ejecución asíncrona de la captura de imágenes de html2canvas
        const pdfBase64 = await generarPDF();
        
        // Acoplamos el string codificado al paquete listo
        datos.pdfBase64 = pdfBase64;

        console.log("Enviando paquete definitivo al servidor de FastAPI...");
        await enviarDatos(datos);
        
        alert('Acta guardada correctamente.');
        
        // F5 nativo para limpiar inputs y reiniciar variables de estado globales
        location.reload();

    } catch (error) {
        console.error("Error crítico detectado en la canalización del submit:", error);
        alert("Hubo un error al guardar el acta. Por favor intente de nuevo.");
    }
});