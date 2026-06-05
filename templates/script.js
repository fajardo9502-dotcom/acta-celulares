// 1. Buscamos la pizarra (canvas) y el botón de borrar
const canvas = document.getElementById('pizarra');
const ctx = canvas.getContext('2d');
const btnLimpiar = document.getElementById('btnLimpiar');
const inputFirma = document.getElementById('Firma_Digital');
let dibujando = false;

// 2. Función para empezar a dibujar
canvas.addEventListener('mousedown', () => dibujando = true);
canvas.addEventListener('mouseup', () => {
    dibujando = false;
    ctx.beginPath();
    // ¡Ojo! Aquí guardamos el dibujo como texto secreto para Python
    inputFirma.value = canvas.toDataURL(); 
});

// 3. El movimiento del mouse
canvas.addEventListener('mousemove', (event) => {
    if (!dibujando) return;
    ctx.lineWidth = 2;
    ctx.lineCap = 'round';
    ctx.strokeStyle = 'black';

    // Dibujamos donde esté el mouse
    ctx.lineTo(event.offsetX, event.offsetY);
    ctx.stroke();
});

// 4. Botón para borrar si te queda fea la firma
btnLimpiar.addEventListener('click', () => {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    inputFirma.value = ""; // Borramos el dato secreto también
});
async function generarPDF() {
    const { jsPDF } = window.jspdf;
    const paginas = document.getElementById('formActa').querySelectorAll('.pagina');
    
    // Reemplazar inputs por spans con el valor
   // 1. Buscamos todos los inputs para convertirlos en texto para el PDF
const inputs = document.querySelectorAll('.pagina input');

inputs.forEach(input => {
    // ¡ESTA ES LA CLAVE! 
    // Si el input es el de la firma, lo saltamos para que no se convierta en texto visible
    if (input.id === 'Firma_Digital') {
        return; 
    }

    // El resto de los inputs (Cédula, Nombre, etc.) sí se convierten en span
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

    // Ocultar botón limpiar
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

    // Restaurar inputs
    document.querySelectorAll('.pagina span').forEach(span => {
        if (span._input) {
            span.parentNode.replaceChild(span._input, span);
        }
    });

    if (btnLimpiar) btnLimpiar.style.display = 'inline-block';

    return pdf.output('datauristring');
}

// Dentro de tu script.js, cuando vayas a enviar los datos:
const enviarDatos = async (datos) => {
    const respuesta = await fetch('/api/acta', { // <--- Esta es la ruta de tu main.py
        method: 'PUT',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(datos)
    });
    const resultado = await respuesta.json();
    console.log(resultado.mensaje);
};


// 5. Cuando el usuario hace clic en "Guardar Información"
document.getElementById('formActa').addEventListener('submit', async (event) => {
    event.preventDefault(); // Evita que la página se recargue

    // Verifica que haya firma
    if (!inputFirma.value) {
        alert('Por favor firme el documento antes de guardar.');
        return;
    }

    // Recoge todos los campos del formulario
    const campos = new FormData(document.getElementById('formActa'));

    const limpiarTextoMayuscula = (valor) => {
        if (valor) {
            return valor.toString().trim().toUpperCase();
        }
        return "";
    };

    // --- FUNCIÓN 2: Pasa a formato "Inicial Mayúscula y resto minúsculas" (Tipo Título) ---
    const capitalizarTexto = (valor) => {
        if (valor) {
            return valor.toString().trim().toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
        }
        return "";
    };

    // Arma el paquete de datos para enviar a Python
    const datos = {
        Funcionario:   capitalizarTexto(campos.get('Funcionario')),
        Cedula:        campos.get('Cedula') ? campos.get('Cedula').toString().trim() : "",
        MODELO:        limpiarTextoMayuscula(campos.get('MODELO')),
        marca:         limpiarTextoMayuscula(campos.get('marca')),
        IMEI1:         limpiarTextoMayuscula(campos.get('IMEI1')),
        IMEI2:         limpiarTextoMayuscula(campos.get('IMEI2')),
        Telefono:      limpiarTextoMayuscula(campos.get('Telefono')),
        Supervisor:    capitalizarTexto(campos.get('Supervisor')),
        Zona_o_Cargo:  limpiarTextoMayuscula(campos.get('Zona o Cargo')),
        Codigo:        limpiarTextoMayuscula(campos.get('Código')),
        firma_digital: inputFirma.value,           // Se mantiene intacta
        Novedades:  capitalizarTexto(campos.get('novedades')),
    };

  try {
        console.log("Generando PDF y enviando datos...");
        
        // 1. Genera el PDF (oculta y muestra el botón limpiar internamente)
        const pdfBase64 = await generarPDF();
        datos.pdfBase64 = pdfBase64;

        // 2. Envía los datos a la API de Python
        await enviarDatos(datos);
        
        // 3. Alerta de éxito al usuario
        alert('Acta guardada correctamente.');

        // 4. Recarga la página por completo (F5 automático)
        // Esto limpia el formulario, borra el canvas y revive el botón de limpiar de forma nativa
        location.reload();

    } catch (error) {
        console.error("Error crítico en el proceso:", error);
        alert("Hubo un error al guardar el acta. Por favor intente de nuevo.");
    }
});
