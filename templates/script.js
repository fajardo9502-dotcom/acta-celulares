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
    const pdf = new jsPDF('p', 'mm', 'a4');
    let primera = true;

    for (const pagina of paginas) {
        const canvasPDF = await html2canvas(pagina, { scale: 2 });
        const imgData = canvasPDF.toDataURL('image/png');
        const ancho = pdf.internal.pageSize.getWidth();
        const alto = (canvasPDF.height * ancho) / canvasPDF.width;

        if (!primera) pdf.addPage();
        pdf.addImage(imgData, 'PNG', 0, 0, ancho, alto);
        primera = false;
    }

    return pdf.output('datauristring');
}

// Dentro de tu script.js, cuando vayas a enviar los datos:
const enviarDatos = async (datos) => {
    const respuesta = await fetch('/api/acta', { // <--- Esta es la ruta de tu main.py
        method: 'POST',
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

    // Arma el paquete de datos para enviar a Python
    const datos = {
        Funcionario:   campos.get('Funcionario')  || "",
        Cedula:        campos.get('Cedula')        || "",
        MODELO:        campos.get('MODELO')        || "",
        marca:         campos.get('marca')         || "",
        IMEI1:         campos.get('IMEI1')         || "",
        IMEI2:         campos.get('IMEI2')         || "",
        Telefono:      campos.get('Telefono')      || "",
        Supervisor:    campos.get('Supervisor')    || "",
        Zona_o_Cargo:  campos.get('Zona o Cargo')  || "",
        Codigo:        campos.get('Código')        || "",
        firma_digital: inputFirma.value,           // la firma del canvas
    };

   // Genera el PDF y lo agrega a los datos
    const pdfBase64 = await generarPDF();
    datos.pdfBase64 = pdfBase64;

    // Envía los datos a Python
    await enviarDatos(datos);
    alert('Acta guardada correctamente.');
    document.getElementById('formActa').reset();
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    inputFirma.value = "";
});