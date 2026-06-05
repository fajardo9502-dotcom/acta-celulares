from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional
import base64
import os
from datetime import date
from openpyxl import load_workbook, Workbook
import threading

app = FastAPI(title="Acta Celulares API")

# --- CORS ---
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "put","OPTIONS"],
    allow_headers=["Content-Type"],
)

# --- CONFIGURACIÓN ---
EXCEL_PATH = r"C:\Users\1030650138\OneDrive - Colombiana de Comercio S.A\base actas de entrega.xlsx"
PDF_FOLDER = "PDFs"
os.makedirs(PDF_FOLDER, exist_ok=True)

# Lock para escritura thread-safe del Excel (igual que synchronized en Java)
excel_lock = threading.Lock()


# --- MODELO DE DATOS (equivalente al parsearJson de Java) ---
class DatosActa(BaseModel):
    Telefono: Optional[str] = ""
    IMEI1: Optional[str] = ""
    IMEI2: Optional[str] = ""
    MODELO: Optional[str] = ""
    marca: Optional[str] = ""
    C_Costos: Optional[str] = ""        # "C. Costos" → C_Costos (FastAPI/Pydantic)
    Supervisor: Optional[str] = ""
    Zona_o_Cargo: Optional[str] = ""    # "Zona o Cargo" → Zona_o_Cargo
    Codigo: Optional[str] = ""          # "Código"
    Cedula: Optional[str] = ""
    Funcionario: Optional[str] = ""
    Bateria: Optional[str] = "No"
    Cargador: Optional[str] = ""
    TipoEquipo: Optional[str] = ""
    Novedades: Optional[str] = ""
    Tipo_Plan: Optional[str] = ""       # "Tipo Plan"
    Costo_Plan: Optional[str] = ""      # "Costo Plan"
    Cuenta: Optional[str] = ""
    Nombre_Cuenta: Optional[str] = ""   # "Nombre Cuenta"
    Tipo_Cargo: Optional[str] = ""      # "Tipo Cargo"
    Tipo_Logia: Optional[str] = ""      # "Tipo Logia"
    pdfBase64: Optional[str] = ""
    firma_digital: Optional[str] = ""


def generar_nombre_pdf(datos: DatosActa) -> str:
    """Equivalente a generarNombrePdf() en Java."""
    cedula = "".join(c for c in (datos.Cedula or "sin_cedula") if c.isalnum())
    fecha = date.today().strftime("%Y%m%d")
    return f"acta_{cedula}_{fecha}.pdf"


def actualizar_fila_excel_actas(datos: DatosActa, ruta_pdf: str):
    """
    [MÉTODO PUT] Recorre el Excel buscando la cédula del funcionario.
    Si la encuentra, actualiza la información técnica y la ruta del PDF en esa misma fila.
    Si NO la encuentra, la crea al final para evitar pérdida de datos.
    """
    with excel_lock:
        if os.path.exists(EXCEL_PATH):
            wb = load_workbook(EXCEL_PATH)
        else:
            wb = Workbook()
            wb.active.title = "base"

        ws = wb.active
        
        cedula_a_buscar = str(datos.Cedula).strip() if datos.Cedula else ""
        empleado_encontrado = False

        def limpiar_mayuscula(texto):
            if texto:
                return str(texto).strip().upper()
            return ""

        # --- FUNCIÓN AYUDANTE 2: INICIAL MAYÚSCULA (Tipo Título) ---
        def limpiar_titulo(texto):
            if texto:
                return str(texto).strip().title() # El método .title() de Python hace la magia
            return ""

        # Si el Excel tiene datos, buscamos desde la fila 2 (la 1 tiene los encabezados)
        if ws.max_row >= 2 and cedula_a_buscar:
            for row in range(2, ws.max_row + 1):
                # Columna 13 es la 'M' (Cédula) según tu mapeo original
                cedula_celda = str(ws.cell(row=row, column=13).value).strip() if ws.cell(row=row, column=13).value else ""
                
                if cedula_celda == cedula_a_buscar:
                    # Aplicamos las funciones según corresponda cada columna
                    ws.cell(row=row, column=2, value=limpiar_mayuscula(datos.Telefono))       
                    ws.cell(row=row, column=3, value=limpiar_mayuscula(datos.IMEI1))          
                    ws.cell(row=row, column=4, value=limpiar_mayuscula(datos.IMEI2))          
                    ws.cell(row=row, column=5, value=limpiar_mayuscula(f"{datos.marca} - {datos.MODELO}")) 
                    ws.cell(row=row, column=6, value=limpiar_mayuscula(datos.C_Costos))       
                    
                    # 👤 Aquí usamos limpiar_titulo para los nombres
                    ws.cell(row=row, column=10, value=limpiar_titulo(datos.Supervisor))     
                    ws.cell(row=row, column=12, value=limpiar_mayuscula(datos.Codigo))        
                    ws.cell(row=row, column=14, value=limpiar_titulo(datos.Funcionario))   
                    
                    ws.cell(row=row, column=15, value=limpiar_mayuscula(datos.Bateria))       
                    ws.cell(row=row, column=16, value=limpiar_mayuscula(datos.Cargador))      
                    ws.cell(row=row, column=17, value=limpiar_mayuscula(datos.TipoEquipo))    
                    ws.cell(row=row, column=19, value=limpiar_mayuscula(datos.Novedades))     
                    ws.cell(row=row, column=22, value=limpiar_mayuscula(datos.Tipo_Plan))     
                    ws.cell(row=row, column=23, value=limpiar_mayuscula(datos.Costo_Plan))    
                    ws.cell(row=row, column=24, value=limpiar_mayuscula(datos.Cuenta))        
                    ws.cell(row=row, column=25, value=limpiar_mayuscula(datos.Nombre_Cuenta)) 
                    ws.cell(row=row, column=26, value=limpiar_mayuscula(datos.Tipo_Cargo))    
                    ws.cell(row=row, column=27, value=limpiar_mayuscula(datos.Tipo_Logia))  
                    ws.cell(row=row, column=11, value=limpiar_mayuscula(datos.Zona_o_Cargo))   
                    ws.cell(row=row, column=29, value=ruta_pdf)            
                    
                    print(f"--> [PUT] Fila {row} ACTUALIZADA (Mixto) para: {cedula_a_buscar}")
                    empleado_encontrado = True
                    break 

        # Caso alternativo (Fila Nueva): Hacemos lo mismo al armar la lista
        if not empleado_encontrado:
            numero_acta = ws.max_row + 1
            fecha_hoy = date.today().strftime("%d/%m/%Y")
            nueva_fila = [
                numero_acta, limpiar_mayuscula(datos.Telefono), limpiar_mayuscula(datos.IMEI1), limpiar_mayuscula(datos.IMEI2),
                limpiar_mayuscula(f"{datos.marca} - {datos.MODELO}"), limpiar_mayuscula(datos.C_Costos), fecha_hoy,
                "", "", limpiar_titulo(datos.Supervisor), limpiar_titulo(datos.Zona_o_Cargo), limpiar_mayuscula(datos.Codigo),
                cedula_a_buscar, limpiar_titulo(datos.Funcionario), limpiar_mayuscula(datos.Bateria), limpiar_mayuscula(datos.Cargador),
                limpiar_mayuscula(datos.TipoEquipo), "", limpiar_mayuscula(datos.Novedades), "ACTIVO", "",
                limpiar_mayuscula(datos.Tipo_Plan), limpiar_mayuscula(datos.Costo_Plan), limpiar_mayuscula(datos.Cuenta), limpiar_mayuscula(datos.Nombre_Cuenta),
                limpiar_mayuscula(datos.Tipo_Cargo), limpiar_mayuscula(datos.Tipo_Logia), "", ruta_pdf
            ]
            ws.append(nueva_fila)
            print(f"--> [PUT] Fila NUEVA creada (Mixto) para: {cedula_a_buscar}")

        wb.save(EXCEL_PATH)
        wb.close()

# --- ENDPOINT PRINCIPAL (equivalente a RecibirActa en Java) ---
@app.put("/api/acta")
async def recibir_acta(datos: DatosActa):
    try:
        # 1. Guardar PDF
        nombre_pdf = generar_nombre_pdf(datos)
        ruta_pdf = ""

        if datos.pdfBase64:
            pdf_b64 = datos.pdfBase64
            if "," in pdf_b64:
                pdf_b64 = pdf_b64.split(",", 1)[1]
            pdf_bytes = base64.b64decode(pdf_b64)
            ruta_pdf = os.path.join(PDF_FOLDER, nombre_pdf)
            with open(ruta_pdf, "wb") as f:
                f.write(pdf_bytes)
            print(f"  PDF guardado: {ruta_pdf}")

        # 2. Agregar fila al Excel
        actualizar_fila_excel_actas(datos, ruta_pdf)
        print(f"  Excel actualizado: {EXCEL_PATH}")

        return {"mensaje": "OK: Registro actualizado y PDF guardado"}

    except Exception as e:
        import traceback
        print(f"Error procesando acta: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


# --- ARCHIVOS ESTÁTICOS (equivalente a ArchivoEstatico en Java) ---
# Sirve la carpeta actual como archivos estáticos (HTML, CSS, JS, imágenes)
app.mount("/", StaticFiles(directory="templates", html=True), name="static")




# =====================================================================
#  PASO 1: AQUÍ PEGAS EL NUEVO MODELO DE DATOS DE DESCUENTOS
# =====================================================================
class DatosDescuento(BaseModel):
    Cedula: Optional[str] = ""
    Funcionario: Optional[str] = ""
    IMEI1: Optional[str] = ""
    MODELO: Optional[str] = ""
    ValorDescuento: Optional[str] = ""
    Novedades: Optional[str] = ""


# =====================================================================
#  PASO 2: AQUÍ PEGAS LA FUNCIÓN QUE ESCRIBE EN EL EXCEL
# =====================================================================
def agregar_fila_descuento(datos: DatosDescuento):
    with excel_lock:
        if os.path.exists(EXCEL_PATH):
            wb = load_workbook(EXCEL_PATH)
        else:
            wb = Workbook()
            wb.active.title = "base"
        
        if "Descuentos" in wb.sheetnames:
            ws = wb["Descuentos"]
        else:
            ws = wb.create_sheet(title="Descuentos")
            ws.append(["Fecha", "Cédula", "Funcionario", "IMEI", "Modelo", "Valor Descuento", "Motivo"])

        fecha_hoy = date.today().strftime("%d/%m/%Y")
        nueva_fila = [
            fecha_hoy, datos.Cedula, datos.Funcionario, 
            datos.IMEI1, datos.MODELO, datos.ValorDescuento, datos.Novedades
        ]
        ws.append(nueva_fila)
        wb.save(EXCEL_PATH)
        wb.close()


# =====================================================================
#  PASO 3: AQUÍ PEGAS EL ENDPOINT DE DESCUENTOS (El Recepcionista)
# =====================================================================
@app.put("/api/descuento")
async def recibir_descuento(datos: DatosDescuento):
    try:
        print(f"--> Petición recibida en /api/descuento para: {datos.Funcionario}")
        agregar_fila_descuento(datos)
        return {"mensaje": "OK: Descuento registrado en el sistema"}
    except Exception as e:
        import traceback
        print(f"Error procesando el descuento: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


# --- ARCHIVOS ESTÁTICOS ---
# (Esto ya lo tenías, déjalo justo debajo del nuevo endpoint)
app.mount("/", StaticFiles(directory="templates", html=True), name="static")

# --- ARRANQUE DEL SERVIDOR ---
if __name__ == "__main__":
    import uvicorn
    print("----------------------------------------------")
    print("SERVIDOR LISTO EN: http://localhost:8080")
    print("----------------------------------------------")
    uvicorn.run(app, host="0.0.0.0", port=8080)
