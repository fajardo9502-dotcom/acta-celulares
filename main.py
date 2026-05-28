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

        # Si el Excel tiene datos, buscamos desde la fila 2 (la 1 tiene los encabezados)
        if ws.max_row >= 2 and cedula_a_buscar:
            for row in range(2, ws.max_row + 1):
                # Columna 13 es la 'M' (Cédula) según tu mapeo original
                cedula_celda = str(ws.cell(row=row, column=13).value).strip() if ws.cell(row=row, column=13).value else ""
                
                if cedula_celda == cedula_a_buscar:
                    # ¡LO ENCONTRÓ! Reescribimos los datos técnicos en esta misma fila
                    ws.cell(row=row, column=2, value=datos.Telefono)       # B - Telefono
                    ws.cell(row=row, column=3, value=datos.IMEI1)          # C - IMEI1
                    ws.cell(row=row, column=4, value=datos.IMEI2)          # D - IMEI2
                    ws.cell(row=row, column=5, value=f"{datos.marca} - {datos.MODELO}") # E - Modelo
                    ws.cell(row=row, column=6, value=datos.C_Costos)       # F - C.Costos
                    ws.cell(row=row, column=10, value=datos.Supervisor)    # J - Supervisor
                    ws.cell(row=row, column=11, value=datos.Zona_o_Cargo)  # K - Zona o Cargo
                    ws.cell(row=row, column=12, value=datos.Codigo)        # L - Código
                    ws.cell(row=row, column=14, value=datos.Funcionario)   # N - Funcionario
                    ws.cell(row=row, column=15, value=datos.Bateria)       # O - Bateria
                    ws.cell(row=row, column=16, value=datos.Cargador)      # P - Cargador
                    ws.cell(row=row, column=17, value=datos.TipoEquipo)    # Q - TipoEquipo
                    ws.cell(row=row, column=19, value=datos.Novedades)     # S - Novedades
                    ws.cell(row=row, column=22, value=datos.Tipo_Plan)     # V - Tipo Plan
                    ws.cell(row=row, column=23, value=datos.Costo_Plan)    # W - Costo Plan
                    ws.cell(row=row, column=24, value=datos.Cuenta)        # X - Cuenta
                    ws.cell(row=row, column=25, value=datos.Nombre_Cuenta) # Y - Nombre Cuenta
                    ws.cell(row=row, column=26, value=datos.Tipo_Cargo)    # Z - Tipo Cargo
                    ws.cell(row=row, column=27, value=datos.Tipo_Logia)    # AA - Tipo Logia
                    ws.cell(row=row, column=29, value=ruta_pdf)            # AC - Ruta PDF (Columna 29)
                    
                    print(f"--> [PUT] Fila {row} ACTUALIZADA exitosamente para la cédula: {cedula_a_buscar}")
                    empleado_encontrado = True
                    break # Detener la búsqueda porque ya lo actualizamos

        # Caso alternativo: Si no se encuentra la cédula, se añade como nueva fila para que no falle
        if not empleado_encontrado:
            numero_acta = ws.max_row + 1
            fecha_hoy = date.today().strftime("%d/%m/%Y")
            nueva_fila = [
                numero_acta, datos.Telefono, datos.IMEI1, datos.IMEI2,
                f"{datos.marca} - {datos.MODELO}", datos.C_Costos, fecha_hoy,
                "", "", datos.Supervisor, datos.Zona_o_Cargo, datos.Codigo,
                datos.Cedula, datos.Funcionario, datos.Bateria, datos.Cargador,
                datos.TipoEquipo, "", datos.Novedades, "Activo", "",
                datos.Tipo_Plan, datos.Costo_Plan, datos.Cuenta, datos.Nombre_Cuenta,
                datos.Tipo_Cargo, datos.Tipo_Logia, "", ruta_pdf
            ]
            ws.append(nueva_fila)
            print(f"--> [PUT] Cédula no localizada. Se insertó registro nuevo al final para: {cedula_a_buscar}")

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
        agregar_fila_excel(datos, ruta_pdf)
        print(f"  Excel actualizado: {EXCEL_PATH}")

        return {"mensaje": "OK: PDF guardado y Excel actualizado"}

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
