# Librerías estándar de Python
import base64
import os
import threading
from datetime import date
from typing import Optional

# Librerías externas
import pandas as pd  # <-- Añadido para el cruce automático
from openpyxl import Workbook, load_workbook
from pydantic import BaseModel

# FastAPI
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

app = FastAPI(title="Acta Celulares API")

# --- CORS ---
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "PUT", "OPTIONS"],
    allow_headers=["Content-Type"],
)

# --- CONFIGURACIÓN ---
EXCEL_PATH = r"C:\Users\POLLO\OneDrive - Colombiana de Comercio S.A\base actas de entrega.xlsx"

#  RUTA DE TU SEGUNDO EXCEL (Ajusta el nombre real del archivo si es diferente a 'Libro2.xlsx')
RUTA_BASE_ALIMENTADORA = r"C:\Users\POLLO\OneDrive - Colombiana de Comercio S.A\report_Planta_de_Personal_Completa__activos_e_inactivos_--colombiana--1010236537--_93bb93e0-54b6-46de-9477-06f6b4e6b9c3 (1).xlsx" 

PDF_FOLDER = "PDFs"
os.makedirs(PDF_FOLDER, exist_ok=True)

# Lock para escritura thread-safe del Excel
excel_lock = threading.Lock()


# --- MODELOS DE DATOS ---
class DatosActa(BaseModel):
    Telefono: Optional[str] = ""
    IMEI1: Optional[str] = ""
    IMEI2: Optional[str] = ""
    MODELO: Optional[str] = ""
    marca: Optional[str] = ""
    C_Costos: Optional[str] = ""        
    Supervisor: Optional[str] = ""
    Zona_o_Cargo: Optional[str] = ""    
    Codigo: Optional[str] = ""          
    Cedula: Optional[str] = ""
    Funcionario: Optional[str] = ""
    Bateria: Optional[str] = "No"
    Cargador: Optional[str] = ""
    TipoEquipo: Optional[str] = ""
    Novedades: Optional[str] = ""
    Tipo_Plan: Optional[str] = ""       
    Costo_Plan: Optional[str] = ""      
    Cuenta: Optional[str] = ""
    Nombre_Cuenta: Optional[str] = ""   
    Tipo_Cargo: Optional[str] = ""      
    Tipo_Logia: Optional[str] = ""      
    pdfBase64: Optional[str] = ""
    firma_digital: Optional[str] = ""


class DatosDescuento(BaseModel):
    Cedula: Optional[str] = ""
    Funcionario: Optional[str] = ""
    IMEI1: Optional[str] = ""
    MODELO: Optional[str] = ""
    ValorDescuento: Optional[str] = ""
    Novedades: Optional[str] = ""


# --- FUNCIÓN DE CRUCE DE DATOS (PANDAS) ---
def procesar_columna_u_matriz(valor_columna_u):
    """
    Toma el valor de la columna U (ej: '104-944058') de la matriz de personal,
    lo pica por el guion y devuelve las partes separadas junto con el nombre UN.
    """
    if not os.path.exists(RUTA_BASE_ALIMENTADORA):
        print(f" Alerta: No se encontró la base matriz alimentadora en {RUTA_BASE_ALIMENTADORA}")
        return None

    try:
        texto_u = str(valor_columna_u).strip()
        if '-' not in texto_u:
            return None

        # Separamos los 3 dígitos del número largo
        un2_parte, largo_parte = texto_u.split('-', 1)
        un2_extraido = un2_parte.strip()
        cc_largo_extraido = largo_parte.strip()

        # Ahora leemos el Excel de personal para buscar el nombre de la unidad (UN)
        df_matriz = pd.read_excel(RUTA_BASE_ALIMENTADORA, dtype=str)
        df_matriz.columns = df_matriz.columns.str.strip()
        
        columna_llave = 'Centro de costos Código'
        if columna_llave in df_matriz.columns:
            df_matriz[columna_llave] = df_matriz[columna_llave].str.strip()
            
            # Buscamos la fila exacta que tiene ese código completo en la columna U
            coincidencia = df_matriz[df_matriz[columna_llave] == texto_u]
            
            un_val = ""
            if not coincidencia.empty:
                fila = coincidencia.iloc[0]
                un_val = "" if pd.isna(fila.get('UN')) else str(fila.get('UN')).strip()
            
            print(f" [Cruce Exitoso] Código U: {texto_u} -> UN2: {un2_extraido} | F: {cc_largo_extraido} | UN: {un_val}")
            return {
                "UN2": un2_extraido,
                "C_COSTOS_LARGO": cc_largo_extraido,
                "UN": un_val
            }
            
    except Exception as e:
        print(f" Error en el procesamiento de romper columna U: {e}")
        
    return None

def generar_nombre_pdf(datos: DatosActa) -> str:
    cedula = "".join(c for c in (datos.Cedula or "sin_cedula") if c.isalnum())
    fecha = date.today().strftime("%Y%m%d")
    return f"acta_{cedula}_{fecha}.pdf"

def procesar_cruce_por_columna_u(valor_combinado_u):
    """
    Toma el valor combinado con guion (ej: '104-944058') que ya estaba en el Excel de Actas,
    lo pica en dos partes y busca en la matriz de personal el nombre de la unidad (UN).
    """
    if not os.path.exists(RUTA_BASE_ALIMENTADORA):
        print(f"⚠️ Alerta: No se encontró la base matriz alimentadora en {RUTA_BASE_ALIMENTADORA}")
        return {"UN2": "", "C_COSTOS_LARGO": "", "UN": ""}

    try:
        texto_u = str(valor_combinado_u).strip()
        if '-' not in texto_u:
            # Si por alguna razón no tiene guion, devolvemos vacío para no romper el programa
            return {"UN2": "", "C_COSTOS_LARGO": texto_u, "UN": ""}

        # Separamos los 3 dígitos del número largo usando el guion '-'
        un2_parte, largo_parte = texto_u.split('-', 1)
        un2_extraido = un2_parte.strip()
        cc_largo_extraido = largo_parte.strip()

        # Leemos la planta de personal para sacar la columna 'UN' (Nombre de unidad)
        df_matriz = pd.read_excel(RUTA_BASE_ALIMENTADORA, dtype=str)
        df_matriz.columns = df_matriz.columns.str.strip()
        
        columna_llave = 'Centro de costos Código'
        un_val = ""
        
        if columna_llave in df_matriz.columns:
            df_matriz[columna_llave] = df_matriz[columna_llave].str.strip()
            coincidencia = df_matriz[df_matriz[columna_llave] == texto_u]
            
            if not coincidencia.empty:
                fila = coincidencia.iloc[0]
                un_val = "" if pd.isna(fila.get('UN')) else str(fila.get('UN')).strip()

        print(f"🎯 [Pandas] Separación exitosa -> UN2: {un2_extraido} | C.Costos: {cc_largo_extraido} | UN: {un_val}")
        return {
            "UN2": un2_extraido,
            "C_COSTOS_LARGO": cc_largo_extraido,
            "UN": un_val
        }

    except Exception as e:
        print(f"🚨 Error procesando el split de la columna U: {e}")
        return {"UN2": "", "C_COSTOS_LARGO": "", "UN": ""}


def actualizar_fila_excel_actas(datos: DatosActa, ruta_pdf: str):
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
            if texto: return str(texto).strip().upper()
            return ""

        def limpiar_titulo(texto):
            if texto: return str(texto).strip().title()
            return ""

        # Si el Excel tiene datos, buscamos por cédula
        if ws.max_row >= 2 and cedula_a_buscar:
            for row in range(2, ws.max_row + 1):
                cedula_celda = str(ws.cell(row=row, column=13).value).strip() if ws.cell(row=row, column=13).value else ""
                
                if cedula_celda == cedula_a_buscar:
                    # 🔍 1. LEER lo que ya estaba guardado en la columna 6 (F) antes de modificar nada
                    valor_u_original = ws.cell(row=row, column=6).value
                    print(f"🔎 Valor original encontrado en Columna F (Fila {row}): '{valor_u_original}'")

                    # ⚙️ 2. PROCESAR el texto original para picar el guion y traer los datos correctos
                    info_unidades = procesar_cruce_por_columna_u(valor_u_original)

                    # 📝 3. ACTUALIZACIÓN de las columnas normales que vienen del formulario web
                    ws.cell(row=row, column=2, value=limpiar_mayuscula(datos.Telefono))       
                    ws.cell(row=row, column=3, value=limpiar_mayuscula(datos.IMEI1))          
                    ws.cell(row=row, column=4, value=limpiar_mayuscula(datos.IMEI2))          
                    ws.cell(row=row, column=5, value=limpiar_mayuscula(f"{datos.marca} - {datos.MODELO}")) 
                    
                    # 🔴 4. REESCRITURA INTELIGENTE de las columnas de costos basadas en el split
                    # Columna 6 (F) -> Reemplaza el '104-944058' dejando solo el número LARGO '944058'
                    ws.cell(row=row, column=6, value=info_unidades["C_COSTOS_LARGO"]) 
                    
                    # Columna 8 (H) -> Guarda los 3 dígitos cortos extraídos ('104')
                    ws.cell(row=row, column=8, value=info_unidades["UN2"])
                    
                    # Columna 9 (I) -> Guarda el nombre de la unidad (UN) traído con Pandas
                    ws.cell(row=row, column=9, value=limpiar_mayuscula(info_unidades["UN"]))
                    
                    # 📝 5. Guardar el resto de los datos del formulario
                    ws.cell(row=row, column=10, value=limpiar_titulo(datos.Supervisor))     
                    ws.cell(row=row, column=11, value=limpiar_mayuscula(datos.Zona_o_Cargo))   
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
                    ws.cell(row=row, column=29, value=ruta_pdf)            
                    
                    print(f"--> [PUT] Fila {row} PROCESADA Y DISTRIBUIDA por split para: {cedula_a_buscar}")
                    empleado_encontrado = True
                    break 

        # Caso fila Nueva (Si no existía previamente, no tiene código original, usa lo que envíe la web)
        if not empleado_encontrado:
            numero_acta = ws.max_row + 1
            fecha_hoy = date.today().strftime("%d/%m/%Y")
            
            # Como es nueva, intenta procesar lo que venga en datos.C_Costos por si acaso tiene guion
            info_unidades = procesar_cruce_por_columna_u(datos.C_Costos) if datos.C_Costos else {"UN2": "", "C_COSTOS_LARGO": "", "UN": ""}
            
            nueva_fila = [
                numero_acta, limpiar_mayuscula(datos.Telefono), limpiar_mayuscula(datos.IMEI1), limpiar_mayuscula(datos.IMEI2),
                limpiar_mayuscula(f"{datos.marca} - {datos.MODELO}"), info_unidades["C_COSTOS_LARGO"], fecha_hoy,
                info_unidades["UN2"],                   # Columna H (8)
                limpiar_mayuscula(info_unidades["UN"]),  # Columna I (9)
                limpiar_titulo(datos.Supervisor), limpiar_titulo(datos.Zona_o_Cargo), limpiar_mayuscula(datos.Codigo),
                cedula_a_buscar, limpiar_titulo(datos.Funcionario), limpiar_mayuscula(datos.Bateria), limpiar_mayuscula(datos.Cargador),
                limpiar_mayuscula(datos.TipoEquipo), "", limpiar_mayuscula(datos.Novedades), "ACTIVO", "",
                limpiar_mayuscula(datos.Tipo_Plan), limpiar_mayuscula(datos.Costo_Plan), limpiar_mayuscula(datos.Cuenta), limpiar_mayuscula(datos.Nombre_Cuenta),
                limpiar_mayuscula(datos.Tipo_Cargo), limpiar_mayuscula(datos.Tipo_Logia), "", ruta_pdf
            ]
            ws.append(nueva_fila)
            print(f"--> [PUT] Fila NUEVA creada para: {cedula_a_buscar}")

        wb.save(EXCEL_PATH)
        wb.close()


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


# --- ENDPOINTS API ---
@app.put("/api/acta")
async def recibir_acta(datos: DatosActa):
    try:
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

        actualizar_fila_excel_actas(datos, ruta_pdf)
        print(f"  Excel actualizado: {EXCEL_PATH}")
        return {"mensaje": "OK: Registro actualizado y PDF guardado"}

    except Exception as e:
        import traceback
        print(f"Error procesando acta: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


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
app.mount("/", StaticFiles(directory="templates", html=True), name="static")


# --- ARRANQUE ---
if __name__ == "__main__":
    import uvicorn
    print("----------------------------------------------")
    print("SERVIDOR LISTO EN: http://localhost:8080")
    print("----------------------------------------------")
    uvicorn.run(app, host="0.0.0.0", port=8080)