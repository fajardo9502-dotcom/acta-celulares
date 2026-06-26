# Librerías estándar de Python
import base64
import os
import threading
from datetime import date
from typing import Optional

# Librerías externas
import pandas as pd  # Usado para el cruce automático de datos corporativos
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
EXCEL_PATH = r"C:\Users\1030650138\OneDrive - Colombiana de Comercio S.A\base actas de entrega.xlsx"
RUTA_BASE_ALIMENTADORA = r"C:\Users\1030650138\OneDrive - Colombiana de Comercio S.A\report_Planta_de_Personal_Completa__activos_e_inactivos_--colombiana--1010236537--_93bb93e0-54b6-46de-9477-06f6b4e6b9c3 (1).xlsx"
PDF_FOLDER = "PDFs"
os.makedirs(PDF_FOLDER, exist_ok=True)

# Lock para escritura thread-safe del Excel
excel_lock = threading.Lock()


# ================================================================
# UTILIDAD: Limpieza estricta de cédula
# ================================================================
def limpiar_cedula(valor) -> str:
    """
    Convierte cualquier valor de cédula a texto numérico estricto.
    Elimina espacios, el '.0' que mete Excel al leer números,
    y la notación científica que puede generar Pandas.
    Ejemplos:
        "1030650138.0"  -> "1030650138"
        " 1030650138 "  -> "1030650138"
        1030650138.0    -> "1030650138"
    """
    if valor is None:
        return ""
    texto = str(valor).strip()
    try:
        # Esto convierte "1030650138.0" o 1.03e+09 a entero limpio
        texto = str(int(float(texto)))
    except (ValueError, OverflowError):
        pass
    # Quitar puntos de miles y espacios residuales
    texto = texto.replace(".", "").replace(",", "").replace(" ", "")
    return texto


# ================================================================
# UTILIDADES DE TEXTO
# ================================================================
def limpiar_mayuscula(texto) -> str:
    if texto:
        return str(texto).strip().upper()
    return ""

def limpiar_titulo(texto) -> str:
    if texto:
        return str(texto).strip().title()
    return ""


# ================================================================
# MODELOS DE DATOS
# ================================================================
class DatosActa(BaseModel):
    Telefono: Optional[str] = ""
    IMEI1: Optional[str] = ""
    IMEI2: Optional[str] = ""
    MODELO: Optional[str] = ""
    marca: Optional[str] = ""
    C_Costos: Optional[str] = ""
    UN: Optional[str] = ""
    Supervisor: Optional[str] = ""
    Zona_o_Cargo: Optional[str] = ""
    Codigo: Optional[str] = ""
    Cedula: Optional[str] = ""
    Funcionario: Optional[str] = ""
    Bateria: Optional[str] = "No"
    Cargador: Optional[str] = ""
    TIPOEQUIPO: Optional[str] = ""
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


# ================================================================
# FUNCIÓN DE CRUCE POR CÉDULA (Base Alimentadora)
# ================================================================
def procesar_cruce_por_cedula(cedula_usuario: str) -> dict:
    """
    Busca al empleado en la base alimentadora por su cédula.
    Retorna un diccionario con cuatro claves:
        - "UN2":            parte corta del centro de costos (ej: "104")
        - "C_COSTOS_LARGO": parte larga del centro de costos (ej: "944058")
        - "UN":             nombre de la unidad organizativa
        - "Tipo_Logia":     Nombre del cargo extraído de la columna H (Base Alimentadora)

    IMPORTANTE: Si NO encuentra al empleado, las claves vienen
    como None para que la función de escritura active la salvaguarda de datos.
    """
    # Agregamos "Tipo_Logia" inicializado en None para la salvaguarda
    resultado = {"UN2": None, "C_COSTOS_LARGO": None, "UN": None, "Tipo_Logia": None}

    cedula_limpia = limpiar_cedula(cedula_usuario)
    if not cedula_limpia:
        print("   Cedula vacia, se omite el cruce.")
        return resultado

    if not os.path.exists(RUTA_BASE_ALIMENTADORA):
        print(f"   Base alimentadora no encontrada en: {RUTA_BASE_ALIMENTADORA}")
        return resultado

    try:
        df = pd.read_excel(RUTA_BASE_ALIMENTADORA, sheet_name="Planta de Personal_Completa (ac", dtype=str)
        df.columns = df.columns.str.strip()

        columna_cedula = 'Nombre de usuario'

        if columna_cedula not in df.columns:
            print(f"  No se encontro columna de cedula. Columnas disponibles: {list(df.columns)}")
            return resultado

        df["_cedula_limpia"] = df[columna_cedula].apply(limpiar_cedula)
        coincidencia = df[df["_cedula_limpia"] == cedula_limpia]

        if coincidencia.empty:
            print(f"   Cedula '{cedula_limpia}' NO encontrada en la base alimentadora.")
            return resultado  

        fila = coincidencia.iloc[0]

        # --- [CRUCE ORIGINAL] Centro de Costos ---
        columna_cc = "Centro de costos Código"
        if columna_cc in df.columns:
            valor_cc = str(fila.get(columna_cc, "")).strip()
            if valor_cc.lower() in ("nan", "none", ""):
                valor_cc = ""

            if "-" in valor_cc:
                un2_parte, largo_parte = valor_cc.split("-", 1)
                resultado["UN2"] = un2_parte.strip()
                resultado["C_COSTOS_LARGO"] = largo_parte.strip()
            elif valor_cc:
                resultado["C_COSTOS_LARGO"] = valor_cc
                resultado["UN2"] = ""
        else:
            print(f"   Columna '{columna_cc}' no encontrada en la base alimentadora.")

        # --- [CRUCE ORIGINAL] Nombre de la Unidad (UN) ---
        columna_un = "Unidad Estratégica de Negocio Nombre"
        if columna_un in df.columns:
            valor_un = str(fila.get(columna_un, "")).strip()
            resultado["UN"] = "" if valor_un.lower() in ("nan", "none") else valor_un.upper()
        else:
            resultado["UN"] = ""

        # --- [NUEVO CRUCE] Extraer columna H (Cargo) ---
        columna_cargo = "Código de cargo Nombre de cargo (1)"
        if columna_cargo in df.columns:
            valor_cargo = str(fila.get(columna_cargo, "")).strip()
            resultado["Tipo_Logia"] = "" if valor_cargo.lower() in ("nan", "none") else valor_cargo.upper()
        else:
            print(f"   Columna '{columna_cargo}' no encontrada en la base alimentadora. Se asigna vacio.")
            resultado["Tipo_Logia"] = ""

        print(f"   [Cruce OK] Cedula: {cedula_limpia} -> "
              f"UN2: {resultado['UN2']} | C.Costos: {resultado['C_COSTOS_LARGO']} | Cargo (AA): {resultado['Tipo_Logia']}")
        return resultado

    except Exception as e:
        print(f"   Error en procesar_cruce_por_cedula: {e}")
        return resultado

# ================================================================
# GENERADOR DE NOMBRE DE PDF
# ================================================================
def generar_nombre_pdf(datos: DatosActa) -> str:
    cedula = "".join(c for c in (datos.Cedula or "sin_cedula") if c.isalnum())
    fecha = date.today().strftime("%Y%m%d")
    return f"acta_{cedula}_{fecha}.pdf"


# ================================================================
# ACTUALIZACIÓN DEL EXCEL DE ACTAS (hoja principal "base")
# ================================================================
 
def actualizar_fila_excel_actas(datos: DatosActa, ruta_pdf: str):
    """
    Actualiza o crea una fila en el Excel de actas basada en la cédula del empleado.
    
    GARANTÍAS:
    1. Novedades SIEMPRE va a columna 19 (S)
    2. Tipo_Logia va a columna 27 (AA):
       - Si cruce exitoso Y cargo en base: usa base alimentadora
       - Si cruce falla O cargo vacío: usa datos.Tipo_Logia del formulario
    3. Salvaguarda activada: Si cédula no existe en planta, F y H NO se tocan
    """
    with excel_lock:
        if os.path.exists(EXCEL_PATH):
            wb = load_workbook(EXCEL_PATH)
        else:
            wb = Workbook()
            wb.active.title = "base"
 
        ws = wb.active
        cedula_a_buscar = limpiar_cedula(datos.Cedula)
        empleado_encontrado = False
 
        # === PASO 1: CRUCE CON BASE ALIMENTADORA ===
        info_unidades = procesar_cruce_por_cedula(cedula_a_buscar)
        cruce_exitoso = info_unidades["UN2"] is not None
 
        # === PASO 2: DETERMINAR TIPO_LOGIA FINAL ===
        # Prioridad:
        # 1. Si cruce exitoso Y base alimentadora tiene cargo → usar base alimentadora
        # 2. Si no → usar lo que viene del formulario web
        tipo_logia_final = ""
        if cruce_exitoso and info_unidades.get("Tipo_Logia"):
            tipo_logia_final = info_unidades["Tipo_Logia"]
            print(f"   [Tipo_Logia] Usando valor de base alimentadora: {tipo_logia_final}")
        elif datos.Tipo_Logia:
            tipo_logia_final = datos.Tipo_Logia
            print(f"   [Tipo_Logia] Usando valor del formulario web: {tipo_logia_final}")
        else:
            print(f"   [Tipo_Logia] Sin valor disponible (quedará vacío)")
 
        # === PASO 3: BUSCAR FILA EXISTENTE POR CÉDULA ===
        if ws.max_row >= 2 and cedula_a_buscar:
            for row in range(2, ws.max_row + 1):
                cedula_celda = limpiar_cedula(ws.cell(row=row, column=13).value)
 
                if cedula_celda == cedula_a_buscar:
                    # FILA ENCONTRADA: ACTUALIZAR EXISTENTE
                    fecha_hoy = date.today().strftime("%d/%m/%Y")
 
                    # Datos del formulario web (SIEMPRE se actualizan)
                    ws.cell(row=row, column=2,  value=limpiar_mayuscula(datos.Telefono))
                    ws.cell(row=row, column=3,  value=limpiar_mayuscula(datos.IMEI1))
                    ws.cell(row=row, column=4,  value=limpiar_mayuscula(datos.IMEI2))
                    ws.cell(row=row, column=5,  value=limpiar_mayuscula(f"{datos.marca} - {datos.MODELO}"))
                    ws.cell(row=row, column=7,  value=fecha_hoy)
                    ws.cell(row=row, column=9,  value="DIBOG")
                    ws.cell(row=row, column=10, value=limpiar_titulo(datos.Supervisor))
                    ws.cell(row=row, column=11, value=limpiar_mayuscula(datos.Zona_o_Cargo))
                    ws.cell(row=row, column=12, value=limpiar_mayuscula(datos.Codigo))
                    ws.cell(row=row, column=14, value=limpiar_titulo(datos.Funcionario))
                    ws.cell(row=row, column=15, value=limpiar_mayuscula(datos.Bateria))
                    ws.cell(row=row, column=16, value=limpiar_mayuscula(datos.Cargador))
                    ws.cell(row=row, column=18, value=limpiar_mayuscula(datos.TIPOEQUIPO))
                    
                    # NOVEDADES EN COLUMNA 19 (S) - GARANTIZADO
                    ws.cell(row=row, column=19, value=limpiar_mayuscula(datos.Novedades))
                    print(f"   [Columna S] Novedades guardadas: {datos.Novedades}")
                    
                    ws.cell(row=row, column=22, value=limpiar_mayuscula(datos.Tipo_Plan))
                    ws.cell(row=row, column=23, value=limpiar_mayuscula(datos.Costo_Plan))
                    ws.cell(row=row, column=24, value=limpiar_mayuscula(datos.Cuenta))
                    ws.cell(row=row, column=25, value=limpiar_mayuscula(datos.Nombre_Cuenta))
                    ws.cell(row=row, column=26, value=limpiar_mayuscula(datos.Tipo_Cargo))
                    
                    # TIPO_LOGIA EN COLUMNA 27 (AA) - GARANTIZADO
                    ws.cell(row=row, column=27, value=limpiar_mayuscula(tipo_logia_final))
                    print(f"   [Columna AA] Tipo_Logia guardado: {tipo_logia_final}")
                    
                    ws.cell(row=row, column=29, value=ruta_pdf)
 
                    # SALVAGUARDA: Solo tocar F y H si cruce fue exitoso
                    if cruce_exitoso:
                        ws.cell(row=row, column=6, value=info_unidades["C_COSTOS_LARGO"])  # F
                        ws.cell(row=row, column=8, value=info_unidades["UN2"])              # H
                        print(f"   [Cruce Exitoso] Columnas F, H actualizadas.")
                    else:
                        print(f"    [Salvaguarda] Cédula no en planta. Columnas F, H preservadas.")
 
                    print(f"   Fila {row} ACTUALIZADA para cédula: {cedula_a_buscar}")
                    empleado_encontrado = True
                    break
 
        # === PASO 4: CREAR FILA NUEVA SI NO EXISTE ===
        if not empleado_encontrado:
            numero_acta = ws.max_row + 1
            fecha_hoy = date.today().strftime("%d/%m/%Y")
 
            # Armar la lista nueva_fila con TODOS los datos en orden exacto
            nueva_fila = [
                numero_acta,                                                     # A  (1)
                limpiar_mayuscula(datos.Telefono),                               # B  (2)
                limpiar_mayuscula(datos.IMEI1),                                  # C  (3)
                limpiar_mayuscula(datos.IMEI2),                                  # D  (4)
                limpiar_mayuscula(f"{datos.marca} - {datos.MODELO}"),            # E  (5)
                info_unidades["C_COSTOS_LARGO"] if cruce_exitoso else "",        # F  (6)
                fecha_hoy,                                                       # G  (7)
                info_unidades["UN2"] if cruce_exitoso else "",                   # H  (8)
                "DIBOG",                                                         # I  (9)
                limpiar_titulo(datos.Supervisor),                                # J  (10)
                limpiar_mayuscula(datos.Zona_o_Cargo),                           # K  (11)
                limpiar_mayuscula(datos.Codigo),                                 # L  (12)
                cedula_a_buscar,                                                 # M  (13)
                limpiar_titulo(datos.Funcionario),                               # N  (14)
                limpiar_mayuscula(datos.Bateria),                                # O  (15)
                limpiar_mayuscula(datos.Cargador),                               # P  (16)
                "",                                                              # Q  (17)
                limpiar_mayuscula(datos.TIPOEQUIPO),                             # R  (18)
                limpiar_mayuscula(datos.Novedades),                              # S  (19) ← GARANTIZADO
                "",                                                              # T  (20)
                "",                                                              # U  (21)
                limpiar_mayuscula(datos.Tipo_Plan),                              # V  (22)
                limpiar_mayuscula(datos.Costo_Plan),                             # W  (23)
                limpiar_mayuscula(datos.Cuenta),                                 # X  (24)
                limpiar_mayuscula(datos.Nombre_Cuenta),                          # Y  (25)
                limpiar_mayuscula(datos.Tipo_Cargo),                             # Z  (26)
                limpiar_mayuscula(tipo_logia_final),                             # AA (27) ← GARANTIZADO
                "",                                                              # AB (28)
                ruta_pdf,                                                        # AC (29)
            ]
            
            ws.append(nueva_fila)
            print(f"    Fila NUEVA creada para cédula: {cedula_a_buscar}")
            print(f"      - Novedades (S): {datos.Novedades}")
            print(f"      - Tipo_Logia (AA): {tipo_logia_final}")
 
        # === PASO 5: GUARDAR ===
        wb.save(EXCEL_PATH)
        wb.close()
        print(f"    Excel guardado exitosamente")

# ================================================================
# REGISTRO DE DESCUENTOS (hoja "Descuentos")
# ================================================================
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
            ws.append(["Fecha", "Cedula", "Funcionario", "IMEI", "Modelo", "Valor Descuento", "Motivo"])

        fecha_hoy = date.today().strftime("%d/%m/%Y")
        ws.append([
            fecha_hoy,
            limpiar_cedula(datos.Cedula),
            limpiar_titulo(datos.Funcionario),
            datos.IMEI1,
            datos.MODELO,
            datos.ValorDescuento,
            datos.Novedades,
        ])
        wb.save(EXCEL_PATH)
        wb.close()


# ================================================================
# ENDPOINTS API
# ================================================================
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
        print(f"--> Peticion recibida en /api/descuento para: {datos.Funcionario}")
        agregar_fila_descuento(datos)
        return {"mensaje": "OK: Descuento registrado en el sistema"}
    except Exception as e:
        import traceback
        print(f"Error procesando el descuento: {e}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=str(e))


# ================================================================
# ARCHIVOS ESTÁTICOS
# ================================================================
app.mount("/", StaticFiles(directory="templates", html=True), name="static")


# ================================================================
# ARRANQUE
# ================================================================
if __name__ == "__main__":
    import uvicorn
    print("----------------------------------------------")
    print("SERVIDOR LISTO EN: http://localhost:8080")
    print("----------------------------------------------")
    uvicorn.run(app, host="0.0.0.0", port=8080)