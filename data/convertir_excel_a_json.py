import pandas as pd
import json
import re
from pathlib import Path
from datetime import datetime, date


# ============================================================
# RUTAS DEL PROYECTO
# ============================================================

# Este archivo está dentro de la carpeta /data.
# Por eso usamos parent.parent para llegar a la carpeta principal del proyecto.
BASE_DIR = Path(__file__).resolve().parent.parent

ARCHIVO_EXCEL = BASE_DIR / "Base_Datos_CAGI_Kaishan_Air_Cooled_Completo.xlsx"
CARPETA_SALIDA = BASE_DIR / "data"
ARCHIVO_SALIDA = CARPETA_SALIDA / "equipos.json"


# ============================================================
# FUNCIONES DE LIMPIEZA
# ============================================================

def normalizar_nombre_columna(texto):
    """
    Convierte nombres de columnas a formato simple.
    Ejemplo:
    'Presión F.L. CAGI (psig)' -> 'presion_f_l_cagi_psig'
    """
    texto = str(texto).strip().lower()

    reemplazos = {
        "á": "a",
        "é": "e",
        "í": "i",
        "ó": "o",
        "ú": "u",
        "ñ": "n"
    }

    for original, nuevo in reemplazos.items():
        texto = texto.replace(original, nuevo)

    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    texto = re.sub(r"_+", "_", texto)

    return texto.strip("_")


def limpiar_valor(valor):
    """
    Convierte valores de Excel a valores compatibles con JSON.
    """
    if pd.isna(valor):
        return None

    if isinstance(valor, str):
        valor = valor.strip()
        return valor if valor else None

    if isinstance(valor, (datetime, date)):
        return valor.isoformat()

    # Convierte valores tipo numpy a valores nativos de Python
    if hasattr(valor, "item"):
        valor = valor.item()

    if isinstance(valor, float):
        if valor.is_integer():
            return int(valor)
        return round(valor, 4)

    return valor


def obtener_valor(fila, *columnas):
    """
    Busca el primer valor disponible dentro de una lista de posibles columnas.
    """
    for columna in columnas:
        valor = fila.get(columna)
        if valor not in [None, ""]:
            return valor
    return None


def normalizar_modelo(texto):
    """
    Normaliza el modelo para crear un ID único.
    """
    if not texto:
        return None

    texto = str(texto).strip().upper()
    texto = re.sub(r"[^A-Z0-9]+", "-", texto)
    texto = re.sub(r"-+", "-", texto)

    return texto.strip("-")


def generar_id_equipo(fila):
    modelo = obtener_valor(
        fila,
        "modelo_cagi",
        "modelo_base",
        "modelo"
    )

    modelo_normalizado = normalizar_modelo(modelo)

    if not modelo_normalizado:
        return None

    return f"KAISHAN-{modelo_normalizado}"


# ============================================================
# FORMATO PARA COTIZACIÓN
# ============================================================

def generar_descripcion_comercial(fila):
    modelo = obtener_valor(fila, "modelo_cagi", "modelo_base", "modelo")
    fabricante = obtener_valor(fila, "fabricante") or "KAISHAN"
    tipo_compresor = obtener_valor(fila, "tipo_compresor") or "Screw"
    refrigeracion = obtener_valor(fila, "refrigeracion") or "Air-cooled"

    descripcion = (
        f"COMPRESOR DE TORNILLO ESTACIONARIO {fabricante}, modelo {modelo}, "
        f"equipo tipo {tipo_compresor}, refrigerado por aire. "
        f"Equipo diseñado para aplicaciones industriales de aire comprimido."
    )

    return descripcion


def generar_caracteristicas_cotizacion(fila):
    """
    Genera el texto que se copiará al campo 'Características técnicas'
    dentro de la app.
    """

    modelo_cagi = obtener_valor(fila, "modelo_cagi")
    modelo_base = obtener_valor(fila, "modelo_base")
    fabricante = obtener_valor(fila, "fabricante")
    tipo_velocidad = obtener_valor(fila, "tipo_velocidad_cagi", "tipo_velocidad")
    refrigeracion = obtener_valor(fila, "refrigeracion")
    tipo_compresor = obtener_valor(fila, "tipo_compresor")
    etapas = obtener_valor(fila, "etapas")
    lubricacion = obtener_valor(fila, "lubricacion")

    motor_hp = obtener_valor(
        fila,
        "motor_nominal_hp",
        "nominal_motor_hp",
        "motor_hp"
    )

    presion = obtener_valor(
        fila,
        "presion_f_l_cagi_psig",
        "presion_fl_cagi_psig",
        "full_load_operating_pressure_psig"
    )

    capacidad = obtener_valor(
        fila,
        "rated_capacity_at_full_load_operating_pressure_acfm",
        "capacidad_acfm",
        "rated_capacity_acfm"
    )

    potencia = obtener_valor(
        fila,
        "total_package_input_power_at_rated_capacity_kw",
        "total_package_input_power_kw",
        "potencia_full_load_kw"
    )

    url_cagi = obtener_valor(
        fila,
        "url_cagi",
        "link_cagi",
        "cagi"
    )

    lineas = []

    if modelo_cagi:
        lineas.append(f"Modelo CAGI: {modelo_cagi}")

    if modelo_base:
        lineas.append(f"Modelo base: {modelo_base}")

    if fabricante:
        lineas.append(f"Fabricante: {fabricante}")

    if tipo_velocidad:
        lineas.append(f"Tipo de velocidad: {tipo_velocidad}")

    if refrigeracion:
        lineas.append(f"Refrigeración: {refrigeracion}")

    if tipo_compresor:
        lineas.append(f"Tipo de compresor: {tipo_compresor}")

    if etapas:
        lineas.append(f"Etapas: {etapas}")

    if lubricacion:
        lineas.append(f"Lubricación: {lubricacion}")

    if motor_hp:
        lineas.append(f"Motor nominal: {motor_hp} HP")

    if presion:
        lineas.append(f"Presión F.L.: {presion} PSIG")

    if capacidad:
        lineas.append(f"Capacidad: {capacidad} ACFM")

    if potencia:
        lineas.append(f"Potencia a plena carga: {potencia} kW")

    if url_cagi:
        lineas.append(f"URL CAGI: {url_cagi}")

    return "\n".join(lineas)


def generar_texto_busqueda(fila):
    campos_busqueda = [
        "id_equipo",
        "familia",
        "modelo_base",
        "modelo_cagi",
        "fabricante",
        "tipo_velocidad_cagi",
        "refrigeracion"
    ]

    textos = []

    for campo in campos_busqueda:
        valor = fila.get(campo)
        if valor:
            textos.append(str(valor))

    return " ".join(textos).upper()


# ============================================================
# CONVERSIÓN PRINCIPAL
# ============================================================

def convertir_excel_a_json():
    if not ARCHIVO_EXCEL.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo Excel:\n{ARCHIVO_EXCEL}\n\n"
            "Verifica que el Excel esté en la carpeta principal del proyecto, "
            "al mismo nivel que index.html."
        )

    CARPETA_SALIDA.mkdir(exist_ok=True)

    hojas = pd.read_excel(
        ARCHIVO_EXCEL,
        sheet_name=None,
        engine="openpyxl"
    )

    equipos = []

    for nombre_hoja, df in hojas.items():
        df = df.dropna(how="all")

        if df.empty:
            continue

        df.columns = [normalizar_nombre_columna(col) for col in df.columns]

        for _, row in df.iterrows():
            fila = {}

            for columna in df.columns:
                fila[columna] = limpiar_valor(row[columna])

            # Evita filas completamente vacías
            if not any(fila.values()):
                continue

            modelo_cagi = obtener_valor(fila, "modelo_cagi")
            modelo_base = obtener_valor(fila, "modelo_base", "modelo")

            # Si no hay modelo, no sirve para el cotizador
            if not modelo_cagi and not modelo_base:
                continue

            fila["hoja_origen"] = nombre_hoja
            fila["id_equipo"] = generar_id_equipo(fila)
            fila["descripcion_comercial"] = generar_descripcion_comercial(fila)
            fila["caracteristicas_cotizacion"] = generar_caracteristicas_cotizacion(fila)
            fila["texto_busqueda"] = generar_texto_busqueda(fila)

            equipos.append(fila)

    equipos = sorted(
        equipos,
        key=lambda x: str(x.get("modelo_cagi") or x.get("modelo_base") or "")
    )

    with open(ARCHIVO_SALIDA, "w", encoding="utf-8") as archivo_json:
        json.dump(equipos, archivo_json, ensure_ascii=False, indent=2)

    print("============================================")
    print("CONVERSIÓN COMPLETADA")
    print("============================================")
    print(f"Archivo Excel leído: {ARCHIVO_EXCEL}")
    print(f"Equipos exportados: {len(equipos)}")
    print(f"Archivo generado: {ARCHIVO_SALIDA}")
    print("============================================")

    if equipos:
        print("Ejemplo del primer equipo exportado:")
        print(json.dumps(equipos[0], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    convertir_excel_a_json()