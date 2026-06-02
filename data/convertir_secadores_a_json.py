import pandas as pd
import json
import re
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent.parent

ARCHIVO_EXCEL = BASE_DIR / "Base_Datos_Secador.xlsx"
ARCHIVO_SALIDA = BASE_DIR / "data" / "secadores.json"


def normalizar_columna(texto):
    texto = str(texto).strip().lower()
    texto = texto.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
    texto = texto.replace("ñ", "n")
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    texto = re.sub(r"_+", "_", texto)
    return texto.strip("_")


def limpiar_valor(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, str):
        valor = valor.strip()
        return valor if valor else None

    if hasattr(valor, "item"):
        valor = valor.item()

    if isinstance(valor, float):
        if valor.is_integer():
            return int(valor)
        return round(valor, 4)

    return valor


def obtener(fila, *nombres):
    for nombre in nombres:
        valor = fila.get(nombre)
        if valor not in [None, ""]:
            return valor
    return None


def generar_modelo_secador(equipo):
    if equipo is None:
        return ""

    texto = str(equipo).strip()

    if texto.upper().startswith("MKE"):
        return texto.upper()

    if texto.endswith(".0"):
        texto = texto[:-2]

    return f"MKE {texto}"


def generar_id_secador(modelo):
    modelo_id = str(modelo).upper()
    modelo_id = re.sub(r"[^A-Z0-9]+", "-", modelo_id)
    modelo_id = re.sub(r"-+", "-", modelo_id).strip("-")
    return f"SECADOR-{modelo_id}"


def generar_descripcion(secador):
    modelo = secador["modelo"]

    return (
        f"SECADOR REFRIGERATIVO MIKROPOR, modelo {modelo}, "
        f"refrigerado por aire, diseñado para tratamiento de aire comprimido industrial. "
        f"Equipo apto para reducir la humedad del sistema y entregar aire seco bajo condiciones de operación especificadas."
    )


def generar_caracteristicas(secador):
    lineas = []

    def agregar(etiqueta, valor, unidad=""):
        if valor not in [None, ""]:
            if unidad:
                lineas.append(f"{etiqueta}: {valor} {unidad}")
            else:
                lineas.append(f"{etiqueta}: {valor}")

    agregar("Modelo", secador.get("modelo"))
    agregar("Caudal nominal", secador.get("caudal_nominal_m3h"), "m³/h")
    agregar("Tipo de enfriamiento", secador.get("tipo_enfriamiento"))
    agregar("Clase de calidad / punto de rocío", secador.get("clase_calidad"))
    agregar("Presión máxima de trabajo", secador.get("presion_maxima_barg"), "barg")
    agregar("Temperatura máxima de entrada", secador.get("temperatura_maxima_entrada_c"), "°C")
    agregar("Temperatura ambiente máxima", secador.get("temperatura_ambiente_maxima_c"), "°C")
    agregar("Caída de presión", secador.get("caida_presion_mbar"), "mbar")
    agregar("Refrigerante", secador.get("refrigerante"))
    agregar("Voltaje", secador.get("voltaje"))
    agregar("Potencia instalada", secador.get("potencia_instalada_kw"), "kW")

    largo = secador.get("largo_mm")
    ancho = secador.get("ancho_mm")
    alto = secador.get("alto_mm")

    if largo and ancho and alto:
        lineas.append(f"Dimensiones: {largo} x {ancho} x {alto} mm")

    agregar("Peso", secador.get("peso_kg"), "kg")
    agregar("Conexión de aire", secador.get("conexion_aire"))
    agregar("Conexión de drenaje", secador.get("conexion_drenaje"))

    return "\n".join(lineas)


def convertir_secadores():
    if not ARCHIVO_EXCEL.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {ARCHIVO_EXCEL}")

    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=0, engine="openpyxl")

    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")

    df.columns = [normalizar_columna(c) for c in df.columns]

    secadores = []

    for _, row in df.iterrows():
        fila = {}

        for columna in df.columns:
            fila[columna] = limpiar_valor(row[columna])

        equipo = obtener(fila, "equipo")
        modelo = generar_modelo_secador(equipo)

        if not modelo:
            continue

        secador = {
            "tipo_equipo": "secador",
            "familia": "MKE",
            "modelo": modelo,
            "modelo_base": modelo,
            "id_equipo": generar_id_secador(modelo),

            "caudal_nominal_m3h": obtener(
                fila,
                "max_rated_air_flow_35_c_inlet_7_bar_25_c_ambient_temperature_m_h",
                "equipo"
            ),

            "tipo_enfriamiento": obtener(fila, "type_of_cooling"),
            "flujo_aire_condensador_m3h": obtener(fila, "condenser_air_flow_m_h"),
            "capacidad_rechazo_calor_w": obtener(fila, "heat_rejection_capacity_45_c_w_max"),
            "refrigerante": obtener(fila, "refrigerant_type"),
            "cantidad_refrigerante_kg": obtener(fila, "refrigerant_quantitiy_kg"),
            "tipo_compresor_interno": obtener(fila, "compressor_type"),
            "tipo_control_capacidad": obtener(fila, "capacity_control_cycling_or_non_cycling"),
            "tipo_expansion": obtener(fila, "expansion_type"),
            "tipo_intercambiador": obtener(fila, "heat_of_exchanger_type"),
            "tipo_drenaje": obtener(fila, "drain_type"),
            "control_drenaje": obtener(fila, "drain_control_type"),
            "nivel_ruido_db": obtener(fila, "noise_level_db"),
            "filtro": obtener(fila, "filter_allocation_and_type"),

            "humedad_relativa_pct": obtener(fila, "relative_humidity"),
            "temperatura_ambiente_maxima_c": obtener(fila, "max_ambient_temp_c"),
            "temperatura_ambiente_minima_c": obtener(fila, "min_ambient_temp_c"),
            "temperatura_maxima_entrada_c": obtener(fila, "max_inlet_temp_c"),
            "presion_maxima_barg": obtener(fila, "max_working_pressure_barg"),

            "clase_calidad": obtener(fila, "humidity_and_liquid_water_class"),
            "caida_presion_mbar": obtener(fila, "pressure_drop_mbar"),
            "temperatura_salida_aire_c": obtener(fila, "at_35_c_inlet_air_temperature_at_dryer_outlet_c"),

            "voltaje": obtener(fila, "voltage_volt_phase_hz"),
            "potencia_instalada_kw": obtener(fila, "total_installed_power_kw"),
            "corriente_nominal_a": obtener(fila, "nominal_ampcity_a"),
            "mca_a": obtener(fila, "mca_a"),
            "lra_a": obtener(fila, "locked_rotor_amper_a"),
            "controlador": obtener(fila, "controller_type"),
            "proteccion_electrica": obtener(fila, "electrical_protection_class_according_iec"),
            "fusible_a": obtener(fila, "fuse_a"),

            "largo_mm": obtener(fila, "length_mm"),
            "ancho_mm": obtener(fila, "width_mm"),
            "alto_mm": obtener(fila, "height_mm"),
            "peso_kg": obtener(fila, "weight_kg"),
            "conexion_aire": obtener(fila, "connection_size"),
            "conexion_drenaje": obtener(fila, "drain_connection_size"),
        }

        secador["descripcion_comercial"] = generar_descripcion(secador)
        secador["caracteristicas_cotizacion"] = generar_caracteristicas(secador)

        secador["texto_busqueda"] = " ".join([
            secador.get("id_equipo", ""),
            secador.get("familia", ""),
            secador.get("modelo", ""),
            str(secador.get("caudal_nominal_m3h", "")),
            str(secador.get("voltaje", "")),
            str(secador.get("conexion_aire", "")),
        ]).upper()

        secadores.append(secador)

    secadores = sorted(secadores, key=lambda x: x.get("caudal_nominal_m3h") or 0)

    with open(ARCHIVO_SALIDA, "w", encoding="utf-8") as f:
        json.dump(secadores, f, ensure_ascii=False, indent=2)

    print("========================================")
    print("SECADORES CONVERTIDOS CORRECTAMENTE")
    print("========================================")
    print(f"Archivo generado: {ARCHIVO_SALIDA}")
    print(f"Secadores exportados: {len(secadores)}")
    print("========================================")

    if secadores:
        print(json.dumps(secadores[0], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    convertir_secadores()