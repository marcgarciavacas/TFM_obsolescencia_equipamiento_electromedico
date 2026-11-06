"""Lectura, validación y cálculo de los resultados del modelo."""


# 1. LIBRERÍAS


from pathlib import Path
from io import BytesIO
from decimal import Decimal, ROUND_HALF_UP
from numbers import Real
import math
import re
import pandas as pd
from openpyxl import load_workbook


# 2. HOJAS Y CRITERIOS DEL MODELO


HOJAS = {
    "equipos": "1. Equipos", "mantenimiento": "2. Mantenimiento",
    "evaluacion": "3. Evaluacion", "resultados": "4. Resultados",
    "vida_util": "B. Vida_util",
}


TECNICOS = ["antiguedad", "estado_funcional", "intervenciones_correctivas",
            "soporte_tecnico", "disponibilidad_recambios",
            "obsolescencia_tecnologica_software"]


CRITICIDAD = ["impacto_actividad", "disponibilidad_alternativas",
              "impacto_seguridad_continuidad"]


# 3. LECTURA Y VINCULACIÓN DE LOS DATOS


def vacio(valor):
    return valor is None or (isinstance(valor, str) and not valor.strip()) or pd.isna(valor)


def identificador(valor):
    if vacio(valor):
        return ""
    if isinstance(valor, (int, float)) and float(valor).is_integer():
        return str(int(valor))
    # Los identificadores que ya son texto conservan sus ceros iniciales.
    return str(valor).strip()


def fecha(valor):
    return None if vacio(valor) else pd.Timestamp(valor).normalize()


def leer_tabla(hoja):
    """Convierte los encabezados y registros de la hoja en una tabla de pandas."""
    filas = list(hoja.values)
    columnas = [c for c in filas[0] if c is not None]
    return pd.DataFrame([r[:len(columnas)] for r in filas[1:]
                         if any(v is not None for v in r[:len(columnas)])], columns=columnas)


def leer_cortes(formula):
    """Recupera los límites de clasificación escritos en las fórmulas de Excel."""
    cortes = sorted({float(v) for v in re.findall(r"<(\d+(?:\.\d+)?)", str(formula))})
    if len(cortes) != 3:
        raise ValueError("No se reconocen los tres umbrales en las fórmulas de niveles.")
    return cortes


def leer_excel(origen):
    """Lee las tablas, los parámetros y los resultados guardados en el libro."""
    contenido = origen if isinstance(origen, bytes) else Path(origen).read_bytes()
    # Se leen por separado los valores guardados y las fórmulas. openpyxl no recalcula el libro.
    libro = load_workbook(BytesIO(contenido), read_only=True, data_only=True)
    formulas = load_workbook(BytesIO(contenido), read_only=True, data_only=False)
    try:
        faltan = set(HOJAS.values()) | {"A. Parametros_modelo"}
        faltan -= set(libro.sheetnames)
        if faltan:
            raise ValueError("Faltan hojas: " + ", ".join(sorted(faltan)))
        datos = {clave: leer_tabla(libro[hoja]) for clave, hoja in HOJAS.items()}
        desconocidos = {"desconocida", "desconocido", "no consta", "no disponible", "nd"}
        # Se normalizan los códigos y las fechas antes de relacionar las tablas.
        for clave, tabla in datos.items():
            for columna in tabla.columns:
                if columna in {"codigo_activo", "numero_serie", "numero_serie_equipo"}:
                    tabla[columna] = tabla[columna].map(identificador)
                if columna.startswith("fecha_"):
                    original = tabla[columna]
                    limpia = original.map(lambda v: None if vacio(v) or
                                          str(v).strip().lower() in desconocidos else v)
                    convertida = pd.to_datetime(limpia, errors="coerce", format="mixed", dayfirst=True)
                    mal = limpia.notna() & convertida.isna()
                    if mal.any():
                        raise ValueError(f"Fecha no reconocida en {HOJAS[clave]} / {columna}.")
                    tabla[columna] = convertida.dt.normalize()

        parametros = list(libro["A. Parametros_modelo"].values)
        config = {}
        for fila in parametros:
            if fila[0] in {"fecha_referencia", "inicio_mantenimiento", "fin_mantenimiento",
                           "ventana_correctivos_recientes", "umbral_correctivos_1", "umbral_correctivos_2",
                           "completitud_tecnica_provisional", "cobertura_mantenimiento"}:
                config[fila[0]] = fila[1]
        for c in ["fecha_referencia", "inicio_mantenimiento", "fin_mantenimiento"]:
            config[c] = fecha(config.get(c))
        config["pesos_tecnicos"] = {r[0]: r[1] for r in parametros if r[0] in TECNICOS}
        config["pesos_criticidad"] = {r[5]: r[6] for r in parametros if r[5] in CRITICIDAD}
        resultado = formulas["4. Resultados"]
        # Los cortes se recuperan de las fórmulas de clasificación de la hoja de resultados.
        config["cortes"] = leer_cortes(resultado["Z2"].value)
        if config["cortes"] != leer_cortes(resultado["AF2"].value):
            raise ValueError("Los umbrales técnicos y de criticidad son distintos: revisar el modelo.")
        matriz = libro["A. Parametros_modelo"]
        config["matriz"] = {
            matriz.cell(f, 6).value: {matriz.cell(24, c).value: matriz.cell(f, c).value
                                    for c in range(7, 11)} for f in range(25, 29)
        }
        datos["config"] = config
        return datos
    finally:
        libro.close()
        formulas.close()


def unir_datos(datos):
    """Relaciona inventario, evaluación y resultados mediante el código de activo."""
    base = datos["resultados"].copy()
    for nombre in ["equipos", "evaluacion"]:
        tabla = datos[nombre]
        nuevas = [c for c in tabla if c not in base or c == "codigo_activo"]
        base = base.merge(tabla[nuevas], on="codigo_activo", how="left", validate="one_to_one")
    return base


# 4. VALIDACIÓN DE LOS DATOS


def validar_datos(datos):
    """Comprueba la estructura y los parámetros; devuelve los avisos de revisión."""
    # Los errores impiden continuar; las diferencias de orden o de serie se devuelven como avisos.
    avisos = []
    requeridas = {
        "equipos": ["codigo_activo", "descripcion_equipo", "familia_tecnologica", "centro", "servicio",
                    "fecha_instalacion", "fecha_adquisicion", "fecha_baja", "fecha_alta_inventario", "numero_serie"],
        "evaluacion": ["codigo_activo", "fecha_valoracion", "unidad_revisora"] + TECNICOS[1:2] + TECNICOS[3:] + CRITICIDAD,
        "mantenimiento": ["id_parte", "codigo_activo", "fecha_apertura", "fecha_cierre",
                          "tipo_intervencion_normalizado", "estado_parte_normalizado", "archivo_origen", "numero_serie_equipo"],
        "resultados": ["codigo_activo", "nivel_tecnico", "nivel_criticidad", "prioridad_orientativa", "estado_resultado"],
        "vida_util": ["tipo_equipo", "limite_inferior_años", "limite_superior_años"],
    }
    for nombre, columnas in requeridas.items():
        faltan = set(columnas) - set(datos[nombre].columns)
        if faltan:
            raise ValueError(f"Faltan columnas en {nombre}: {', '.join(sorted(faltan))}")
    codigos = set(datos["equipos"]["codigo_activo"])
    for nombre in ["equipos", "evaluacion", "resultados"]:
        claves = datos[nombre]["codigo_activo"]
        if claves.empty or claves.eq("").any() or claves.duplicated().any():
            raise ValueError(f"Códigos vacíos o duplicados en {nombre}.")
        if set(claves) != codigos:
            raise ValueError(f"Los códigos de {nombre} no coinciden con los del inventario.")
    partes = datos["mantenimiento"]
    if not set(partes["codigo_activo"]).issubset(codigos):
        raise ValueError("Hay partes vinculados a códigos que no existen en el inventario.")
    cerrados = partes["estado_parte_normalizado"].eq("Cerrado")
    if partes.loc[cerrados, ["fecha_apertura", "fecha_cierre"]].isna().any().any():
        raise ValueError("Hay partes cerrados sin fechas válidas de apertura o cierre.")
    if (partes["fecha_cierre"] < partes["fecha_apertura"]).any():
        raise ValueError("Hay partes con cierre anterior a su apertura.")
    if partes.duplicated(["id_parte", "archivo_origen"]).any():
        raise ValueError("Hay partes duplicados (identificador y archivo de origen). Revisar antes de contar.")
    vida = datos["vida_util"]
    if vida["tipo_equipo"].str.strip().str.casefold().duplicated().any():
        raise ValueError("Hay tipos de equipo duplicados en la tabla de vida útil.")
    for _, r in vida.iterrows():
        if not vacio(r["limite_inferior_años"]) and not vacio(r["limite_superior_años"]):
            if not 0 < r["limite_inferior_años"] <= r["limite_superior_años"]:
                raise ValueError("Intervalo de vida útil no válido: " + str(r["tipo_equipo"]))
    cfg = datos["config"]
    for nombre, criterios in [("pesos_tecnicos", TECNICOS), ("pesos_criticidad", CRITICIDAD)]:
        pesos = cfg[nombre]
        if set(pesos) != set(criterios) or any(not isinstance(p, (int, float)) or p < 0 for p in pesos.values()):
            raise ValueError("Pesos incompletos o no válidos: " + nombre)
        if not math.isclose(sum(pesos.values()), 1, abs_tol=1e-9):
            raise ValueError("Los pesos deben sumar 1: " + nombre)
    if any(cfg.get(c) is None for c in ["fecha_referencia", "inicio_mantenimiento", "fin_mantenimiento"]):
        raise ValueError("Faltan fechas de referencia del modelo.")
    if cfg["inicio_mantenimiento"] > cfg["fin_mantenimiento"]:
        raise ValueError("El periodo de mantenimiento está invertido.")
    if not 0 < cfg["umbral_correctivos_1"] < cfg["umbral_correctivos_2"]:
        raise ValueError("Umbrales de correctivos no válidos.")
    if not 0 <= cfg["completitud_tecnica_provisional"] <= 1 or cfg["ventana_correctivos_recientes"] < 1:
        raise ValueError("Completitud o ventana de correctivos no válida.")
    if cfg["cobertura_mantenimiento"] not in {"Completa", "Parcial", "No disponible"}:
        raise ValueError("Cobertura de mantenimiento no reconocida.")
    if datos["equipos"]["codigo_activo"].tolist() != datos["evaluacion"]["codigo_activo"].tolist():
        avisos.append("El orden de evaluación difiere del inventario. Revisar las referencias por fila del Excel.")
    series = datos["equipos"].set_index("codigo_activo")["numero_serie"]
    distintos = set()
    for _, r in partes.iterrows():
        a, b = str(series[r["codigo_activo"]]).casefold(), str(r["numero_serie_equipo"]).casefold()
        if a and b and a != b:
            distintos.add(r["codigo_activo"])
    if distintos:
        avisos.append("Diferencias de serie entre inventario y partes en los códigos: " + ", ".join(sorted(distintos)))
    if datos["resultados"]["estado_resultado"].map(vacio).any():
        raise ValueError("Faltan resultados guardados. Recalcula y guarda el archivo en Excel.")
    return avisos


# 5. PUNTUACIONES E ÍNDICES


def puntuacion(valor):
    """Obtiene los puntos de una valoración y conserva ND como dato ausente."""
    if vacio(valor) or str(valor).strip().upper().startswith("ND"):
        return None
    if isinstance(valor, (int, float)) and valor in (0, 1, 2, 3):
        return int(valor)
    coincidencia = re.match(r"^([0-3])(?:\s*-|$)", str(valor).strip())
    if not coincidencia:
        raise ValueError("Puntuación no reconocida: " + str(valor))
    return int(coincidencia[1])


def puntuar_correctivos(valor, cfg):
    if vacio(valor):
        return None
    if valor == 0:
        return 0
    if valor <= cfg["umbral_correctivos_1"]:
        return 1
    return 2 if valor <= cfg["umbral_correctivos_2"] else 3


def nivel(indice, cortes):
    if vacio(indice):
        return "Pendiente"
    # El redondeo evita cambios de nivel debidos únicamente a la representación decimal.
    indice = round(float(indice), 10)
    for limite, etiqueta in zip(cortes, ["Baja", "Moderada", "Alta"]):
        if indice < limite:
            return etiqueta
    return "Muy alta"


def indice_tecnico(puntos, cfg):
    """Calcula los extremos del índice técnico y el nivel que puede asignarse."""
    pesos = cfg["pesos_tecnicos"]
    # Los criterios ausentes se consideran con 0 y 3 puntos en los extremos inferior y superior.
    # Sus pesos no se redistribuyen.
    inferior = 100 * sum(pesos[c] * p / 3 for c, p in puntos.items() if p is not None)
    superior = inferior + 100 * sum(pesos[c] for c, p in puntos.items() if p is None)
    completitud = round(sum(pesos[c] for c, p in puntos.items() if p is not None), 12)
    bajo, alto = nivel(inferior, cfg["cortes"]), nivel(superior, cfg["cortes"])
    # Se exige la completitud mínima y que ambos extremos correspondan al mismo nivel.
    clasificacion = bajo if completitud >= cfg["completitud_tecnica_provisional"] and bajo == alto else "Pendiente"
    return inferior, superior, completitud, clasificacion


# 6. CÁLCULO DE LOS RESULTADOS


def calcular_resultados(datos, config=None):
    """Calcula los indicadores, índices, niveles y prioridad de cada equipo."""
    cfg = config if config is not None else datos["config"]
    evaluaciones = datos["evaluacion"].set_index("codigo_activo")
    vidas = {str(r["tipo_equipo"]).strip().casefold(): r for _, r in datos["vida_util"].iterrows()}
    partes = datos["mantenimiento"]
    filas = []
    for _, equipo in datos["equipos"].iterrows():
        codigo = equipo["codigo_activo"]
        ev = evaluaciones.loc[codigo]
        # 6.1. FECHAS Y ANTIGÜEDAD
        # Se toma la instalación como origen y, si falta, la adquisición.
        origen = fecha(equipo["fecha_instalacion"])
        if origen is None:
            origen = fecha(equipo["fecha_adquisicion"])
        baja = fecha(equipo["fecha_baja"])
        antiguedad = None
        if origen is not None and origen <= cfg["fecha_referencia"]:
            años = (cfg["fecha_referencia"] - origen).days / 365.25
            antiguedad = float(Decimal(str(años)).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP))
        vida = vidas.get(str(equipo["descripcion_equipo"]).strip().casefold())
        inferior = superior = medio = punt_edad = None
        if vida is not None:
            inferior = vida["limite_inferior_años"]
            superior = vida["limite_superior_años"]
            inferior = None if vacio(inferior) or inferior <= 0 else float(inferior)
            superior = None if vacio(superior) or superior <= 0 else float(superior)
        if inferior is not None and superior is not None:
            medio = (inferior + superior) / 2
            if antiguedad is not None:
                punt_edad = 0 if antiguedad < inferior else 1 if antiguedad < medio else 2 if antiguedad < superior else 3

        # 6.2. INDICADORES DE MANTENIMIENTO
        # El periodo individual se limita por el origen, la baja y las fechas del análisis.
        inicio = max(cfg["inicio_mantenimiento"], origen or cfg["inicio_mantenimiento"])
        fin = min(cfg["fin_mantenimiento"], baja or cfg["fin_mantenimiento"])
        exposicion = max(0, (fin - inicio).days + 1) / 365.25 if origen is not None else None
        # Se seleccionan partes cerrados con apertura en el periodo individual
        # y cierre no posterior al corte general de mantenimiento.
        cerrados = partes.loc[
            partes["codigo_activo"].eq(codigo) & partes["estado_parte_normalizado"].eq("Cerrado") &
            partes["fecha_apertura"].ge(inicio) & partes["fecha_apertura"].lt(fin + pd.Timedelta(days=1)) &
            partes["fecha_cierre"].lt(cfg["fin_mantenimiento"] + pd.Timedelta(days=1))]
        correctivos = cerrados.loc[cerrados["tipo_intervencion_normalizado"].eq("Correctivo")]
        desde_recientes = max(inicio, cfg["fin_mantenimiento"] - pd.Timedelta(days=cfg["ventana_correctivos_recientes"] - 1))
        numero = len(correctivos)
        recientes = int(correctivos["fecha_apertura"].ge(desde_recientes).sum())
        # Una cobertura no disponible deja los recuentos sin determinar; no equivale a cero intervenciones.
        if cfg["cobertura_mantenimiento"] == "No disponible":
            numero = recientes = None
        frecuencia = numero / exposicion if numero is not None and exposicion is not None and exposicion > 0 else None
        punt_frecuencia = puntuar_correctivos(frecuencia, cfg)
        punt_recientes = puntuar_correctivos(recientes, cfg)
        # Se conserva la mayor puntuación entre la frecuencia anual y los correctivos recientes.
        punt_correctivos = max(punt_frecuencia, punt_recientes) if punt_frecuencia is not None and punt_recientes is not None else None
        # 6.3. ÍNDICES TÉCNICO Y DE CRITICIDAD
        puntos = {c: puntuacion(ev[c]) for c in TECNICOS if c not in {"antiguedad", "intervenciones_correctivas"}}
        puntos["antiguedad"] = punt_edad
        puntos["intervenciones_correctivas"] = punt_correctivos
        ind_inf, ind_sup, completo_t, nivel_t = indice_tecnico(puntos, cfg)
        crit = {c: puntuacion(ev[c]) for c in CRITICIDAD}
        # La criticidad exige sus tres criterios; el índice técnico admite información parcial según sus límites.
        completo_c = sum(p is not None for p in crit.values()) / 3
        indice_c = 100 * sum(cfg["pesos_criticidad"][c] * p / 3 for c, p in crit.items()) if completo_c == 1 else None
        nivel_c = nivel(indice_c, cfg["cortes"])
        # 6.4. PRIORIDAD Y ESTADO DEL RESULTADO
        # Los niveles se cruzan en la matriz y las bajas quedan fuera de la prioridad.
        if baja is not None and baja <= cfg["fecha_referencia"]:
            prioridad = "Fuera de alcance (baja)"
        elif "Pendiente" in (nivel_t, nivel_c):
            prioridad = "Pendiente"
        else:
            prioridad = cfg["matriz"][nivel_t][nivel_c]
        estado = "Pendiente"
        if "Pendiente" not in (nivel_t, nivel_c):
            estado = "Completo" if completo_t == 1 and completo_c == 1 and cfg["cobertura_mantenimiento"] == "Completa" else "Provisional"

        # 6.5. ALERTAS DE REVISIÓN
        # Las alertas no modifican por sí mismas los índices ni la prioridad.
        alertas = []
        if puntos["estado_funcional"] == 3:
            alertas.append("Función principal comprometida")
        if puntos["soporte_tecnico"] == 3 and puntos["disponibilidad_recambios"] == 3:
            alertas.append("Sin soporte ni recambios")
        if crit["impacto_seguridad_continuidad"] == 3 and crit["disponibilidad_alternativas"] is not None and crit["disponibilidad_alternativas"] >= 2:
            alertas.append("Continuidad crítica sin alternativa suficiente")
        if completo_t < 1:
            alertas.append("Evaluación técnica incompleta")
        if cfg["cobertura_mantenimiento"] == "Parcial":
            alertas.append("Mantenimiento vinculado parcialmente")
        if origen is None:
            alertas.append("Falta fecha de origen")
        alta = fecha(equipo["fecha_alta_inventario"])
        if alta is not None and alta > cfg["fecha_referencia"]:
            alertas.append("Confirmar presencia al corte histórico")
        # Si la validación de Ingeniería está pendiente se avisa de ello. En el resto de casos
        # se revisan también la ausencia de registros y el periodo de observación corto.
        if ev["unidad_revisora"] == "Pendiente de validación por Ingeniería":
            alertas.append("Validación de Ingeniería pendiente")
        else:
            if numero == 0:
                alertas.append("Ausencia de registros de mantenimiento en el periodo" if cerrados.empty else "Ausencia de registros correctivos en el periodo")
            if exposicion is not None and exposicion < 1:
                alertas.append("Periodo de observación corto (<1 año)")

        # 6.6. REGISTRO DE LOS RESULTADOS
        # Los campos se conservan con los mismos nombres que en la hoja de resultados de Excel.
        filas.append({
            "codigo_activo": codigo, "descripcion_equipo": equipo["descripcion_equipo"],
            "familia_tecnologica": equipo["familia_tecnologica"], "centro": equipo["centro"], "servicio": equipo["servicio"],
            "fecha_origen": origen, "antiguedad_años": antiguedad,
            "vida_util_inferior_años": inferior, "vida_util_superior_años": superior, "punto_medio_vida_util": medio,
            "puntuacion_antiguedad": punt_edad, "exposicion_mantenimiento_años": exposicion,
            "numero_correctivos_observados": numero, "frecuencia_anual_correctivos_observada": frecuencia,
            "correctivos_observados_ultimos_12_meses": recientes, "puntuacion_frecuencia": punt_frecuencia,
            "puntuacion_correctivos_12m": punt_recientes, "puntuacion_correctivos": punt_correctivos,
            "puntuacion_estado_funcional": puntos["estado_funcional"], "puntuacion_soporte_tecnico": puntos["soporte_tecnico"],
            "puntuacion_recambios": puntos["disponibilidad_recambios"],
            "puntuacion_obsolescencia_tecnologica": puntos["obsolescencia_tecnologica_software"],
            "indice_tecnico_inferior": ind_inf, "indice_tecnico_superior": ind_sup, "completitud_tecnica": completo_t,
            "nivel_tecnico": nivel_t, "puntuacion_impacto_actividad": crit["impacto_actividad"],
            "puntuacion_alternativas": crit["disponibilidad_alternativas"], "puntuacion_seguridad_continuidad": crit["impacto_seguridad_continuidad"],
            "indice_criticidad": indice_c, "completitud_criticidad": completo_c, "nivel_criticidad": nivel_c,
            "prioridad_orientativa": prioridad, "alertas": "; ".join(alertas) + (";" if alertas else ""), "estado_resultado": estado,
        })
    return pd.DataFrame(filas)


# 7. COMPARACIÓN CON EXCEL


def iguales(a, b):
    """Compara valores con tolerancia numérica y tratamiento de datos ausentes."""
    if vacio(a) or vacio(b):
        return vacio(a) and vacio(b)
    if isinstance(a, Real) and isinstance(b, Real):
        # Se admiten diferencias numéricas mínimas sin considerarlas discrepancias.
        return math.isclose(float(a), float(b), rel_tol=1e-9, abs_tol=1e-8)
    if isinstance(a, pd.Timestamp) or isinstance(b, pd.Timestamp):
        return pd.Timestamp(a) == pd.Timestamp(b)
    return str(a).strip() == str(b).strip()


def comparar_resultados(datos):
    """Contrasta los cálculos de Python con los resultados guardados en Excel."""
    excel = datos["resultados"].set_index("codigo_activo")
    python = calcular_resultados(datos).set_index("codigo_activo")
    if set(excel.index) != set(python.index):
        raise ValueError("Los códigos calculados no coinciden con los del Excel.")
    if set(excel.columns) != set(python.columns):
        raise ValueError("Las columnas de resultados han cambiado. Revisar el lector y el modelo.")
    diferencias = []
    for codigo in excel.index:
        for campo in excel.columns:
            a, b = excel.at[codigo, campo], python.at[codigo, campo]
            if not iguales(a, b):
                diferencias.append({"codigo_activo": codigo, "campo": campo, "valor_excel": a, "valor_python": b})
    return pd.DataFrame(diferencias, columns=["codigo_activo", "campo", "valor_excel", "valor_python"])
