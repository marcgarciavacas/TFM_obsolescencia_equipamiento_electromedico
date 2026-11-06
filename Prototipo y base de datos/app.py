"""Interfaz de consulta del equipamiento electromédico."""


# 1. LIBRERÍAS


from pathlib import Path
from html import escape
from zipfile import BadZipFile
import pandas as pd
import streamlit as st
from modelo import (leer_excel, unir_datos, vacio, TECNICOS, CRITICIDAD,
                    puntuacion, validar_datos, comparar_resultados)


# 2. CONFIGURACIÓN DE LA APLICACIÓN


CARPETA = Path(__file__).resolve().parent


EXCEL_INICIAL = CARPETA / "Base_datos.xlsx"


PRIORIDADES = ["P1. Muy alta", "P2. Alta", "P3. Media", "P4. Baja", "Pendiente", "Fuera de alcance (baja)"]


NIVELES = ["Baja", "Moderada", "Alta", "Muy alta", "Pendiente"]


ETIQUETAS = {
    "codigo_activo": "Código",
    "antiguedad_años": "Antigüedad (años)",
    "exposicion_mantenimiento_años": "Exposición (años)",
    "numero_correctivos_observados": "Correctivos observados",
    "frecuencia_anual_correctivos_observada": "Correctivos/año observado",
    "correctivos_observados_ultimos_12_meses": "Correctivos en 12 meses",
    "duracion_trabajo_h": "Trabajo registrado (h)",
    "obsolescencia_tecnologica_software": "Tecnología y software",
    "intervenciones_correctivas": "Intervenciones correctivas",
    "disponibilidad_recambios": "Disponibilidad de recambios",
    "impacto_seguridad_continuidad": "Seguridad y continuidad",
}


# 3. ESTILOS DE LA INTERFAZ


def aplicar_estilo():
    """Combina los colores de Streamlit con las reglas de estilo.css."""
    opciones = {"fondo": ("backgroundColor", "#ffffff"),
                "panel": ("secondaryBackgroundColor", "#f2f6f8"),
                "texto": ("textColor", "#173344"), "acento": ("primaryColor", "#176b83")}
    variables = ";".join(f"--obs-{nombre}:{st.get_option('theme.' + opcion) or defecto}"
                         for nombre, (opcion, defecto) in opciones.items())
    css = (CARPETA / "estilo.css").read_text(encoding="utf-8")
    st.html("<style>.stApp{" + variables + "}" + css + "</style>")


# 4. FORMATO Y PRESENTACIÓN DE LOS DATOS


def texto_html(valor):
    """Escapa los valores de Excel antes de incorporarlos al HTML."""
    return escape("ND" if vacio(valor) else str(valor))


def prioridad_html(valor):
    clase = str(valor)[:2].lower() if str(valor).startswith(("P1.", "P2.", "P3.", "P4.")) else "nd"
    return f'<span class="obs-prioridad obs-{clase}">{texto_html(valor)}</span>'


def etiqueta(nombre):
    return ETIQUETAS.get(nombre, nombre.replace("_", " ").capitalize())


def numero(valor, decimales=1):
    return "ND" if vacio(valor) else f"{valor:.{decimales}f}".replace(".", ",")


def fecha_texto(valor):
    return "No disponible" if vacio(valor) else pd.Timestamp(valor).strftime("%d/%m/%Y")


def tabla_html(tabla):
    """Presenta tablas de consulta con fechas, números y títulos formateados."""
    cabecera = "".join(f"<th scope='col'>{texto_html(etiqueta(c))}</th>" for c in tabla)
    filas = []
    for _, fila in tabla.iterrows():
        celdas = []
        for columna, valor in fila.items():
            if columna.startswith("fecha_"):
                celda = texto_html(fecha_texto(valor))
            elif columna in {"numero_correctivos_observados", "correctivos_observados_ultimos_12_meses"}:
                celda = numero(valor, 0)
            elif columna == "antiguedad_años":
                celda = numero(valor, 1)
            elif isinstance(valor, float) and not vacio(valor):
                celda = numero(valor, 2)
            else:
                celda = texto_html(valor)
            celdas.append("<td>" + celda + "</td>")
        filas.append("<tr>" + "".join(celdas) + "</tr>")
    st.html('<div class="obs-tabla"><table><thead><tr>' + cabecera +
            "</tr></thead><tbody>" + "".join(filas) + "</tbody></table></div>")


def mostrar_tabla(tabla):
    """Presenta las diferencias encontradas entre los resultados de Excel y Python."""
    vista = tabla.copy()
    configuracion = {}
    for columna in vista:
        if pd.api.types.is_datetime64_any_dtype(vista[columna]):
            configuracion[etiqueta(columna)] = st.column_config.DateColumn(format="DD/MM/YYYY")
            vista[columna] = vista[columna].dt.date
    vista = vista.rename(columns=etiqueta)
    st.dataframe(vista, hide_index=True, width="stretch", column_config=configuracion)


# 5. CARGA Y COMPROBACIÓN DEL EXCEL


@st.cache_data(show_spinner=False)
def cargar(contenido):
    """Lee y valida el libro, contrasta los cálculos y reúne las tablas para la consulta."""
    # La caché reutiliza la carga mientras no cambie el contenido del libro.
    datos = leer_excel(contenido)
    avisos = validar_datos(datos)
    diferencias = comparar_resultados(datos)
    return datos, unir_datos(datos), avisos, diferencias


# 6. NAVEGACIÓN Y FILTROS


def abrir_ficha(codigo):
    st.session_state["activo_consulta"] = codigo
    st.session_state["apartado"] = "Ficha de equipo"


def guardar_activo():
    st.session_state["activo_consulta"] = st.session_state["selector_activo"]


def mover_pagina(clave, cambio):
    st.session_state[clave] = max(0, st.session_state.get(clave, 0) + cambio)


def lista_activos(tabla, alertas=False):
    """Muestra páginas de doce equipos y permite abrir la ficha del activo seleccionado."""
    nombre = "alertas" if alertas else "equipos"
    # Cada listado conserva su propia página en el estado de sesión.
    clave = "pagina_" + nombre
    consulta = tuple(tabla["codigo_activo"])
    # Un cambio de consulta reinicia la paginación.
    if st.session_state.get("consulta_" + nombre) != consulta:
        st.session_state[clave] = 0
        st.session_state["consulta_" + nombre] = consulta
    pagina = min(st.session_state.get(clave, 0), max(0, (len(tabla) - 1) // 12))
    inicio = pagina * 12
    visibles = tabla.iloc[inicio:inicio + 12]
    st.caption(f"{len(tabla)} activos · Selecciona un código para abrir su ficha · "
               f"{inicio + 1 if len(tabla) else 0}–{min(inicio + 12, len(tabla))}")
    anchos = [0.5, 2.3, 1.2, 3] if alertas else [0.5, 2.5, 1.6, 1.1, 1.3]
    titulos = ["Código", "Equipo y modelo", "Estado", "Alertas"] if alertas else ["Código", "Equipo y modelo", "Serie", "Estado", "Prioridad"]
    with st.container(key="cabecera_" + nombre):
        for columna, titulo in zip(st.columns(anchos, gap="small"), titulos):
            columna.caption(titulo)
    for _, r in visibles.iterrows():
        with st.container(key=f"fila_{nombre}_{r['codigo_activo']}"):
            c = st.columns(anchos, gap="small", vertical_alignment="center")
            c[0].button(str(r["codigo_activo"]), type="tertiary", key=f"abrir_{nombre}_{r['codigo_activo']}",
                        help="Abrir ficha del activo " + str(r["codigo_activo"]),
                        on_click=abrir_ficha, args=(r["codigo_activo"],))
            with c[1]:
                st.html(f'<div class="obs-equipo">{texto_html(r["descripcion_equipo"])}'
                        f'<span>{texto_html(r["fabricante"])} · {texto_html(r["modelo"])}</span></div>')
            if alertas:
                with c[2]:
                    st.html('<span class="obs-celda">' + texto_html(r["estado_inventario"]) + "</span>")
                with c[3]:
                    st.html('<span class="obs-celda">' + texto_html(r["alertas"]) + "</span>")
            else:
                with c[2]:
                    st.html('<span class="obs-celda">' + texto_html(r["numero_serie"]) + "</span>")
                with c[3]:
                    st.html('<span class="obs-celda">' + texto_html(r["estado_inventario"]) + "</span>")
                with c[4]:
                    st.html(prioridad_html(r["prioridad_orientativa"]))
    if tabla.empty:
        st.info("No hay equipos con estos filtros.")
    anterior, _, siguiente = st.columns([1, 2, 1])
    anterior.button("Anterior", key="anterior_" + nombre, disabled=pagina == 0,
                     on_click=mover_pagina, args=(clave, -1))
    siguiente.button("Siguiente", key="siguiente_" + nombre, disabled=inicio + 12 >= len(tabla),
                      on_click=mover_pagina, args=(clave, 1))


def filtrar(base):
    """Aplica los filtros y la opción de incluir equipos dados de baja."""
    # Los filtros se combinan; una selección vacía conserva todos los valores de ese campo.
    vista = base.copy()
    for campo, titulo in [("familia_tecnologica", "Familia"), ("servicio", "Servicio"),
                           ("estado_inventario", "Estado de inventario"), ("prioridad_orientativa", "Prioridad")]:
        valores = sorted(base[campo].dropna().astype(str).unique())
        elegidos = st.sidebar.multiselect(titulo, valores, key="filtro_" + campo, placeholder="Todos")
        if elegidos:
            vista = vista.loc[vista[campo].isin(elegidos)]
    incluir = st.sidebar.checkbox("Incluir bajas fuera de alcance", value=False)
    if not incluir:
        vista = vista.loc[vista["prioridad_orientativa"].ne("Fuera de alcance (baja)")]
    return vista


# 7. VISTAS DE CONSULTA


def grafico_categorias(tabla, campo, orden):
    """Representa los recuentos con barras proporcionales a la categoría más numerosa."""
    conteos = tabla[campo].value_counts().reindex(orden, fill_value=0)
    maximo = max(1, int(conteos.max()))
    filas = []
    for categoria, cantidad in conteos.items():
        if cantidad:
            filas.append(f'<div class="obs-barra-fila"><span>{texto_html(categoria)}</span>'
                         f'<div class="obs-pista"><div class="obs-barra" style="width:{100 * cantidad / maximo:.3f}%"></div></div>'
                         f'<span class="obs-numero">{cantidad}</span></div>')
    descripcion = "; ".join(f"{c}: {n} equipos" for c, n in conteos.items() if n)
    st.html('<div class="obs-barras" role="img" aria-label="' + texto_html(descripcion) + '">' + "".join(filas) + "</div>")


def resumen(base, vista):
    """Distingue los recuentos del archivo completo de los correspondientes a la consulta."""
    st.header("Resumen del inventario")
    a, b, c, d = st.columns(4)
    a.metric("Activos en el archivo", len(base))
    b.metric("Activos en la consulta", len(vista))
    c.metric("Bajas fuera de alcance", int(base["prioridad_orientativa"].eq("Fuera de alcance (baja)").sum()))
    d.metric("Con alertas en la consulta", int(vista["alertas"].fillna("").str.strip().ne("").sum()))
    incluidos = vista.loc[vista["prioridad_orientativa"].ne("Fuera de alcance (baja)")]
    if incluidos.empty:
        st.info("No hay activos dentro del alcance con estos filtros.")
        return
    st.subheader("Prioridad orientativa de renovación")
    grafico_categorias(incluidos, "prioridad_orientativa", PRIORIDADES)
    st.write("")
    izquierda, derecha = st.columns(2)
    with izquierda:
        st.subheader("Obsolescencia técnica y funcional")
        grafico_categorias(incluidos, "nivel_tecnico", NIVELES)
    with derecha:
        st.subheader("Criticidad asistencial")
        grafico_categorias(incluidos, "nivel_criticidad", NIVELES)


def equipos(vista):
    """Prepara el listado ordenado por descripción y código del activo."""
    st.header("Equipos")
    columnas = ["codigo_activo", "descripcion_equipo", "fabricante", "modelo", "numero_serie", "servicio", "estado_inventario",
                "nivel_tecnico", "nivel_criticidad", "prioridad_orientativa", "estado_resultado", "alertas"]
    tabla = vista[columnas].copy()
    tabla["orden_codigo"] = pd.to_numeric(tabla["codigo_activo"], errors="coerce")
    tabla = tabla.sort_values(["descripcion_equipo", "orden_codigo", "codigo_activo"]).drop(columns="orden_codigo")
    lista_activos(tabla)


def ficha(datos, vista):
    """Reúne la identificación, la evaluación y los partes vinculados al equipo elegido."""
    st.header("Ficha de equipo")
    if vista.empty:
        st.info("No hay equipos con estos filtros.")
        return
    por_codigo = vista.set_index("codigo_activo")
    def nombre(codigo):
        r = por_codigo.loc[codigo]
        return f"{codigo} · {r['descripcion_equipo']} · {r['fabricante']} {r['modelo']} · Serie {r['numero_serie']}"
    opciones = por_codigo.index.tolist()
    preferido = st.session_state.get("activo_consulta", opciones[0])
    if preferido not in opciones:
        preferido = opciones[0]
    if st.session_state.get("selector_activo") != preferido:
        st.session_state["selector_activo"] = preferido
    codigo = st.selectbox("Seleccionar activo", opciones, format_func=nombre,
                         key="selector_activo", on_change=guardar_activo)
    r = por_codigo.loc[codigo]
    st.subheader(f"{r['descripcion_equipo']} · {r['fabricante']} {r['modelo']}")
    detalles = [("Código", codigo), ("Fabricante / serie", f"{r['fabricante']} / {r['numero_serie']}"),
                ("Servicio", r["servicio"]), ("Ubicación", r["ubicacion"]), ("Estado de inventario", r["estado_inventario"])]
    st.html('<dl class="obs-detalles">' + "".join(f"<dt>{texto_html(k)}</dt><dd>{texto_html(v)}</dd>" for k, v in detalles) + "</dl>")
    a, b, c, d = st.columns(4)
    # Se muestran ambos extremos técnicos cuando los datos ausentes generan un intervalo.
    indice_t = numero(r["indice_tecnico_inferior"])
    if not vacio(r["indice_tecnico_inferior"]) and r["indice_tecnico_inferior"] != r["indice_tecnico_superior"]:
        indice_t += " a " + numero(r["indice_tecnico_superior"])
    a.metric("Índice técnico (0–100)", indice_t)
    b.metric("Índice de criticidad (0–100)", numero(r["indice_criticidad"]))
    c.metric("Prioridad", r["prioridad_orientativa"])
    d.metric("Resultado", r["estado_resultado"])
    st.caption(f"Nivel técnico: {r['nivel_tecnico']} · Criticidad: {r['nivel_criticidad']}")
    for aviso in str(r["alertas"] if not vacio(r["alertas"]) else "").split(";"):
        if aviso.strip():
            st.html('<div class="obs-aviso">' + texto_html(aviso.strip()) + "</div>")
    st.subheader("Antigüedad y mantenimiento")
    tabla_html(pd.DataFrame([{"fecha_origen": r["fecha_origen"], "antiguedad_años": r["antiguedad_años"],
                                "vida_util_de_referencia": numero(r["vida_util_inferior_años"], 0) + "–" + numero(r["vida_util_superior_años"], 0) + " años"}]))
    tabla_html(pd.DataFrame([{"exposicion_mantenimiento_años": r["exposicion_mantenimiento_años"],
                                "numero_correctivos_observados": r["numero_correctivos_observados"],
                                "frecuencia_anual_correctivos_observada": r["frecuencia_anual_correctivos_observada"],
                                "correctivos_observados_ultimos_12_meses": r["correctivos_observados_ultimos_12_meses"]}]))
    st.subheader("Criterios y pesos")
    # Edad y correctivos proceden del cálculo; los demás criterios se recuperan de la evaluación.
    fuentes = {"antiguedad": "puntuacion_antiguedad", "intervenciones_correctivas": "puntuacion_correctivos"}
    criterios = []
    for grupo, nombres in [("Técnica y funcional", TECNICOS), ("Criticidad asistencial", CRITICIDAD)]:
        pesos = datos["config"]["pesos_tecnicos" if grupo == "Técnica y funcional" else "pesos_criticidad"]
        for criterio in nombres:
            valor = r[fuentes[criterio]] if criterio in fuentes else puntuacion(r[criterio])
            detalle = ("Calculada con antigüedad y vida útil" if criterio == "antiguedad" else
                       "Máximo entre frecuencia anual y correctivos recientes" if criterio == "intervenciones_correctivas" else str(r[criterio]))
            criterios.append({"dimensión": grupo, "criterio": etiqueta(criterio), "puntuación": "ND" if vacio(valor) else str(int(valor)),
                               "peso (%)": round(100 * pesos[criterio], 2), "valoración": detalle})
    tabla_html(pd.DataFrame(criterios))
    st.caption(f"Valoración: {fecha_texto(r['fecha_valoracion'])}. Unidad revisora: {r['unidad_revisora']}.")
    partes = datos["mantenimiento"].loc[datos["mantenimiento"]["codigo_activo"].eq(codigo)].sort_values("fecha_apertura", ascending=False)
    st.subheader(f"Partes vinculados ({len(partes)})")
    columnas = ["id_parte", "fecha_apertura", "fecha_cierre", "tipo_intervencion_normalizado", "estado_parte_normalizado", "duracion_trabajo_h", "archivo_origen"]
    tabla_html(partes[columnas])


def alertas(vista, avisos):
    """Presenta las alertas de los equipos y los avisos generales del archivo."""
    st.header("Alertas y comprobaciones")
    tabla = vista.loc[vista["alertas"].fillna("").str.strip().ne(""),
                      ["codigo_activo", "descripcion_equipo", "fabricante", "modelo", "estado_inventario", "prioridad_orientativa", "alertas"]]
    lista_activos(tabla, alertas=True)
    st.subheader("Comprobaciones del archivo completo")
    for aviso in avisos:
        st.warning(aviso)
    if not avisos:
        st.info("Sin avisos estructurales detectados.")


# 8. INICIO DE LA APLICACIÓN


def main():
    """Carga la base de datos y organiza las vistas de consulta."""
    st.set_page_config(page_title="Obsolescencia del equipamiento electromédico", layout="wide")
    aplicar_estilo()
    st.html('<div class="obs-centro">Badalona Serveis Assistencials S.A. - Hospital Municipal de Badalona</div>')
    st.title("Obsolescencia del equipamiento electromédico")
    try:
        contenido = EXCEL_INICIAL.read_bytes()
        with st.spinner("Leyendo y contrastando los cálculos…"):
            datos, base, avisos, diferencias = cargar(contenido)
    except (ValueError, KeyError, TypeError, OSError, BadZipFile) as error:
        st.error(f"No se puede analizar el archivo: {error}")
        st.stop()
    # La consulta se detiene cuando los resultados guardados no coinciden con el cálculo de Python.
    if not diferencias.empty:
        st.error(f"Se han encontrado {len(diferencias)} diferencias entre Excel y Python. Recalcula y guarda el libro en Excel. Si persisten, revisa las reglas antes de continuar.")
        mostrar_tabla(diferencias)
        st.stop()
    st.sidebar.html('<div class="obs-kicker">Departamento de Ingeniería Clínica</div><div class="obs-marca">Equipamiento electromédico</div><span class="obs-etiqueta">Consulta del inventario</span>')
    apartados = ["Resumen", "Equipos", "Ficha de equipo", "Alertas"]
    if st.session_state.get("apartado") not in apartados:
        st.session_state["apartado"] = "Resumen"
    pagina = st.sidebar.radio("Apartado", apartados, key="apartado", label_visibility="collapsed")
    vista = filtrar(base)
    st.sidebar.divider()
    st.sidebar.caption(f"{len(base)} activos registrados")
    cfg = datos["config"]
    st.caption(f"Referencia: {fecha_texto(cfg['fecha_referencia'])}. Periodo de partes de Mantenimiento: {fecha_texto(cfg['inicio_mantenimiento'])}–{fecha_texto(cfg['fin_mantenimiento'])}")
    st.html('<hr class="obs-separador">')
    if pagina == "Resumen":
        resumen(base, vista)
    elif pagina == "Equipos":
        equipos(vista)
    elif pagina == "Ficha de equipo":
        ficha(datos, vista)
    elif pagina == "Alertas":
        alertas(vista, avisos)


if __name__ == "__main__":
    main()
