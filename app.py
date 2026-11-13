import io
import os
import re
import time
from datetime import datetime
from typing import Dict, List

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from google import genai
from google.genai import types
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------
# CONFIGURACIÓN GENERAL Y ESTILOS
# ---------------------------------------------------------------------
st.set_page_config(
    page_title="Plataforma SSOMA IA - Chile",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

MODELO_GEMINI = "gemini-3.8-flash"

CATALOGO: Dict[str, List[str]] = {
    "1. Operativos": [
        "MIPER / Matriz de Identificación de Peligros y Evaluación de Riesgos",
        "PTS / Procedimiento de Trabajo Seguro",
        "ATS / Análisis de Trabajo Seguro",
        "IRL / Información de Riesgos Laborales",
        "Inspección planeada de seguridad",
        "Checklist de equipos, herramientas o vehículos",
        "Investigación de incidente o accidente",
        "Programa de trabajo preventivo",
        "Permiso de trabajo",
        "Estándar de seguridad operacional",
    ],
    "2. Sistemas de Gestión (SIG/ISO)": [
        "Procedimiento del Sistema de Gestión Integrado",
        "Matriz de requisitos legales y otros requisitos",
        "Matriz de aspectos e impactos ambientales",
        "Programa de objetivos y metas SIG (ISO 45001/14001/9001)",
        "Programa de auditoría interna",
        "Plan de acción / tratamiento de hallazgos",
        "Informe de auditoría interna",
        "Acta de Revisión por la Dirección",
        "Control operacional integrado",
        "Evaluación de cumplimiento legal",
    ],
    "3. Salud Ocupacional": [
        "Programa de vigilancia de la salud ocupacional",
        "Protocolo de vigilancia ocupacional (MINSAL)",
        "Evaluación de exposición ocupacional",
        "Programa de ergonomía (TMERT / MMC)",
        "Programa PREXOR (Ruido Ocupacional)",
        "Programa de higiene ocupacional",
        "Procedimiento de factores psicosociales y Ley Karin (21.643)",
        "Plan de intervención de riesgos para la salud",
        "Charla de Seguridad (5 Minutos)",
    ],
    "4. Legal/CPHS": [
        "RIOHS / Reglamento Interno de Orden, Higiene y Seguridad",
        "Acta de reunión del Comité Paritario (CPHS - DS 44)",
        "Constitución y funcionamiento del CPHS",
        "Registro de elección de representantes CPHS",
        "Protocolo de prevención Ley Karin (21.643)",
        "Procedimiento de investigación Ley Karin",
        "Registro de entrega de información preventiva (ODI/IRL)",
        "Plan anual de prevención de riesgos",
        "Matriz de obligaciones legales SST",
    ],
    "5. Emergencias/GRD": [
        "Plan de emergencia y evacuación",
        "Plan de Gestión del Riesgo de Desastres (GRD)",
        "Procedimiento ante emergencias (Incendio/Derrames/Sismos)",
        "Plan de continuidad operacional",
        "Matriz de amenazas y vulnerabilidades",
        "Programa e informe de simulacros",
    ],
}

SECTORES = [
    "Transporte y logística",
    "Construcción y obras civiles",
    "Minería y Faenas",
    "Industria y manufactura",
    "Servicios Sanitarios y Ambientales",
    "Salud y Clínicas",
    "Educación",
    "Comercio y Retail",
    "Agricultura y agroindustria",
    "Energía y Electricidad",
    "Administración pública",
    "Otro / No especificado",
]

MARCO_NORMATIVO = """
MARCO NORMATIVO CHILENO VIGENTE
- Ley N° 16.744 sobre accidentes del trabajo y enfermedades profesionales.
- Decreto Supremo N° 44/2024 del Ministerio del Trabajo y Previsión Social: Reglamento sobre gestión preventiva de los riesgos laborales (Vigente desde el 1 de febrero de 2025, reemplaza al DS 40 y DS 54).
- Decreto Supremo N° 594 del Ministerio de Salud (Condiciones sanitarias y ambientales básicas en los lugares de trabajo).
- Ley N° 21.643 (Ley Karin), Código del Trabajo y normativa en prevención, investigación y sanción del acoso laboral, sexual y violencia en el trabajo.
- Normas ISO 45001:2018 (SST), ISO 14001:2015 (Medio Ambiente) e ISO 9001:2015 (Calidad).
- Protocolos de Vigilancia del MINSAL (PREXOR, TMERT, MMC, ERGO, SILICE, PSICOSOCIAL).
"""

PROMPT_MAESTRO = """
Actúa como un equipo técnico multidisciplinario senior en Seguridad y Salud en el Trabajo (SST), prevención de riesgos, higiene ocupacional, medio ambiente, gestión de calidad y cumplimiento legal en Chile.

Tu tarea es elaborar un BORRADOR TÉCNICO DOCUMENTAL, específico y utilizable, con lenguaje profesional chileno, estructura clara, controles verificables y enfoque preventivo.

REGLAS CRÍTICAS DE RIGOR NORMATIVO:
1. No inventes leyes, decretos, artículos ni límites permisibles.
2. Aplica explícitamente el Decreto Supremo N° 44/2024 como marco general vigente de gestión preventiva en Chile desde febrero de 2025 (no cite DS 40 ni DS 54 como vigentes).
3. Adapta los peligros, riesgos, consecuencias y controles al sector y tarea solicitada. Aplica la jerarquía de controles: Eliminación, Sustitución, Ingeniería, Administrativos y EPP.
4. Para matrices (MIPER, EPP, Aspectos Ambientales), entrega obligatoriamente una TABLA MARKDOWN bien formateada con columnas claras.
5. Para procedimientos (PTS, RIOHS, Planes), incorpora objetivo, alcance, responsabilidades, paso a paso operativo, prohibiciones, EPP y emergencias.
6. Aplica Ley Karin (21.643) y protocolos MINSAL vigentes según corresponda al sector.

FORMATO DE SALIDA:
- Entrega el documento en Markdown, listo para revisión.
- Incluye una tabla de identificación documental (Código, Versión, Fecha, Elaborado, Aprobado).
- No agregues comentarios conversacionales fuera del documento técnico.

DATOS DEL REQUERIMIENTO:
Categoría: {categoria}
Documento solicitado: {documento}
Sector económico: {sector}
Actividad o puesto de trabajo: {actividad}
Detalles operativos adicionales: {detalles}

{marco_normativo}
"""


def obtener_api_key() -> str:
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")
    except Exception:
        key = ""
    return key or os.getenv("GEMINI_API_KEY", "")


def generar_documento(
    api_key: str,
    categoria: str,
    documento: str,
    sector: str,
    actividad: str,
    detalles: str,
) -> str:
    client = genai.Client(api_key=api_key)
    prompt = PROMPT_MAESTRO.format(
        categoria=categoria,
        documento=documento,
        sector=sector,
        actividad=actividad,
        detalles=detalles.strip()
        or "No se proporcionaron detalles adicionales.",
        marco_normativo=MARCO_NORMATIVO,
    )
    
    # Sistema de reintentos automáticos ante saturación temporal (Error 503)
    ultimo_error = None
    for intento in range(3):
        try:
            response = client.models.generate_content(
                model=MODELO_GEMINI,
                contents=prompt,
                config=types.GenerateContentConfig(
                    temperature=0.25,
                    top_p=0.85,
                    max_output_tokens=8192,
                ),
            )
            texto = (response.text or "").strip()
            if texto:
                return texto
        except Exception as e:
            ultimo_error = e
            error_str = str(e)
            if "503" in error_str or "UNAVAILABLE" in error_str or "high demand" in error_str:
                time.sleep(3 * (intento + 1))  # Pausa progresiva de espera: 3s, 6s, 9s
                continue
            else:
                raise e
                
    raise RuntimeError(f"Los servidores están experimentando alta demanda (Error 503). Por favor, intenta nuevamente en un momento. Detalle: {ultimo_error}")


def limpiar_nombre_archivo(texto: str) -> str:
    texto = texto.lower().strip()
    texto = re.sub(r"[^a-z0-9áéíóúñ]+", "_", texto)
    texto = texto.strip("_")
    return texto[:70] or "documento_ssoma"


def markdown_a_docx(markdown_text: str, metadata: dict) -> bytes:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(9.5)
    styles["Normal"].paragraph_format.space_after = Pt(4)

    title = doc.add_heading("DOCUMENTACIÓN TÉCNICA SSOMA", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    for key, value in metadata.items():
        cells = table.add_row().cells
        cells[0].text = str(key)
        cells[1].text = str(value)
        for run in cells[0].paragraphs[0].runs:
            run.bold = True
    doc.add_paragraph()

    in_table = False
    table_rows = []

    def flush_table():
        nonlocal table_rows
        if not table_rows:
            return
        parsed = [
            [c.strip() for c in row.strip().strip("|").split("|")]
            for row in table_rows
        ]
        parsed = [
            r
            for r in parsed
            if not all(re.fullmatch(r":?-{3,}:?", c or "-") for c in r)
        ]
        if parsed:
            cols = max(len(r) for r in parsed)
            t = doc.add_table(rows=0, cols=cols)
            t.style = "Table Grid"
            for ridx, row in enumerate(parsed):
                cells = t.add_row().cells
                for cidx in range(cols):
                    # LIMPIEZA DE ETIQUETAS HTML <br> PARA WORD
                    texto_celda = row[cidx] if cidx < len(row) else ""
                    texto_celda = texto_celda.replace("<br>", "\n")
                    
                    cells[cidx].text = texto_celda
                    if ridx == 0:
                        for run in cells[cidx].paragraphs[0].runs:
                            run.bold = True
        table_rows = []

    for raw_line in markdown_text.splitlines():
        line = raw_line.strip()
        if line.startswith("|") and line.endswith("|"):
            in_table = True
            table_rows.append(line)
            continue
        if in_table:
            flush_table()
            in_table = False

        if not line:
            doc.add_paragraph()
        elif line.startswith("### "):
            doc.add_heading(line[4:], level=3)
        elif line.startswith("## "):
            doc.add_heading(line[3:], level=2)
        elif line.startswith("# "):
            doc.add_heading(line[2:], level=1)
        elif line.startswith("- ") or line.startswith("* "):
            doc.add_paragraph(line[2:], style="List Bullet")
        elif line in ("---", "***"):
            p = doc.add_paragraph("─" * 50)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            plain = re.sub(r"\*\*(.*?)\*\*", r"\1", line)
            plain = re.sub(r"\*(.*?)\*", r"\1", plain)
            doc.add_paragraph(plain)

    if in_table:
        flush_table()

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def generar_excel(markdown_text: str, metadata: dict) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_meta = pd.DataFrame(
            [{"Campo": k, "Valor": v} for k, v in metadata.items()]
        )
        df_meta.to_excel(writer, index=False, sheet_name="Identificación")

        lineas = markdown_text.splitlines()
        lineas_tabla = [l for l in lineas if l.startswith("|") and l.endswith("|")]

        if len(lineas_tabla) >= 2:
            headers = [
                c.strip() for c in lineas_tabla[0].split("|")[1:-1] if c.strip() != ""
            ]
            filas = []
            for linea in lineas_tabla[1:]:
                if "---" in linea:
                    continue
                cols = [c.strip() for c in linea.split("|")[1:-1]]
                if len(cols) == len(headers):
                    filas.append(cols)
            if filas:
                df_matriz = pd.DataFrame(filas, columns=headers)
                df_matriz.to_excel(writer, index=False, sheet_name="Matriz_SSOMA")

        filas_doc = []
        for num, l in enumerate(lineas, start=1):
            limpio = l.strip()
            if limpio:
                filas_doc.append({"N°": num, "Contenido": limpio})
        pd.DataFrame(filas_doc).to_excel(
            writer, index=False, sheet_name="Texto_Completo"
        )

        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F4E78")
            for col in ws.columns:
                max_len = max(
                    (len(str(c.value)) for c in col if c.value is not None), default=10
                )
                ws.column_dimensions[get_column_letter(col[0].column)].width = min(
                    max(max_len + 2, 12), 60
                )

    return buffer.getvalue()


# ---------------------------------------------------------------------
# INTERFAZ WEB STREAMLIT
# ---------------------------------------------------------------------
st.title("🛡️ Plataforma SSOMA | Generador Documental Profesional")
st.caption(
    "Plataforma de asistencia técnica SSOMA adaptada al Decreto Supremo N°"
    " 44/2024, Ley Karin y Normativa Legal de Chile."
)

with st.sidebar:
    st.header("⚙️ Configuración")
    st.markdown(f"**Motor:** `Google Gemini 3.8 Flash`")
    st.info(
        "Los documentos son borradores técnicos y requieren revisión del"
        " Experto en Prevención de Riesgos antes de su aprobación."
    )

col1, col2 = st.columns([1, 1])

with col1:
    # Agregamos la opción "Otro / Personalizado" al final de la lista de categorías
    categorias_con_otro = list(CATALOGO.keys()) + ["Otro / Personalizado"]
    categoria_sel = st.selectbox("1. Categoría documental", categorias_con_otro)
    
    if categoria_sel == "Otro / Personalizado":
        categoria = st.text_input("Especifica la categoría documental:")
        # Permitir cualquier documento personalizado si elige categoría libre
        documento = st.text_input("2. Nombre del documento específico requerido:")
    else:
        categoria = categoria_sel
        # Agregamos la opción "Otro / Especificar" al catálogo específico
        lista_docs = CATALOGO[categoria] + ["Otro / Especificar documento"]
        documento_sel = st.selectbox("2. Documento específico", lista_docs)
        
        if documento_sel == "Otro / Especificar documento":
            documento = st.text_input("2. Escribe el documento específico que necesitas:")
        else:
            documento = documento_sel

with col2:
    sector_sel = st.selectbox("3. Sector económico", SECTORES + ["Otro / Especificar"])
    if sector_sel == "Otro / Especificar":
        sector = st.text_input("Especifica el sector económico:")
    else:
        sector = sector_sel
        
    actividad = st.text_input(
        "4. Actividad o puesto de trabajo",
        placeholder="Ej: Mantención preventiva de camiones en taller",
    )

detalles = st.text_area(
    "5. Detalles operativos adicionales",
    placeholder=(
        "Ingresa turnos, herramientas utilizadas, uso de sustancias"
        " químicas, trabajos en altura o condiciones del terreno..."
    ),
    height=140,
)

api_key = obtener_api_key()
if not api_key:
    st.warning("Configura tu API Key de Gemini para operar.")
    api_key = st.text_input("Google Gemini API Key", type="password")

generar = st.button(
    "🚀 Generar Documento SSOMA",
    type="primary",
    use_container_width=True,
    disabled=not bool(api_key.strip()) or not bool(actividad.strip()),
)

if generar:
    with st.spinner("Generando documento bajo normativa chilena vigente..."):
        try:
            resultado = generar_documento(
                api_key=api_key.strip(),
                categoria=categoria,
                documento=documento,
                sector=sector,
                actividad=actividad.strip(),
                detalles=detalles,
            )
            st.session_state["documento_ssoma"] = resultado
            st.session_state["metadata_ssoma"] = {
                "Categoría": categoria,
                "Documento": documento,
                "Sector": sector,
                "Actividad": actividad.strip(),
                "Fecha": datetime.now().strftime("%d-%m-%Y %H:%M"),
            }
        except Exception as exc:
            st.error(f"Error al generar documento: {exc}")

if "documento_ssoma" in st.session_state:
    st.divider()
    st.subheader("Documento Generado (Editable en pantalla)")

    resultado_editado = st.text_area(
        "Revisa o modifica el texto antes de exportar:",
        value=st.session_state["documento_ssoma"],
        height=500,
    )
    st.session_state["documento_ssoma"] = resultado_editado
    metadata = st.session_state.get("metadata_ssoma", {})
    nombre_base = limpiar_nombre_archivo(
        f"{metadata.get('Documento', 'doc')}_{metadata.get('Actividad', 'ssoma')}"
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button(
            "📄 Descargar TXT / Markdown",
            data=resultado_editado.encode("utf-8"),
            file_name=f"{nombre_base}.txt",
            mime="text/plain",
            use_container_width=True,
        )
    with c2:
        docx_bytes = markdown_a_docx(resultado_editado, metadata)
        st.download_button(
            "📝 Descargar Word (.docx)",
            data=docx_bytes,
            file_name=f"{nombre_base}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
            use_container_width=True,
        )
    with c3:
        xlsx_bytes = generar_excel(resultado_editado, metadata)
        st.download_button(
            "📊 Descargar Excel (.xlsx)",
            data=xlsx_bytes,
            file_name=f"{nombre_base}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )
