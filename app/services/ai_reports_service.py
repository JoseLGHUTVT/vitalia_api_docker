import json
import logging
from typing import Any, Dict, List, Optional
import httpx

from app.core.config import settings

logger = logging.getLogger("vitalia.reportes_ia")
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"

SYSTEM_PROMPT_CLINICO = (
    "Eres el Asistente de Gestión Clínica de VITALIA. Ayudas al equipo médico y administrativo "
    "de una clínica a entender su funcionamiento general a partir de las cifras de atención: "
    "citas médicas, médicos, consultorios, tiempos de espera, medicamentos y recetas.\n"
    "REGLAS ESTRICTAS DE LENGUAJE Y ANÁLISIS:\n"
    "1. Usa un lenguaje claro, profesional, empático y médico-administrativo. Hablas con médicos, enfermeros y administradores de clínicas.\n"
    "2. PROHIBIDO USAR TECNICISMOS DE SISTEMAS: No menciones palabras como 'BI', 'Business Intelligence', 'Groq', 'algoritmos', "
    "'machine learning', 'APIs', 'paralización total', 'umbral crítico <40%', ni nombres de variables de código como 'total_citas', 'horas_ocupadas', etc.\n"
    "3. Si en el periodo no hay citas registradas (cero consultas o cero datos), NO uses lenguaje catastrófico o alarmista. "
    "Simplemente explica con amabilidad que no hay citas ni actividad registrada en las fechas seleccionadas e invita a consultar un rango mayor (como los últimos 30 días o todo el historial).\n"
    "4. En 'score_eficiencia', devuelve un número entero del 0 al 100 que represente el nivel de actividad y atención de la clínica. Si no hay actividad en el periodo, asigna 50 indicando 'Periodo sin registros'.\n"
    "5. En 'alertas_anomalias', destaca situaciones prácticas de la clínica: tiempos de espera de pacientes, necesidad de reabastecer medicamentos, uso de consultorios o distribución de consultas médicas.\n"
    "6. En 'recomendaciones', proporciona exactamente 3 sugerencias útiles y concretas para el cuidado de los pacientes, la agenda de los doctores o la farmacia de la clínica.\n"
    "7. Responde SIEMPRE en español en formato JSON estricto con la siguiente estructura:\n"
    "{\n"
    '  "resumen_ejecutivo": "Explicación clara de 3-4 líneas sobre cómo estuvo la atención en la clínica durante este periodo...",\n'
    '  "score_eficiencia": 80,\n'
    '  "alertas_anomalias": [\n'
    '    {\n'
    '      "titulo": "Título de la observación",\n'
    '      "descripcion": "Detalle en lenguaje médico y claro sin tecnicismos",\n'
    '      "prioridad": "critico|oportunidad|optimo",\n'
    '      "metrica": "cifra de referencia (ej. 15 min de espera)"\n'
    "    }\n"
    "  ],\n"
    '  "recomendaciones": [\n'
    '    {\n'
    '      "titulo": "Nombre de la sugerencia",\n'
    '      "accion": "Paso práctico a realizar en la clínica",\n'
    '      "impacto_esperado": "Beneficio esperado para los pacientes o la clínica",\n'
    '      "area": "Consultorios|Tiempos|Medicamentos|Consultas|General"\n'
    "    }\n"
    "  ]\n"
    "}"
)

SYSTEM_PROMPT_CONSULTOR = (
    "Eres el Asistente Clínico de VITALIA. Respondes preguntas de médicos, enfermeros y administradores "
    "de la clínica sobre el funcionamiento de citas, pacientes, disponibilidad de consultorios, recetas y medicamentos. "
    "Sé conciso (máximo 120 palabras), cordial, claro y enfocado en la atención médica y el bienestar de los pacientes. "
    "Cita cifras sencillas basadas en los datos de la clínica. No uses tecnicismos de sistemas, programación ni inteligencia artificial."
)


def _obtener_modelos_groq() -> List[str]:
    candidatos = []
    if settings.GROQ_MODEL:
        candidatos.append(settings.GROQ_MODEL)
    for m in ("openai/gpt-oss-120b", "openai/gpt-oss-20b", "llama-3.3-70b-versatile"):
        if m not in candidatos:
            candidatos.append(m)
    return candidatos


def _calcular_score_algoritmico(datos: Dict[str, Any]) -> int:
    citas_totales = datos.get("total_citas", 0)
    horas_ocupadas = datos.get("horas_ocupadas_totales", 0)

    # Si no hay ninguna actividad en el periodo
    if citas_totales == 0 and horas_ocupadas == 0:
        return 50

    score = 85
    espera_prom = datos.get("espera_promedio_minutos", 0)
    if espera_prom > 30:
        score -= 20
    elif espera_prom > 20:
        score -= 10
    elif espera_prom > 0 and espera_prom <= 12:
        score += 5

    inventario_bajo = datos.get("total_medicamentos_stock_bajo", 0)
    if inventario_bajo > 5:
        score -= 15
    elif inventario_bajo > 0:
        score -= 5

    citas_canceladas = datos.get("citas_canceladas", 0)
    if citas_totales > 0:
        tasa_cancelacion = (citas_canceladas / citas_totales) * 100
        if tasa_cancelacion > 25:
            score -= 12
        elif tasa_cancelacion < 8:
            score += 5

    return max(40, min(98, score))


def generar_analisis_por_reglas(datos: Dict[str, Any], clinica_nombre: str) -> Dict[str, Any]:
    citas_totales = datos.get("total_citas", 0)
    horas_ocupadas = datos.get("horas_ocupadas_totales", 0)
    espera_prom = datos.get("espera_promedio_minutos", 0)
    inventario_bajo = datos.get("total_medicamentos_stock_bajo", 0)
    total_recetas = datos.get("total_recetas", 0)

    # Caso: Periodo sin registros o con cero actividad
    if citas_totales == 0 and horas_ocupadas == 0:
        return {
            "resumen_ejecutivo": (
                f"En las fechas seleccionadas no se registraron citas médicas ni atención en consultorios en {clinica_nombre}. "
                "Te sugerimos consultar un rango mayor de fechas, como los últimos 30 días o todo el historial, para revisar la actividad general de la clínica."
            ),
            "score_eficiencia": 50,
            "alertas_anomalias": [
                {
                    "titulo": "Sin consultas en las fechas elegidas",
                    "descripcion": "No se encontraron citas programadas en este rango de fechas.",
                    "prioridad": "oportunidad",
                    "metrica": "0 citas",
                },
                {
                    "titulo": "Consultorios disponibles",
                    "descripcion": "Los consultorios no tuvieron consultas registradas en este periodo.",
                    "prioridad": "oportunidad",
                    "metrica": "0 horas",
                },
                {
                    "titulo": "Estado de medicamentos",
                    "descripcion": f"Se registran {inventario_bajo} medicamentos con existencias mínimas en farmacia.",
                    "prioridad": "critico" if inventario_bajo > 0 else "optimo",
                    "metrica": f"{inventario_bajo} por reabastecer",
                },
            ],
            "recomendaciones": [
                {
                    "titulo": "Revisar un periodo con actividad",
                    "accion": "Elige la opción de últimos 30 días o todo el historial para ver las citas y médicos activos.",
                    "impacto_esperado": "Consulta completa de la actividad de la clínica",
                    "area": "General",
                },
                {
                    "titulo": "Verificar agenda de consultas",
                    "accion": "Confirmar que los horarios y turnos de los médicos estén agendados en el sistema.",
                    "impacto_esperado": "Mayor visibilidad en la atención médica",
                    "area": "Consultas",
                },
                {
                    "titulo": "Revisión de insumos en farmacia",
                    "accion": "Monitorear el stock de medicamentos para asegurar insumos suficientes para los pacientes.",
                    "impacto_esperado": "Disponibilidad oportuna de medicamentos",
                    "area": "Medicamentos",
                },
            ],
            "origen": "reglas",
            "modelo": "clinico_vitalia",
        }

    score = _calcular_score_algoritmico(datos)
    alertas = []
    recomendaciones = []

    if espera_prom > 20:
        alertas.append({
            "titulo": "Tiempo de espera en recepción",
            "descripcion": f"El tiempo promedio de espera de los pacientes es de {espera_prom} minutos.",
            "prioridad": "critico",
            "metrica": f"{espera_prom} min prom.",
        })
        recomendaciones.append({
            "titulo": "Agilizar la recepción de pacientes",
            "accion": "Organizar la llegada escalonada de pacientes para evitar tiempos prolongados en sala de espera.",
            "impacto_esperado": "Menor tiempo de espera y mayor satisfacción de los pacientes.",
            "area": "Tiempos",
        })
    else:
        alertas.append({
            "titulo": "Atención fluida y oportuna",
            "descripcion": f"Los pacientes esperaron un promedio de {espera_prom} minutos antes de ser atendidos.",
            "prioridad": "optimo",
            "metrica": f"{espera_prom} min prom.",
        })

    if inventario_bajo > 0:
        alertas.append({
            "titulo": "Medicamentos con existencias bajas",
            "descripcion": f"Hay {inventario_bajo} medicamentos que están por agotarse o bajo su nivel mínimo de seguridad.",
            "prioridad": "critico",
            "metrica": f"{inventario_bajo} por reabastecer",
        })
        recomendaciones.append({
            "titulo": "Reabastecimiento de medicamentos",
            "accion": "Solicitar surtido de los medicamentos que tienen pocas unidades en almacén.",
            "impacto_esperado": "Garantizar que los pacientes reciban sus tratamientos sin demoras.",
            "area": "Medicamentos",
        })
    else:
        alertas.append({
            "titulo": "Medicamentos con buen nivel de existencias",
            "descripcion": "Todos los medicamentos de farmacia cuentan con existencias adecuadas.",
            "prioridad": "optimo",
            "metrica": "En orden",
        })

    if horas_ocupadas < 10 and citas_totales > 0:
        alertas.append({
            "titulo": "Disponibilidad de consultorios",
            "descripcion": f"Los consultorios acumulan {horas_ocupadas} horas de atención, con espacios disponibles para programar más consultas.",
            "prioridad": "oportunidad",
            "metrica": f"{horas_ocupadas} hrs",
        })
        recomendaciones.append({
            "titulo": "Aprovechar espacios en consultorios",
            "accion": "Ofrecer horarios adicionales de consulta médica para atender a más pacientes.",
            "impacto_esperado": "Mayor atención de pacientes en la clínica.",
            "area": "Consultorios",
        })
    else:
        recomendaciones.append({
            "titulo": "Recordatorios de citas",
            "accion": "Enviar recordatorios de consulta a los pacientes para asegurar su asistencia puntual.",
            "impacto_esperado": "Disminuir ausencias y mejorar el seguimiento médico.",
            "area": "Consultas",
        })

    resumen = (
        f"En {clinica_nombre} se atendieron {citas_totales} citas médicas durante el periodo, con "
        f"{horas_ocupadas} horas de uso en consultorios. El tiempo promedio de espera fue de {espera_prom} minutos, "
        f"se emitieron {total_recetas} recetas y se reportan {inventario_bajo} medicamentos en nivel bajo de existencias."
    )

    return {
        "resumen_ejecutivo": resumen,
        "score_eficiencia": score,
        "alertas_anomalias": alertas[:4],
        "recomendaciones": recomendaciones[:3],
        "origen": "reglas",
        "modelo": "clinico_vitalia",
    }


async def analizar_reporte_con_ia(
    datos_agregados: Dict[str, Any],
    clinica_nombre: str,
    tipo_reporte: str = "general",
) -> Dict[str, Any]:
    """Genera diagnóstico estructurado amigable para personal médico con respaldo determinístico."""
    if not settings.GROQ_API_KEY:
        logger.info("GROQ_API_KEY no configurada. Generando análisis clínico determinístico.")
        return generar_analisis_por_reglas(datos_agregados, clinica_nombre)

    payload_ia = {
        "clinica": clinica_nombre,
        "tipo_reporte": tipo_reporte,
        "metricas": datos_agregados,
    }

    ultimo_error = None
    for modelo in _obtener_modelos_groq():
        cuerpo = {
            "model": modelo,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT_CLINICO},
                {
                    "role": "user",
                    "content": f"Revisa los siguientes datos de atención de la clínica y ofrece tus observaciones:\n{json.dumps(payload_ia, ensure_ascii=False)}",
                },
            ],
            "temperature": 0.2,
            "response_format": {"type": "json_object"},
        }
        try:
            async with httpx.AsyncClient(timeout=30.0) as cliente:
                res = await cliente.post(
                    GROQ_API_URL,
                    headers={
                        "Authorization": f"Bearer {settings.GROQ_API_KEY}",
                        "Content-Type": "application/json",
                    },
                    json=cuerpo,
                )
                if res.status_code == 200:
                    data = res.json()
                    contenido = data["choices"][0]["message"]["content"]
                    parsed = json.loads(contenido)
                    parsed["origen"] = "ia"
                    parsed["modelo"] = modelo
                    # Asegurar campos requeridos
                    if "resumen_ejecutivo" in parsed and "score_eficiencia" in parsed:
                        return parsed
                else:
                    ultimo_error = f"HTTP {res.status_code}: {res.text}"
                    logger.warning("Error de respuesta al consultar modelo %s: %s", modelo, ultimo_error)
        except Exception as e:
            ultimo_error = str(e)
            logger.warning("Fallo al conectar con modelo %s: %s", modelo, e)

    logger.warning("No se pudo obtener respuesta con modelos disponibles (%s). Usando reglas de contingencia.", ultimo_error)
    return generar_analisis_por_reglas(datos_agregados, clinica_nombre)


async def responder_consulta_ia(
    pregunta: str,
    contexto_datos: Optional[Dict[str, Any]] = None,
    clinica_nombre: str = "",
    tipo_reporte: str = "general",
    contexto_numerico: Optional[Dict[str, Any]] = None,
) -> str:
    """Responde consultas sobre el funcionamiento de la clínica en lenguaje médico y comprensible."""
    ctx = contexto_datos if contexto_datos is not None else (contexto_numerico or {})
    if not settings.GROQ_API_KEY:
        citas = ctx.get("total_citas", 0) or ctx.get("citas", {}).get("total_citas", 0)
        espera = ctx.get("espera_promedio_minutos", 0) or ctx.get("tiempos_espera", {}).get("promedio_minutos", 0)
        meds = ctx.get("total_medicamentos_stock_bajo", 0) or ctx.get("inventario", {}).get("items_en_riesgo", 0)
        return (
            f"En {clinica_nombre}, los datos registrados muestran {citas} citas, "
            f"{espera} minutos de espera en promedio y "
            f"{meds} medicamentos con existencias bajas."
        )

    prompt_usuario = (
        f"Clínica: {clinica_nombre}\n"
        f"Módulo: {tipo_reporte}\n"
        f"Datos de la clínica: {json.dumps(ctx, ensure_ascii=False)}\n\n"
        f"Pregunta del personal de la clínica: {pregunta}"
    )

    for modelo in _obtener_modelos_groq():
        try:
            async with httpx.AsyncClient(timeout=25.0) as cliente:
                res = await cliente.post(
                    GROQ_API_URL,
                    headers={
                        "Authorization": f"Bearer {settings.GROQ_API_KEY}",
                        "Content-Type": "application/json",
                    },
                    json={
                        "model": modelo,
                        "messages": [
                            {"role": "system", "content": SYSTEM_PROMPT_CONSULTOR},
                            {"role": "user", "content": prompt_usuario},
                        ],
                        "temperature": 0.25,
                    },
                )
                if res.status_code == 200:
                    data = res.json()
                    return data["choices"][0]["message"]["content"].strip()
        except Exception as e:
            logger.warning("Error en asistente de la clínica con modelo %s: %s", modelo, e)

    return (
        f"En {clinica_nombre}, se registraron {contexto_datos.get('total_citas', 0)} citas médicas, "
        f"un promedio de espera de {contexto_datos.get('espera_promedio_minutos', 0)} minutos y "
        f"{contexto_datos.get('total_medicamentos_stock_bajo', 0)} medicamentos en nivel bajo."
    )
