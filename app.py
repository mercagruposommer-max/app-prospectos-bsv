import os
import json
from flask import Flask, request, render_template_string

app = Flask(__name__)

# --- CONFIGURACIÓN DE SHAREPOINT ---
SITE_URL = "https://tu-empresa.sharepoint.com/sites/tu-sitio"
USERNAME = "usuario@tu-empresa.com"
PASSWORD = "TuPassword123"

# Base de datos en memoria para el prototipo
REGISTROS_PROSPECTOS = []

# MATRIZ COMPLETA DE TOLLGATES (TG0 A TG13)
TOLLGATES_DATA = {
  "TG0": {
    "objeto": "Lead", "fase": "BD → MO",
    "secciones": [{
      "nombre": "Datos del Contacto",
      "campos": [
        {"id": "nombre", "campo": "Nombre", "tipo": "Texto", "req": True, "ayuda": "Nombre del contacto.", "notas": ""},
        {"id": "apellidos", "campo": "Apellidos", "tipo": "Texto", "req": True, "ayuda": "Apellidos del contacto.", "notas": ""},
        {"id": "empresa", "campo": "Empresa / Razón Social", "tipo": "Texto", "req": True, "ayuda": "Nombre de la empresa del lead.", "notas": ""},
        {"id": "cargo", "campo": "Cargo / Título", "tipo": "Texto", "req": True, "ayuda": "Cargo o posición del contacto.", "notas": ""},
        {"id": "email", "campo": "Email Corporativo", "tipo": "Email", "req": True, "ayuda": "Email corporativo del contacto.", "notas": ""},
        {"id": "telefono", "campo": "Teléfono Contacto", "tipo": "Teléfono", "req": False, "ayuda": "Teléfono directo o móvil.", "notas": ""},
        {"id": "pais_region", "campo": "País / Región", "tipo": "Lista (picklist)", "req": True, "ayuda": "País o región donde opera el lead.", "notas": "", "opts": ["Mexico", "USA"]}
      ]
    }, {
      "nombre": "Origen y Asignación",
      "campos": [
        {"id": "notas_adicionales", "campo": "Notas Adicionales", "tipo": "Texto largo", "req": False, "ayuda": "Notas generales sobre el lead.", "notas": "NO REQUERIDO"}
      ]
    }]
  },
  "TG1": {
    "objeto": "Lead", "fase": "MO",
    "secciones": [{
      "nombre": "Segmentación y Fit",
      "campos": [
        {"id": "macro_segmento", "campo": "Macro Segmento", "tipo": "Lista (picklist)", "req": True, "ayuda": "Macro segmento al que pertenece el lead.", "notas": "", "opts": ["Industrial", "Automotriz", "Alimentos y Bebidas"]},
        {"id": "sub_segmento", "campo": "Sub-Segmento", "tipo": "Lista (picklist)", "req": False, "ayuda": "Sub-segmento específico.", "notas": "", "opts": ["Manufactura", "Empaque", "Ensamblaje"]},
        {"id": "geografia", "campo": "Geografía", "tipo": "Lista (picklist)", "req": True, "ayuda": "Región geográfica donde opera el cliente.", "notas": "", "opts": ["Norte", "Centro", "Bajio", "Occidente", "Golfo"]},
        {"id": "relevancia", "campo": "Relevancia del Portafolio", "tipo": "Lista (picklist)", "req": True, "ayuda": "Nivel de relevancia del portafolio.", "notas": "", "opts": ["Alta", "Media", "Baja"]},
        {"id": "tamano_empresa", "campo": "Tamaño de Empresa", "tipo": "Lista (picklist)", "req": True, "ayuda": "Estimación del tamaño por ventas anuales.", "notas": "", "opts": ["Micro — <$10M", "Pequeña — $10-50M", "Mediana — $50-200M", "Grande — $200M-$1B", "Enterprise — >$1B"]},
        {"id": "banda", "campo": "Banda Asignada", "tipo": "Lista (picklist)", "req": True, "ayuda": "Banda de clasificación BSV del cliente.", "notas": "si es K o A, entonces es BSV - Normal", "opts": ["K", "A", "B", "C", "D"]}
      ]
    }]
  },
  "TG2": {
    "objeto": "Lead", "fase": "MO",
    "secciones": [{
      "nombre": "Engagement y Señales",
      "campos": [
        {"id": "rol_contacto", "campo": "Rol del Contacto", "tipo": "Lista (picklist)", "req": True, "ayuda": "Rol del contacto en la decisión.", "notas": "", "opts": ["Decision Maker", "Influencer", "Evaluador técnico", "Usuario final", "Desconocido"]}
      ]
    }, {
      "nombre": "Enriquecimiento",
      "campos": [
        {"id": "num_plantas", "campo": "Número de Plantas / Ubicaciones", "tipo": "Número", "req": False, "ayuda": "Número estimado de plantas operativas.", "notas": ""},
        {"id": "contactos_adic", "campo": "Contactos Adicionales Identificados", "tipo": "Texto largo", "req": False, "ayuda": "Nombre, cargo y email de contactos adicionales.", "notas": ""}
      ]
    }]
  },
  "TG3": {
    "objeto": "Lead", "fase": "MO",
    "secciones": [{
      "nombre": "Conversión MQL",
      "campos": [
        {"id": "codigo_xz", "campo": "Código XZ", "tipo": "Lista (picklist)", "req": True, "ayuda": "Código de clasificación XZ.", "notas": "Consumo", "opts": ["Directo, Indirecto"]}
      ]
    }, {
      "nombre": "Notas de Transferencia",
      "campos": [
        {"id": "puntos_interes", "campo": "Puntos de Interés Detectados", "tipo": "Texto largo", "req": True, "ayuda": "Hipótesis de necesidades o pain points.", "notas": ""},
        {"id": "calidad_contacto", "campo": "Calidad del Contacto", "tipo": "Lista (picklist)", "req": True, "ayuda": "Evaluación subjetiva de la calidad del contacto.", "notas": "", "opts": ["Alta", "Media", "Baja"]},
        {"id": "comentarios_ventas", "campo": "Comentarios Adicionales para Ventas", "tipo": "Texto largo", "req": False, "ayuda": "Notas adicionales para Ventas.", "notas": ""}
      ]
    }]
  },
  "TG4": {
    "objeto": "Lead", "fase": "MO → WO",
    "secciones": [{
      "nombre": "Validación ERP",
      "campos": [
        {"id": "ventas_12m", "campo": "Ventas Últimos 12 Meses", "tipo": "Lista (picklist)", "req": True, "ayuda": "Rango de ventas de los últimos 12 meses.", "notas": "", "opts": ["Sin historial", "$0", "<$20k"]},
        {"id": "actividad_lammsa", "campo": "Actividad Lammsa Previa", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si tiene historial con productos Sommer.", "notas": "", "opts": ["Nunca ha comprado", "Histórico inactivo", "Activo — producto Sommer"]}
      ]
    }, {
      "nombre": "Clasificación Comercial",
      "campos": [
        {"id": "clasif_comercial", "campo": "Clasificación Comercial", "tipo": "Lista (picklist)", "req": True, "ayuda": "Clasificación comercial del lead.", "notas": "", "opts": ["K, A, B, C, D"]},
        {"id": "justif_clasif", "campo": "Justificación de Clasificación", "tipo": "Texto largo", "req": True, "ayuda": "Evidencia de la lógica aplicada.", "notas": ""}
      ]
    }, {
      "nombre": "Asignación y SLA",
      "campos": [
        {"id": "osp_asignado", "campo": "OSP Asignado", "tipo": "Texto", "req": True, "ayuda": "Nombre del OSP que recibe la cuenta.", "notas": ""},
        {"id": "gerente_area", "campo": "Gerente de Área", "tipo": "Texto", "req": True, "ayuda": "Nombre del Gerente responsable.", "notas": ""},
        {"id": "fecha_asignacion", "campo": "Fecha de Asignación", "tipo": "Fecha", "req": True, "ayuda": "Fecha de asignación formal.", "notas": ""},
        {"id": "fecha_comp_1ra_inter", "campo": "Fecha Compromiso 1ra Interacción", "tipo": "Fecha", "req": True, "ayuda": "Fecha comprometida para primera interacción (<=5 días).", "notas": "Convierte a cuenta"}
      ]
    }]
  },
  "TG5": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "Datos de la WO",
      "campos": [
        {"id": "tipo_oportunidad", "campo": "Tipo de Oportunidad", "tipo": "Lista (picklist)", "req": True, "ayuda": "Clasificación comercial de la oportunidad.", "notas": "", "opts": ["XR — Reactivación", "XP — Prospecto", "XS — Cross Sell Sommer"]},
        {"id": "aoi_estimado", "campo": "AOI Estimado (%)", "tipo": "Porcentaje (%)", "req": True, "ayuda": "Porcentaje de ahorro o ingreso estimado.", "notas": ""},
        {"id": "primera_inter_ok", "campo": "Primera Interacción Completada", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si se completó la primera interacción.", "notas": "Enviar Template de Agradecimiento", "opts": ["Sí", "No — pendiente"]},
        {"id": "fecha_1ra_inter", "campo": "Fecha de Primera Interacción", "tipo": "Fecha", "req": False, "ayuda": "Fecha real de la primera interacción.", "notas": ""},
        {"id": "formato_1ra_inter", "campo": "Formato de Primera Interacción", "tipo": "Lista (picklist)", "req": False, "ayuda": "Canal o formato utilizado.", "notas": "", "opts": ["Presencial — planta", "Presencial — oficina", "Virtual"]}
      ]
    }, {
      "nombre": "Plan BSV",
      "campos": [
        {"id": "estado_plan_bsv", "campo": "Estado del Plan BSV", "tipo": "Lista (picklist)", "req": True, "ayuda": "Estado del Plan BSV para la oportunidad.", "notas": "", "opts": ["Sí — completado", "En proceso", "No — pendiente"]}
      ]
    }]
  },
  "TG6": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "Plan BSV",
      "campos": [
        {"id": "fecha_est_plan_bsv", "campo": "Fecha Estimada Presentación Plan BSV", "tipo": "Fecha", "req": False, "ayuda": "Fecha estimada para presentar el Plan BSV.", "notas": ""},
        {"id": "notas_plan_bsv", "campo": "Notas del Plan BSV", "tipo": "Texto largo", "req": False, "ayuda": "Consideraciones iniciales del Plan BSV.", "notas": ""}
      ]
    }, {
      "nombre": "Decisión GO / NO-GO",
      "campos": [
        {"id": "decision_go", "campo": "Decisión GO / NO-GO", "tipo": "Lista (picklist)", "req": True, "ayuda": "Decisión formal de continuidad.", "notas": "Flujo de Aprobación Gerencia", "opts": ["GO — continuar con investigación BSV", "NO-GO — salida", "TICKLER — re-evaluar en fecha futura"]},
        {"id": "justif_decision", "campo": "Justificación de la Decisión", "tipo": "Texto largo", "req": True, "ayuda": "Justificación obligatoria.", "notas": ""},
        {"id": "siguiente_paso", "campo": "Siguiente Paso Concreto", "tipo": "Texto", "req": True, "ayuda": "Siguiente acción específica.", "notas": ""},
        {"id": "fecha_sig_paso", "campo": "Fecha del Siguiente Paso", "tipo": "Fecha", "req": True, "ayuda": "Fecha comprometida para el siguiente paso.", "notas": ""}
      ]
    }, {
      "nombre": "MEDDICC",
      "campos": [
        {"id": "meddicc_i_score", "campo": "MEDDICC I — Score (Dolor Implicado)", "tipo": "Número", "req": True, "ayuda": "Score Implicated Pain (1-10). Gate: >=4.", "notas": ""},
        {"id": "meddicc_i_notas", "campo": "MEDDICC I — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Descripción del dolor y costo de inacción.", "notas": ""},
        {"id": "meddicc_m_score", "campo": "MEDDICC M — Score (Métricas)", "tipo": "Número", "req": True, "ayuda": "Score Metrics (1-10). Gate: >=5.", "notas": ""},
        {"id": "meddicc_m_notas", "campo": "MEDDICC M — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Métricas cuantificadas de impacto.", "notas": ""},
        {"id": "meddicc_c2_score", "campo": "MEDDICC C2 — Score (Competencia)", "tipo": "Número", "req": True, "ayuda": "Score Competition (1-10).", "notas": ""},
        {"id": "meddicc_c2_notas", "campo": "MEDDICC C2 — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Proveedor actual y ventajas de LAMMSA.", "notas": ""}
      ]
    }]
  },
  "TG7": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "MEDDICC",
      "campos": [
        {"id": "meddicc_c1_score", "campo": "MEDDICC C1 — Score (Campeón)", "tipo": "Número", "req": True, "ayuda": "Score Champion (1-10).", "notas": "Agregar en contacto"},
        {"id": "meddicc_c1_notas", "campo": "MEDDICC C1 — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Nombre del Campeón y evidencia.", "notas": "Requiere interacción previa"},
        {"id": "meddicc_d1_score", "campo": "MEDDICC D1 — Score (Criterios de Decisión)", "tipo": "Número", "req": True, "ayuda": "Score Decision Criteria (1-10).", "notas": ""},
        {"id": "meddicc_d1_notas", "campo": "MEDDICC D1 — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Criterios de evaluación del cliente.", "notas": ""}
      ]
    }]
  },
  "TG8": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "MEDDICC",
      "campos": [
        {"id": "meddicc_e_score", "campo": "MEDDICC E — Score (Comprador Económico)", "tipo": "Número", "req": True, "ayuda": "Score Economic Buyer (1-10).", "notas": "Agregar en contacto"},
        {"id": "meddicc_e_notas", "campo": "MEDDICC E — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Nombre y nivel de soporte del Economic Buyer.", "notas": "Requiere interacción previa"},
        {"id": "meddicc_d2_score", "campo": "MEDDICC D2 — Score (Proceso de Decisión)", "tipo": "Número", "req": True, "ayuda": "Score Decision Process (1-10).", "notas": ""},
        {"id": "meddicc_d2_notas", "campo": "MEDDICC D2 — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Pasos del proceso de decisión y timeline.", "notas": ""},
        {"id": "meddicc_p_score", "campo": "MEDDICC P — Score (Proceso de Contratación)", "tipo": "Número", "req": True, "ayuda": "Score Paper Process (1-10).", "notas": ""},
        {"id": "meddicc_p_notas", "campo": "MEDDICC P — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Pasos de revisión legal y firmas.", "notas": ""}
      ]
    }]
  },
  "TG9": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "Investigación BSV",
      "campos": [
        {"id": "resumen_bsv", "campo": "Resumen Ejecutivo BSV", "tipo": "Texto largo", "req": True, "ayuda": "Resumen de la situación y propuesta.", "notas": ""},
        {"id": "mapa_stakeholders", "campo": "Mapa de Stakeholders y Coaches", "tipo": "Texto largo", "req": True, "ayuda": "Mapa de actores clave.", "notas": ""},
        {"id": "hipotesis_fit", "campo": "Hipótesis de Fit de Negocio", "tipo": "Texto largo", "req": True, "ayuda": "Cómo LAMMSA resuelve el problema.", "notas": ""},
        {"id": "aoi_cuantificado", "campo": "AOI Cuantificado ($)", "tipo": "Moneda ($)", "req": True, "ayuda": "Valor en dólares del AOI.", "notas": ""},
        {"id": "analisis_incumbentes", "campo": "Análisis de Incumbentes", "tipo": "Texto largo", "req": True, "ayuda": "Análisis de competidores actuales.", "notas": ""},
        {"id": "posicionamiento_diferenciador", "campo": "Posicionamiento Diferenciador LAMMSA", "tipo": "Texto largo", "req": True, "ayuda": "Argumento diferenciador de LAMMSA.", "notas": ""}
      ]
    }, {
      "nombre": "Aprobación Gerencial",
      "campos": [
        {"id": "aprobacion_summit", "campo": "Aprobación para Summit", "tipo": "Lista (picklist)", "req": True, "ayuda": "Aprobación para el Summit.", "notas": "Subir deck BSV", "opts": ["Aprobado para Summit", "Pendiente de revisión", "No aprobado"]},
        {"id": "fecha_rev_gerencial", "campo": "Fecha de Revisión Gerencial", "tipo": "Fecha", "req": False, "ayuda": "Fecha de revisión gerencial.", "notas": ""},
        {"id": "notas_aprobacion", "campo": "Notas de Aprobación", "tipo": "Texto largo", "req": False, "ayuda": "Notas del Gerente de Área.", "notas": ""}
      ]
    }]
  },
  "TG10": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "Datos del Summit",
      "campos": [
        {"id": "fecha_bsv_summit", "campo": "Fecha del BSV Summit", "tipo": "Fecha", "req": True, "ayuda": "Fecha acordada para el Summit.", "notas": "Mail de confirmación requerido"},
        {"id": "participantes_cliente", "campo": "Participantes del Cliente", "tipo": "Texto largo", "req": True, "ayuda": "Nombre y cargo de asistentes del cliente.", "notas": ""},
        {"id": "participantes_lammsa", "campo": "Participantes LAMMSA", "tipo": "Texto largo", "req": True, "ayuda": "Representantes de LAMMSA presentes.", "notas": ""}
      ]
    }]
  },
  "TG11": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "Datos del Summit",
      "campos": [
        {"id": "deck_presentado", "campo": "Deck Presentado", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si el deck fue presentado.", "notas": "", "opts": ["Sí", "No — pendiente"]},
        {"id": "roundtable_ejecutado", "campo": "Round-Table Ejecutado", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si se ejecutó el round-table.", "notas": "", "opts": ["Sí", "Parcial", "No"]},
        {"id": "acuerdos_summit", "campo": "Acuerdos y Próximos Pasos del Summit", "tipo": "Texto largo", "req": True, "ayuda": "Acuerdos formalizados.", "notas": "Mail de agradecimiento requerido"}
      ]
    }, {
      "nombre": "Resultado del Summit",
      "campos": [
        {"id": "resultado_summit", "campo": "Resultado del Summit", "tipo": "Lista (picklist)", "req": True, "ayuda": "Resultado formal del Summit.", "notas": "", "opts": ["Compromiso verbal — proceden", "Compromiso escrito — NDA / LOI firmado", "Interés sin compromiso — más información", "Sin compromiso — no avanzan"]},
        {"id": "notas_resultado_summit", "campo": "Notas del Resultado del Summit", "tipo": "Texto largo", "req": False, "ayuda": "Notas y objeciones del Summit.", "notas": ""}
      ]
    }, {
      "nombre": "Alcance de la Solución",
      "campos": [
        {"id": "plantas_scope", "campo": "Plantas en Scope", "tipo": "Texto", "req": True, "ayuda": "Plantas incluidas en el alcance.", "notas": ""},
        {"id": "categorias_solucion", "campo": "Categorías Incluidas en la Solución", "tipo": "Texto largo", "req": True, "ayuda": "Categorías de producto/servicio.", "notas": ""},
        {"id": "modelo_servicio", "campo": "Modelo de Servicio", "tipo": "Lista (picklist)", "req": True, "ayuda": "Modelo de servicio propuesto.", "notas": "", "opts": ["Full service — inventario LAMMSA", "Consignación", "Mixto", "A definir"]},
        {"id": "plan_implementacion", "campo": "Plan de Implementación por Fases", "tipo": "Texto largo", "req": True, "ayuda": "Entregables, fechas y hitos.", "notas": ""}
      ]
    }, {
      "nombre": "Economía de la Propuesta",
      "campos": [
        {"id": "calculadora_aoi_ok", "campo": "Calculadora AOI Completada", "tipo": "Lista (picklist)", "req": True, "ayuda": "Validación por Finanzas.", "notas": "", "opts": ["Sí — validada", "No — pendiente"]},
        {"id": "aoi_proyectado_ano1", "campo": "AOI Proyectado Año 1 ($)", "tipo": "Moneda ($)", "req": True, "ayuda": "AOI proyectado primer año ($USD).", "notas": ""},
        {"id": "aoi_proyectado_anos2_3", "campo": "AOI Proyectado Años 2-3 ($)", "tipo": "Moneda ($)", "req": True, "ayuda": "AOI proyectado años 2-3 ($USD).", "notas": ""},
        {"id": "margen_vs_floor", "campo": "Margen vs Floor Mínimo", "tipo": "Lista (picklist)", "req": True, "ayuda": "Evaluación de margen vs floor.", "notas": "Bajo el floor no procede", "opts": ["Sobre el floor — aprobado", "En el floor — requiere excepción", "Bajo el floor — no procede"]}
      ]
    }]
  },
  "TG12": {
    "objeto": "Opportunity", "fase": "WO",
    "secciones": [{
      "nombre": "Aprobaciones Internas",
      "campos": [
        {"id": "aprobacion_ops", "campo": "Aprobación Ops / Supply Chain", "tipo": "Lista (picklist)", "req": True, "ayuda": "Aprobación formal de Operaciones.", "notas": "", "opts": ["Aprobado", "Pendiente", "Rechazado"]},
        {"id": "aprobacion_finanzas", "campo": "Aprobación Finanzas", "tipo": "Lista (picklist)", "req": True, "ayuda": "Aprobación formal de Finanzas.", "notas": "", "opts": ["Aprobado", "Pendiente", "Rechazado"]},
        {"id": "propuesta_aprobada_int", "campo": "Propuesta Internamente Aprobada", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indicador global de aprobaciones.", "notas": "", "opts": ["Sí — aprobación completa", "No — pendiente aprobaciones"]},
        {"id": "fecha_presentacion_cliente", "campo": "Fecha de Presentación al Cliente", "tipo": "Fecha", "req": True, "ayuda": "Fecha confirmada de presentación.", "notas": "Mail de confirmación requerido"}
      ]
    }]
  },
  "TG13": {
    "objeto": "Opportunity", "fase": "WO → SO",
    "secciones": [{
      "nombre": "Presentación y Negociación",
      "campos": [
        {"id": "fecha_presentacion_prop", "campo": "Fecha de Presentación de Propuesta", "tipo": "Fecha", "req": True, "ayuda": "Fecha real de presentación.", "notas": "Mail de agradecimiento requerido"},
        {"id": "participantes_pres_cliente", "campo": "Participantes del Cliente", "tipo": "Texto largo", "req": True, "ayuda": "Asistentes a la presentación.", "notas": ""},
        {"id": "ce_presente", "campo": "¿Economic Buyer Presente?", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si el CE estuvo presente.", "notas": "", "opts": ["Sí", "No — representado por Campeón", "No — sin representación ejecutiva"]},
        {"id": "objeciones_manejo", "campo": "Objeciones Planteadas y Manejo", "tipo": "Texto largo", "req": False, "ayuda": "Objeciones y respuestas.", "notas": ""},
        {"id": "estado_negociacion", "campo": "Estado de la Negociación", "tipo": "Lista (picklist)", "req": True, "ayuda": "Estado de negociación.", "notas": "", "opts": ["Sin objeciones — procede", "Objeciones menores en proceso", "Objeciones mayores — negociación activa", "Stall — cliente no responde"]}
      ]
    }, {
      "nombre": "Forecast",
      "campos": [
        {"id": "forecast_categoria", "campo": "Categoría de Forecast BSV", "tipo": "Lista (picklist)", "req": True, "ayuda": "Categoría de forecast BSV.", "notas": "Commit requiere CE + Campeón + Paper Process", "opts": ["Pipeline", "Best Case", "Commit Candidato", "Commit", "Ganado — PO emitida"]}
      ]
    }, {
      "nombre": "Resultado Final",
      "campos": [
        {"id": "resultado_final", "campo": "Resultado Final de la Oportunidad", "tipo": "Lista (picklist)", "req": True, "ayuda": "Resultado final.", "notas": "", "opts": ["GANADO — PO emitida", "GANADO — Contrato firmado", "TICKLER — Postergado con fecha", "PERDIDO — XO definitivo"]},
        {"id": "notas_resultado_final", "campo": "Notas del Resultado Final", "tipo": "Texto largo", "req": False, "ayuda": "Notas y aprendizajes del cierre.", "notas": ""}
      ]
    }]
  }
}

# PLANTILLA HTML SALESFORCE LIGHTNING SYSTEM
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Salesforce — Prospectos BSV</title>
    <style>
        :root {
            --sf-brand: #0176d3;
            --sf-brand-dark: #005fb2;
            --sf-path-blue: #00396b;
            --sf-green: #2e844a;
            --sf-bg: #b0c4df;
            --sf-border: #dddbda;
            --sf-text-main: #181818;
            --sf-text-muted: #514f4d;
            --sf-required: #ea001e;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background-color: var(--sf-bg);
            margin: 0; padding: 0;
            color: var(--sf-text-main);
        }

        .sf-global-header {
            background-color: #ffffff; border-bottom: 1px solid var(--sf-border);
            padding: 0 16px; display: flex; align-items: center; justify-content: space-between; height: 50px;
        }
        .sf-app-launcher { display: flex; align-items: center; gap: 12px; }
        .sf-waffle { display: grid; grid-template-columns: repeat(3, 4px); gap: 3px; cursor: pointer; }
        .sf-waffle div { width: 4px; height: 4px; background-color: var(--sf-brand); border-radius: 1px; }
        .sf-app-title { font-weight: 700; font-size: 16px; color: #181818; }
        
        .sf-nav-menu { display: flex; align-items: center; gap: 4px; height: 100%; }
        .sf-nav-item {
            padding: 0 12px; height: 100%; display: flex; align-items: center;
            font-size: 13px; color: #514f4d; cursor: pointer; text-decoration: none; border-bottom: 3px solid transparent;
        }
        .sf-nav-item.active { color: var(--sf-brand); font-weight: 700; border-bottom-color: var(--sf-brand); }

        .sf-container {
            max-width: 1400px; margin: 12px auto; background: #ffffff;
            border-radius: 6px; box-shadow: 0 2px 8px rgba(0,0,0,0.15);
            overflow: hidden; border: 1px solid var(--sf-border);
        }

        /* VISTA 1: TABLA VISTOS RECIENTEMENTE */
        .sf-list-header {
            padding: 16px 24px; background: #ffffff; border-bottom: 1px solid var(--sf-border);
            display: flex; align-items: center; justify-content: space-between;
        }
        .sf-list-title { display: flex; align-items: center; gap: 12px; }
        .sf-lead-icon {
            width: 32px; height: 32px; background-color: #4bca81; border-radius: 4px;
            display: flex; align-items: center; justify-content: center; color: white; font-size: 18px;
        }
        .sf-btn-nuevo {
            background-color: var(--sf-brand); color: #ffffff; border: 1px solid var(--sf-brand);
            padding: 8px 20px; border-radius: 4px; font-weight: 700; font-size: 13px; cursor: pointer;
        }
        .sf-btn-nuevo:hover { background-color: var(--sf-brand-dark); }
        .sf-btn-sub {
            background-color: #ffffff; color: var(--sf-brand); border: 1px solid var(--sf-border);
            padding: 7px 14px; border-radius: 4px; font-size: 12px; font-weight: 600; cursor: pointer;
        }

        .sf-table { width: 100%; border-collapse: collapse; font-size: 13px; }
        .sf-table th { background-color: #fafafa; border-bottom: 2px solid var(--sf-border); padding: 10px 14px; text-align: left; color: #514f4d; font-weight: 700; }
        .sf-table td { padding: 12px 14px; border-bottom: 1px solid var(--sf-border); color: #181818; }
        .sf-table tr:hover { background-color: #f3f3f3; cursor: pointer; }

        /* VISTA 2: DETALLE / CAPTURA SPLIT */
        .sf-detail-top-bar {
            padding: 12px 24px; background: #ffffff; border-bottom: 1px solid var(--sf-border);
            display: flex; align-items: center; justify-content: space-between;
        }
        .sf-highlights-grid {
            display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px;
            padding: 12px 24px; background-color: #fafafa; border-bottom: 1px solid var(--sf-border);
        }
        .sf-highlight-item span { display: block; font-size: 11px; color: var(--sf-text-muted); font-weight: 600; }
        .sf-highlight-item strong { font-size: 13px; color: var(--sf-text-main); }

        /* ESTILOS DE PESTAÑAS Y BLOQUEO PROGRESIVO DE CHEVRONS */
        .sf-path-bar { display: flex; background-color: #f3f3f3; padding: 8px 16px; border-bottom: 1px solid var(--sf-border); overflow-x: auto; white-space: nowrap; }
        
        .sf-chevron {
            padding: 8px 16px; font-size: 12px; font-weight: 700; text-align: center; border: 1px solid var(--sf-border); border-radius: 4px; margin-right: 6px;
            transition: all 0.2s ease;
        }
        .sf-chevron.active {
            background-color: var(--sf-brand); color: #ffffff; border-color: var(--sf-brand); cursor: pointer;
        }
        .sf-chevron.completed {
            background-color: var(--sf-green); color: #ffffff; border-color: var(--sf-green); cursor: pointer;
        }
        .sf-chevron.disabled {
            background-color: #e0e0e0; color: #888888; border-color: #cccccc; cursor: not-allowed; opacity: 0.45; pointer-events: none;
        }

        .sf-split-layout { display: grid; grid-template-columns: 7fr 3fr; background-color: #b0c4df; gap: 12px; padding: 12px; }
        @media (max-width: 992px) { .sf-split-layout { grid-template-columns: 1fr; } }

        .sf-main-col { background: #ffffff; border-radius: 4px; border: 1px solid var(--sf-border); }
        .sf-side-col { display: flex; flex-direction: column; gap: 12px; }

        .sf-tabs { display: flex; border-bottom: 1px solid var(--sf-border); background: #ffffff; padding-left: 16px; }
        .sf-tab { padding: 12px 20px; font-size: 13px; font-weight: 600; color: #514f4d; cursor: pointer; border-bottom: 3px solid transparent; }
        .sf-tab.active { color: var(--sf-brand); border-bottom-color: var(--sf-brand); }

        .sf-card-section { background: #ffffff; border: 1px solid var(--sf-border); border-radius: 4px; margin: 16px; overflow: hidden; }
        .sf-card-header { background: #f3f3f3; padding: 10px 16px; font-size: 13px; font-weight: 700; border-bottom: 1px solid var(--sf-border); }
        .sf-field-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px 24px; padding: 16px; }

        .sf-field-group { display: flex; flex-direction: column; }
        .sf-label { font-size: 12px; font-weight: 600; color: #514f4d; margin-bottom: 4px; }
        .sf-req { color: var(--sf-required); font-weight: bold; }
        .sf-input, .sf-select, .sf-textarea { padding: 7px 10px; border: 1px solid var(--sf-border); border-radius: 4px; font-size: 13px; box-sizing: border-box; width: 100%; }
        .sf-help-text { font-size: 11px; color: #514f4d; margin-top: 3px; }

        .sf-side-card { background: #ffffff; border: 1px solid var(--sf-border); border-radius: 4px; padding: 14px; }
        .sf-drop-box { border: 2px dashed var(--sf-border); border-radius: 4px; padding: 20px; text-align: center; background: #fafafa; margin-top: 8px; }

        .alert-success { background-color: #d4edda; color: #155724; padding: 12px 16px; border-radius: 4px; margin: 12px; border: 1px solid #c3e6cb; font-size: 13px; font-weight:600; }
    </style>
</head>
<body>

<div class="sf-global-header">
    <div class="sf-app-launcher">
        <div class="sf-waffle"><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div></div>
        <span class="sf-app-title">Ventas</span>
    </div>
    <div class="sf-nav-menu">
        <a class="sf-nav-item">Inicio</a>
        <a class="sf-nav-item">Cuentas ▾</a>
        <a class="sf-nav-item">Contactos ▾</a>
        <a class="sf-nav-item active">Prospectos ▾</a>
        <a class="sf-nav-item">Oportunidades ▾</a>
        <a class="sf-nav-item">Tareas ▾</a>
        <a class="sf-nav-item">Calendario</a>
        <a class="sf-nav-item">Reportes ▾</a>
    </div>
    <div style="font-size:12px; color:#514f4d;">Salesforce CRM BSV</div>
</div>

<div class="sf-container">

    <!-- VISTA 1: TABLA VISTOS RECIENTEMENTE -->
    <div id="vista-lista" style="display: {% if mostrar_detalle %}none{% else %}block{% endif %};">
        <div class="sf-list-header">
            <div class="sf-list-title">
                <div class="sf-lead-icon">★</div>
                <div>
                    <span style="font-size:11px; color:#514f4d; font-weight:600;">Prospectos</span>
                    <h2 style="margin:0; font-size:18px;">Vistos recientemente ▾</h2>
                </div>
            </div>
            <div>
                <button class="sf-btn-nuevo" onclick="abrirNuevoFormulario()">+ Nuevo</button>
            </div>
        </div>

        <div style="overflow-x: auto;">
            <table class="sf-table">
                <thead>
                    <tr>
                        <th width="30"><input type="checkbox"></th>
                        <th>Nombre completo</th>
                        <th>Cargo</th>
                        <th>Compañía</th>
                        <th>Teléfono</th>
                        <th>Email</th>
                        <th>Estado de prospecto</th>
                    </tr>
                </thead>
                <tbody>
                    {% if registros %}
                        {% for r in registros %}
                            <tr>
                                <td><input type="checkbox"></td>
                                <td><a style="color:var(--sf-brand); font-weight:600;">{{ r.nombre }} {{ r.apellidos }}</a></td>
                                <td>{{ r.cargo or '—' }}</td>
                                <td>{{ r.empresa or '—' }}</td>
                                <td>{{ r.telefono or '—' }}</td>
                                <td>{{ r.email or '—' }}</td>
                                <td><span style="background:#eef4fe; color:#0176d3; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600;">Nuevo</span></td>
                            </tr>
                        {% endfor %}
                    {% else %}
                        <tr>
                            <td colspan="7" style="text-align: center; padding: 35px; color: #514f4d;">
                                No hay prospectos registrados aún. Haz clic en el botón <strong>"+ Nuevo"</strong> para agregar un registro.
                            </td>
                        </tr>
                    {% endif %}
                </tbody>
            </table>
        </div>
    </div>

    <!-- VISTA 2: FORMULARIO Y DETALLE SPLIT (70% / 30%) -->
    <div id="vista-detalle" style="display: {% if mostrar_detalle %}block{% else %}none{% endif %};">
        
        <div class="sf-detail-top-bar">
            <div style="display:flex; align-items:center; gap:12px;">
                <button class="sf-btn-sub" onclick="volverALista()">← Volver a la Lista</button>
                <div class="sf-lead-icon">★</div>
                <div>
                    <span style="font-size:11px; color:#514f4d; font-weight:600;" id="header-objeto-fase">Objeto SF: Lead | Fase: BD → MO</span>
                    <h1 id="dyn-lead-title" style="margin:0; font-size:18px;">— Sin registrar —</h1>
                </div>
            </div>
            <div>
                <button class="sf-btn-sub">+ Seguir</button>
            </div>
        </div>

        <div class="sf-highlights-grid">
            <div class="sf-highlight-item">
                <span>Empresa / Razón Social</span>
                <strong id="dyn-empresa">—</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Cargo</span>
                <strong id="dyn-cargo">—</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Teléfono</span>
                <strong id="dyn-telefono">—</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Email</span>
                <strong id="dyn-email">—</strong>
            </div>
        </div>

        <!-- BARRA DE CHEVRONS CON LÓGICA DE BLOQUEO PROGRESIVO -->
        <div class="sf-path-bar">
            {% for tg_key in tg_keys %}
                {% set tg_idx = loop.index0 %}
                <div class="sf-chevron {% if tg_key == active_tg %}active{% elif tg_idx < unlocked_idx %}completed{% else %}disabled{% endif %}"
                     id="tab-btn-{{ tg_key }}"
                     onclick="{% if tg_idx <= unlocked_idx %}activarTollgate('{{ tg_key }}', {{ tg_idx }}){% else %}return false;{% endif %}">
                    {% if tg_idx < unlocked_idx %}✓ {% endif %}{{ tg_key }}
                </div>
            {% endfor %}
        </div>

        <!-- FORMULARIO PRINCIPAL -->
        <form method="POST" id="form-prospecto">
            <input type="hidden" name="current_active_tg" id="current_active_tg" value="{{ active_tg }}">
            <input type="hidden" name="unlocked_idx" id="unlocked_idx" value="{{ unlocked_idx }}">

            <div class="sf-split-layout">
                
                <div class="sf-main-col">
                    <div class="sf-tabs">
                        <div class="sf-tab active">Detalles de Captura</div>
                        <div class="sf-tab">Actividad</div>
                        <div class="sf-tab">Chatter</div>
                    </div>

                    {% if mensaje %}
                        <div class="alert-success">{{ mensaje }}</div>
                    {% endif %}

                    {% for tg_key, tg_val in tg_data.items() %}
                        <div id="pantalla-{{ tg_key }}" class="tg-screen" style="{% if tg_key != active_tg %}display:none;{% endif %}">
                            {% for seccion in tg_val['secciones'] %}
                                <div class="sf-card-section">
                                    <div class="sf-card-header">▼ {{ seccion['nombre'] }}</div>
                                    <div class="sf-field-grid">
                                        {% for field in seccion['campos'] %}
                                            <div class="sf-field-group" {% if field['tipo'] == 'Texto largo' %}style="grid-column: span 2;"{% endif %}>
                                                <div class="sf-label">
                                                    {% if field['req'] %}<span class="sf-req">*</span>{% endif %}{{ field['campo'] }}
                                                </div>

                                                {% if field['tipo'] == 'Lista (picklist)' %}
                                                    <select name="{{ field['id'] }}" class="sf-select" {% if field['req'] and tg_key == active_tg %}required{% endif %}>
                                                        <option value="">--Seleccione {{ field['campo'] }}--</option>
                                                        {% if field['opts'] %}
                                                            {% for opt in field['opts'] %}
                                                                <option value="{{ opt }}">{{ opt }}</option>
                                                            {% endfor %}
                                                        {% endif %}
                                                    </select>
                                                {% elif field['tipo'] == 'Texto largo' %}
                                                    <textarea name="{{ field['id'] }}" class="sf-textarea" rows="3" {% if field['req'] and tg_key == active_tg %}required{% endif %}></textarea>
                                                {% elif field['tipo'] == 'Fecha' %}
                                                    <input type="date" name="{{ field['id'] }}" class="sf-input" {% if field['req'] and tg_key == active_tg %}required{% endif %}>
                                                {% elif field['tipo'] == 'Email' %}
                                                    <input type="email" id="input-{{ field['id'] }}" name="{{ field['id'] }}" class="sf-input" {% if field['req'] and tg_key == active_tg %}required{% endif %} oninput="actualizarHighlights()">
                                                {% elif field['tipo'] == 'Teléfono' %}
                                                    <input type="tel" id="input-{{ field['id'] }}" name="{{ field['id'] }}" class="sf-input" {% if field['req'] and tg_key == active_tg %}required{% endif %} oninput="actualizarHighlights()">
                                                {% elif field['tipo'] in ['Número', 'Porcentaje (%)', 'Moneda ($)'] %}
                                                    <input type="number" step="any" name="{{ field['id'] }}" class="sf-input" {% if field['req'] and tg_key == active_tg %}required{% endif %}>
                                                {% else %}
                                                    <input type="text" id="input-{{ field['id'] }}" name="{{ field['id'] }}" class="sf-input" {% if field['req'] and tg_key == active_tg %}required{% endif %} oninput="actualizarHighlights()">
                                                {% endif %}

                                                {% if field['ayuda'] %}
                                                    <span class="sf-help-text">{{ field['ayuda'] }}</span>
                                                {% endif %}
                                            </div>
                                        {% endfor %}
                                    </div>
                                </div>
                            {% endfor %}
                        </div>
                    {% endfor %}

                    <!-- BOTÓN PRINCIPAL DE AVANZAR -->
                    <div style="padding: 16px; text-align: right; background: #ffffff; border-top: 1px solid var(--sf-border);">
                        <button type="button" class="sf-btn-sub" onclick="volverALista()" style="margin-right: 8px;">Cancelar</button>
                        <button type="submit" class="sf-btn-nuevo">
                            {% if active_tg == 'TG13' %}Avanzar y Finalizar{% else %}Avanzar ➔{% endif %}
                        </button>
                    </div>
                </div>

                <!-- BARRA LATERAL DERECHA (30%) -->
                <div class="sf-side-col">
                    <div class="sf-side-card" style="background:#fff3cd; border-color:#ffeeba;">
                        <span style="font-size:12px; color:#856404; font-weight:600;">⚠ Verificación</span>
                        <p style="margin:4px 0 0 0; font-size:12px; color:#856404;">No encontramos duplicados potenciales de este Prospecto.</p>
                    </div>

                    <div class="sf-side-card">
                        <div style="font-size:13px; font-weight:700; margin-bottom:8px;">Vínculos rápidos</div>
                        <ul style="margin:0; padding-left:16px; font-size:12px; color:var(--sf-brand);">
                            <li style="margin-bottom:4px;"><a>Historial de aprobaciones (0)</a></li>
                            <li style="margin-bottom:4px;"><a>Archivos (0)</a></li>
                            <li><a>Notas (0)</a></li>
                        </ul>
                    </div>

                    <div class="sf-side-card">
                        <div style="display:flex; justify-content:space-between; align-items:center;">
                            <span style="font-size:13px; font-weight:700;">Archivos (0)</span>
                            <button type="button" class="sf-btn-sub" style="padding:2px 8px; font-size:11px;">Cargar</button>
                        </div>
                        <div class="sf-drop-box">
                            <span style="font-size:12px; color:#514f4d;">Suelte archivos aquí</span>
                        </div>
                    </div>
                </div>

            </div>
        </form>
    </div>

</div>

<script>
    const tgMetadatos = {{ tg_meta_json|safe }};
    const unlockedIndexGlobal = {{ unlocked_idx }};

    function abrirNuevoFormulario() {
        document.getElementById('form-prospecto').reset();
        document.getElementById('unlocked_idx').value = 0;
        document.getElementById('current_active_tg').value = 'TG0';
        actualizarHighlights();
        document.getElementById('vista-lista').style.display = 'none';
        document.getElementById('vista-detalle').style.display = 'block';
        activarTollgate('TG0', 0);
    }

    function volverALista() {
        document.getElementById('vista-detalle').style.display = 'none';
        document.getElementById('vista-lista').style.display = 'block';
    }

    function activarTollgate(tgId, tgIdx) {
        if (tgIdx > unlockedIndexGlobal) {
            return false; // Bloqueado: no se puede avanzar sin presionar 'Avanzar'
        }

        document.querySelectorAll('.tg-screen').forEach(el => el.style.display = 'none');
        document.querySelectorAll('.sf-chevron').forEach(el => el.classList.remove('active'));

        const pantallaTarget = document.getElementById('pantalla-' + tgId);
        const tabTarget = document.getElementById('tab-btn-' + tgId);

        if (pantallaTarget) pantallaTarget.style.display = 'block';
        if (tabTarget) tabTarget.classList.add('active');

        document.getElementById('current_active_tg').value = tgId;

        if (tgMetadatos[tgId]) {
            document.getElementById('header-objeto-fase').innerText = 'Objeto SF: ' + tgMetadatos[tgId].objeto + ' | Fase: ' + tgMetadatos[tgId].fase;
        }
    }

    function actualizarHighlights() {
        const nombreElem = document.getElementById('input-nombre');
        const apellidosElem = document.getElementById('input-apellidos');
        const empresaElem = document.getElementById('input-empresa');
        const cargoElem = document.getElementById('input-cargo');
        const telefonoElem = document.getElementById('input-telefono');
        const emailElem = document.getElementById('input-email');

        const nombre = nombreElem ? nombreElem.value.trim() : '';
        const apellidos = apellidosElem ? apellidosElem.value.trim() : '';
        const empresa = empresaElem ? empresaElem.value.trim() : '';
        const cargo = cargoElem ? cargoElem.value.trim() : '';
        const telefono = telefonoElem ? telefonoElem.value.trim() : '';
        const email = emailElem ? emailElem.value.trim() : '';

        const nombreCompleto = (nombre || apellidos) ? (nombre + ' ' + apellidos) : '— Sin registrar —';
        document.getElementById('dyn-lead-title').innerText = nombreCompleto;
        document.getElementById('dyn-empresa').innerText = empresa || '—';
        document.getElementById('dyn-cargo').innerText = cargo || '—';
        document.getElementById('dyn-telefono').innerText = telefono || '—';
        document.getElementById('dyn-email').innerText = email || '—';
    }
</script>

</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def home():
    mensaje = None
    mostrar_detalle = False
    active_tg = 'TG0'
    unlocked_idx = 0
    tg_keys = list(TOLLGATES_DATA.keys())

    if request.method == 'POST':
        mostrar_detalle = True
        current_active_tg = request.form.get('current_active_tg', 'TG0')
        unlocked_idx = int(request.form.get('unlocked_idx', 0))

        nombre = request.form.get('nombre', '')
        apellidos = request.form.get('apellidos', '')
        empresa = request.form.get('empresa', '')
        cargo = request.form.get('cargo', '')
        email = request.form.get('email', '')
        telefono = request.form.get('telefono', '')

        # 1. GUARDADO EN LA BASE DE DATOS LOCAL
        nuevo_registro = {
            'nombre': nombre,
            'apellidos': apellidos,
            'empresa': empresa,
            'cargo': cargo,
            'email': email,
            'telefono': telefono
        }
        REGISTROS_PROSPECTOS.append(nuevo_registro)

        # 2. GUARDADO EN SHAREPOINT
        try:
            from office365.runtime.auth.user_credential import UserCredential
            from office365.sharepoint.client_context import ClientContext
            ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
            target_list = ctx.web.lists.get_by_title("BSV_Leads")
            
            target_list.add_item({
                "Title": f"{nombre} {apellidos}".strip(),
                "BSV_Empresa___Razon_Social__c": empresa,
                "BSV_Cargo___Titulo__c": cargo,
                "BSV_Email_Corporativo__c": email,
                "BSV_Telefono_Contacto__c": telefono
            })
            ctx.execute_query()
            mensaje = f"¡Datos de {current_active_tg} guardados en SharePoint correctamente!"
        except Exception as e:
            mensaje = f"¡Datos de {current_active_tg} guardados localmente! (Avanzando al siguiente Tollgate)"

        # 3. LÓGICA DE AVANCE Y DESBLOQUEO AL SIGUIENTE TOLLGATE
        current_idx = tg_keys.index(current_active_tg) if current_active_tg in tg_keys else 0
        if current_idx < len(tg_keys) - 1:
            next_idx = current_idx + 1
            unlocked_idx = max(unlocked_idx, next_idx)
            active_tg = tg_keys[next_idx]
        else:
            active_tg = current_active_tg
            mostrar_detalle = False
            mensaje = f"¡Captura completa de los 14 Tollgates finalizada exitosamente para {nombre} {apellidos}!"

    tg_meta_json = json.dumps({k: {"objeto": v["objeto"], "fase": v["fase"]} for k, v in TOLLGATES_DATA.items()})

    return render_template_string(
        HTML_TEMPLATE,
        tg_data=TOLLGATES_DATA,
        tg_keys=tg_keys,
        active_tg=active_tg,
        unlocked_idx=unlocked_idx,
        mostrar_detalle=mostrar_detalle,
        tg_meta_json=tg_meta_json,
        registros=REGISTROS_PROSPECTOS,
        mensaje=mensaje
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
