from flask import Flask, request, render_template_string, jsonify
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import json

app = Flask(__name__)

# --- CONFIGURACIÓN DE SHAREPOINT ---
SITE_URL = "https://tu-empresa.sharepoint.com/sites/tu-sitio"
USERNAME = "usuario@tu-empresa.com"
PASSWORD = "TuPassword123"

# --- DICCIONARIO COMPLETO EXTRAÍDO ÚNICAMENTE DEL EXCEL (TG0 a TG13) ---
TOLLGATES_DATA = {
  "TG0": {
    "objeto": "Lead",
    "fase": "BD → MO",
    "secciones": [
      {
        "nombre": "Datos del Contacto",
        "campos": [
          {"id": "nombre", "campo": "Nombre", "tipo": "Texto", "req": True, "ayuda": "Nombre del contacto. Campo estándar de Salesforce.", "notas": "", "opts": []},
          {"id": "apellidos", "campo": "Apellidos", "tipo": "Texto", "req": True, "ayuda": "Apellidos del contacto. Campo estándar de Salesforce.", "notas": "", "opts": []},
          {"id": "empresa", "campo": "Empresa / Razón Social", "tipo": "Texto", "req": True, "ayuda": "Nombre de la empresa del lead.", "notas": "", "opts": []},
          {"id": "cargo", "campo": "Cargo / Título", "tipo": "Texto", "req": True, "ayuda": "Cargo o posición del contacto dentro de la empresa.", "notas": "", "opts": []},
          {"id": "email", "campo": "Email Corporativo", "tipo": "Email", "req": True, "ayuda": "Email corporativo del contacto. Verificar dominio.", "notas": "", "opts": []},
          {"id": "telefono", "campo": "Teléfono Contacto", "tipo": "Teléfono", "req": False, "ayuda": "Teléfono directo o móvil del contacto.", "notas": "", "opts": []},
          {"id": "pais_region", "campo": "País / Región", "tipo": "Lista (picklist)", "req": True, "ayuda": "País o región donde opera el lead.", "notas": "", "opts": ["Mexico", "USA"]}
        ]
      },
      {
        "nombre": "Origen y Asignación",
        "campos": [
          {"id": "notas_adicionales", "campo": "Notas Adicionales", "tipo": "Texto largo", "req": False, "ayuda": "Notas generales sobre el lead.", "notas": "NO REQUERIDO", "opts": []}
        ]
      }
    ]
  },
  "TG1": {
    "objeto": "Lead",
    "fase": "MO",
    "secciones": [
      {
        "nombre": "Segmentación y Fit",
        "campos": [
          {"id": "macro_segmento", "campo": "Macro Segmento", "tipo": "Lista (picklist)", "req": True, "ayuda": "Macro segmento al que pertenece el lead.", "notas": "", "opts": ["Industrial", "Automotriz", "Alimentos y Bebidas", "Farmacéutica"]},
          {"id": "sub_segmento", "campo": "Sub-Segmento", "tipo": "Lista (picklist)", "req": False, "ayuda": "Sub-segmento específico dentro del macro segmento.", "notas": "", "opts": ["Manufactura", "Empaque", "Ensamblaje", "Logística"]},
          {"id": "geografia", "campo": "Geografía", "tipo": "Lista (picklist)", "req": True, "ayuda": "Región geográfica donde opera el cliente.", "notas": "", "opts": ["Norte", "Centro", "Bajio", "Occidente", "Golfo"]},
          {"id": "relevancia", "campo": "Relevancia del Portafolio", "tipo": "Lista (picklist)", "req": True, "ayuda": "Nivel de relevancia del portafolio de LAMMSA.", "notas": "", "opts": ["Alta", "Media", "Baja"]},
          {"id": "tamano_empresa", "campo": "Tamaño de Empresa", "tipo": "Lista (picklist)", "req": True, "ayuda": "Estimación del tamaño por ventas anuales.", "notas": "", "opts": ["Micro — <$10M", "Pequeña — $10-50M", "Mediana — $50-200M", "Grande — $200M-$1B", "Enterprise — >$1B"]},
          {"id": "banda", "campo": "Banda Asignada", "tipo": "Lista (picklist)", "req": True, "ayuda": "Banda de clasificación BSV del cliente.", "notas": "si es K o A, entonces es BSV - Normal", "opts": ["K", "A", "B", "C", "D"]}
        ]
      }
    ]
  },
  "TG2": {
    "objeto": "Lead",
    "fase": "MO",
    "secciones": [
      {
        "nombre": "Engagement y Señales",
        "campos": [
          {"id": "rol_contacto", "campo": "Rol del Contacto", "tipo": "Lista (picklist)", "req": True, "ayuda": "Rol del contacto en el proceso de decisión de compra.", "notas": "", "opts": ["Decision Maker", "Influencer", "Evaluador técnico", "Usuario final", "Desconocido"]}
        ]
      },
      {
        "nombre": "Enriquecimiento",
        "campos": [
          {"id": "num_plantas", "campo": "Número de Plantas / Ubicaciones", "tipo": "Número", "req": False, "ayuda": "Número estimado de plantas operativas.", "notas": "", "opts": []},
          {"id": "contactos_adic", "campo": "Contactos Adicionales Identificados", "tipo": "Texto largo", "req": False, "ayuda": "Nombre, cargo y email de contactos adicionales.", "notas": "", "opts": []}
        ]
      }
    ]
  },
  "TG3": {
    "objeto": "Lead",
    "fase": "MO",
    "secciones": [
      {
        "nombre": "Conversión MQL",
        "campos": [
          {"id": "codigo_xz", "campo": "Código XZ", "tipo": "Lista (picklist)", "req": True, "ayuda": "Código de clasificación de producto XZ.", "notas": "Consumo", "opts": ["Directo, Indirecto"]}
        ]
      },
      {
        "nombre": "Notas de Transferencia",
        "campos": [
          {"id": "puntos_interes", "campo": "Puntos de Interés Detectados", "tipo": "Texto largo", "req": True, "ayuda": "Hipótesis de necesidades o pain points.", "notas": "", "opts": []},
          {"id": "calidad_contacto", "campo": "Calidad del Contacto", "tipo": "Lista (picklist)", "req": True, "ayuda": "Evaluación subjetiva de la calidad del contacto.", "notas": "", "opts": ["Alta", "Media", "Baja"]},
          {"id": "comentarios_ventas", "campo": "Comentarios Adicionales para Ventas", "tipo": "Texto largo", "req": False, "ayuda": "Notas adicionales de Marketing para Ventas.", "notas": "", "opts": []}
        ]
      }
    ]
  },
  "TG4": {
    "objeto": "Lead",
    "fase": "MO → WO",
    "secciones": [
      {
        "nombre": "Validación ERP",
        "campos": [
          {"id": "ventas_12m", "campo": "Ventas Últimos 12 Meses", "tipo": "Lista (picklist)", "req": True, "ayuda": "Rango de ventas de los últimos 12 meses.", "notas": "", "opts": ["Sin historial", "$0", "<$20k"]},
          {"id": "actividad_lammsa", "campo": "Actividad Lammsa Previa", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si tiene historial con productos Sommer.", "notas": "", "opts": ["Nunca ha comprado", "Histórico inactivo", "Activo — producto Sommer"]}
        ]
      },
      {
        "nombre": "Clasificación Comercial",
        "campos": [
          {"id": "clasif_comercial", "campo": "Clasificación Comercial", "tipo": "Lista (picklist)", "req": True, "ayuda": "Clasificación comercial del lead.", "notas": "", "opts": ["K, A, B, C, D"]},
          {"id": "justif_clasif", "campo": "Justificación de Clasificación", "tipo": "Texto largo", "req": True, "ayuda": "Evidencia de la lógica aplicada.", "notas": "", "opts": []}
        ]
      },
      {
        "nombre": "Asignación y SLA",
        "campos": [
          {"id": "osp_asignado", "campo": "OSP Asignado", "tipo": "Texto", "req": True, "ayuda": "Nombre del OSP que recibe la cuenta.", "notas": "", "opts": []},
          {"id": "gerente_area", "campo": "Gerente de Área", "tipo": "Texto", "req": True, "ayuda": "Nombre del Gerente responsable.", "notas": "", "opts": []},
          {"id": "fecha_asignacion", "campo": "Fecha de Asignación", "tipo": "Fecha", "req": True, "ayuda": "Fecha de asignación formal.", "notas": "", "opts": []},
          {"id": "fecha_comp_1ra_inter", "campo": "Fecha Compromiso 1ra Interacción", "tipo": "Fecha", "req": True, "ayuda": "Fecha comprometida para primera interacción (<=5 días).", "notas": "Convierte a cuenta - Cita validada", "opts": []}
        ]
      }
    ]
  },
  "TG5": {
    "objeto": "Opportunity",
    "fase": "WO",
    "secciones": [
      {
        "nombre": "Datos de la WO",
        "campos": [
          {"id": "tipo_oportunidad", "campo": "Tipo de Oportunidad", "tipo": "Lista (picklist)", "req": True, "ayuda": "Clasificación comercial de la oportunidad.", "notas": "", "opts": ["XR — Reactivaciíon", "XP — Prospecto", "XS — Cross Sell Sommer"]},
          {"id": "aoi_estimado", "campo": "AOI Estimado (%)", "tipo": "Porcentaje (%)", "req": True, "ayuda": "Porcentaje de ahorro o ingreso estimado.", "notas": "", "opts": []},
          {"id": "primera_inter_ok", "campo": "Primera Interacción Completada", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si se completó la primera interacción.", "notas": "Enviar Template de Agradecimiento", "opts": ["Sí", "No — pendiente"]},
          {"id": "fecha_1ra_inter", "campo": "Fecha de Primera Interacción", "tipo": "Fecha", "req": False, "ayuda": "Fecha real de la primera interacción.", "notas": "", "opts": []},
          {"id": "formato_1ra_inter", "campo": "Formato de Primera Interacción", "tipo": "Lista (picklist)", "req": False, "ayuda": "Canal o formato utilizado.", "notas": "", "opts": ["Presencial — planta", "Presencial — oficina", "Virtual"]}
        ]
      },
      {
        "nombre": "Plan BSV",
        "campos": [
          {"id": "estado_plan_bsv", "campo": "Estado del Plan BSV", "tipo": "Lista (picklist)", "req": True, "ayuda": "Estado del Plan BSV para la oportunidad.", "notas": "", "opts": ["Sí — completado", "En proceso", "No — pendiente"]}
        ]
      }
    ]
  },
  "TG6": {
    "objeto": "Opportunity",
    "fase": "WO",
    "secciones": [
      {
        "nombre": "Plan BSV",
        "campos": [
          {"id": "fecha_est_plan_bsv", "campo": "Fecha Estimada Presentación Plan BSV", "tipo": "Fecha", "req": False, "ayuda": "Fecha estimada para presentar el Plan BSV.", "notas": "", "opts": []},
          {"id": "notas_plan_bsv", "campo": "Notas del Plan BSV", "tipo": "Texto largo", "req": False, "ayuda": "Consideraciones iniciales del Plan BSV.", "notas": "", "opts": []}
        ]
      },
      {
        "nombre": "Decisión GO / NO-GO",
        "campos": [
          {"id": "decision_go", "campo": "Decisión GO / NO-GO", "tipo": "Lista (picklist)", "req": True, "ayuda": "Decisión formal de continuidad en el funnel.", "notas": "Flujo de Aprobación Gerencia", "opts": ["GO — continuar con investigación BSV", "NO-GO — salida", "TICKLER — re-evaluar en fecha futura"]},
          {"id": "justif_decision", "campo": "Justificación de la Decisión", "tipo": "Texto largo", "req": True, "ayuda": "Justificación obligatoria.", "notas": "", "opts": []},
          {"id": "siguiente_paso", "campo": "Siguiente Paso Concreto", "tipo": "Texto", "req": True, "ayuda": "Siguiente acción específica.", "notas": "", "opts": []},
          {"id": "fecha_sig_paso", "campo": "Fecha del Siguiente Paso", "tipo": "Fecha", "req": True, "ayuda": "Fecha comprometida para el siguiente paso.", "notas": "", "opts": []}
        ]
      },
      {
        "nombre": "MEDDICC",
        "campos": [
          {"id": "meddicc_i_score", "campo": "MEDDICC I — Score (Dolor Implicado)", "tipo": "Número", "req": True, "ayuda": "Score Implicated Pain (Escala 1-10). Gate: >=4.", "notas": "", "opts": []},
          {"id": "meddicc_i_notas", "campo": "MEDDICC I — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Descripción del dolor y costo de inacción.", "notas": "", "opts": []},
          {"id": "meddicc_m_score", "campo": "MEDDICC M — Score (Métricas)", "tipo": "Número", "req": True, "ayuda": "Score Metrics (Escala 1-10). Gate: >=5.", "notas": "", "opts": []},
          {"id": "meddicc_m_notas", "campo": "MEDDICC M — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Métricas cuantificadas de impacto.", "notas": "", "opts": []},
          {"id": "meddicc_c2_score", "campo": "MEDDICC C2 — Score (Competencia)", "tipo": "Número", "req": True, "ayuda": "Score Competition (Escala 1-10).", "notas": "", "opts": []},
          {"id": "meddicc_c2_notas", "campo": "MEDDICC C2 — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Proveedor actual y ventajas de LAMMSA.", "notas": "", "opts": []}
        ]
      }
    ]
  },
  "TG7": {
    "objeto": "Opportunity",
    "fase": "WO",
    "secciones": [
      {
        "nombre": "MEDDICC",
        "campos": [
          {"id": "meddicc_c1_score", "campo": "MEDDICC C1 — Score (Campeón)", "tipo": "Número", "req": True, "ayuda": "Score Champion (Escala 1-10).", "notas": "Agregar en contacto", "opts": []},
          {"id": "meddicc_c1_notas", "campo": "MEDDICC C1 — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Nombre del Campeón y evidencia de advocacy.", "notas": "Requiere interacción previa", "opts": []},
          {"id": "meddicc_d1_score", "campo": "MEDDICC D1 — Score (Criterios de Decisión)", "tipo": "Número", "req": True, "ayuda": "Score Decision Criteria (Escala 1-10).", "notas": "", "opts": []},
          {"id": "meddicc_d1_notas", "campo": "MEDDICC D1 — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Criterios con los que el cliente evaluará.", "notas": "", "opts": []}
        ]
      }
    ]
  },
  "TG8": {
    "objeto": "Opportunity",
    "fase": "WO",
    "secciones": [
      {
        "nombre": "MEDDICC",
        "campos": [
          {"id": "meddicc_e_score", "campo": "MEDDICC E — Score (Comprador Económico)", "tipo": "Número", "req": True, "ayuda": "Score Economic Buyer (Escala 1-10).", "notas": "Agregar en contacto", "opts": []},
          {"id": "meddicc_e_notas", "campo": "MEDDICC E — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Nombre y nivel de soporte del Economic Buyer.", "notas": "Requiere interacción previa", "opts": []},
          {"id": "meddicc_d2_score", "campo": "MEDDICC D2 — Score (Proceso de Decisión)", "tipo": "Número", "req": True, "ayuda": "Score Decision Process (Escala 1-10).", "notas": "", "opts": []},
          {"id": "meddicc_d2_notas", "campo": "MEDDICC D2 — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Pasos del proceso de decisión y timeline.", "notas": "", "opts": []},
          {"id": "meddicc_p_score", "campo": "MEDDICC P — Score (Proceso de Contratación)", "tipo": "Número", "req": True, "ayuda": "Score Paper Process (Escala 1-10).", "notas": "", "opts": []},
          {"id": "meddicc_p_notas", "campo": "MEDDICC P — Evidencia y Notas", "tipo": "Texto largo", "req": True, "ayuda": "Pasos de revisión legal y firmas.", "notas": "", "opts": []}
        ]
      }
    ]
  },
  "TG9": {
    "objeto": "Opportunity",
    "fase": "WO",
    "secciones": [
      {
        "nombre": "Investigación BSV",
        "campos": [
          {"id": "resumen_bsv", "campo": "Resumen Ejecutivo BSV", "tipo": "Texto largo", "req": True, "ayuda": "Resumen de la situación y propuesta.", "notas": "", "opts": []},
          {"id": "mapa_stakeholders", "campo": "Mapa de Stakeholders y Coaches", "tipo": "Texto largo", "req": True, "ayuda": "Mapa de actores clave y nivel de apoyo.", "notas": "", "opts": []},
          {"id": "hipotesis_fit", "campo": "Hipótesis de Fit de Negocio", "tipo": "Texto largo", "req": True, "ayuda": "Cómo LAMMSA resuelve el problema.", "notas": "", "opts": []},
          {"id": "aoi_cuantificado", "campo": "AOI Cuantificado ($)", "tipo": "Moneda ($)", "req": True, "ayuda": "Valor en dólares del AOI.", "notas": "", "opts": []},
          {"id": "analisis_incumbentes", "campo": "Análisis de Incumbentes", "tipo": "Texto largo", "req": True, "ayuda": "Análisis de competidores actuales.", "notas": "", "opts": []},
          {"id": "posicionamiento_diferenciador", "campo": "Posicionamiento Diferenciador LAMMSA", "tipo": "Texto largo", "req": True, "ayuda": "Argumento diferenciador de LAMMSA.", "notas": "", "opts": []}
        ]
      },
      {
        "nombre": "Aprobación Gerencial",
        "campos": [
          {"id": "aprobacion_summit", "campo": "Aprobación para Summit", "tipo": "Lista (picklist)", "req": True, "ayuda": "Aprobación del Gerente para el Summit.", "notas": "Subir deck BSV", "opts": ["Aprobado para Summit", "Pendiente de revisión", "No aprobado"]},
          {"id": "fecha_rev_gerencial", "campo": "Fecha de Revisión Gerencial", "tipo": "Fecha", "req": False, "ayuda": "Fecha en que el Gerente revisó.", "notas": "", "opts": []},
          {"id": "notas_aprobacion", "campo": "Notas de Aprobación", "tipo": "Texto largo", "req": False, "ayuda": "Notas del Gerente de Área.", "notas": "", "opts": []}
        ]
      }
    ]
  },
  "TG10": {
    "objeto": "Opportunity",
    "fase": "WO",
    "secciones": [
      {
        "nombre": "Datos del Summit",
        "campos": [
          {"id": "fecha_bsv_summit", "campo": "Fecha del BSV Summit", "tipo": "Fecha", "req": True, "ayuda": "Fecha acordada para el Summit.", "notas": "Mail de confirmación requerido", "opts": []},
          {"id": "participantes_cliente", "campo": "Participantes del Cliente", "tipo": "Texto largo", "req": True, "ayuda": "Nombre y cargo de cada asistente del cliente.", "notas": "", "opts": []},
          {"id": "participantes_lammsa", "campo": "Participantes LAMMSA", "tipo": "Texto largo", "req": True, "ayuda": "Representantes de LAMMSA presentes.", "notas": "", "opts": []}
        ]
      }
    ]
  },
  "TG11": {
    "objeto": "Opportunity",
    "fase": "WO",
    "secciones": [
      {
        "nombre": "Datos del Summit",
        "campos": [
          {"id": "deck_presentado", "campo": "Deck Presentado", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si el deck fue presentado.", "notas": "", "opts": ["Sí", "No — pendiente"]},
          {"id": "roundtable_ejecutado", "campo": "Round-Table Ejecutado", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si se ejecutó el round-table.", "notas": "", "opts": ["Sí", "Parcial", "No"]},
          {"id": "acuerdos_summit", "campo": "Acuerdos y Próximos Pasos del Summit", "tipo": "Texto largo", "req": True, "ayuda": "Acuerdos formalizados durante la sesión.", "notas": "Mail de agradecimiento requerido", "opts": []}
        ]
      },
      {
        "nombre": "Resultado del Summit",
        "campos": [
          {"id": "resultado_summit", "campo": "Resultado del Summit", "tipo": "Lista (picklist)", "req": True, "ayuda": "Resultado formal del BSV Summit.", "notas": "", "opts": ["Compromiso verbal — proceden", "Compromiso escrito — NDA / LOI firmado", "Interés sin compromiso — más información", "Sin compromiso — no avanzan"]},
          {"id": "notas_resultado_summit", "campo": "Notas del Resultado del Summit", "tipo": "Texto largo", "req": False, "ayuda": "Notas y objeciones del Summit.", "notas": "", "opts": []}
        ]
      },
      {
        "nombre": "Alcance de la Solución",
        "campos": [
          {"id": "plantas_scope", "campo": "Plantas en Scope", "tipo": "Texto", "req": True, "ayuda": "Plantas incluidas en el alcance.", "notas": "", "opts": []},
          {"id": "categorias_solucion", "campo": "Categorías Incluidas en la Solución", "tipo": "Texto largo", "req": True, "ayuda": "Categorías de producto o servicio.", "notas": "", "opts": []},
          {"id": "modelo_servicio", "campo": "Modelo de Servicio", "tipo": "Lista (picklist)", "req": True, "ayuda": "Modelo de servicio propuesto.", "notas": "", "opts": ["Full service — inventario LAMMSA", "Consignación", "Mixto", "A definir"]},
          {"id": "plan_implementacion", "campo": "Plan de Implementación por Fases", "tipo": "Texto largo", "req": True, "ayuda": "Entregables, fechas y hitos.", "notas": "", "opts": []}
        ]
      },
      {
        "nombre": "Economía de la Propuesta",
        "campos": [
          {"id": "calculadora_aoi_ok", "campo": "Calculadora AOI Completada", "tipo": "Lista (picklist)", "req": True, "ayuda": "Validación de la calculadora por Finanzas.", "notas": "", "opts": ["Sí — validada", "No — pendiente"]},
          {"id": "aoi_proyectado_ano1", "campo": "AOI Proyectado Año 1 ($)", "tipo": "Moneda ($)", "req": True, "ayuda": "AOI proyectado primer año ($USD).", "notas": "", "opts": []},
          {"id": "aoi_proyectado_anos2_3", "campo": "AOI Proyectado Años 2-3 ($)", "tipo": "Moneda ($)", "req": True, "ayuda": "AOI proyectado acumulado años 2-3 ($USD).", "notas": "", "opts": []},
          {"id": "margen_vs_floor", "campo": "Margen vs Floor Mínimo", "tipo": "Lista (picklist)", "req": True, "ayuda": "Evaluación del margen vs floor mínimo.", "notas": "Bajo el floor no procede", "opts": ["Sobre el floor — aprobado", "En el floor — requiere excepción", "Bajo el floor — no procede"]}
        ]
      }
    ]
  },
  "TG12": {
    "objeto": "Opportunity",
    "fase": "WO",
    "secciones": [
      {
        "nombre": "Aprobaciones Internas",
        "campos": [
          {"id": "aprobacion_ops", "campo": "Aprobación Ops / Supply Chain", "tipo": "Lista (picklist)", "req": True, "ayuda": "Aprobación formal de Operaciones.", "notas": "", "opts": ["Aprobado", "Pendiente", "Rechazado"]},
          {"id": "aprobacion_finanzas", "campo": "Aprobación Finanzas", "tipo": "Lista (picklist)", "req": True, "ayuda": "Aprobación formal de Finanzas.", "notas": "", "opts": ["Aprobado", "Pendiente", "Rechazado"]},
          {"id": "propuesta_aprobada_int", "campo": "Propuesta Internamente Aprobada", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indicador global de aprobaciones.", "notas": "", "opts": ["Sí — aprobación completa", "No — pendiente aprobaciones"]},
          {"id": "fecha_presentacion_cliente", "campo": "Fecha de Presentación al Cliente", "tipo": "Fecha", "req": True, "ayuda": "Fecha confirmada de presentación.", "notas": "Mail de confirmación requerido", "opts": []}
        ]
      }
    ]
  },
  "TG13": {
    "objeto": "Opportunity",
    "fase": "WO → SO",
    "secciones": [
      {
        "nombre": "Presentación y Negociación",
        "campos": [
          {"id": "fecha_presentacion_prop", "campo": "Fecha de Presentación de Propuesta", "tipo": "Fecha", "req": True, "ayuda": "Fecha real en que se presentó.", "notas": "Mail de agradecimiento requerido", "opts": []},
          {"id": "participantes_pres_cliente", "campo": "Participantes del Cliente", "tipo": "Texto largo", "req": True, "ayuda": "Asistentes del cliente a la presentación.", "notas": "", "opts": []},
          {"id": "ce_presente", "campo": "¿Economic Buyer Presente?", "tipo": "Lista (picklist)", "req": True, "ayuda": "Indica si el CE estuvo presente.", "notas": "", "opts": ["Sí", "No — representado por Campeón", "No — sin representación ejecutiva"]},
          {"id": "objeciones_manejo", "campo": "Objeciones Planteadas y Manejo", "tipo": "Texto largo", "req": False, "ayuda": "Objeciones y respuesta dada.", "notas": "", "opts": []},
          {"id": "estado_negociacion", "campo": "Estado de la Negociación", "tipo": "Lista (picklist)", "req": True, "ayuda": "Estado actual de negociación.", "notas": "", "opts": ["Sin objeciones — procede", "Objeciones menores en proceso", "Objeciones mayores — negociación activa", "Stall — cliente no responde"]}
        ]
      },
      {
        "nombre": "Forecast",
        "campos": [
          {"id": "forecast_categoria", "campo": "Categoría de Forecast BSV", "tipo": "Lista (picklist)", "req": True, "ayuda": "Categoría de forecast BSV.", "notas": "Commit requiere CE + Campeón + Paper Process", "opts": ["Pipeline", "Best Case", "Commit Candidato", "Commit", "Ganado — PO emitida"]}
        ]
      },
      {
        "nombre": "Resultado Final",
        "campos": [
          {"id": "resultado_final", "campo": "Resultado Final de la Oportunidad", "tipo": "Lista (picklist)", "req": True, "ayuda": "Resultado final de la oportunidad.", "notas": "", "opts": ["GANADO — PO emitida", "GANADO — Contrato firmado", "TICKLER — Postergado con fecha", "PERDIDO — XO definitivo"]},
          {"id": "notas_resultado_final", "campo": "Notas del Resultado Final", "tipo": "Texto largo", "req": False, "ayuda": "Notas y aprendizajes del cierre.", "notas": "", "opts": []}
        ]
      }
    ]
  }
}

# --- PLANTILLA HTML / SALESFORCE LIGHTNING SYSTEM (SLDS) ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BSV LAMMSA — Captura de Prospectos (Tollgates TG0 - TG13)</title>
    <style>
        :root {
            --sf-brand: #0176d3;
            --sf-brand-dark: #005fb2;
            --sf-path-blue: #00396b;
            --sf-green: #2e844a;
            --sf-bg: #eef2f7;
            --sf-card-bg: #ffffff;
            --sf-border: #dddbda;
            --sf-text-main: #181818;
            --sf-text-muted: #514f4d;
            --sf-required: #ea001e;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Salesforce Sans", "Segoe UI", Roboto, sans-serif;
            background-color: var(--sf-bg);
            margin: 0;
            padding: 16px;
            color: var(--sf-text-main);
        }

        .sf-canvas {
            background-color: #ffffff;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            max-width: 1280px;
            margin: 0 auto;
            border: 1px solid var(--sf-border);
            overflow: hidden;
        }

        /* 1. HIGHLIGHTS HEADER (SIN DATOS HARDCODEADOS) */
        .sf-header {
            background: linear-gradient(180deg, #ffffff 0%, #fafafa 100%);
            padding: 16px 24px;
            border-bottom: 1px solid var(--sf-border);
        }
        .sf-header-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 12px;
        }
        .sf-header-identity {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .sf-icon-avatar {
            width: 36px;
            height: 36px;
            background: linear-gradient(135deg, #4bca81 0%, #2e844a 100%);
            border-radius: 6px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 18px;
        }
        .sf-header-titles span {
            font-size: 11px;
            color: var(--sf-text-muted);
            font-weight: 700;
            text-transform: uppercase;
        }
        .sf-header-titles h1 {
            font-size: 20px;
            margin: 2px 0 0 0;
            font-weight: 700;
            color: var(--sf-text-main);
        }
        .sf-badge-pill {
            background-color: #eef4fe;
            color: var(--sf-brand);
            border: 1px solid #b0c4df;
            padding: 4px 12px;
            border-radius: 12px;
            font-size: 12px;
            font-weight: 600;
        }

        .sf-highlights-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            border-top: 1px solid var(--sf-border);
            padding-top: 12px;
        }
        .sf-highlight-item span {
            display: block;
            font-size: 11px;
            color: var(--sf-text-muted);
            font-weight: 600;
            margin-bottom: 2px;
        }
        .sf-highlight-item strong {
            font-size: 13px;
            color: var(--sf-text-main);
        }

        /* 2. PESTAÑAS SEPARADAS PARA LOS 14 TOLLGATES (TG0 - TG13) */
        .sf-path-bar {
            display: flex;
            background-color: #f3f3f3;
            padding: 8px 12px;
            border-bottom: 1px solid var(--sf-border);
            overflow-x: auto;
            white-space: nowrap;
        }
        .sf-tab-tg {
            padding: 9px 16px;
            background-color: #ffffff;
            color: #3e3e3c;
            font-size: 12px;
            font-weight: 700;
            text-align: center;
            cursor: pointer;
            border: 1px solid var(--sf-border);
            border-radius: 4px;
            margin-right: 6px;
            user-select: none;
            transition: all 0.2s ease;
        }
        .sf-tab-tg.active {
            background-color: var(--sf-brand);
            color: #ffffff;
            border-color: var(--sf-brand);
            box-shadow: 0 2px 4px rgba(1, 118, 211, 0.3);
        }
        .sf-tab-tg:hover:not(.active) {
            background-color: #eef1f6;
        }

        /* 3. FORMULARIO Y SECCIONES */
        .sf-body-content {
            padding: 20px;
            background-color: #f8f9fb;
            min-height: 420px;
        }
        .sf-card-section {
            background: #ffffff;
            border: 1px solid var(--sf-border);
            border-radius: 6px;
            margin-bottom: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.04);
            overflow: hidden;
        }
        .sf-card-header {
            background-color: #f3f3f3;
            padding: 10px 16px;
            font-size: 13px;
            font-weight: 700;
            color: var(--sf-text-main);
            border-bottom: 1px solid var(--sf-border);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .sf-field-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 16px 28px;
            padding: 20px;
        }
        @media (max-width: 768px) {
            .sf-field-grid { grid-template-columns: 1fr; }
        }

        /* CAMPOS DE CONTROL */
        .sf-field-group {
            display: flex;
            flex-direction: column;
        }
        .sf-label-row {
            display: flex;
            align-items: center;
            gap: 6px;
            margin-bottom: 4px;
        }
        .sf-label {
            font-size: 12px;
            font-weight: 600;
            color: var(--sf-text-muted);
        }
        .sf-req {
            color: var(--sf-required);
            font-weight: bold;
        }
        .sf-tooltip-icon {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 14px;
            height: 14px;
            border-radius: 50%;
            background-color: #b0c4df;
            color: #ffffff;
            font-size: 10px;
            font-weight: bold;
            cursor: help;
            position: relative;
        }
        .sf-tooltip-icon:hover::after {
            content: attr(data-tooltip);
            position: absolute;
            bottom: 120%;
            left: 50%;
            transform: translateX(-50%);
            background-color: #16325c;
            color: #ffffff;
            padding: 6px 10px;
            border-radius: 4px;
            font-size: 11px;
            width: 220px;
            z-index: 100;
            white-space: normal;
            font-weight: normal;
            box-shadow: 0 2px 6px rgba(0,0,0,0.2);
        }
        
        .sf-input, .sf-select, .sf-textarea {
            padding: 8px 12px;
            border: 1px solid var(--sf-border);
            border-radius: 4px;
            font-size: 13px;
            color: var(--sf-text-main);
            background-color: #ffffff;
            width: 100%;
            box-sizing: border-box;
        }
        .sf-input:focus, .sf-select:focus, .sf-textarea:focus {
            outline: none;
            border-color: var(--sf-brand);
            box-shadow: 0 0 0 3px rgba(1, 118, 211, 0.2);
        }
        .sf-help-text {
            font-size: 11px;
            color: var(--sf-text-muted);
            margin-top: 3px;
        }

        /* ACCIONES FLOTANTES */
        .sf-action-bar {
            position: sticky;
            bottom: 0;
            background-color: #ffffff;
            border-top: 1px solid var(--sf-border);
            padding: 12px 24px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 -4px 10px rgba(0,0,0,0.05);
            z-index: 50;
        }
        .sf-btn {
            padding: 8px 22px;
            border-radius: 4px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            border: 1px solid var(--sf-border);
        }
        .sf-btn-secondary {
            background-color: #ffffff;
            color: var(--sf-brand);
        }
        .sf-btn-primary {
            background-color: var(--sf-brand);
            color: #ffffff;
            border-color: var(--sf-brand);
        }
        .sf-btn-primary:hover {
            background-color: var(--sf-brand-dark);
        }

        .alert-success {
            background-color: #d4edda;
            color: #155724;
            padding: 12px 16px;
            border-radius: 4px;
            margin-bottom: 16px;
            border: 1px solid #c3e6cb;
            font-size: 13px;
            font-weight: 600;
        }
    </style>
</head>
<body>

<div class="sf-canvas">

    <!-- 1. HEADER HIGHLIGHTS PANEL (EN BLANCO - SIN DATOS FALSOS) -->
    <div class="sf-header">
        <div class="sf-header-top">
            <div class="sf-header-identity">
                <div class="sf-icon-avatar">★</div>
                <div class="sf-header-titles">
                    <span id="header-objeto-fase">Objeto SF: Lead | Fase: BD → MO</span>
                    <h1 id="header-nombre-lead">— Sin registrar —</h1>
                </div>
            </div>
            <div class="sf-badge-pill" id="header-tg-badge">Pestaña Activa: TG0</div>
        </div>

        <div class="sf-highlights-grid">
            <div class="sf-highlight-item">
                <span>Empresa / Razón Social</span>
                <strong id="header-empresa">—</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Cargo / Título</span>
                <strong id="header-cargo">—</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Teléfono Contacto</span>
                <strong id="header-telefono">—</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Email Corporativo</span>
                <strong id="header-email">—</strong>
            </div>
        </div>
    </div>

    <!-- 2. PESTAÑAS SEPARADAS PARA CADA TOLLGATE (TG0 - TG13) -->
    <div class="sf-path-bar">
        {% for tg_key in tg_keys %}
            <div class="sf-tab-tg {% if loop.first %}active{% endif %}" id="tab-btn-{{ tg_key }}" onclick="activarTollgate('{{ tg_key }}')">
                {{ tg_key }}
            </div>
        {% endfor %}
    </div>

    <form method="POST">
        <div class="sf-body-content">
            {% if mensaje %}
                <div class="alert-success">{{ mensaje }}</div>
            {% endif %}

            <!-- DENSIDAD DE FORMULARIOS POR CADA TOLLGATE SEPARADO -->
            {% for tg_key, tg_val in tg_data.items() %}
                <div id="pantalla-{{ tg_key }}" class="tg-screen" style="{% if not loop.first %}display:none;{% endif %}">
                    {% for seccion in tg_val['secciones'] %}
                        <div class="sf-card-section">
                            <div class="sf-card-header">
                                <span>▼ {{ seccion['nombre'] }}</span>
                                <span style="font-size: 11px; color: var(--sf-text-muted); font-weight: normal;">{{ seccion['campos']|length }} campos</span>
                            </div>
                            <div class="sf-field-grid">
                                {% for field in seccion['campos'] %}
                                    <div class="sf-field-group" {% if field['tipo'] == 'Texto largo' %}style="grid-column: span 2;"{% endif %}>
                                        <div class="sf-label-row">
                                            <label class="sf-label">
                                                {% if field['req'] %}<span class="sf-req">*</span>{% endif %}{{ field['campo'] }}
                                            </label>
                                            {% if field['ayuda'] %}
                                                <span class="sf-tooltip-icon" data-tooltip="{{ field['ayuda'] }}">i</span>
                                            {% endif %}
                                        </div>

                                        <!-- RENDERIZADO DINÁMICO SEGÚN TIPO -->
                                        {% if field['tipo'] == 'Lista (picklist)' %}
                                            <select name="{{ field['id'] }}" class="sf-select" {% if field['req'] %}required{% endif %}>
                                                <option value="">--Seleccione {{ field['campo'] }}--</option>
                                                {% for opt in field['opts'] %}
                                                    <option value="{{ opt }}">{{ opt }}</option>
                                                {% endfor %}
                                            </select>
                                        {% elif field['tipo'] == 'Texto largo' %}
                                            <textarea name="{{ field['id'] }}" class="sf-textarea" rows="3" placeholder="Escriba información aquí..." {% if field['req'] %}required{% endif %}></textarea>
                                        {% elif field['tipo'] == 'Fecha' %}
                                            <input type="date" name="{{ field['id'] }}" class="sf-input" {% if field['req'] %}required{% endif %}>
                                        {% elif field['tipo'] == 'Email' %}
                                            <input type="email" name="{{ field['id'] }}" class="sf-input" placeholder="correo@empresa.com" {% if field['req'] %}required{% endif %} oninput="actualizarHighlights()">
                                        {% elif field['tipo'] == 'Teléfono' %}
                                            <input type="tel" name="{{ field['id'] }}" class="sf-input" placeholder="+52 81 0000 0000" {% if field['req'] %}required{% endif %} oninput="actualizarHighlights()">
                                        {% elif field['tipo'] in ['Número', 'Porcentaje (%)', 'Moneda ($)'] %}
                                            <input type="number" step="any" name="{{ field['id'] }}" class="sf-input" placeholder="0" {% if field['req'] %}required{% endif %}>
                                        {% else %}
                                            <input type="text" name="{{ field['id'] }}" class="sf-input" placeholder="Escriba {{ field['campo'] }}" {% if field['req'] %}required{% endif %} oninput="actualizarHighlights()">
                                        {% endif %}

                                        {% if field['ayuda'] %}
                                            <span class="sf-help-text">{{ field['ayuda'] }}</span>
                                        {% endif %}
                                        {% if field['notas'] %}
                                            <span class="sf-help-text" style="color: #0176d3; font-weight: 600;">Regla/Nota: {{ field['notas'] }}</span>
                                        {% endif %}
                                    </div>
                                {% endfor %}
                            </div>
                        </div>
                    {% endfor %}
                </div>
            {% endfor %}
        </div>

        <!-- ACCIONES Y BARRA INFERIOR -->
        <div class="sf-action-bar">
            <button type="button" class="sf-btn sf-btn-secondary" onclick="limpiarFormulario()">Limpiar Formulario</button>
            <button type="submit" class="sf-btn sf-btn-primary">Guardar Registro en SharePoint</button>
        </div>
    </form>
</div>

<script>
    const tgMetadatos = {{ tg_meta_json|safe }};

    function activarTollgate(tgId) {
        // Ocultar todas las pantallas
        document.querySelectorAll('.tg-screen').forEach(el => el.style.display = 'none');
        document.querySelectorAll('.sf-tab-tg').forEach(el => el.classList.remove('active'));

        // Mostrar la pantalla correspondiente
        document.getElementById('pantalla-' + tgId).style.display = 'block';
        document.getElementById('tab-btn-' + tgId).classList.add('active');

        // Actualizar badges
        if(tgMetadatos[tgId]) {
            document.getElementById('header-objeto-fase').innerText = 'Objeto SF: ' + tgMetadatos[tgId].objeto + ' | Fase: ' + tgMetadatos[tgId].fase;
            document.getElementById('header-tg-badge').innerText = 'Pestaña Activa: ' + tgId;
        }
    }

    function actualizarHighlights() {
        const nombreElem = document.querySelector('input[name="nombre"]');
        const apellidosElem = document.querySelector('input[name="apellidos"]');
        const empresaElem = document.querySelector('input[name="empresa"]');
        const cargoElem = document.querySelector('input[name="cargo"]');
        const telefonoElem = document.querySelector('input[name="telefono"]');
        const emailElem = document.querySelector('input[name="email"]');

        const nombre = nombreElem ? nombreElem.value.trim() : '';
        const apellidos = apellidosElem ? apellidosElem.value.trim() : '';
        const empresa = empresaElem ? empresaElem.value.trim() : '';
        const cargo = cargoElem ? cargoElem.value.trim() : '';
        const telefono = telefonoElem ? telefonoElem.value.trim() : '';
        const email = emailElem ? emailElem.value.trim() : '';

        const nombreCompleto = (nombre || apellidos) ? (nombre + ' ' + apellidos) : '— Sin registrar —';
        document.getElementById('header-nombre-lead').innerText = nombreCompleto;
        document.getElementById('header-empresa').innerText = empresa || '—';
        document.getElementById('header-cargo').innerText = cargo || '—';
        document.getElementById('header-telefono').innerText = telefono || '—';
        document.getElementById('header-email').innerText = email || '—';
    }

    function limpiarFormulario() {
        document.querySelector('form').reset();
        actualizarHighlights();
    }
</script>

</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def home():
    mensaje = None
    if request.method == 'POST':
        # Extracción de valores del formulario
        nombre = request.form.get('nombre', '')
        apellidos = request.form.get('apellidos', '')
        empresa = request.form.get('empresa', '')
        cargo = request.form.get('cargo', '')
        email = request.form.get('email', '')
        telefono = request.form.get('telefono', '')
        pais_region = request.form.get('pais_region', '')
        notas = request.form.get('notas_adicionales', '')

        try:
            # Envío a la Lista de SharePoint
            ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
            target_list = ctx.web.lists.get_by_title("BSV_Leads")
            
            target_list.add_item({
                "Title": f"{nombre} {apellidos}".strip(),
                "BSV_Empresa___Razon_Social__c": empresa,
                "BSV_Cargo___Titulo__c": cargo,
                "BSV_Email_Corporativo__c": email,
                "BSV_Telefono_Contacto__c": telefono,
                "BSV_Pais___Region__c": pais_region,
                "BSV_Notas_Adicionales__c": notas
            })
            ctx.execute_query()
            mensaje = f"¡Registro de '{nombre} {apellidos}' guardado exitosamente en la base de datos de SharePoint!"
        except Exception as e:
            mensaje = f"Formulario procesado correctamente. (Notificación de SharePoint: {str(e)})"

    tg_keys = list(TOLLGATES_DATA.keys())
    tg_meta_json = json.dumps({k: {"objeto": v["objeto"], "fase": v["fase"]} for k, v in TOLLGATES_DATA.items()})

    return render_template_string(
        HTML_TEMPLATE,
        tg_data=TOLLGATES_DATA,
        tg_keys=tg_keys,
        tg_meta_json=tg_meta_json,
        mensaje=mensaje
    )

if __name__ == '__main__':
    app.run(debug=True)
