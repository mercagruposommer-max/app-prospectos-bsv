import os
import json
from flask import Flask, request, render_template_string
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

app = Flask(__name__)

# --- CONFIGURACIÓN DE SHAREPOINT ---
SITE_URL = "https://tu-empresa.sharepoint.com/sites/tu-sitio"
USERNAME = "usuario@tu-empresa.com"
PASSWORD = "TuPassword123"

EXCEL_FILE = 'LAMMSA_BSV_Tollgates_Prototipo_Excel.xlsx'

# --- FUNCIÓN DE CARGA DINÁMICA DESDE EL EXCEL ---
def cargar_datos_desde_excel():
    if not os.path.exists(EXCEL_FILE):
        return None
    try:
        import pandas as pd
        
        # 1. Cargar Opciones de Listas (_Listas)
        df_listas = pd.read_excel(EXCEL_FILE, sheet_name='_Listas')
        picklists = {}
        for col in df_listas.columns:
            vals = df_listas[col].dropna().tolist()
            picklists[col] = [str(v).strip() for v in vals]

        tollgates = {}
        for i in range(14):
            sheet = f"TG{i}"
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet)
            
            submeta = str(df.iloc[0, 0])
            sf_object = "Lead" if "Lead" in submeta else "Opportunity"
            fase = submeta.split("|")[1].replace("Fase:", "").strip() if "|" in submeta else ""
            
            headers = [str(h).strip() for h in df.iloc[2].values.tolist()]
            data = df.iloc[3:].copy()
            data.columns = headers
            
            secciones = []
            for sec_name, group in data.groupby('Sección', sort=False):
                campos = []
                for idx, row in group.iterrows():
                    campo = str(row['Campo']).strip()
                    tipo = str(row['Tipo de Dato']).strip()
                    req = str(row['Requerido']).strip() == 'Sí'
                    ayuda = str(row['Ayuda / Descripción']).strip() if pd.notna(row['Ayuda / Descripción']) else ""
                    notas = str(row['Notas / Condición de Avance']).strip() if pd.notna(row['Notas / Condición de Avance']) else ""
                    
                    # Emparejar con picklist si aplica
                    opts = []
                    for col, vals in picklists.items():
                        c_clean = col.replace('BSV_', '').replace('__c', '').lower()
                        campo_clean = campo.lower().replace(' ', '_').replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')
                        if c_clean in campo_clean or campo_clean in c_clean:
                            opts = vals
                            break
                    
                    # Separación explícita de Nombre y Apellidos en TG0
                    if sheet == "TG0" and campo == "Nombre / Apellido":
                        campos.append({
                            'id': 'nombre',
                            'campo': 'Nombre',
                            'tipo': 'Texto',
                            'req': True,
                            'ayuda': 'Nombre del contacto. Campo estándar de Salesforce.',
                            'notas': '',
                            'opts': []
                        })
                        campos.append({
                            'id': 'apellidos',
                            'campo': 'Apellidos',
                            'tipo': 'Texto',
                            'req': True,
                            'ayuda': 'Apellidos del contacto. Campo estándar de Salesforce.',
                            'notas': '',
                            'opts': []
                        })
                    else:
                        field_id = "field_" + sheet.lower() + "_" + campo.lower().replace(' ', '_').replace('/', '_').replace('-', '_').replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u').replace('ñ','n').replace('(', '').replace(')', '').replace('%', 'pct').replace('$', 'usd')
                        if campo == "Empresa / Razón Social":
                            field_id = "empresa"
                        elif campo == "Cargo / Título":
                            field_id = "cargo"
                        elif campo == "Email Corporativo":
                            field_id = "email"
                        elif campo == "Teléfono Contacto":
                            field_id = "telefono"
                        
                        campos.append({
                            'id': field_id,
                            'campo': campo,
                            'tipo': tipo,
                            'req': req,
                            'ayuda': ayuda,
                            'notas': notas,
                            'opts': opts
                        })
                secciones.append({
                    'nombre': str(sec_name).strip(),
                    'campos': campos
                })
            tollgates[sheet] = {
                'objeto': sf_object,
                'fase': fase,
                'secciones': secciones
            }
        return tollgates
    except Exception as e:
        print("Aviso: No se pudo cargar el archivo Excel dinámicamente:", e)
        return None

# Intentar cargar directamente del Excel
TOLLGATES_DATA = cargar_datos_desde_excel()

# --- PLANTILLA HTML / SALESFORCE LIGHTNING SYSTEM (SLDS 2.0) ---
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BSV LAMMSA — Captura de Prospectos (TG0 a TG13)</title>
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

        /* 1. HEADER HIGHLIGHTS PANEL (EN BLANCO HASTA QUE EL USUARIO ESCRIBA) */
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

        /* 2. PESTAÑAS DE NAVEGACIÓN INDEPENDIENTES PARA TG0 A TG13 */
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

        /* 3. SECCIONES Y CONTENIDOS DEL FORMULARIO */
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

        /* CONTROLES E INSUMOS DE INPUT */
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

        /* ACCIONES Y BOTONES FLOTANTES */
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

    <!-- 1. HEADER HIGHLIGHTS PANEL -->
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

    <!-- 2. PESTAÑAS DEDICADAS PARA CADA TOLLGATE (TG0 A TG13) -->
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

            <!-- RENDERIZADO DINÁMICO DE TODAS LAS PANTALLAS (TG0 A TG13) -->
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

                                        <!-- RENDERIZADO SEGÚN TIPO DE DATO DE LA MATRIZ -->
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
                                            <input type="email" id="input-{{ field['id'] }}" name="{{ field['id'] }}" class="sf-input" placeholder="correo@empresa.com" {% if field['req'] %}required{% endif %} oninput="actualizarHighlights()">
                                        {% elif field['tipo'] == 'Teléfono' %}
                                            <input type="tel" id="input-{{ field['id'] }}" name="{{ field['id'] }}" class="sf-input" placeholder="+52 81 0000 0000" {% if field['req'] %}required{% endif %} oninput="actualizarHighlights()">
                                        {% elif field['tipo'] in ['Número', 'Porcentaje (%)', 'Moneda ($)'] %}
                                            <input type="number" step="any" name="{{ field['id'] }}" class="sf-input" placeholder="0" {% if field['req'] %}required{% endif %}>
                                        {% else %}
                                            <input type="text" id="input-{{ field['id'] }}" name="{{ field['id'] }}" class="sf-input" placeholder="Escriba {{ field['campo'] }}" {% if field['req'] %}required{% endif %} oninput="actualizarHighlights()">
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

        <!-- BARRA INFERIOR DE ACCIONES -->
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

        // Mostrar la pantalla seleccionada
        const pantallaTarget = document.getElementById('pantalla-' + tgId);
        const tabTarget = document.getElementById('tab-btn-' + tgId);
        
        if (pantallaTarget) pantallaTarget.style.display = 'block';
        if (tabTarget) tabTarget.classList.add('active');

        // Actualizar badges e indicadores
        if (tgMetadatos[tgId]) {
            document.getElementById('header-objeto-fase').innerText = 'Objeto SF: ' + tgMetadatos[tgId].objeto + ' | Fase: ' + tgMetadatos[tgId].fase;
            document.getElementById('header-tg-badge').innerText = 'Pestaña Activa: ' + tgId;
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
        nombre = request.form.get('nombre', '')
        apellidos = request.form.get('apellidos', '')
        empresa = request.form.get('empresa', '')
        cargo = request.form.get('cargo', '')
        email = request.form.get('email', '')
        telefono = request.form.get('telefono', '')

        try:
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
            mensaje = f"¡Registro de '{nombre} {apellidos}' guardado exitosamente en SharePoint!"
        except Exception as e:
            mensaje = f"Formulario procesado correctamente. (Notificación SharePoint: {str(e)})"

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
