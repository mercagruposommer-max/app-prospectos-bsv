import os
import json
from flask import Flask, request, render_template_string

app = Flask(__name__)

# --- CONFIGURACIÓN DE SHAREPOINT ---
SITE_URL = "https://tu-empresa.sharepoint.com/sites/tu-sitio"
USERNAME = "usuario@tu-empresa.com"
PASSWORD = "TuPassword123"

EXCEL_FILE = 'LAMMSA_BSV_Tollgates_Prototipo_Excel.xlsx'

# --- CARGA DINÁMICA DE CAMPOS DESDE EL EXCEL ---
def cargar_datos_excel():
    if not os.path.exists(EXCEL_FILE):
        return None
    try:
        import pandas as pd
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
                    
                    opts = []
                    for col, vals in picklists.items():
                        c_clean = col.replace('BSV_', '').replace('__c', '').lower()
                        campo_clean = campo.lower().replace(' ', '_').replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u')
                        if c_clean in campo_clean or campo_clean in c_clean:
                            opts = vals
                            break
                    
                    # Separación de Nombre y Apellidos en TG0
                    if sheet == "TG0" and campo == "Nombre / Apellido":
                        campos.append({'id': 'nombre', 'campo': 'Nombre', 'tipo': 'Texto', 'req': True, 'ayuda': 'Nombre del contacto.', 'notas': '', 'opts': []})
                        campos.append({'id': 'apellidos', 'campo': 'Apellidos', 'tipo': 'Texto', 'req': True, 'ayuda': 'Apellidos del contacto.', 'notas': '', 'opts': []})
                    else:
                        field_id = "field_" + sheet.lower() + "_" + campo.lower().replace(' ', '_').replace('/', '_').replace('-', '_').replace('á','a').replace('é','e').replace('í','i').replace('ó','o').replace('ú','u').replace('ñ','n').replace('(', '').replace(')', '').replace('%', 'pct').replace('$', 'usd')
                        if campo == "Empresa / Razón Social": field_id = "empresa"
                        elif campo == "Cargo / Título": field_id = "cargo"
                        elif campo == "Email Corporativo": field_id = "email"
                        elif campo == "Teléfono Contacto": field_id = "telefono"
                        
                        campos.append({'id': field_id, 'campo': campo, 'tipo': tipo, 'req': req, 'ayuda': ayuda, 'notas': notas, 'opts': opts})
                
                secciones.append({'nombre': str(sec_name).strip(), 'campos': campos})
            
            tollgates[sheet] = {'objeto': sf_object, 'fase': fase, 'secciones': secciones}
        return tollgates
    except Exception as e:
        print("Error al leer Excel:", e)
        return None

TOLLGATES_DATA = cargar_datos_excel()

# Base de datos temporal en memoria para la lista "Vistos recientemente"
REGISTROS_PROSPECTOS = []

# --- PLANTILLA HTML / SALESFORCE LIGHTNING SYSTEM COMPLETO ---
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
            padding: 0;
            color: var(--sf-text-main);
        }

        /* 1. BARRA NAVEGACIÓN SUPERIOR GLOBAL SALESFORCE */
        .sf-global-header {
            background-color: #ffffff;
            border-bottom: 1px solid var(--sf-border);
            padding: 0 16px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            height: 50px;
        }
        .sf-app-launcher {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .sf-waffle {
            display: grid;
            grid-template-columns: repeat(3, 4px);
            gap: 3px;
            cursor: pointer;
        }
        .sf-waffle div { width: 4px; height: 4px; background-color: var(--sf-brand); border-radius: 1px; }
        .sf-app-title { font-weight: 700; font-size: 16px; color: #181818; }
        
        .sf-nav-menu {
            display: flex;
            align-items: center;
            gap: 4px;
            height: 100%;
        }
        .sf-nav-item {
            padding: 0 12px;
            height: 100%;
            display: flex;
            align-items: center;
            font-size: 13px;
            color: #514f4d;
            cursor: pointer;
            border-bottom: 3px solid transparent;
            text-decoration: none;
        }
        .sf-nav-item.active {
            color: var(--sf-brand);
            font-weight: 700;
            border-bottom-color: var(--sf-brand);
        }

        .sf-search-bar input {
            background-color: #f3f3f3;
            border: 1px solid var(--sf-border);
            border-radius: 4px;
            padding: 6px 12px;
            width: 280px;
            font-size: 12px;
        }

        .sf-container {
            max-width: 1400px;
            margin: 12px auto;
            background: #ffffff;
            border-radius: 6px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
            overflow: hidden;
            border: 1px solid var(--sf-border);
        }

        /* 2. VISTA 1: TABLA VISTOS RECIENTEMENTE (EDITED-IMAGE.PNG) */
        .sf-list-header {
            padding: 16px 24px;
            background: #ffffff;
            border-bottom: 1px solid var(--sf-border);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .sf-list-title {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .sf-lead-icon {
            width: 32px;
            height: 32px;
            background-color: #4bca81;
            border-radius: 4px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 18px;
        }
        .sf-list-actions {
            display: flex;
            gap: 8px;
        }
        .sf-btn-nuevo {
            background-color: var(--sf-brand);
            color: #ffffff;
            border: 1px solid var(--sf-brand);
            padding: 7px 18px;
            border-radius: 4px;
            font-weight: 700;
            font-size: 13px;
            cursor: pointer;
            box-shadow: 0 0 0 2px rgba(1,118,211,0.2);
        }
        .sf-btn-nuevo:hover { background-color: var(--sf-brand-dark); }
        .sf-btn-sub {
            background-color: #ffffff;
            color: var(--sf-brand);
            border: 1px solid var(--sf-border);
            padding: 7px 14px;
            border-radius: 4px;
            font-size: 12px;
            font-weight: 600;
            cursor: pointer;
        }

        .sf-table-wrapper {
            overflow-x: auto;
            background-color: #ffffff;
        }
        .sf-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }
        .sf-table th {
            background-color: #fafafa;
            border-bottom: 2px solid var(--sf-border);
            padding: 10px 14px;
            text-align: left;
            color: #514f4d;
            font-weight: 700;
        }
        .sf-table td {
            padding: 12px 14px;
            border-bottom: 1px solid var(--sf-border);
            color: #181818;
        }
        .sf-table tr:hover { background-color: #f3f3f3; cursor: pointer; }
        .sf-link-name { color: var(--sf-brand); font-weight: 600; text-decoration: none; }

        /* 3. VISTA 2: VISTA DE DETALLE / FORMULARIO (IMAGE_F1E5FA.PNG) */
        .sf-detail-top-bar {
            padding: 12px 24px;
            background: #ffffff;
            border-bottom: 1px solid var(--sf-border);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .sf-highlights-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 20px;
            padding: 12px 24px;
            background-color: #fafafa;
            border-bottom: 1px solid var(--sf-border);
        }
        .sf-highlight-item span { display: block; font-size: 11px; color: var(--sf-text-muted); font-weight: 600; }
        .sf-highlight-item strong { font-size: 13px; color: var(--sf-text-main); }

        /* CHEVRONS PATH BAR */
        .sf-path-bar {
            display: flex;
            background-color: #f3f3f3;
            padding: 8px 16px;
            border-bottom: 1px solid var(--sf-border);
            overflow-x: auto;
        }
        .sf-chevron {
            flex: 1;
            padding: 8px 12px 8px 22px;
            background-color: #eef1f6;
            color: #3e3e3c;
            font-size: 12px;
            font-weight: 600;
            text-align: center;
            position: relative;
            cursor: pointer;
            margin-right: -12px;
            clip-path: polygon(0% 0%, 88% 0%, 100% 50%, 88% 100%, 0% 100%, 12% 50%);
            white-space: nowrap;
        }
        .sf-chevron.active { background-color: var(--sf-path-blue); color: #ffffff; }

        /* ESTRUCTURA SPLIT 70% / 30% */
        .sf-split-layout {
            display: grid;
            grid-template-columns: 7fr 3fr;
            background-color: #b0c4df;
            gap: 12px;
            padding: 12px;
        }
        @media (max-width: 992px) { .sf-split-layout { grid-template-columns: 1fr; } }

        .sf-main-col { background: #ffffff; border-radius: 4px; border: 1px solid var(--sf-border); }
        .sf-side-col { display: flex; flex-direction: column; gap: 12px; }

        /* TABS LEFT COL */
        .sf-tabs { display: flex; border-bottom: 1px solid var(--sf-border); background: #ffffff; padding-left: 16px; }
        .sf-tab { padding: 12px 20px; font-size: 13px; font-weight: 600; color: #514f4d; cursor: pointer; border-bottom: 3px solid transparent; }
        .sf-tab.active { color: var(--sf-brand); border-bottom-color: var(--sf-brand); }

        /* FORM GRID */
        .sf-card-section { background: #ffffff; border: 1px solid var(--sf-border); border-radius: 4px; margin: 16px; overflow: hidden; }
        .sf-card-header { background: #f3f3f3; padding: 10px 16px; font-size: 13px; font-weight: 700; border-bottom: 1px solid var(--sf-border); }
        .sf-field-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px 24px; padding: 16px; }

        .sf-field-group { display: flex; flex-direction: column; }
        .sf-label { font-size: 12px; font-weight: 600; color: #514f4d; margin-bottom: 4px; }
        .sf-req { color: var(--sf-required); font-weight: bold; }
        .sf-input, .sf-select, .sf-textarea { padding: 7px 10px; border: 1px solid var(--sf-border); border-radius: 4px; font-size: 13px; box-sizing: border-box; width: 100%; }
        .sf-help-text { font-size: 11px; color: #514f4d; margin-top: 3px; }

        /* SIDEBAR CARDS (RIGHT COL) */
        .sf-side-card { background: #ffffff; border: 1px solid var(--sf-border); border-radius: 4px; padding: 14px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .sf-side-card-title { font-size: 13px; font-weight: 700; color: #181818; margin-bottom: 8px; display: flex; align-items: center; justify-content: space-between; }
        .sf-drop-box { border: 2px dashed var(--sf-border); border-radius: 4px; padding: 20px; text-align: center; background: #fafafa; margin-top: 8px; }

        .alert-success { background-color: #d4edda; color: #155724; padding: 12px 16px; border-radius: 4px; margin: 12px; border: 1px solid #c3e6cb; font-size: 13px; }
    </style>
</head>
<body>

<!-- BARRA NAVEGACIÓN GLOBAL SALESFORCE -->
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
    <div class="sf-search-bar">
        <input type="text" placeholder="Buscar en Salesforce...">
    </div>
</div>

<div class="sf-container">

    <!-- ================= VISTA 1: TABLA VISTOS RECIENTEMENTE ================= -->
    <div id="vista-lista" style="display: block;">
        <div class="sf-list-header">
            <div class="sf-list-title">
                <div class="sf-lead-icon">★</div>
                <div>
                    <span style="font-size:11px; color:#514f4d; font-weight:600;">Prospectos</span>
                    <h2 style="margin:0; font-size:18px;">Vistos recientemente ▾</h2>
                </div>
            </div>
            <div class="sf-list-actions">
                <!-- BOTÓN DESTACADO EN ROJO DE LA IMAGEN EDITED-IMAGE.PNG -->
                <button class="sf-btn-nuevo" onclick="abrirNuevoFormulario()">+ Nuevo</button>
                <button class="sf-btn-sub">Vista de inteligencia</button>
                <button class="sf-btn-sub">Importar</button>
                <button class="sf-btn-sub">Cambiar estado</button>
            </div>
        </div>

        <div class="sf-table-wrapper">
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
                <tbody id="tabla-prospectos-body">
                    {% if registros %}
                        {% for r in registros %}
                            <tr onclick="verDetalleProspecto('{{ loop.index0 }}')">
                                <td><input type="checkbox"></td>
                                <td><a class="sf-link-name">{{ r.nombre }} {{ r.apellidos }}</a></td>
                                <td>{{ r.cargo or '—' }}</td>
                                <td>{{ r.empresa or '—' }}</td>
                                <td>{{ r.telefono or '—' }}</td>
                                <td>{{ r.email or '—' }}</td>
                                <td><span style="background:#eef4fe; color:#0176d3; padding:2px 8px; border-radius:10px; font-size:11px; font-weight:600;">Nuevo</span></td>
                            </tr>
                        {% endfor %}
                    {% else %}
                        <tr>
                            <td colspan="7" style="text-align: center; padding: 30px; color: #514f4d;">
                                No hay prospectos registrados aún. Haz clic en el botón <strong>"+ Nuevo"</strong> para agregar el primero.
                            </td>
                        </tr>
                    {% endif %}
                </tbody>
            </table>
        </div>
    </div>

    <!-- ================= VISTA 2: DETALLE Y CAPTURA DE PROSPECTO (IMAGE_F1E5FA.PNG) ================= -->
    <div id="vista-detalle" style="display: none;">
        
        <!-- HEADER DE DETALLE Y ACCIONES -->
        <div class="sf-detail-top-bar">
            <div style="display:flex; align-items:center; gap:12px;">
                <button class="sf-btn-sub" onclick="volverALista()">← Volver a la Lista</button>
                <div class="sf-lead-icon">★</div>
                <div>
                    <span style="font-size:11px; color:#514f4d; font-weight:600;">Prospecto</span>
                    <h1 id="dyn-lead-title" style="margin:0; font-size:18px;">— Sin registrar —</h1>
                </div>
            </div>
            <div style="display:flex; gap:8px;">
                <button class="sf-btn-sub">+ Seguir</button>
                <button class="sf-btn-sub">Modificar</button>
                <button class="sf-btn-sub">Eliminar</button>
            </div>
        </div>

        <!-- HIGHLIGHTS PANEL -->
        <div class="sf-highlights-grid">
            <div class="sf-highlight-item">
                <span>Compañía</span>
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

        <!-- CHEVRONS PATH BAR -->
        <div class="sf-path-bar">
            {% for tg_key in tg_keys %}
                <div class="sf-chevron {% if loop.first %}active{% endif %}" id="tab-btn-{{ tg_key }}" onclick="activarTollgate('{{ tg_key }}')">
                    {{ tg_key }}
                </div>
            {% endfor %}
        </div>

        <!-- LAYOUT SPLIT 70% / 30% -->
        <form method="POST" id="form-prospecto">
            <div class="sf-split-layout">
                
                <!-- COLUMNA PRINCIPAL IZQUIERDA (70%) -->
                <div class="sf-main-col">
                    <div class="sf-tabs">
                        <div class="sf-tab active">Detalles</div>
                        <div class="sf-tab">Actividad</div>
                        <div class="sf-tab">Chatter</div>
                    </div>

                    {% if mensaje %}
                        <div class="alert-success">{{ mensaje }}</div>
                    {% endif %}

                    <!-- PANTALLAS TOLLGATES (TG0 - TG13) -->
                    {% for tg_key, tg_val in tg_data.items() %}
                        <div id="pantalla-{{ tg_key }}" class="tg-screen" style="{% if not loop.first %}display:none;{% endif %}">
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
                                                    <select name="{{ field['id'] }}" class="sf-input" {% if field['req'] %}required{% endif %}>
                                                        <option value="">--Seleccione--</option>
                                                        {% for opt in field['opts'] %}
                                                            <option value="{{ opt }}">{{ opt }}</option>
                                                        {% endfor %}
                                                    </select>
                                                {% elif field['tipo'] == 'Texto largo' %}
                                                    <textarea name="{{ field['id'] }}" class="sf-textarea" rows="3" {% if field['req'] %}required{% endif %}></textarea>
                                                {% elif field['tipo'] == 'Fecha' %}
                                                    <input type="date" name="{{ field['id'] }}" class="sf-input" {% if field['req'] %}required{% endif %}>
                                                {% elif field['tipo'] == 'Email' %}
                                                    <input type="email" id="input-{{ field['id'] }}" name="{{ field['id'] }}" class="sf-input" {% if field['req'] %}required{% endif %} oninput="actualizarHighlights()">
                                                {% elif field['tipo'] == 'Teléfono' %}
                                                    <input type="tel" id="input-{{ field['id'] }}" name="{{ field['id'] }}" class="sf-input" {% if field['req'] %}required{% endif %} oninput="actualizarHighlights()">
                                                {% elif field['tipo'] in ['Número', 'Porcentaje (%)', 'Moneda ($)'] %}
                                                    <input type="number" step="any" name="{{ field['id'] }}" class="sf-input" {% if field['req'] %}required{% endif %}>
                                                {% else %}
                                                    <input type="text" id="input-{{ field['id'] }}" name="{{ field['id'] }}" class="sf-input" {% if field['req'] %}required{% endif %} oninput="actualizarHighlights()">
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

                    <div style="padding: 16px; text-align: right; background: #ffffff; border-top: 1px solid var(--sf-border);">
                        <button type="button" class="sf-btn-sub" onclick="volverALista()" style="margin-right: 8px;">Cancelar</button>
                        <button type="submit" class="sf-btn-nuevo">Guardar Prospecto</button>
                    </div>
                </div>

                <!-- COLUMNA DERECHA SIDEBAR (30%) DE RELACIONADOS -->
                <div class="sf-side-col">
                    <div class="sf-side-card" style="background:#fff3cd; border-color:#ffeeba;">
                        <span style="font-size:12px; color:#856404; font-weight:600;">⚠ Verificación</span>
                        <p style="margin:4px 0 0 0; font-size:12px; color:#856404;">No encontramos duplicados potenciales de este Prospecto.</p>
                    </div>

                    <div class="sf-side-card">
                        <div class="sf-side-card-title">
                            <span>Vínculos rápidos de lista relacionada</span>
                        </div>
                        <ul style="margin:0; padding-left:16px; font-size:12px; color:var(--sf-brand);">
                            <li style="margin-bottom:4px;"><a style="color:var(--sf-brand);">Historial de aprobaciones (0)</a></li>
                            <li style="margin-bottom:4px;"><a style="color:var(--sf-brand);">Archivos (0)</a></li>
                            <li><a style="color:var(--sf-brand);">Notas (0)</a></li>
                        </ul>
                    </div>

                    <div class="sf-side-card">
                        <div class="sf-side-card-title">
                            <span>Archivos (0)</span>
                            <button type="button" class="sf-btn-sub" style="padding:2px 8px; font-size:11px;">Cargar archivos</button>
                        </div>
                        <div class="sf-drop-box">
                            <span style="font-size:12px; color:#514f4d;">O suelte archivos aquí</span>
                        </div>
                    </div>

                    <div class="sf-side-card">
                        <div class="sf-side-card-title">
                            <span>Notas (0)</span>
                            <button type="button" class="sf-btn-sub" style="padding:2px 8px; font-size:11px;">Nueva</button>
                        </div>
                    </div>
                </div>

            </div>
        </form>
    </div>

</div>

<script>
    const tgMetadatos = {{ tg_meta_json|safe }};

    function abrirNuevoFormulario() {
        document.getElementById('form-prospecto').reset();
        actualizarHighlights();
        document.getElementById('vista-lista').style.display = 'none';
        document.getElementById('vista-detalle').style.display = 'block';
        activarTollgate('TG0');
    }

    function volverALista() {
        document.getElementById('vista-detalle').style.display = 'none';
        document.getElementById('vista-lista').style.display = 'block';
    }

    function activarTollgate(tgId) {
        document.querySelectorAll('.tg-screen').forEach(el => el.style.display = 'none');
        document.querySelectorAll('.sf-chevron').forEach(el => el.classList.remove('active'));

        const pantallaTarget = document.getElementById('pantalla-' + tgId);
        const tabTarget = document.getElementById('tab-btn-' + tgId);

        if (pantallaTarget) pantallaTarget.style.display = 'block';
        if (tabTarget) tabTarget.classList.add('active');
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
    if request.method == 'POST':
        nombre = request.form.get('nombre', '')
        apellidos = request.form.get('apellidos', '')
        empresa = request.form.get('empresa', '')
        cargo = request.form.get('cargo', '')
        email = request.form.get('email', '')
        telefono = request.form.get('telefono', '')

        nuevo_registro = {
            'nombre': nombre,
            'apellidos': apellidos,
            'empresa': empresa,
            'cargo': cargo,
            'email': email,
            'telefono': telefono
        }
        
        # Guardar en la lista en memoria para visualización en la tabla de Vistos Recientemente
        REGISTROS_PROSPECTOS.append(nuevo_registro)

        # Intento de envío a SharePoint
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
            mensaje = f"¡Prospecto '{nombre} {apellidos}' guardado exitosamente en la base de datos de SharePoint!"
        except Exception as e:
            mensaje = f"Prospecto guardado localmente en la lista del prototipo."

    tg_keys = list(TOLLGATES_DATA.keys()) if TOLLGATES_DATA else []
    tg_meta_json = json.dumps({k: {"objeto": v["objeto"], "fase": v["fase"]} for k, v in TOLLGATES_DATA.items()}) if TOLLGATES_DATA else "{}"

    return render_template_string(
        HTML_TEMPLATE,
        tg_data=TOLLGATES_DATA,
        tg_keys=tg_keys,
        tg_meta_json=tg_meta_json,
        registros=REGISTROS_PROSPECTOS,
        mensaje=mensaje
    )

if __name__ == '__main__':
    app.run(debug=True)
