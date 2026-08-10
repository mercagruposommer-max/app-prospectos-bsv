from flask import Flask, request, render_template_string
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

app = Flask(__name__)

# --- CONFIGURACIÓN DE SHAREPOINT ---
SITE_URL = "https://tu-empresa.sharepoint.com/sites/tu-sitio"
USERNAME = "usuario@tu-empresa.com"
PASSWORD = "TuPassword123"

# --- PLANTILLA HTML/CSS: SALESFORCE LIGHTNING SYSTEM 2.0 (SLDS) ---
HTML_SALESFORCE_EXCEL_ONLY = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BSV LAMMSA — Captura de Prospectos</title>
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
            --sf-shadow: 0 2px 8px rgba(0,0,0,0.08);
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
            box-shadow: var(--sf-shadow);
            max-width: 1240px;
            margin: 0 auto;
            border: 1px solid var(--sf-border);
            overflow: hidden;
        }

        /* 1. HEADER / HIGHLIGHTS PANEL DINO-DINÁMICO */
        .sf-header {
            background: linear-gradient(180deg, #ffffff 0%, #fafafa 100%);
            padding: 16px 24px;
            border-bottom: 1px solid var(--sf-border);
        }
        .sf-header-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 14px;
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
            box-shadow: 0 2px 4px rgba(46,132,74,0.3);
        }
        .sf-header-titles span {
            font-size: 11px;
            color: var(--sf-text-muted);
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
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
            word-break: break-all;
        }

        /* 2. BARRA DE CHEVRONS (TOLLGATE PATH) */
        .sf-path-bar {
            display: flex;
            background-color: #f3f3f3;
            padding: 8px 16px;
            border-bottom: 1px solid var(--sf-border);
            overflow-x: auto;
        }
        .sf-chevron {
            flex: 1;
            padding: 9px 12px 9px 24px;
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
            user-select: none;
            transition: all 0.2s ease;
        }
        .sf-chevron:first-child {
            clip-path: polygon(0% 0%, 88% 0%, 100% 50%, 88% 100%, 0% 100%);
            padding-left: 16px;
        }
        .sf-chevron.completed {
            background-color: var(--sf-green);
            color: #ffffff;
        }
        .sf-chevron.active {
            background-color: var(--sf-path-blue);
            color: #ffffff;
            box-shadow: inset 0 -2px 0 #ffffff;
        }
        .sf-chevron:hover:not(.active) {
            background-color: #dddbda;
        }

        /* 3. PESTAÑAS Y METADATOS DE NAVEGACIÓN */
        .sf-nav-strip {
            display: flex;
            align-items: center;
            justify-content: space-between;
            background-color: #ffffff;
            border-bottom: 1px solid var(--sf-border);
            padding: 0 16px;
        }
        .sf-tabs {
            display: flex;
        }
        .sf-tab {
            padding: 12px 20px;
            font-size: 13px;
            font-weight: 600;
            color: var(--sf-text-muted);
            cursor: pointer;
            border-bottom: 3px solid transparent;
            transition: all 0.2s ease;
        }
        .sf-tab.active {
            color: var(--sf-brand);
            border-bottom-color: var(--sf-brand);
        }

        /* 4. SECCIONES Y CONTENIDOS DE FORMULARIO */
        .sf-body-content {
            padding: 20px;
            background-color: #f8f9fb;
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

        /* CONTROLES DE CAMPO CON TOOLTIPS */
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
            white-space: normal;
            width: 200px;
            z-index: 100;
            box-shadow: 0 2px 6px rgba(0,0,0,0.2);
            font-weight: normal;
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
            transition: all 0.2s ease;
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

        /* BOTONES FLOTANTES */
        .sf-action-bar {
            position: sticky;
            bottom: 0;
            background-color: #ffffff;
            border-top: 1px solid var(--sf-border);
            padding: 12px 24px;
            display: flex;
            justify-content: flex-end;
            gap: 12px;
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
            transition: all 0.2s ease;
        }
        .sf-btn-secondary {
            background-color: #ffffff;
            color: var(--sf-brand);
        }
        .sf-btn-secondary:hover {
            background-color: #f3f3f3;
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

    <!-- 1. HEADER HIGHLIGHTS PANEL (DATOS DINÁMICOS EXCLUSIVOS DE EXCEL) -->
    <div class="sf-header">
        <div class="sf-header-top">
            <div class="sf-header-identity">
                <div class="sf-icon-avatar">★</div>
                <div class="sf-header-titles">
                    <span>Objeto SF: Lead | Fase: BD → MO</span>
                    <h1 id="dyn-lead-title">Nuevo Prospecto (Por llenar)</h1>
                </div>
            </div>
            <div class="sf-badge-pill">TG0 — Origen y Asignación</div>
        </div>

        <div class="sf-highlights-grid">
            <div class="sf-highlight-item">
                <span>Empresa / Razón Social</span>
                <strong id="dyn-empresa">—</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Cargo / Título</span>
                <strong id="dyn-cargo">—</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Teléfono Contacto</span>
                <strong id="dyn-telefono">—</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Email Corporativo</span>
                <strong id="dyn-email">—</strong>
            </div>
        </div>
    </div>

    <!-- 2. RUTA DE TOLLGATES (EXCEL PROTOTIPO TG0 - TG13) -->
    <div class="sf-path-bar">
        <div class="sf-chevron active" id="btn-tg0" onclick="cambiarPaso('tg0')">✓ TG0 — Origen</div>
        <div class="sf-chevron" id="btn-tg1" onclick="cambiarPaso('tg1')">TG1 — Fit</div>
        <div class="sf-chevron" id="btn-tg2" onclick="cambiarPaso('tg2')">TG2 — Engagement</div>
        <div class="sf-chevron" id="btn-tg3" onclick="cambiarPaso('tg3')">TG3 — MQL</div>
        <div class="sf-chevron" id="btn-tg4" onclick="cambiarPaso('tg4')">TG4 — Clasificación</div>
        <div class="sf-chevron" id="btn-tg5" onclick="cambiarPaso('tg5')">TG5 — WO</div>
    </div>

    <!-- 3. STRIP DE PESTAÑAS -->
    <div class="sf-nav-strip">
        <div class="sf-tabs">
            <div class="sf-tab active">Detalles de la Captura</div>
            <div class="sf-tab">Instrucciones y Reglas</div>
            <div class="sf-tab">Historial</div>
        </div>
        <span style="font-size: 11px; color: var(--sf-text-muted); font-weight: 600;">Paso 1 de 14</span>
    </div>

    <form method="POST">
        <div class="sf-body-content">
            
            {% if mensaje %}
                <div class="alert-success">{{ mensaje }}</div>
            {% endif %}

            <!-- ================= PANTALLA TG0 ================= -->
            <div id="pantalla-tg0" class="tg-screen">
                
                <!-- SECCIÓN 1: DATOS DEL CONTACTO (CAMPOS SEPARADOS E INDEPENDIENTES) -->
                <div class="sf-card-section">
                    <div class="sf-card-header">
                        <span>▼ Datos del Contacto</span>
                        <span style="font-size: 11px; color: var(--sf-text-muted); font-weight: normal;">7 campos obligatorios/opcionales</span>
                    </div>
                    <div class="sf-field-grid">
                        
                        <!-- 1. NOMBRE (INDEPENDIENTE) -->
                        <div class="sf-field-group">
                            <div class="sf-label-row">
                                <label class="sf-label"><span class="sf-req">*</span>Nombre</label>
                                <span class="sf-tooltip-icon" data-tooltip="Nombre del contacto. Campo estándar de Salesforce.">i</span>
                            </div>
                            <input type="text" name="nombre" class="sf-input" placeholder="Ingrese nombre" required oninput="actualizarHighlights()">
                            <span class="sf-help-text">Campo estándar de Salesforce.</span>
                        </div>

                        <!-- 2. APELLIDOS (INDEPENDIENTE) -->
                        <div class="sf-field-group">
                            <div class="sf-label-row">
                                <label class="sf-label"><span class="sf-req">*</span>Apellidos</label>
                                <span class="sf-tooltip-icon" data-tooltip="Apellidos del contacto. Campo estándar de Salesforce.">i</span>
                            </div>
                            <input type="text" name="apellidos" class="sf-input" placeholder="Ingrese apellidos" required oninput="actualizarHighlights()">
                            <span class="sf-help-text">Campo estándar de Salesforce.</span>
                        </div>

                        <!-- 3. EMPRESA / RAZÓN SOCIAL -->
                        <div class="sf-field-group">
                            <div class="sf-label-row">
                                <label class="sf-label"><span class="sf-req">*</span>Empresa / Razón Social</label>
                                <span class="sf-tooltip-icon" data-tooltip="Nombre de la empresa del lead.">i</span>
                            </div>
                            <input type="text" name="empresa" class="sf-input" placeholder="Ingrese razón social" required oninput="actualizarHighlights()">
                            <span class="sf-help-text">Nombre de la empresa del lead.</span>
                        </div>

                        <!-- 4. CARGO / TÍTULO -->
                        <div class="sf-field-group">
                            <div class="sf-label-row">
                                <label class="sf-label"><span class="sf-req">*</span>Cargo / Título</label>
                                <span class="sf-tooltip-icon" data-tooltip="Cargo o posición del contacto dentro de la empresa.">i</span>
                            </div>
                            <input type="text" name="cargo" class="sf-input" placeholder="Ej. Director / Gerente de Planta" required oninput="actualizarHighlights()">
                            <span class="sf-help-text">Cargo o posición del contacto en la empresa.</span>
                        </div>

                        <!-- 5. EMAIL CORPORATIVO -->
                        <div class="sf-field-group">
                            <div class="sf-label-row">
                                <label class="sf-label"><span class="sf-req">*</span>Email Corporativo</label>
                                <span class="sf-tooltip-icon" data-tooltip="Email corporativo del contacto. Verificar dominio.">i</span>
                            </div>
                            <input type="email" name="email" class="sf-input" placeholder="correo@empresa.com" required oninput="actualizarHighlights()">
                            <span class="sf-help-text">Verificar que pertenece al dominio de la empresa.</span>
                        </div>

                        <!-- 6. TELÉFONO CONTACTO -->
                        <div class="sf-field-group">
                            <div class="sf-label-row">
                                <label class="sf-label">Teléfono Contacto</label>
                                <span class="sf-tooltip-icon" data-tooltip="Teléfono directo o móvil del contacto.">i</span>
                            </div>
                            <input type="tel" name="telefono" class="sf-input" placeholder="+52 81 0000 0000" oninput="actualizarHighlights()">
                            <span class="sf-help-text">Teléfono directo o móvil del contacto.</span>
                        </div>

                        <!-- 7. PAÍS / REGIÓN (PICKLIST EXCEL) -->
                        <div class="sf-field-group" style="grid-column: span 2;">
                            <div class="sf-label-row">
                                <label class="sf-label"><span class="sf-req">*</span>País / Región</label>
                                <span class="sf-tooltip-icon" data-tooltip="Determina si la geografía está dentro del scope de LAMMSA.">i</span>
                            </div>
                            <select name="pais_region" class="sf-select" required>
                                <option value="">--Seleccione País / Región--</option>
                                <option value="Mexico">Mexico</option>
                                <option value="USA">USA</option>
                            </select>
                            <span class="sf-help-text">País o región donde opera el lead.</span>
                        </div>

                    </div>
                </div>

                <!-- SECCIÓN 2: ORIGEN Y ASIGNACIÓN -->
                <div class="sf-card-section">
                    <div class="sf-card-header">
                        <span>▼ Origen y Asignación</span>
                        <span style="font-size: 11px; color: #18a0fb; font-weight: 600;">NO REQUERIDO</span>
                    </div>
                    <div class="sf-field-grid">
                        <div class="sf-field-group" style="grid-column: span 2;">
                            <div class="sf-label-row">
                                <label class="sf-label">Notas Adicionales</label>
                                <span class="sf-tooltip-icon" data-tooltip="Notas generales sobre el lead. Campo estándar de Salesforce.">i</span>
                            </div>
                            <textarea name="notas_adicionales" class="sf-textarea" rows="3" placeholder="Escriba aquí observaciones iniciales sobre el prospecto..."></textarea>
                            <span class="sf-help-text">Notas generales sobre el lead. Campo estándar de Salesforce.</span>
                        </div>
                    </div>
                </div>

            </div>

            <!-- ================= PANTALLA TG1 (SEGMENTACIÓN Y FIT) ================= -->
            <div id="pantalla-tg1" class="tg-screen" style="display:none;">
                <div class="sf-card-section">
                    <div class="sf-card-header">
                        <span>▼ TG1 — Segmentación y Fit</span>
                        <span style="font-size: 11px; color: var(--sf-text-muted);">Taxonomía BSV LAMMSA</span>
                    </div>
                    <div class="sf-field-grid">
                        <div class="sf-field-group">
                            <div class="sf-label-row">
                                <label class="sf-label"><span class="sf-req">*</span>Geografía</label>
                            </div>
                            <select name="geografia" class="sf-select">
                                <option value="">--Seleccione Geografía--</option>
                                <option value="Norte">Norte</option>
                                <option value="Centro">Centro</option>
                                <option value="Bajio">Bajío</option>
                                <option value="Occidente">Occidente</option>
                                <option value="Golfo">Golfo</option>
                            </select>
                        </div>

                        <div class="sf-field-group">
                            <div class="sf-label-row">
                                <label class="sf-label"><span class="sf-req">*</span>Relevancia del Portafolio</label>
                            </div>
                            <select name="relevancia" class="sf-select">
                                <option value="">--Seleccione Relevancia--</option>
                                <option value="Alta">Alta</option>
                                <option value="Media">Media</option>
                                <option value="Baja">Baja</option>
                            </select>
                        </div>

                        <div class="sf-field-group">
                            <div class="sf-label-row">
                                <label class="sf-label"><span class="sf-req">*</span>Tamaño de Empresa</label>
                            </div>
                            <select name="tamano_empresa" class="sf-select">
                                <option value="">--Seleccione Tamaño--</option>
                                <option value="Micro — <$10M">Micro — &lt;$10M</option>
                                <option value="Pequeña — $10-50M">Pequeña — $10-50M</option>
                                <option value="Mediana — $50-200M">Mediana — $50-200M</option>
                                <option value="Grande — $200M-$1B">Grande — $200M-$1B</option>
                                <option value="Enterprise — >$1B">Enterprise — &gt;$1B</option>
                            </select>
                        </div>

                        <div class="sf-field-group">
                            <div class="sf-label-row">
                                <label class="sf-label"><span class="sf-req">*</span>Banda Asignada</label>
                            </div>
                            <select name="banda" class="sf-select">
                                <option value="">--Seleccione Banda--</option>
                                <option value="K">K</option>
                                <option value="A">A</option>
                                <option value="B">B</option>
                                <option value="C">C</option>
                                <option value="D">D</option>
                            </select>
                            <span class="sf-help-text">Si es K o A, entonces es BSV - Normal.</span>
                        </div>
                    </div>
                </div>
            </div>

        </div>

        <!-- 5. ACCIONES FLOTANTES -->
        <div class="sf-action-bar">
            <button type="button" class="sf-btn sf-btn-secondary" onclick="limpiarFormulario()">Limpiar Campos</button>
            <button type="submit" class="sf-btn sf-btn-primary">Guardar Prospecto en SharePoint</button>
        </div>
    </form>
</div>

<script>
    function actualizarHighlights() {
        const nombre = document.querySelector('input[name="nombre"]').value.trim();
        const apellidos = document.querySelector('input[name="apellidos"]').value.trim();
        const empresa = document.querySelector('input[name="empresa"]').value.trim();
        const cargo = document.querySelector('input[name="cargo"]').value.trim();
        const telefono = document.querySelector('input[name="telefono"]').value.trim();
        const email = document.querySelector('input[name="email"]').value.trim();

        // Actualizar título principal
        const tituloCompleto = (nombre || apellidos) ? (nombre + ' ' + apellidos) : 'Nuevo Prospecto (Por llenar)';
        document.getElementById('dyn-lead-title').innerText = tituloCompleto;

        // Actualizar resumen de highlights
        document.getElementById('dyn-empresa').innerText = empresa || '—';
        document.getElementById('dyn-cargo').innerText = cargo || '—';
        document.getElementById('dyn-telefono').innerText = telefono || '—';
        document.getElementById('dyn-email').innerText = email || '—';
    }

    function cambiarPaso(pasoId) {
        document.querySelectorAll('.tg-screen').forEach(el => el.style.display = 'none');
        document.querySelectorAll('.sf-chevron').forEach(el => el.classList.remove('active'));
        
        document.getElementById('pantalla-' + pasoId).style.display = 'block';
        document.getElementById('btn-' + pasoId).classList.add('active');
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
        nombre = request.form.get('nombre')
        apellidos = request.form.get('apellidos')
        empresa = request.form.get('empresa')
        cargo = request.form.get('cargo')
        email = request.form.get('email')
        telefono = request.form.get('telefono')
        pais_region = request.form.get('pais_region')
        notas = request.form.get('notas_adicionales')

        try:
            # Inserción directa en la Lista de SharePoint 'BSV_Leads'
            ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
            target_list = ctx.web.lists.get_by_title("BSV_Leads")
            
            target_list.add_item({
                "Title": f"{nombre} {apellidos}",
                "BSV_Empresa___Razon_Social__c": empresa,
                "BSV_Cargo___Titulo__c": cargo,
                "BSV_Email_Corporativo__c": email,
                "BSV_Telefono_Contacto__c": telefono,
                "BSV_Pais___Region__c": pais_region,
                "BSV_Notas_Adicionales__c": notas
            })
            ctx.execute_query()
            mensaje = f"¡Prospecto '{nombre} {apellidos}' de la empresa '{empresa}' guardado exitosamente en SharePoint!"
        except Exception as e:
            mensaje = f"Notificación: Formulario validado correctamente localmente. (Conexión SharePoint: {str(e)})"

    return render_template_string(HTML_SALESFORCE_EXCEL_ONLY, mensaje=mensaje)

if __name__ == '__main__':
    app.run(debug=True)
