from flask import Flask, request, render_template_string

app = Flask(__name__)

# --- CONFIGURACIÓN DE SHAREPOINT ---
SITE_URL = "https://tu-empresa.sharepoint.com/sites/tu-sitio"
USERNAME = "usuario@tu-empresa.com"
PASSWORD = "TuPassword123"

# --- PLANTILLA HTML/CSS/JS: SALESFORCE LIGHTNING SYSTEM ULTRA-POLISHED ---
HTML_SALESFORCE_PREMIUM = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Salesforce BSV — Captura de Prospectos</title>
    <style>
        :root {
            --sf-brand: #0176d3;
            --sf-brand-hover: #005fb2;
            --sf-path-active: #00396b;
            --sf-success: #2e844a;
            --sf-bg: #b0c4df;
            --sf-card-bg: #ffffff;
            --sf-border: #dddbda;
            --sf-header-bg: #f3f3f3;
            --sf-text-main: #181818;
            --sf-text-muted: #444444;
            --sf-text-light: #706e6b;
            --sf-required: #ea001e;
            --sf-shadow: 0 2px 8px rgba(0, 0, 0, 0.12);
        }

        * { box-sizing: border-box; }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--sf-bg);
            margin: 0;
            padding: 16px;
            color: var(--sf-text-main);
            -webkit-font-smoothing: antialiased;
        }

        /* CONTENEDOR PRINCIPAL TIPO CANVAS */
        .sf-canvas {
            background-color: #f3f3f3;
            border-radius: 6px;
            box-shadow: var(--sf-shadow);
            max-width: 1300px;
            margin: 0 auto;
            border: 1px solid var(--sf-border);
            overflow: hidden;
        }

        /* 1. HIGHLIGHTS PANEL (HEADER PREMIUM) */
        .sf-header {
            background-color: #ffffff;
            padding: 16px 24px;
            border-bottom: 1px solid var(--sf-border);
        }
        .sf-header-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 14px;
        }
        .sf-header-left {
            display: flex;
            align-items: center;
            gap: 14px;
        }
        .sf-lead-avatar {
            width: 38px;
            height: 38px;
            background: linear-gradient(135deg, #4bca81, #2e844a);
            border-radius: 6px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.15);
        }
        .sf-lead-avatar svg {
            width: 22px;
            height: 22px;
            fill: white;
        }
        .sf-header-titles span {
            font-size: 11px;
            color: var(--sf-text-light);
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
        .sf-badge-status {
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
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            border-top: 1px solid #f3f3f3;
            padding-top: 12px;
        }
        .sf-highlight-item span {
            display: block;
            font-size: 11px;
            color: var(--sf-text-light);
            margin-bottom: 2px;
        }
        .sf-highlight-item strong {
            font-size: 13px;
            color: var(--sf-text-main);
        }
        .sf-highlight-item a {
            color: var(--sf-brand);
            text-decoration: none;
            font-size: 13px;
            font-weight: 500;
        }

        /* 2. BARRA DE PATH / CHEVRONS (TOLLGATES) */
        .sf-path-container {
            background-color: #ffffff;
            padding: 8px 16px;
            border-bottom: 1px solid var(--sf-border);
            overflow-x: auto;
        }
        .sf-path-bar {
            display: flex;
            align-items: center;
            min-width: 900px;
        }
        .sf-chevron {
            flex: 1;
            padding: 9px 12px 9px 26px;
            background-color: #f3f3f3;
            color: #3e3e3c;
            font-size: 12px;
            font-weight: 600;
            text-align: center;
            position: relative;
            cursor: pointer;
            margin-right: -14px;
            clip-path: polygon(0% 0%, 88% 0%, 100% 50%, 88% 100%, 0% 100%, 12% 50%);
            white-space: nowrap;
            transition: all 0.2s ease;
            user-select: none;
        }
        .sf-chevron:first-child {
            clip-path: polygon(0% 0%, 88% 0%, 100% 50%, 88% 100%, 0% 100%);
            padding-left: 16px;
        }
        .sf-chevron.completed {
            background-color: var(--sf-success);
            color: #ffffff;
        }
        .sf-chevron.active {
            background-color: var(--sf-path-active);
            color: #ffffff;
            font-weight: 700;
        }
        .sf-chevron:hover:not(.active) {
            background-color: #e5e5e5;
        }

        /* 3. PESTAÑAS (TABS) */
        .sf-tabs {
            display: flex;
            background-color: #ffffff;
            border-bottom: 1px solid var(--sf-border);
            padding-left: 20px;
        }
        .sf-tab {
            padding: 12px 24px;
            font-size: 13px;
            font-weight: 600;
            color: var(--sf-text-light);
            cursor: pointer;
            border-bottom: 3px solid transparent;
            transition: color 0.2s, border-color 0.2s;
        }
        .sf-tab.active {
            color: var(--sf-brand);
            border-bottom-color: var(--sf-brand);
        }

        /* 4. SECCIONES ACCORDION Y FORMULARIO */
        .sf-body-content {
            padding: 20px;
        }
        .sf-tab-pane {
            display: none;
        }
        .sf-tab-pane.active {
            display: block;
        }

        .sf-accordion {
            background-color: #ffffff;
            border: 1px solid var(--sf-border);
            border-radius: 4px;
            margin-bottom: 18px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            overflow: hidden;
        }
        .sf-accordion-header {
            background-color: #f3f3f3;
            padding: 10px 16px;
            font-size: 13px;
            font-weight: 700;
            color: var(--sf-text-main);
            border-bottom: 1px solid var(--sf-border);
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: space-between;
            user-select: none;
        }
        .sf-accordion-header:hover {
            background-color: #eef1f6;
        }
        .sf-accordion-title {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .sf-icon-arrow {
            transition: transform 0.2s ease;
            font-size: 10px;
        }
        .sf-accordion.collapsed .sf-icon-arrow {
            transform: rotate(-90deg);
        }
        .sf-accordion.collapsed .sf-form-grid {
            display: none;
        }

        /* FORM GRID & CAMPOS */
        .sf-form-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 14px 28px;
            padding: 20px;
        }
        @media (max-width: 768px) {
            .sf-form-grid { grid-template-columns: 1fr; }
        }

        .sf-field {
            display: flex;
            flex-direction: column;
        }
        .sf-field.full-width {
            grid-column: span 2;
        }
        @media (max-width: 768px) {
            .sf-field.full-width { grid-column: span 1; }
        }

        .sf-label {
            font-size: 12px;
            font-weight: 600;
            color: var(--sf-text-muted);
            margin-bottom: 4px;
        }
        .sf-req {
            color: var(--sf-required);
            font-weight: bold;
            margin-right: 3px;
        }
        .sf-input, .sf-select, .sf-textarea {
            padding: 7px 11px;
            border: 1px solid var(--sf-border);
            border-radius: 4px;
            font-size: 13px;
            background-color: #ffffff;
            box-sizing: border-box;
            width: 100%;
            color: var(--sf-text-main);
            transition: border-color 0.2s, box-shadow 0.2s;
        }
        .sf-input:focus, .sf-select:focus, .sf-textarea:focus {
            outline: none;
            border-color: var(--sf-brand);
            box-shadow: 0 0 0 3px rgba(1, 118, 211, 0.2);
        }
        .sf-help-text {
            font-size: 11px;
            color: var(--sf-text-light);
            margin-top: 3px;
        }
        .sf-sublink {
            font-size: 11px;
            color: var(--sf-brand);
            text-decoration: none;
            margin-top: 3px;
            display: inline-block;
        }
        .sf-sublink:hover { text-decoration: underline; }

        /* FOOTER CON BOTONES FLOTANTES */
        .sf-action-bar {
            position: sticky;
            bottom: 0;
            background-color: #ffffff;
            border-top: 1px solid var(--sf-border);
            padding: 12px 24px;
            display: flex;
            justify-content: flex-end;
            gap: 12px;
            box-shadow: 0 -3px 8px rgba(0,0,0,0.06);
            z-index: 100;
        }
        .sf-btn {
            padding: 8px 22px;
            border-radius: 4px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            border: 1px solid var(--sf-border);
            transition: background-color 0.2s;
        }
        .sf-btn-cancel {
            background-color: #ffffff;
            color: var(--sf-brand);
        }
        .sf-btn-cancel:hover {
            background-color: #f4f6f9;
        }
        .sf-btn-save {
            background-color: var(--sf-brand);
            color: #ffffff;
            border-color: var(--sf-brand);
        }
        .sf-btn-save:hover {
            background-color: var(--sf-brand-hover);
        }

        .alert-box {
            background-color: #d4edda;
            color: #155724;
            padding: 12px 18px;
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
            <div class="sf-header-left">
                <div class="sf-lead-avatar">
                    <svg viewBox="0 0 24 24"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>
                </div>
                <div class="sf-header-titles">
                    <span>Prospecto BSV</span>
                    <h1 id="header-nombre-completo">Khaled Parra Chá</h1>
                </div>
            </div>
            <div class="sf-badge-status">Proceso Comercial Activo</div>
        </div>

        <div class="sf-highlights-grid">
            <div class="sf-highlight-item">
                <span>Compañía</span>
                <strong id="header-compania">Hermosillo</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Cargo</span>
                <strong id="header-cargo">Director de Operaciones Noreste</strong>
            </div>
            <div class="sf-highlight-item">
                <span>Teléfono</span>
                <a href="#" id="header-telefono">+52 8183 424811</a>
            </div>
            <div class="sf-highlight-item">
                <span>Email Corporativo</span>
                <a href="#" id="header-email">khaled.parra@hermosillo.com</a>
            </div>
        </div>
    </div>

    <!-- 2. BARRA DE CHEVRONS / TOLLGATES -->
    <div class="sf-path-container">
        <div class="sf-path-bar">
            <div class="sf-chevron completed">✓</div>
            <div class="sf-chevron active" onclick="seleccionarEtapa(this)">Nuevo (TG0)</div>
            <div class="sf-chevron" onclick="seleccionarEtapa(this)">Evaluando TG1</div>
            <div class="sf-chevron" onclick="seleccionarEtapa(this)">Trabajando TG2</div>
            <div class="sf-chevron" onclick="seleccionarEtapa(this)">MQL Transferido (TG3-4)</div>
            <div class="sf-chevron" onclick="seleccionarEtapa(this)">WO en Proceso (TG5-12)</div>
            <div class="sf-chevron" onclick="seleccionarEtapa(this)">Cierre / Negotiation (TG13)</div>
        </div>
    </div>

    <!-- 3. PESTAÑAS (TAB BAR) -->
    <div class="sf-tabs">
        <div class="sf-tab active" onclick="cambiarTab('detalles', this)">Detalles</div>
        <div class="sf-tab" onclick="cambiarTab('actividad', this)">Actividad</div>
        <div class="sf-tab" onclick="cambiarTab('chatter', this)">Chatter</div>
    </div>

    <form method="POST">
        <div class="sf-body-content">
            
            {% if mensaje %}
                <div class="alert-box">{{ mensaje }}</div>
            {% endif %}

            <!-- TAB 1: DETALLES -->
            <div id="tab-detalles" class="sf-tab-pane active">

                <!-- ACORDEÓN 1: TG0 - DATOS DEL CONTACTO -->
                <div class="sf-accordion">
                    <div class="sf-accordion-header" onclick="toggleAccordion(this)">
                        <div class="sf-accordion-title">
                            <span class="sf-icon-arrow">▼</span>
                            <span>TG0 — Datos del Contacto, Origen y Asignación</span>
                        </div>
                        <span style="font-size:11px; color:#706e6b;">Objeto SF: Lead | Fase: BD → MO</span>
                    </div>

                    <div class="sf-form-grid">
                        <!-- CAMPOS TOTALMENTE INDEPENDIENTES DE NOMBRE Y APELLIDO -->
                        <div class="sf-field">
                            <label class="sf-label">Tratamiento</label>
                            <select name="tratamiento" class="sf-select">
                                <option value="">--Ninguno--</option>
                                <option value="Sr.">Sr.</option>
                                <option value="Sra.">Sra.</option>
                                <option value="Ing.">Ing.</option>
                                <option value="Lic.">Lic.</option>
                                <option value="Dr.">Dr.</option>
                            </select>
                        </div>

                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>Estado de prospecto</label>
                            <select name="estado_prospecto" class="sf-select" required>
                                <option value="Nuevo" selected>Nuevo</option>
                                <option value="Evaluando TG1">Evaluando TG1</option>
                                <option value="Trabajando TG2">Trabajando TG2</option>
                                <option value="Convertido MQL">Convertido MQL</option>
                            </select>
                        </div>

                        <!-- NOMBRE INDEPENDIENTE -->
                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>Nombre</label>
                            <input type="text" name="nombre" class="sf-input" value="Khaled" required oninput="actualizarHeader()">
                            <span class="sf-help-text">Nombre del contacto principal.</span>
                        </div>

                        <!-- APELLIDOS INDEPENDIENTES -->
                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>Apellidos</label>
                            <input type="text" name="apellidos" class="sf-input" value="Parra Chá" required oninput="actualizarHeader()">
                            <span class="sf-help-text">Apellidos del contacto principal.</span>
                        </div>

                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>Empresa / Razón Social</label>
                            <input type="text" name="empresa" class="sf-input" value="Hermosillo" required oninput="actualizarHeader()">
                        </div>

                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>Cargo / Título</label>
                            <input type="text" name="cargo" class="sf-input" value="Director de Operaciones Noreste" required oninput="actualizarHeader()">
                        </div>

                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>Email Corporativo</label>
                            <input type="email" name="email" class="sf-input" value="khaled.parra@hermosillo.com" required oninput="actualizarHeader()">
                        </div>

                        <div class="sf-field">
                            <label class="sf-label">Teléfono Contacto</label>
                            <input type="tel" name="telefono" class="sf-input" value="+52 8183 424811" oninput="actualizarHeader()">
                        </div>

                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>País / Región</label>
                            <select name="pais_region" class="sf-select" required>
                                <option value="Mexico" selected>Mexico</option>
                                <option value="USA">USA</option>
                            </select>
                        </div>

                        <div class="sf-field">
                            <label class="sf-label">Propietario del prospecto</label>
                            <input type="text" class="sf-input" value="Elizabeth Torres Oyarzabal" readonly style="background-color: #f3f3f3;">
                        </div>

                        <div class="sf-field full-width">
                            <label class="sf-label">Notas Adicionales</label>
                            <textarea name="notas_tg0" class="sf-textarea" rows="2" placeholder="Notas generales sobre el prospecto..."></textarea>
                        </div>
                    </div>
                </div>

                <!-- ACORDEÓN 2: TG1 - SEGMENTACIÓN Y FIT -->
                <div class="sf-accordion">
                    <div class="sf-accordion-header" onclick="toggleAccordion(this)">
                        <div class="sf-accordion-title">
                            <span class="sf-icon-arrow">▼</span>
                            <span>TG1 — Segmentación y Fit</span>
                        </div>
                        <span style="font-size:11px; color:#706e6b;">Objeto SF: Lead | Fase: MO</span>
                    </div>

                    <div class="sf-form-grid">
                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>Macro Segmento</label>
                            <select name="macro_segmento" class="sf-select" required>
                                <option value="">--Seleccionar--</option>
                                <option value="Industrial" selected>Industrial</option>
                                <option value="Automotriz">Automotriz</option>
                                <option value="Manufactura">Manufactura</option>
                            </select>
                            <a href="#" class="sf-sublink">Ver todas las dependencias</a>
                        </div>

                        <div class="sf-field">
                            <label class="sf-label">Sub-Segmento</label>
                            <select name="sub_segmento" class="sf-select">
                                <option value="">--Seleccionar--</option>
                                <option value="Construcción e Infraestructura" selected>Construcción e Infraestructura</option>
                                <option value="Metalmecánica">Metalmecánica</option>
                            </select>
                            <a href="#" class="sf-sublink">Ver todas las dependencias</a>
                        </div>

                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>Geografía</label>
                            <select name="geografia" class="sf-select" required>
                                <option value="Norte" selected>Norte</option>
                                <option value="Centro">Centro</option>
                                <option value="Bajio">Bajío</option>
                                <option value="Occidente">Occidente</option>
                                <option value="Golfo">Golfo</option>
                            </select>
                        </div>

                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>Relevancia del Portafolio</label>
                            <select name="relevancia" class="sf-select" required>
                                <option value="Alta" selected>Alta</option>
                                <option value="Media">Media</option>
                                <option value="Baja">Baja</option>
                            </select>
                        </div>

                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>Tamaño de Empresa</label>
                            <select name="tamano_empresa" class="sf-select" required>
                                <option value="Micro — <$10M">Micro — &lt;$10M</option>
                                <option value="Pequeña — $10-50M">Pequeña — $10-50M</option>
                                <option value="Mediana — $50-200M" selected>Mediana — $50-200M</option>
                                <option value="Grande — $200M-$1B">Grande — $200M-$1B</option>
                                <option value="Enterprise — >$1B">Enterprise — &gt;$1B</option>
                            </select>
                        </div>

                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>Banda Asignada</label>
                            <select name="banda" class="sf-select" required>
                                <option value="K" selected>K</option>
                                <option value="A">A</option>
                                <option value="B">B</option>
                                <option value="C">C</option>
                                <option value="D">D</option>
                            </select>
                            <span class="sf-help-text">Si es K o A, clasifica como BSV Normal.</span>
                        </div>
                    </div>
                </div>

            </div>

            <!-- TAB 2: ACTIVIDAD -->
            <div id="tab-actividad" class="sf-tab-pane">
                <div class="sf-accordion">
                    <div class="sf-accordion-header">
                        <span>▼ Registro de Actividad y Seguimiento</span>
                    </div>
                    <div style="padding: 20px; font-size: 13px; color: #706e6b;">
                        <p><strong>Llamada Registrada:</strong> Primera interacción telefónica realizada el día de hoy.</p>
                        <p><strong>Correo Enviado:</strong> Template de confirmación de reunión enviado a khaled.parra@hermosillo.com</p>
                    </div>
                </div>
            </div>

            <!-- TAB 3: CHATTER -->
            <div id="tab-chatter" class="sf-tab-pane">
                <div class="sf-accordion">
                    <div class="sf-accordion-header">
                        <span>▼ Muro de Colaboración (Chatter)</span>
                    </div>
                    <div style="padding: 20px;">
                        <textarea class="sf-textarea" rows="3" placeholder="Escribe una actualización para el equipo comercial..."></textarea>
                        <button type="button" class="sf-btn sf-btn-save" style="margin-top: 10px;">Compartir</button>
                    </div>
                </div>
            </div>

        </div>

        <!-- 5. BARRA FLOTANTE DE BOTONES -->
        <div class="sf-action-bar">
            <button type="button" class="sf-btn sf-btn-cancel">Cancelar</button>
            <button type="submit" class="sf-btn sf-btn-save">Guardar en SharePoint</button>
        </div>
    </form>
</div>

<script>
    // Actualizar Header en tiempo real
    function actualizarHeader() {
        const nombre = document.querySelector('input[name="nombre"]').value;
        const apellidos = document.querySelector('input[name="apellidos"]').value;
        const empresa = document.querySelector('input[name="empresa"]').value;
        const cargo = document.querySelector('input[name="cargo"]').value;
        const email = document.querySelector('input[name="email"]').value;
        const telefono = document.querySelector('input[name="telefono"]').value;

        document.getElementById('header-nombre-completo').innerText = nombre + " " + apellidos;
        document.getElementById('header-compania').innerText = empresa;
        document.getElementById('header-cargo').innerText = cargo;
        document.getElementById('header-email').innerText = email;
        document.getElementById('header-email').href = "mailto:" + email;
        document.getElementById('header-telefono').innerText = telefono;
    }

    // Toggle para desplegar/colapsar acordeones
    function toggleAccordion(header) {
        const accordion = header.parentElement;
        accordion.classList.toggle('collapsed');
    }

    // Cambiar Pestañas (Detalles, Actividad, Chatter)
    function cambiarTab(tabId, tabElement) {
        document.querySelectorAll('.sf-tab').forEach(el => el.classList.remove('active'));
        document.querySelectorAll('.sf-tab-pane').forEach(el => el.classList.remove('active'));

        tabElement.classList.add('active');
        document.getElementById('tab-' + tabId).classList.add('active');
    }

    // Seleccionar Etapa en Chevrons
    function seleccionarEtapa(element) {
        document.querySelectorAll('.sf-chevron').forEach(el => el.classList.remove('active'));
        element.classList.add('active');
    }
</script>

</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def home():
    mensaje = None
    if request.method == 'POST':
        mensaje = "¡Prospecto guardado exitosamente en la base de datos de SharePoint!"

    return render_template_string(HTML_SALESFORCE_PREMIUM, mensaje=mensaje)

if __name__ == '__main__':
    app.run(debug=True)
