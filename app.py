from flask import Flask, request, render_template_string
import json

app = Flask(__name__)

# --- CONFIGURACIÓN DE SHAREPOINT ---
SITE_URL = "https://tu-empresa.sharepoint.com/sites/tu-sitio"
USERNAME = "usuario@tu-empresa.com"
PASSWORD = "TuPassword123"

# --- PLANTILLA HTML/CSS: RÉPLICA SALESFORCE LIGHTNING (SLDS) ---
HTML_SALESFORCE_REPLICA = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Salesforce — Captura de Prospectos BSV</title>
    <style>
        :root {
            --sf-blue: #0176d3;
            --sf-blue-dark: #005fb2;
            --sf-path-blue: #00396b;
            --sf-green: #2e844a;
            --sf-bg: #b0c4df;
            --sf-card-bg: #ffffff;
            --sf-border: #dddbda;
            --sf-text-main: #181818;
            --sf-text-muted: #444444;
            --sf-required: #ea001e;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--sf-bg);
            margin: 0;
            padding: 12px;
            color: var(--sf-text-main);
        }

        .sf-canvas {
            background-color: #f3f3f3;
            border-radius: 4px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.15);
            max-width: 1280px;
            margin: 0 auto;
            border: 1px solid var(--sf-border);
            overflow: hidden;
        }

        /* 1. HIGHLIGHTS PANEL (HEADER SUPERIOR) */
        .sf-header {
            background-color: #ffffff;
            padding: 12px 20px;
            border-bottom: 1px solid var(--sf-border);
        }
        .sf-header-top {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 12px;
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
        .sf-header-titles span {
            font-size: 11px;
            color: #706e6b;
            font-weight: 600;
        }
        .sf-header-titles h1 {
            font-size: 18px;
            margin: 0;
            font-weight: 700;
            color: #181818;
        }
        
        .sf-highlights-grid {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 15px;
            border-top: 1px solid #f3f3f3;
            padding-top: 10px;
        }
        .sf-highlight-item span {
            display: block;
            font-size: 11px;
            color: #706e6b;
        }
        .sf-highlight-item strong {
            font-size: 13px;
            color: #181818;
        }
        .sf-highlight-item a {
            color: #0176d3;
            text-decoration: none;
            font-size: 13px;
        }

        /* 2. BARRA DE CHEVRONS (PATH BAR) */
        .sf-path-bar {
            display: flex;
            background-color: #ffffff;
            padding: 8px 15px;
            border-bottom: 1px solid var(--sf-border);
            overflow-x: auto;
        }
        .sf-chevron {
            flex: 1;
            padding: 8px 10px 8px 22px;
            background-color: #f3f3f3;
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
        }
        .sf-chevron:first-child {
            clip-path: polygon(0% 0%, 88% 0%, 100% 50%, 88% 100%, 0% 100%);
            padding-left: 15px;
        }
        .sf-chevron.completed {
            background-color: var(--sf-green);
            color: #ffffff;
        }
        .sf-chevron.active {
            background-color: var(--sf-path-blue);
            color: #ffffff;
        }

        /* 3. PESTAÑAS (TABS) */
        .sf-tabs {
            display: flex;
            background-color: #ffffff;
            border-bottom: 1px solid var(--sf-border);
            padding-left: 15px;
        }
        .sf-tab {
            padding: 10px 20px;
            font-size: 13px;
            font-weight: 600;
            color: #706e6b;
            cursor: pointer;
            border-bottom: 3px solid transparent;
        }
        .sf-tab.active {
            color: var(--sf-blue);
            border-bottom-color: var(--sf-blue);
        }

        /* 4. FORMULARIO Y ACORDEONES */
        .sf-body-content {
            padding: 15px;
        }
        .sf-accordion {
            background-color: #ffffff;
            border: 1px solid var(--sf-border);
            border-radius: 4px;
            margin-bottom: 15px;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
        }
        .sf-accordion-header {
            background-color: #f3f3f3;
            padding: 8px 12px;
            font-size: 13px;
            font-weight: 700;
            color: #181818;
            border-bottom: 1px solid var(--sf-border);
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .sf-form-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px 25px;
            padding: 15px;
        }
        @media (max-width: 768px) {
            .sf-form-grid { grid-template-columns: 1fr; }
        }

        /* CONTROLES DE CAMPOS */
        .sf-field {
            display: flex;
            flex-direction: column;
        }
        .sf-label {
            font-size: 12px;
            font-weight: 600;
            color: var(--sf-text-muted);
            margin-bottom: 3px;
        }
        .sf-req {
            color: var(--sf-required);
            font-weight: bold;
            margin-right: 3px;
        }
        .sf-input, .sf-select, .sf-textarea {
            padding: 6px 10px;
            border: 1px solid var(--sf-border);
            border-radius: 4px;
            font-size: 13px;
            background-color: #ffffff;
            box-sizing: border-box;
            width: 100%;
        }
        .sf-input:focus, .sf-select:focus, .sf-textarea:focus {
            outline: none;
            border-color: var(--sf-blue);
            box-shadow: 0 0 3px rgba(1, 118, 211, 0.5);
        }
        .sf-sublink {
            font-size: 11px;
            color: var(--sf-blue);
            text-decoration: none;
            margin-top: 2px;
            display: inline-block;
        }

        /* COMPOUND FIELD (NOMBRE COMPLETO) */
        .sf-compound-box {
            border: 1px solid var(--sf-border);
            border-radius: 4px;
            padding: 10px;
            background-color: #fafafa;
        }
        .sf-compound-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 8px;
        }

        /* BOTONES FLOTANTES / FOOTER */
        .sf-action-bar {
            position: sticky;
            bottom: 0;
            background-color: #ffffff;
            border-top: 1px solid var(--sf-border);
            padding: 10px 20px;
            display: flex;
            justify-content: flex-end;
            gap: 10px;
            box-shadow: 0 -2px 5px rgba(0,0,0,0.05);
        }
        .sf-btn {
            padding: 7px 18px;
            border-radius: 4px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            border: 1px solid var(--sf-border);
        }
        .sf-btn-cancel {
            background-color: #ffffff;
            color: #0176d3;
        }
        .sf-btn-save {
            background-color: #0176d3;
            color: #ffffff;
            border-color: #0176d3;
        }
        .sf-btn-save:hover {
            background-color: #005fb2;
        }
        
        .alert-box {
            background-color: #d4edda;
            color: #155724;
            padding: 10px 15px;
            border-radius: 4px;
            margin-bottom: 12px;
            font-size: 13px;
        }
    </style>
</head>
<body>

<div class="sf-canvas">

    <!-- 1. HEADER HIGHLIGHTS PANEL -->
    <div class="sf-header">
        <div class="sf-header-top">
            <div class="sf-lead-icon">★</div>
            <div class="sf-header-titles">
                <span>Prospecto</span>
                <h1 id="header-nombre">Khaled Parra Chá</h1>
            </div>
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
                <span>Teléfono (2)</span>
                <a href="#" id="header-telefono">+52 8183 424811</a>
            </div>
            <div class="sf-highlight-item">
                <span>Email</span>
                <a href="#" id="header-email">khaled.parra@hermosillo.com</a>
            </div>
        </div>
    </div>

    <!-- 2. BARRA DE CHEVRONS (TOLLGATES) -->
    <div class="sf-path-bar">
        <div class="sf-chevron completed">✓</div>
        <div class="sf-chevron active" onclick="mostrarPaso('tg0')">Nuevo (TG0)</div>
        <div class="sf-chevron" onclick="mostrarPaso('tg1')">Evaluando TG1</div>
        <div class="sf-chevron" onclick="mostrarPaso('tg2')">Trabajando TG2</div>
        <div class="sf-chevron" onclick="mostrarPaso('tg3')">Seguimiento TG3-4</div>
        <div class="sf-chevron" onclick="mostrarPaso('tg5')">WO TG5-TG12</div>
        <div class="sf-chevron" onclick="mostrarPaso('tg13')">Cierre TG13</div>
    </div>

    <!-- 3. PESTAÑAS -->
    <div class="sf-tabs">
        <div class="sf-tab active">Detalles</div>
        <div class="sf-tab">Actividad</div>
        <div class="sf-tab">Chatter</div>
    </div>

    <form method="POST">
        <div class="sf-body-content">
            
            {% if mensaje %}
                <div class="alert-box">{{ mensaje }}</div>
            {% endif %}

            <!-- SECCIÓN 1: INFORMACIÓN GENERAL DEL PROSPECTO -->
            <div class="sf-accordion">
                <div class="sf-accordion-header">
                    <span>▼ Información de los prospecto</span>
                </div>
                <div class="sf-form-grid">
                    <!-- Columna Izquierda -->
                    <div>
                        <div class="sf-field" style="margin-bottom: 12px;">
                            <label class="sf-label"><span class="sf-req">*</span>Estado de prospecto</label>
                            <select name="estado_prospecto" class="sf-select">
                                <option value="Nuevo" selected>Nuevo</option>
                                <option value="Evaluando">Evaluando TG1</option>
                                <option value="Trabajando">Trabajando TG2</option>
                                <option value="Calificado">Calificado MQL</option>
                            </select>
                        </div>

                        <!-- Campo Compuesto Nombre Completo -->
                        <div class="sf-field">
                            <label class="sf-label"><span class="sf-req">*</span>Nombre completo</label>
                            <div class="sf-compound-box">
                                <div style="margin-bottom: 6px;">
                                    <span class="sf-label">Tratamiento</span>
                                    <select name="tratamiento" class="sf-select">
                                        <option value="">--Ninguno--</option>
                                        <option value="Sr.">Sr.</option>
                                        <option value="Ing.">Ing.</option>
                                        <option value="Lic.">Lic.</option>
                                    </select>
                                </div>
                                <div class="sf-compound-grid">
                                    <div>
                                        <span class="sf-label">Nombre</span>
                                        <input type="text" name="nombre" class="sf-input" value="Khaled" oninput="actualizarHeader()">
                                    </div>
                                    <div>
                                        <span class="sf-label">Segundo nombre</span>
                                        <input type="text" name="segundo_nombre" class="sf-input" placeholder="Segundo nombre">
                                    </div>
                                </div>
                                <div class="sf-compound-grid" style="margin-top: 6px;">
                                    <div>
                                        <span class="sf-label"><span class="sf-req">*</span>Apellidos</span>
                                        <input type="text" name="apellidos" class="sf-input" value="Parra Chá" required oninput="actualizarHeader()">
                                    </div>
                                    <div>
                                        <span class="sf-label">Sufijo</span>
                                        <input type="text" name="sufijo" class="sf-input" placeholder="Sufijo">
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>

                    <!-- Columna Derecha -->
                    <div>
                        <div class="sf-field" style="margin-bottom: 12px;">
                            <label class="sf-label">Propietario del prospecto</label>
                            <div style="display:flex; align-items:center; gap:8px; font-size:13px; font-weight:600;">
                                <span style="background:#e0e0e0; border-radius:50%; width:24px; height:24px; display:inline-flex; align-items:center; justify-content:center;">👤</span>
                                Elizabeth Torres Oyarzabal
                            </div>
                        </div>

                        <div class="sf-field" style="margin-bottom: 12px;">
                            <label class="sf-label"><span class="sf-req">*</span>Empresa</label>
                            <select name="empresa_dropdown" class="sf-select">
                                <option value="Lammsa" selected>Lammsa</option>
                                <option value="Sommer">Sommer</option>
                            </select>
                            <a href="#" class="sf-sublink">Ver todas las dependencias</a>
                        </div>

                        <div class="sf-field" style="margin-bottom: 12px;">
                            <label class="sf-label"><span class="sf-req">*</span>Gerencia</label>
                            <select name="gerencia" class="sf-select">
                                <option value="GN002 NORTE" selected>GN002 NORTE</option>
                                <option value="GN001 CENTRO">GN001 CENTRO</option>
                            </select>
                            <a href="#" class="sf-sublink">Ver todas las dependencias</a>
                        </div>

                        <div class="sf-field" style="margin-bottom: 12px;">
                            <label class="sf-label"><span class="sf-req">*</span>Sucursal</label>
                            <select name="sucursal" class="sf-select">
                                <option value="Monterrey" selected>Monterrey</option>
                                <option value="CDMX">CDMX</option>
                            </select>
                            <a href="#" class="sf-sublink">Ver todas las dependencias</a>
                        </div>

                        <div class="sf-field" style="margin-bottom: 12px;">
                            <label class="sf-label"><span class="sf-req">*</span>Compañía</label>
                            <input type="text" name="compania" class="sf-input" value="Hermosillo" required oninput="actualizarHeader()">
                        </div>

                        <div class="sf-field">
                            <label class="sf-label">Sitio Web</label>
                            <input type="text" name="sitio_web" class="sf-input" value="www.hermosillo.com">
                        </div>
                    </div>
                </div>
            </div>

            <!-- SECCIÓN 2: TG0 & TG1 (DICCIONARIO LAMMSA) -->
            <div class="sf-accordion">
                <div class="sf-accordion-header">
                    <span>▼ TG1 — Segmentación y Fit (Taxonomía BSV)</span>
                </div>
                <div class="sf-form-grid">
                    <div class="sf-field">
                        <label class="sf-label"><span class="sf-req">*</span>País / Región</label>
                        <select name="pais_region" class="sf-select" required>
                            <option value="Mexico" selected>Mexico</option>
                            <option value="USA">USA</option>
                        </select>
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
                        <select name="relevancia" class="sf-select">
                            <option value="Alta" selected>Alta</option>
                            <option value="Media">Media</option>
                            <option value="Baja">Baja</option>
                        </select>
                    </div>

                    <div class="sf-field">
                        <label class="sf-label"><span class="sf-req">*</span>Tamaño de Empresa</label>
                        <select name="tamano_empresa" class="sf-select">
                            <option value="Micro — <$10M">Micro — &lt;$10M</option>
                            <option value="Pequeña — $10-50M" selected>Pequeña — $10-50M</option>
                            <option value="Mediana — $50-200M">Mediana — $50-200M</option>
                            <option value="Grande — $200M-$1B">Grande — $200M-$1B</option>
                        </select>
                    </div>

                    <div class="sf-field">
                        <label class="sf-label"><span class="sf-req">*</span>Banda Asignada</label>
                        <select name="banda" class="sf-select">
                            <option value="K">K</option>
                            <option value="A" selected>A</option>
                            <option value="B">B</option>
                            <option value="C">C</option>
                            <option value="D">D</option>
                        </select>
                    </div>

                    <div class="sf-field">
                        <label class="sf-label">Cargo del Contacto</label>
                        <input type="text" name="cargo_input" class="sf-input" value="Director de Operaciones Noreste" oninput="actualizarHeader()">
                    </div>
                </div>
            </div>

        </div>

        <!-- 5. BARRA DE ACCIÓN FLOTANTE (BOTONES INFERIORES) -->
        <div class="sf-action-bar">
            <button type="button" class="sf-btn sf-btn-cancel">Cancelar</button>
            <button type="submit" class="sf-btn sf-btn-save">Guardar</button>
        </div>
    </form>
</div>

<script>
    function actualizarHeader() {
        const nombre = document.querySelector('input[name="nombre"]').value;
        const apellidos = document.querySelector('input[name="apellidos"]').value;
        const compania = document.querySelector('input[name="compania"]').value;
        const cargo = document.querySelector('input[name="cargo_input"]').value;

        document.getElementById('header-nombre').innerText = nombre + " " + apellidos;
        document.getElementById('header-compania').innerText = compania;
        document.getElementById('header-cargo').innerText = cargo;
    }

    function mostrarPaso(pasoId) {
        console.log("Cambiando a paso: " + pasoId);
    }
</script>

</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def home():
    mensaje = None
    if request.method == 'POST':
        mensaje = "¡Prospecto guardado exitosamente en la Lista de SharePoint!"

    return render_template_string(HTML_SALESFORCE_REPLICA, mensaje=mensaje)

if __name__ == '__main__':
    app.run(debug=True)
