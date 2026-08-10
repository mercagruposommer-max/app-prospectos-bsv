from flask import Flask, request, render_template_string
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

app = Flask(__name__)

# --- CONFIGURACIÓN DE SHAREPOINT ---
SITE_URL = "https://tu-empresa.sharepoint.com/sites/tu-sitio"
USERNAME = "usuario@tu-empresa.com"
PASSWORD = "TuPassword123"

# --- PLANTILLA HTML/CSS ESTILO SALESFORCE LIGHTNING ---
HTML_SALESFORCE = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Captura de Prospectos BSV — Salesforce Replica</title>
    <style>
        :root {
            --sf-blue: #0176d3;
            --sf-blue-dark: #005fb2;
            --sf-bg: #f3f3f3;
            --sf-border: #dddbda;
            --sf-header-bg: #f3f3f3;
            --sf-text: #181818;
            --sf-label: #444444;
            --sf-required: #ea001e;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            background-color: #b0c4df;
            margin: 0;
            padding: 15px;
            color: var(--sf-text);
        }

        .sf-container {
            background-color: #ffffff;
            border-radius: 4px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            max-width: 1200px;
            margin: 0 auto;
            overflow: hidden;
            border: 1px solid var(--sf-border);
        }

        /* Header de Registro Salesforce */
        .sf-header {
            background-color: #ffffff;
            padding: 15px 20px;
            border-bottom: 1px solid var(--sf-border);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .sf-header-title {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .sf-icon {
            width: 32px;
            height: 32px;
            background-color: #4bca81;
            border-radius: 4px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: bold;
            font-size: 18px;
        }
        .sf-header-text h1 {
            font-size: 18px;
            margin: 0;
            color: #181818;
            font-weight: 700;
        }
        .sf-header-text span {
            font-size: 12px;
            color: #706e6b;
            text-transform: uppercase;
        }

        /* Barra de Ruta / Chevrons (Tollgates) */
        .sf-path {
            display: flex;
            background-color: #f3f3f3;
            padding: 10px 15px;
            border-bottom: 1px solid var(--sf-border);
            overflow-x: auto;
            white-space: nowrap;
        }
        .sf-step {
            flex: 1;
            padding: 8px 15px 8px 25px;
            background: #eef1f6;
            color: #3e3e3c;
            font-size: 13px;
            font-weight: 600;
            position: relative;
            text-align: center;
            cursor: pointer;
            border-right: 2px solid white;
            user-select: none;
            clip-path: polygon(0% 0%, 85% 0%, 100% 50%, 85% 100%, 0% 100%, 15% 50%);
            margin-right: -10px;
        }
        .sf-step:first-child {
            clip-path: polygon(0% 0%, 85% 0%, 100% 50%, 85% 100%, 0% 100%);
            padding-left: 15px;
        }
        .sf-step.active {
            background-color: #014486;
            color: white;
        }
        .sf-step.completed {
            background-color: #4bca81;
            color: white;
        }

        /* Pestañas (Detalles, Actividad, Chatter) */
        .sf-tabs {
            display: flex;
            border-bottom: 1px solid var(--sf-border);
            background-color: #ffffff;
            padding-left: 15px;
        }
        .sf-tab {
            padding: 12px 20px;
            font-size: 14px;
            font-weight: 600;
            color: #706e6b;
            cursor: pointer;
            border-bottom: 3px solid transparent;
        }
        .sf-tab.active {
            color: var(--sf-blue);
            border-bottom-color: var(--sf-blue);
        }

        /* Formulario y Grid a 2 Columnas */
        .sf-body {
            padding: 20px;
            background-color: #f3f3f3;
        }
        .sf-section-card {
            background: white;
            border: 1px solid var(--sf-border);
            border-radius: 4px;
            margin-bottom: 20px;
            overflow: hidden;
        }
        .sf-section-header {
            background-color: #f3f3f3;
            padding: 10px 15px;
            font-size: 14px;
            font-weight: 700;
            color: #181818;
            border-bottom: 1px solid var(--sf-border);
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .sf-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px 30px;
            padding: 20px;
        }
        @media (max-width: 768px) {
            .sf-grid { grid-template-columns: 1fr; }
        }

        /* Campos de Entrada */
        .sf-field {
            display: flex;
            flex-direction: column;
        }
        .sf-label {
            font-size: 12px;
            font-weight: 600;
            color: var(--sf-label);
            margin-bottom: 4px;
        }
        .sf-label .required {
            color: var(--sf-required);
            margin-right: 2px;
        }
        .sf-input, .sf-select, .sf-textarea {
            padding: 8px 12px;
            border: 1px solid var(--sf-border);
            border-radius: 4px;
            font-size: 13px;
            color: #181818;
            background-color: #ffffff;
            transition: border-color 0.2s;
        }
        .sf-input:focus, .sf-select:focus, .sf-textarea:focus {
            outline: none;
            border-color: var(--sf-blue);
            box-shadow: 0 0 3px rgba(1, 118, 211, 0.5);
        }
        .sf-help-text {
            font-size: 11px;
            color: #706e6b;
            margin-top: 3px;
        }

        /* Footer con Botones de Acción */
        .sf-footer {
            background-color: #f3f3f3;
            padding: 12px 20px;
            border-top: 1px solid var(--sf-border);
            display: flex;
            justify-content: flex-end;
            gap: 10px;
        }
        .sf-btn {
            padding: 8px 20px;
            border-radius: 4px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            border: 1px solid var(--sf-border);
        }
        .sf-btn-secondary {
            background-color: #ffffff;
            color: var(--sf-blue);
        }
        .sf-btn-primary {
            background-color: var(--sf-blue);
            color: #ffffff;
            border-color: var(--sf-blue);
        }
        .sf-btn-primary:hover {
            background-color: var(--sf-blue-dark);
        }

        .alert-success {
            background-color: #d4edda;
            color: #155724;
            padding: 12px 20px;
            border-bottom: 1px solid #c3e6cb;
            font-size: 13px;
            font-weight: 600;
        }
    </style>
</head>
<body>

<div class="sf-container">
    <!-- Header estilo Salesforce -->
    <div class="sf-header">
        <div class="sf-header-title">
            <div class="sf-icon">★</div>
            <div class="sf-header-text">
                <span>Prospecto / Lead</span>
                <h1>Captura de Prospecto — Sistema BSV LAMMSA</h1>
            </div>
        </div>
    </div>

    <!-- Barra de Tollgates (Chevrons) -->
    <div class="sf-path">
        <div class="sf-step completed" onclick="switchTG('tg0')">✓ TG0 — Origen</div>
        <div class="sf-step active" id="btn-tg1" onclick="switchTG('tg1')">TG1 — Fit</div>
        <div class="sf-step" id="btn-tg2" onclick="switchTG('tg2')">TG2 — Engagement</div>
        <div class="sf-step" id="btn-tg3" onclick="switchTG('tg3')">TG3 — MQL</div>
        <div class="sf-step" id="btn-tg4" onclick="switchTG('tg4')">TG4 — Clasificación</div>
        <div class="sf-step" id="btn-tg5" onclick="switchTG('tg5')">TG5 — WO</div>
    </div>

    <!-- Sub-pestanas -->
    <div class="sf-tabs">
        <div class="sf-tab active">Detalles de Oportunidad</div>
        <div class="sf-tab">Historial y Actividad</div>
        <div class="sf-tab">Chatter / Notas</div>
    </div>

    {% if mensaje %}
        <div class="alert-success">{{ mensaje }}</div>
    {% endif %}

    <form method="POST">
        <div class="sf-body">

            <!-- PANTALLA TG0 -->
            <div id="pantalla-tg0" class="tg-screen">
                <div class="sf-section-card">
                    <div class="sf-section-header">▼ TG0 — Datos del Contacto</div>
                    <div class="sf-grid">
                        <div class="sf-field">
                            <label class="sf-label"><span class="required">*</span>Nombre / Apellido</label>
                            <input type="text" name="nombre" class="sf-input" placeholder="Nombre completo del contacto" required>
                            <span class="sf-help-text">Campo estándar de Salesforce.</span>
                        </div>
                        <div class="sf-field">
                            <label class="sf-label"><span class="required">*</span>Empresa / Razón Social</label>
                            <input type="text" name="empresa" class="sf-input" placeholder="Nombre de la empresa" required>
                        </div>
                        <div class="sf-field">
                            <label class="sf-label"><span class="required">*</span>Cargo / Título</label>
                            <input type="text" name="cargo" class="sf-input" placeholder="Posición en la empresa" required>
                        </div>
                        <div class="sf-field">
                            <label class="sf-label"><span class="required">*</span>Email Corporativo</label>
                            <input type="email" name="email" class="sf-input" placeholder="correo@empresa.com" required>
                        </div>
                        <div class="sf-field">
                            <label class="sf-label">Teléfono Contacto</label>
                            <input type="tel" name="telefono" class="sf-input" placeholder="+52 81 0000 0000">
                        </div>
                        <div class="sf-field">
                            <label class="sf-label"><span class="required">*</span>País / Región</label>
                            <select name="pais_region" class="sf-select" required>
                                <option value="">--Ninguno--</option>
                                <option value="Mexico">Mexico</option>
                                <option value="USA">USA</option>
                            </select>
                        </div>
                    </div>
                </div>

                <div class="sf-section-card">
                    <div class="sf-section-header">▼ Origen y Asignación</div>
                    <div class="sf-grid">
                        <div class="sf-field" style="grid-column: span 2;">
                            <label class="sf-label">Notas Adicionales</label>
                            <textarea name="notas_tg0" class="sf-textarea" rows="3" placeholder="Notas generales sobre el lead..."></textarea>
                        </div>
                    </div>
                </div>
            </div>

            <!-- PANTALLA TG1 -->
            <div id="pantalla-tg1" class="tg-screen" style="display:none;">
                <div class="sf-section-card">
                    <div class="sf-section-header">▼ TG1 — Segmentación y Fit</div>
                    <div class="sf-grid">
                        <div class="sf-field">
                            <label class="sf-label"><span class="required">*</span>Macro Segmento</label>
                            <select name="macro_segmento" class="sf-select">
                                <option value="">--Ninguno--</option>
                                <option value="Industrial">Industrial</option>
                                <option value="Automotriz">Automotriz</option>
                                <option value="Alimentos">Alimentos y Bebidas</option>
                            </select>
                        </div>
                        <div class="sf-field">
                            <label class="sf-label">Sub-Segmento</label>
                            <select name="sub_segmento" class="sf-select">
                                <option value="">--Ninguno--</option>
                                <option value="Manufactura">Manufactura de plásticos</option>
                                <option value="Embalaje">Embalaje</option>
                            </select>
                        </div>
                        <div class="sf-field">
                            <label class="sf-label"><span class="required">*</span>Geografía</label>
                            <select name="geografia" class="sf-select">
                                <option value="">--Ninguno--</option>
                                <option value="Norte">Norte</option>
                                <option value="Centro">Centro</option>
                                <option value="Bajio">Bajío</option>
                                <option value="Occidente">Occidente</option>
                                <option value="Golfo">Golfo</option>
                            </select>
                        </div>
                        <div class="sf-field">
                            <label class="sf-label"><span class="required">*</span>Relevancia del Portafolio</label>
                            <select name="relevancia" class="sf-select">
                                <option value="">--Ninguno--</option>
                                <option value="Alta">Alta</option>
                                <option value="Media">Media</option>
                                <option value="Baja">Baja</option>
                            </select>
                        </div>
                        <div class="sf-field">
                            <label class="sf-label"><span class="required">*</span>Tamaño de Empresa</label>
                            <select name="tamano_empresa" class="sf-select">
                                <option value="">--Ninguno--</option>
                                <option value="Micro — <$10M">Micro — <$10M</option>
                                <option value="Pequeña — $10-50M">Pequeña — $10-50M</option>
                                <option value="Mediana — $50-200M">Mediana — $50-200M</option>
                                <option value="Grande — $200M-$1B">Grande — $200M-$1B</option>
                                <option value="Enterprise — >$1B">Enterprise — >$1B</option>
                            </select>
                        </div>
                        <div class="sf-field">
                            <label class="sf-label"><span class="required">*</span>Banda Asignada</label>
                            <select name="banda" class="sf-select">
                                <option value="">--Ninguno--</option>
                                <option value="K">K</option>
                                <option value="A">A</option>
                                <option value="B">B</option>
                                <option value="C">C</option>
                                <option value="D">D</option>
                            </select>
                            <span class="sf-help-text">Si es K o A, se clasifica como BSV Normal.</span>
                        </div>
                    </div>
                </div>
            </div>

        </div>

        <!-- Botones de Acción -->
        <div class="sf-footer">
            <button type="button" class="sf-btn sf-btn-secondary">Cancelar</button>
            <button type="submit" class="sf-btn sf-btn-primary">Guardar Prospecto</button>
        </div>
    </form>
</div>

<script>
    function switchTG(tgId) {
        // Ocultar todas las pantallas
        document.querySelectorAll('.tg-screen').forEach(el => el.style.display = 'none');
        // Quitar estado activo de los chevrons
        document.querySelectorAll('.sf-step').forEach(el => el.classList.remove('active'));
        
        // Mostrar la pantalla seleccionada
        document.getElementById('pantalla-' + tgId).style.display = 'block';
        document.getElementById('btn-' + tgId).classList.add('active');
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
        empresa = request.form.get('empresa')
        geografia = request.form.get('geografia')

        try:
            # Conexión directa a la Lista de SharePoint
            ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
            target_list = ctx.web.lists.get_by_title("BSV_Leads")
            
            target_list.add_item({
                "Title": nombre,
                "BSV_Empresa___Razon_Social__c": empresa,
                "BSV_Geografia__c": geografia
            })
            ctx.execute_query()
            mensaje = "¡Prospecto guardado exitosamente en la base de datos de SharePoint!"
        except Exception as e:
            mensaje = f"Error al guardar: {str(e)}"

    return render_template_string(HTML_SALESFORCE, mensaje=mensaje)

if __name__ == '__main__':
    app.run(debug=True)
