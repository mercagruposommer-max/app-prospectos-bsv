from flask import Flask, request, render_template_string
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

app = Flask(__name__)

# --- CONFIGURACIÓN DE SHAREPOINT ---
# Reemplaza estas 3 variables con los datos reales de tu empresa
SITE_URL = "https://tu-empresa.sharepoint.com/sites/tu-sitio"
USERNAME = "usuario@tu-empresa.com"
PASSWORD = "TuPassword123"

# --- DISEÑO DEL FORMULARIO (HTML/CSS) ---
HTML_FORM = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Captura de Prospectos BSV</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; margin: 0; padding: 20px; }
        .card { background: white; padding: 25px; border-radius: 8px; max-width: 500px; margin: 0 auto; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        h2 { color: #0070d2; margin-top: 0; border-bottom: 2px solid #eef1f6; padding-bottom: 10px; }
        label { font-weight: 600; display: block; margin-top: 15px; color: #333; }
        input, select { width: 100%; padding: 10px; margin-top: 5px; box-sizing: border-box; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }
        button { background: #0070d2; color: white; border: none; padding: 12px; margin-top: 25px; border-radius: 4px; font-size: 16px; font-weight: bold; cursor: pointer; width: 100%; }
        button:hover { background: #005fb2; }
        .alert { background: #d4edda; color: #155724; padding: 10px; border-radius: 4px; margin-bottom: 15px; border: 1px solid #c3e6cb; }
    </style>
</head>
<body>
    <div class="card">
        <h2>Captura de Prospectos BSV (TG0 / TG1)</h2>
        
        {% if mensaje %}
            <div class="alert">{{ mensaje }}</div>
        {% endif %}

        <form method="POST">
            <label>Nombre Completo (TG0):</label>
            <input type="text" name="nombre" placeholder="Ej. Juan Pérez" required>

            <label>Empresa / Razón Social (TG0):</label>
            <input type="text" name="empresa" placeholder="Ej. Acme Corp" required>

            <label>Geografía (TG1):</label>
            <select name="geografia" required>
                <option value="Norte">Norte</option>
                <option value="Centro">Centro</option>
                <option value="Bajio">Bajío</option>
            </select>

            <button type="submit">Guardar en SharePoint</button>
        </form>
    </div>
</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def home():
    mensaje = None
    if request.method == 'POST':
        # 1. Obtener datos llenados en el formulario
        nombre = request.form.get('nombre')
        empresa = request.form.get('empresa')
        geografia = request.form.get('geografia')

        try:
            # 2. Conectar a SharePoint
            ctx = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
            target_list = ctx.web.lists.get_by_title("BSV_Leads")
            
            # 3. Guardar el registro directamente en la Lista de SharePoint
            target_list.add_item({
                "Title": nombre,
                "BSV_Empresa___Razon_Social__c": empresa,
                "BSV_Geografia__c": geografia
            })
            ctx.execute_query()
            mensaje = "¡Prospecto guardado exitosamente en la base de datos de SharePoint!"
        except Exception as e:
            mensaje = f"Error al guardar en SharePoint: {str(e)}"

    return render_template_string(HTML_FORM, mensaje=mensaje)

if __name__ == '__main__':
    app.run(debug=True)