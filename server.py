from flask import Flask, render_template_string, request, send_file
import pandas as pd
import os
import io

app = Flask(__name__)

def listar_excels():
    return [f for f in os.listdir('.') if f.endswith('.xlsx') and os.path.isfile(f)]

@app.route("/", methods=["GET", "POST"])
def mostrar_excel():
    archivos = listar_excels()
    archivo_seleccionado = None
    tabla_html = ""
    fecha_inicio = ""
    fecha_fin = ""
    df_filtrado = None

    if request.method == "POST":
        archivo_seleccionado = request.form.get("archivo")
        fecha_inicio = request.form.get("fecha_inicio")
        fecha_fin = request.form.get("fecha_fin")

        if archivo_seleccionado in archivos:
            try:
                df = pd.read_excel(archivo_seleccionado, engine="openpyxl")
                # Intentar convertir la primera columna a fechas
                df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], errors='coerce')

                # Aplicar filtro si las fechas están presentes
                if fecha_inicio:
                    df = df[df.iloc[:, 0] >= pd.to_datetime(fecha_inicio)]
                if fecha_fin:
                    df = df[df.iloc[:, 0] <= pd.to_datetime(fecha_fin)]

                df_filtrado = df
                tabla_html = df.to_html(classes="table table-striped", index=False)

            except Exception as e:
                tabla_html = f"<div class='alert alert-danger'>Error al leer el archivo Excel: {e}</div>"
        else:
            tabla_html = "<div class='alert alert-warning'>Archivo no válido seleccionado.</div>"

    return render_template_string("""
        <html>
            <head>
                <title>Ver archivos Excel con fechas</title>
                <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
            </head>
            <body class="p-4">
                <h1>Archivos Excel disponibles</h1>

                <form method="POST" action="/">
                    <div class="mb-3">
                        <label for="archivo" class="form-label">Selecciona un archivo Excel:</label>
                        <select name="archivo" id="archivo" class="form-select" required>
                            <option value="" disabled {{ 'selected' if not archivo_seleccionado else '' }}>-- Elegí un archivo --</option>
                            {% for archivo in archivos %}
                                <option value="{{ archivo }}" {% if archivo == archivo_seleccionado %}selected{% endif %}>{{ archivo }}</option>
                            {% endfor %}
                        </select>
                    </div>

                    <div class="mb-3 row">
                        <div class="col">
                            <label for="fecha_inicio" class="form-label">Fecha desde:</label>
                            <input type="date" class="form-control" name="fecha_inicio" value="{{ fecha_inicio }}">
                        </div>
                        <div class="col">
                            <label for="fecha_fin" class="form-label">Fecha hasta:</label>
                            <input type="date" class="form-control" name="fecha_fin" value="{{ fecha_fin }}">
                        </div>
                    </div>

                    <button type="submit" class="btn btn-primary">Mostrar contenido</button>
                </form>

                {% if archivo_seleccionado and tabla_html %}
                    <form method="POST" action="/descargar" style="margin-top:20px;">
                        <input type="hidden" name="archivo" value="{{ archivo_seleccionado }}">
                        <input type="hidden" name="fecha_inicio" value="{{ fecha_inicio }}">
                        <input type="hidden" name="fecha_fin" value="{{ fecha_fin }}">
                        <button type="submit" class="btn btn-success">Descargar archivo filtrado</button>
                    </form>
                    <hr>
                    <h2>Contenido de {{ archivo_seleccionado }}</h2>
                    <div>{{ tabla_html | safe }}</div>
                {% endif %}
            </body>
        </html>
    """, archivos=archivos, archivo_seleccionado=archivo_seleccionado, tabla_html=tabla_html, 
       fecha_inicio=fecha_inicio, fecha_fin=fecha_fin)

@app.route("/descargar", methods=["POST"])
def descargar_excel():
    archivo = request.form.get("archivo")
    fecha_inicio = request.form.get("fecha_inicio")
    fecha_fin = request.form.get("fecha_fin")
    archivos = listar_excels()

    if archivo not in archivos:
        return "Archivo no válido o no encontrado.", 400
    try:
        df = pd.read_excel(archivo, engine="openpyxl")
        df.iloc[:, 0] = pd.to_datetime(df.iloc[:, 0], errors='coerce')

        if fecha_inicio:
            df = df[df.iloc[:, 0] >= pd.to_datetime(fecha_inicio)]
        if fecha_fin:
            df = df[df.iloc[:, 0] <= pd.to_datetime(fecha_fin)]

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=f"filtrado_{archivo}",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return f"Error al generar el archivo: {e}", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
