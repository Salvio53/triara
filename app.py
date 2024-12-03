from flask import Flask, request, jsonify, render_template
import pandas as pd

app = Flask(__name__)

# Ruta del archivo Excel
ruta_archivo = r"C:\Users\salvi\Documents\Proyectos\Gestión interactivo.xlsm"

@app.route('/')
def home():
    return render_template('index.html')  # Servir la interfaz HTML

@app.route('/buscar', methods=['POST'])
def buscar():
    try:
        datos = request.json
        hoja = datos.get("hoja")
        palabra_clave = datos.get("palabra_clave", "").strip().lower()

        if not hoja or not palabra_clave:
            return jsonify({"error": "Faltan datos para realizar la búsqueda"}), 400

        # Cargar la hoja correspondiente del archivo Excel
        df = pd.read_excel(ruta_archivo, sheet_name=hoja)

        # Filtrar datos por palabra clave
        mask = df.apply(lambda column: column.astype(str).str.contains(palabra_clave, case=False, na=False))
        resultados = df[mask.any(axis=1)]

        # Convertir resultados a formato JSON
        return jsonify({"resultados": resultados.to_dict(orient="records")})

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
