from flask import Flask, render_template, request, redirect, url_for
import os
import pandas as pd
import json

app = Flask(__name__)

# Directorio de archivos Excel
BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "materiales")

# Variable global para almacenar resultados de cada flujo
materiales_finales = []

# Función auxiliar para renombrar columnas
def renombrar_columnas(df):
    df_renombrado = df.rename(
        columns={
            "1. Cód.SAP": "Cód.SAP",
            "2. MATERIAL": "MATERIAL",
            "3. Descripción": "Descripción",
            "5.CONDICIÓN": "CONDICIÓN"
        }
    )
    columnas = ["Cód.SAP", "MATERIAL", "Descripción", "4.CANTIDAD", "CONDICIÓN"]
    columnas_presentes = [col for col in columnas if col in df_renombrado.columns]
    return df_renombrado[columnas_presentes]

# ===================================
# Página de Inicio
# ===================================
@app.route("/")
def index():
    global materiales_finales
    materiales_finales = []  # Reinicia la lista para cada nueva corrida
    return render_template("index.html")


# ===================================
# FLUJO B: Saca Tubing
# ===================================
@app.route("/flujo_b", methods=["GET", "POST"])
def flujo_b():
    if request.method == "POST":
        saca = request.form.get("saca_tubing")
        if saca == "NO":
            return redirect(url_for("flujo_c"))
        elif saca == "SI":
            return redirect(url_for("flujo_b_seleccion"))
        else:
            return "Selecciona una opción.", 400
    return render_template("flujo_b.html")

@app.route("/flujo_b/seleccion", methods=["GET", "POST"])
def flujo_b_seleccion():
    try:
        file_path = os.path.join(BASE_DIR, "saca tubing.xlsx")
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
        unique_diametros = sorted([d for d in df["DIÁMETRO"].dropna().unique() if d.upper() != "TODOS"])
    except Exception as e:
        return f"Error al cargar el Excel: {e}"
    if request.method == "POST":
        selected = request.form.getlist("diametros")
        if not selected:
            return "Selecciona al menos un DIÁMETRO.", 400
        diametros_str = ",".join(selected)
        return redirect(url_for("flujo_b_cantidades", diametros=diametros_str))
    else:
        return render_template("flujo_b_seleccion.html", unique_diametros=unique_diametros)

@app.route("/flujo_b/cantidades", methods=["GET", "POST"])
def flujo_b_cantidades():
    diametros_str = request.args.get("diametros", "")
    selected = diametros_str.split(",") if diametros_str else []
    if request.method == "POST":
        quantities = {}
        for diam in selected:
            qty = request.form.get(f"qty_{diam}", type=float)
            quantities[diam] = qty
        try:
            file_path = os.path.join(BASE_DIR, "saca tubing.xlsx")
            df = pd.read_excel(file_path)
            df.columns = df.columns.str.strip()
        except Exception as e:
            return f"Error: {e}"
        df_filtered = df[(df["DIÁMETRO"].isin(selected)) | (df["DIÁMETRO"].str.upper() == "TODOS")].copy()
        for diam, qty in quantities.items():
            mask = (df_filtered["DIÁMETRO"] == diam) & (df_filtered["4.CANTIDAD"].isna())
            df_filtered.loc[mask, "4.CANTIDAD"] = qty
        df_filtered_renombrado = renombrar_columnas(df_filtered)
        materiales_finales.append(("FLUJO B", df_filtered_renombrado))
        return redirect(url_for("flujo_c"))
    else:
        return render_template("flujo_b_cantidades.html", selected_diametros=selected)

# ===================================
# FLUJO C: Tubería de Baja
# ===================================
from flask import json  # (o usa json estándar)

@app.route("/flujo_c", methods=["GET"])
def flujo_c():
    # Muestra la página inicial del Flujo C (pregunta si baja tubing)
    return render_template("flujo_c.html")

@app.route("/flujo_c/decidir", methods=["POST"])
def flujo_c_decidir():
    baja = request.form.get("baja_tubing")
    if baja == "NO":
        # Si el usuario responde NO, se continúa (por ejemplo, al Flujo D)
        return redirect(url_for("flujo_d"))
    elif baja == "SI":
        return redirect(url_for("flujo_c_seleccion"))
    else:
        return "Selecciona una opción.", 400

@app.route("/flujo_c/seleccion", methods=["GET", "POST"])
def flujo_c_seleccion():
    try:
        file_path = os.path.join(BASE_DIR, "baja tubing.xlsx")
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
        # Se extraen los DIÁMETRO únicos (excluyendo "TODOS")
        unique_diametros = sorted([x for x in df["DIÁMETRO"].dropna().unique() if x != "TODOS"])
    except Exception as e:
        return f"Error al cargar el Excel: {e}"
    
    if request.method == "POST":
        selected = request.form.getlist("diametros")
        if not selected:
            return "Selecciona al menos un DIÁMETRO.", 400
        diametros_str = ",".join(selected)
        return redirect(url_for("flujo_c_tipo", diametros=diametros_str))
    else:
        return render_template("flujo_c_seleccion.html", unique_diametros=unique_diametros)

@app.route("/flujo_c/tipo", methods=["GET", "POST"])
def flujo_c_tipo():
    diametros_str = request.args.get("diametros", "")
    selected_diametros = diametros_str.split(",") if diametros_str else []
    try:
        file_path = os.path.join(BASE_DIR, "baja tubing.xlsx")
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
    except Exception as e:
        return f"Error: {e}"
    # Para cada DIÁMETRO, extraemos las opciones de TIPO (excluyendo "TODOS")
    filtros = {}
    for diam in selected_diametros:
        subset = df[df["DIÁMETRO"] == diam]
        tipos = sorted([x for x in subset["TIPO"].dropna().unique() if x != "TODOS"])
        if not tipos:
            tipos = ["TODOS"]
        filtros[diam] = tipos
    if request.method == "POST":
        selected_tipos_dict = {}
        for diam in selected_diametros:
            sel = request.form.getlist(f"tipo_{diam}")
            if not sel:
                sel = ["TODOS"]
            selected_tipos_dict[diam] = sel
        tipos_json = json.dumps(selected_tipos_dict)
        return redirect(url_for("flujo_c_diacsg", diametros=diametros_str, tipos=tipos_json))
    else:
        return render_template("flujo_c_tipo.html", selected_diametros=selected_diametros, filtros=filtros)

@app.route("/flujo_c/diacsg", methods=["GET", "POST"])
def flujo_c_diacsg():
    diametros_str = request.args.get("diametros", "")
    selected_diametros = diametros_str.split(",") if diametros_str else []
    tipos_json = request.args.get("tipos", "{}")
    selected_tipos_dict = json.loads(tipos_json)
    try:
        file_path = os.path.join(BASE_DIR, "baja tubing.xlsx")
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
    except Exception as e:
        return f"Error: {e}"
    # Se calcula la unión de los TIPO seleccionados, agregando "TODOS"
    union_tipos = set()
    for sel in selected_tipos_dict.values():
        if sel == ["TODOS"]:
            union_tipos.add("TODOS")
        else:
            union_tipos.update(sel)
            union_tipos.add("TODOS")
    diam_filter = ["TODOS"] if selected_diametros == ["TODOS"] else selected_diametros + ["TODOS"]
    df_filtered = df[df["DIÁMETRO"].isin(diam_filter) & df["TIPO"].isin(union_tipos)]
    unique_csg = sorted([x for x in df_filtered["DIÁMETRO CSG"].dropna().unique() if x != "TODOS"])
    if not unique_csg:
        # Si no hay valores para DIÁMETRO CSG, se continúa automáticamente usando "TODOS"
        return redirect(url_for("flujo_c_cantidades", diametros=diametros_str, tipos=tipos_json, diacsg="TODOS"))
    if request.method == "POST":
        selected_csg = request.form.get("diacsg")
        if not selected_csg:
            selected_csg = "TODOS"
        # Se fuerza el filtrado: si se selecciona un valor, se usa [valor, "TODOS"]
        return redirect(url_for("flujo_c_cantidades", diametros=diametros_str, tipos=tipos_json, diacsg=selected_csg))
    else:
        return render_template("flujo_c_diacsg.html", unique_csg=unique_csg)

@app.route("/flujo_c/cantidades", methods=["GET", "POST"])
def flujo_c_cantidades():
    diametros_str = request.args.get("diametros", "")
    selected_diametros = diametros_str.split(",") if diametros_str else []
    tipos_json = request.args.get("tipos", "{}")
    selected_tipos_dict = json.loads(tipos_json)
    # Aquí se recibe el valor seleccionado en DIÁMETRO CSG; se utiliza para filtrar
    diacsg = request.args.get("diacsg", "TODOS")
    try:
        file_path = os.path.join(BASE_DIR, "baja tubing.xlsx")
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
    except Exception as e:
        return f"Error: {e}"
    if request.method == "POST":
        quantities = {}
        for diam in selected_diametros:
            for tipo in selected_tipos_dict.get(diam, []):
                qty = request.form.get(f"qty_{diam}_{tipo}", type=float)
                quantities[(diam, tipo)] = qty
        # Ahora, se aplica el filtrado incluyendo DIÁMETRO, TIPO y DIÁMETRO CSG
        for (diam, tipo), qty in quantities.items():
            condition = (
                df["DIÁMETRO"].isin([diam, "TODOS"]) &
                df["TIPO"].isin([tipo, "TODOS"]) &
                df["DIÁMETRO CSG"].isin([diacsg, "TODOS"])
            )
            df.loc[condition & df["4.CANTIDAD"].isna(), "4.CANTIDAD"] = qty
        final_condition = pd.Series([False]*len(df))
        for diam_value, fdict in selected_tipos_dict.items():
            temp = pd.Series([False]*len(df))
            for tipo_val in fdict:
                cond = (
                    df["DIÁMETRO"].isin([diam_value, "TODOS"]) &
                    df["TIPO"].isin([tipo_val, "TODOS"]) &
                    df["DIÁMETRO CSG"].isin([diacsg, "TODOS"])
                )
                temp = temp | cond
            final_condition = final_condition | temp
        final_df = df[final_condition]
        final_df_renombrado = renombrar_columnas(final_df)
        materiales_finales.append(("FLUJO C", final_df_renombrado))
        return redirect(url_for("flujo_d"))
    else:
        # Prepara una lista de combinaciones para mostrar los campos de cantidad
        combos = []
        for diam in selected_diametros:
            for tipo in selected_tipos_dict.get(diam, []):
                combos.append((diam, tipo))
        return render_template("flujo_c_cantidades.html", combos=combos)

# ===================================
# FLUJO D: Profundiza
# ===================================
@app.route("/flujo_d", methods=["GET"])
def flujo_d():
    return render_template("flujo_d.html")

@app.route("/flujo_d/decidir", methods=["POST"])
def flujo_d_decidir():
    profundizar = request.form.get("profundizar")
    if profundizar == "NO":
        return redirect(url_for("flujo_e"))
    elif profundizar == "SI":
        return redirect(url_for("flujo_d_seleccion"))
    else:
        return "Selecciona una opción.", 400

@app.route("/flujo_d/seleccion", methods=["GET", "POST"])
def flujo_d_seleccion():
    file_path = os.path.join(BASE_DIR, "profundiza.xlsx")
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].astype(str).str.strip()
        # Reemplazar valores 'nan'
        if 'TIPO' in df.columns:
            df['TIPO'] = df['TIPO'].replace('nan', '').fillna('')
        if 'DIÁMETRO CSG' in df.columns:
            df['DIÁMETRO CSG'] = df['DIÁMETRO CSG'].replace('nan', '').fillna('')
    except Exception as e:
        return f"Error al cargar el Excel: {e}"
    
    # Elegir la columna a usar: si existe "DIÁMETRO", se usa; sino, "DIÁMETRO CSG"
    if "DIÁMETRO" in df.columns:
        col = "DIÁMETRO"
    elif "DIÁMETRO CSG" in df.columns:
        col = "DIÁMETRO CSG"
    else:
        return "La columna de DIÁMETRO no se encontró en el Excel."
    
    unique_values = sorted(df[col].dropna().unique().tolist())
    if request.method == "POST":
        selected = request.form.getlist("valores")
        if not selected:
            return "No se seleccionaron valores.", 400
        selected_str = ",".join(selected)
        # Redirigir a la página de ingreso de cantidades, pasando la columna y los valores seleccionados
        return redirect(url_for("flujo_d_cantidades", valores=selected_str, col=col))
    else:
        return render_template("flujo_d_seleccion.html", unique_values=unique_values, col=col)

@app.route("/flujo_d/cantidades", methods=["GET", "POST"])
def flujo_d_cantidades():
    valores_str = request.args.get("valores", "")
    col = request.args.get("col", "")
    selected_values = valores_str.split(",") if valores_str else []
    file_path = os.path.join(BASE_DIR, "profundiza.xlsx")
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].astype(str).str.strip()
    except Exception as e:
        return f"Error al cargar el Excel: {e}"
    if request.method == "POST":
        quantities = {}
        for val in selected_values:
            qty = request.form.get(f"qty_{val}", type=float)
            quantities[val] = qty
        # Filtrar el DataFrame según la columna y los valores seleccionados
        filtered_df = df[df[col].isin(selected_values)]
        for val, qty in quantities.items():
            mask = (filtered_df[col] == val) & (filtered_df["4.CANTIDAD"].isna())
            filtered_df.loc[mask, "4.CANTIDAD"] = qty
        final_df_renombrado = renombrar_columnas(filtered_df)
        materiales_finales.append(("FLUJO D", final_df_renombrado))
        return redirect(url_for("flujo_e"))
    else:
        # Preparar lista de valores para mostrar los campos de cantidad
        combos = selected_values  # Cada valor seleccionado se usará para ingresar cantidad
        return render_template("flujo_d_cantidades.html", combos=combos, col=col)

# ===================================
# FLUJO E: Baja BES
# ===================================
@app.route("/flujo_e", methods=["GET"])
def flujo_e():
    return render_template("flujo_e.html")

@app.route("/flujo_e/decidir", methods=["POST"])
def flujo_e_decidir():
    baja_varilla = request.form.get("baja_BES")
    if baja_varilla == "NO":
        return redirect(url_for("flujo_f"))
    elif baja_varilla == "SI":
        return redirect(url_for("flujo_h"))
    else:
        return "Selecciona una opción.", 400


# ===================================
# FLUJO F: Abandona pozo (Flujo F)
# ===================================

# Página inicial: pregunta "¿Abandono/recupero?"
@app.route("/flujo_f", methods=["GET"])
def flujo_f():
    return render_template("flujo_f.html")

@app.route("/flujo_f/decidir", methods=["POST"])
def flujo_f_decidir():
    abandono = request.form.get("abandono")
    if abandono == "NO":
        # Si responde NO, se salta este flujo y se pasa al siguiente (Flujo G o H)
        return redirect(url_for("flujo_h"))
    elif abandono == "SI":
        return redirect(url_for("flujo_f_filtros"))
    else:
        return "Selecciona una opción.", 400

# Ruta para configurar los filtros (DIÁMETRO y DIÁMETRO CSG)
@app.route("/flujo_f/filtros", methods=["GET", "POST"])
def flujo_f_filtros():
    file_path = os.path.join(BASE_DIR, "abandono-recupero.xlsx")
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
    except Exception as e:
        return f"Error al cargar el Excel: {e}"
    # Verificar que existan las columnas requeridas
    if "DIÁMETRO" not in df.columns:
        return "La columna 'DIÁMETRO' no se encontró en el Excel."
    if "DIÁMETRO CSG" not in df.columns:
        return "La columna 'DIÁMETRO CSG' no se encontró en el Excel."
    # Construir opciones internas (con "TODOS") y opciones para mostrar (sin "TODOS")
    all_diametros = list(df["DIÁMETRO"].dropna().unique())
    opciones_diam = [d for d in all_diametros if d.upper() != "TODOS"]
    # Para el filtrado interno se usa "TODOS" (si no se selecciona nada, se asume "TODOS")
    all_diametros_csg = list(df["DIÁMETRO CSG"].dropna().unique())
    opciones_diacsg = [d for d in all_diametros_csg if d.upper() != "TODOS"]
    
    if request.method == "POST":
        selected_diametros = request.form.getlist("diametros")
        # Si el usuario no selecciona nada, se usará "TODOS" internamente
        if not selected_diametros:
            selected_diametros = ["TODOS"]
        else:
            # Se añade "TODOS" internamente para ampliar el filtrado
            selected_diametros.append("TODOS")
        selected_diacsg = request.form.get("diacsg")
        if not selected_diacsg:
            selected_diacsg = "TODOS"
        # Se preparan los datos (en este ejemplo, simplemente se pasan en query string)
        filtros = {"diametros": selected_diametros, "diacsg": selected_diacsg}
        filtros_json = json.dumps(filtros)
        diametros_str = ",".join(selected_diametros)
        return redirect(url_for("flujo_f_cantidades", diametros=diametros_str, filtros=filtros_json))
    else:
        return render_template("flujo_f_filtros.html", 
                               opciones_diam=opciones_diam, 
                               opciones_diacsg=opciones_diacsg)

# Ruta para ingresar cantidades
@app.route("/flujo_f/cantidades", methods=["GET", "POST"])
def flujo_f_cantidades():
    diametros_str = request.args.get("diametros", "")
    # La lista interna contiene "TODOS", pero para mostrar los inputs se excluye "TODOS"
    selected_diametros = diametros_str.split(",") if diametros_str else []
    display_diametros = [d for d in selected_diametros if d.upper() != "TODOS"]
    filtros_json = request.args.get("filtros", "{}")
    # Para este flujo se usará el Excel "abandono-recupero.xlsx"
    file_path = os.path.join(BASE_DIR, "abandono-recupero.xlsx")
    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
    except Exception as e:
        return f"Error al cargar el Excel: {e}"
    if request.method == "POST":
        quantities = {}
        for diam in display_diametros:
            qty = request.form.get(f"qty_{diam}", type=float)
            quantities[diam] = qty
        filtered_df = df[df["DIÁMETRO"].isin(selected_diametros)]
        for diam, qty in quantities.items():
            mask = (filtered_df["DIÁMETRO"] == diam) & (filtered_df["4.CANTIDAD"].isna())
            filtered_df.loc[mask, "4.CANTIDAD"] = qty
        final_df_renombrado = renombrar_columnas(filtered_df)
        materiales_finales.append(("FLUJO F", final_df_renombrado))
        # En lugar de imprimir, se guarda para consolidar al final
        return redirect(url_for("flujo_h"))
    else:
        if not display_diametros:
            return "No se encontraron DIÁMETRO seleccionados.", 400
        return render_template("flujo_f_cantidades.html", selected_diametros=display_diametros)



# ===================================
# FLUJO H: Material de agregación
# ===================================
@app.route("/flujo_h", methods=["GET"])
def flujo_h():
    return render_template("flujo_h.html")

@app.route("/flujo_h/decidir", methods=["POST"])
def flujo_h_decidir():
    respuesta = request.form.get("agregar_material")
    if respuesta == "SI":
        return redirect(url_for("flujo_h_seleccion"))
    elif respuesta == "NO":
        return redirect(url_for("flujo_final"))
    else:
        return "Selecciona una opción.", 400

@app.route("/flujo_h/seleccion", methods=["GET", "POST"])
def flujo_h_seleccion():
    file_path_H = os.path.join(BASE_DIR, "GENERAL(1).xlsx")
    try:
        df_H = pd.read_excel(file_path_H)
        df_H.columns = df_H.columns.str.strip()
        # Si no existe la columna "4.CANTIDAD", se crea con 0
        if "4.CANTIDAD" not in df_H.columns:
            df_H["4.CANTIDAD"] = 0
        else:
            df_H["4.CANTIDAD"] = pd.to_numeric(df_H["4.CANTIDAD"], errors="coerce")
    except Exception as e:
        return f"Error al cargar el Excel: {e}"
    
    # Extraer la lista de materiales de la columna "2. MATERIAL"
    if "2. MATERIAL" in df_H.columns:
        materiales = df_H["2. MATERIAL"].astype(str).unique().tolist()
    else:
        materiales = []
    # Se genera la tabla HTML para que el usuario la consulte (opcional)
    #df_html = df_H.to_html(classes="table table-bordered", index=False)
    if request.method == "POST":
        seleccionados = request.form.getlist("materiales")
        if not seleccionados:
            return "No se seleccionó ningún material.", 400
        seleccion_str = ",".join(seleccionados)
        return redirect(url_for("flujo_h_cantidades", materiales=seleccion_str))
    else:
        return render_template("flujo_h_seleccion.html", materiales=materiales)

@app.route("/flujo_h/cantidades", methods=["GET", "POST"])
def flujo_h_cantidades():
    materiales_str = request.args.get("materiales", "")
    seleccionados = materiales_str.split(",") if materiales_str else []
    file_path_H = os.path.join(BASE_DIR, "GENERAL(1).xlsx")
    try:
        df_H = pd.read_excel(file_path_H)
        df_H.columns = df_H.columns.str.strip()
        # Crear o convertir la columna "4.CANTIDAD"
        if "4.CANTIDAD" not in df_H.columns:
            df_H["4.CANTIDAD"] = 0
        else:
            df_H["4.CANTIDAD"] = pd.to_numeric(df_H["4.CANTIDAD"], errors="coerce")
    except Exception as e:
        return f"Error al cargar el Excel: {e}"
    if request.method == "POST":
        quantities = {}
        for mat in seleccionados:
            qty = request.form.get(f"qty_{mat}", type=float)
            quantities[mat] = qty
        # Actualizamos el DataFrame: para cada material seleccionado, asignar la cantidad en filas sin valor
        for mat, qty in quantities.items():
            mask = (df_H["2. MATERIAL"].astype(str) == mat) & ((df_H["4.CANTIDAD"].isna()) | (df_H["4.CANTIDAD"] <= 0))
            df_H.loc[mask, "4.CANTIDAD"] = qty
        # Filtramos solo los materiales con cantidad mayor que 0
        assigned_df = df_H[df_H["2. MATERIAL"].astype(str).isin(seleccionados) & (df_H["4.CANTIDAD"] > 0)]
        if not assigned_df.empty:
            assigned_df_renombrado = renombrar_columnas(assigned_df)
            # Guardamos el resultado del Flujo H en la variable global
            materiales_finales.append(("FLUJO H", assigned_df_renombrado))
        else:
            print("No se asignaron cantidades (o todas fueron 0).")
        return redirect(url_for("flujo_final"))
    else:
        if not seleccionados:
            return "No se encontraron materiales seleccionados.", 400
        return render_template("flujo_h_cantidades.html", materiales=seleccionados)

# ===================================
# FLUJO FINAL: Resultados Consolidados
# ===================================
@app.route("/flujo_final", methods=["GET"])
def flujo_final():
    return render_template("flujo_final.html", materiales_finales=materiales_finales)

#====================================
# EXPORTAR AL EXCEL
#====================================

@app.route("/export_excel")
def export_excel():
    import io
    from flask import send_file

    # Combina todos los DataFrames en uno solo, agregando una columna que indique el flujo
    combined_df = pd.concat([df.assign(Flujo=flow) for flow, df in materiales_finales], ignore_index=True)
    
    # Crea un buffer en memoria para guardar el Excel
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    
    # Escribe el DataFrame combinado en una sola hoja
    combined_df.to_excel(writer, sheet_name="Materiales Consolidados", index=False)
    writer.save()
    output.seek(0)
    
    return send_file(
        output,
        attachment_filename="materiales_consolidados.xlsx",
        as_attachment=True,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == "__main__":
    app.run(debug=True)
