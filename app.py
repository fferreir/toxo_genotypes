from flask import Flask, render_template_string, request
import pandas as pd
import os

app = Flask(__name__)

# Caminho absoluto para o arquivo Excel no servidor
FILE_PATH = 'caminho/para/excel'

def get_data():
    if not os.path.exists(FILE_PATH):
        return None, None
    
    # Lê a planilha original
    df = pd.read_excel(FILE_PATH, header=None, engine='openpyxl')
    
    # Selecionamos as colunas de interesse: 0 a 12 (A até M) e a 14 (O)
    # Pulamos o índice 13 (Coluna N)
    cols_to_use = list(range(13)) + [14]
    
    # Cabeçalhos na linha 7 (índice 6)
    headers = df.iloc[6, cols_to_use].tolist()
    
    # Dados da linha 8 em diante (índice 7)
    data = df.iloc[7:, cols_to_use].copy()
    data.columns = headers
    
    # Tratamento de NaN: transforma valores nulos em strings vazias para exibição
    data = data.fillna('').reset_index(drop=True)
    
    return data, headers

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="us">
<head>
    <meta charset="UTF-8">
    <title>Multilocus-PCR-RFLP Genotype Search</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f8f9fa; }
        .container { max-width: 1400px; margin: auto; background: white; padding: 25px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h2 { color: #004a99; border-bottom: 2px solid #eee; padding-bottom: 10px; }
        
        /* Área de Texto do Projeto */
        .project-description { background: #fff; border: 1px solid #e0e0e0; padding: 20px; border-radius: 8px; margin-bottom: 30px; line-height: 1.6; }
        .project-description h3 { margin-top: 0; color: #004a99; }
        
        .instruction { font-size: 13px; color: #555; background: #e7f3ff; padding: 10px; border-radius: 4px; margin-bottom: 20px; border-left: 4px solid #007bff; }
        .grid-inputs { display: grid; grid-template-columns: repeat(auto-fill, minmax(90px, 1fr)); gap: 10px; margin-bottom: 20px; }
        .field label { display: block; font-size: 11px; font-weight: bold; color: #333; margin-bottom: 3px; }
        input { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
        .btn-search { background: #28a745; color: white; border: none; padding: 10px 25px; border-radius: 4px; cursor: pointer; font-weight: bold; }
        .btn-search:hover { background: #218838; }
        .error-msg { color: #dc3545; background: #f8d7da; padding: 10px; border-radius: 4px; margin-bottom: 20px; font-weight: bold; }
        
        table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 12px; }
        th { background: #004a99; color: white; padding: 10px; text-align: left; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background: #f2f2f2; }
        tr:hover { background: #e9ecef; }
    </style>
</head>
<body>
    <div class="container">
        <div class="project-description">
            <h2>About the Project</h2>
            <p>
                This dashboard was developed to facilitate the identification of <i>Toxoplasma gondii</i> genotypes and respective Brazilian strains (isolates, primary samples and clinical samples) already reported using the Multilocus-PCR-RFLP technique (Su et al., 2010). The main objective is to enable researchers to quickly and accurately compare their allele profiles with the project database.
            </p>
            <p><strong>Objectives:</strong> Support for diagnosis and diversity studies.</p>
            <p>Curator: Dr. Hilda FJ Pena (hfpena@usp.br) </p>
            <p>Last update: <strong>March/2026</strong></p>

        </div>

        <h2>Multilocus-PCR-RFLP Genotype Search</h2>
        <div class="instruction">
            <strong>Important:</strong> To perform the search, all 11 fields must be filled in. Use <strong>Ctrl+V</strong> in the first field to paste an entire row from Excel. *CS3 marker was not investigated in all studies and it is NOT used to define the PCR-RFLP genotype.
        </div>

        {% if error %}
        <div class="error-msg">{{ error }}</div>
        {% endif %}
        
        <form method="POST">
            <div class="grid-inputs">
                {% for col in filter_cols %}
                <div class="field">
                    <label>{{ col }}</label>
                    <input type="text" name="{{ col }}" id="f_{{ loop.index0 }}" 
                           list="l_{{ loop.index0 }}" value="{{ request.form.get(col, '') }}"
                           onpaste="handlePaste(event, {{ loop.index0 }})" autocomplete="off">
                    <datalist id="l_{{ loop.index0 }}">
                        {% for val in unique_values[col] %}
                        <option value="{{ val }}">
                        {% endfor %}
                    </datalist>
                </div>
                {% endfor %}
            </div>
            <button type="submit" name="search" class="btn-search">Search</button>
            <a href="/toxo_genotypes" style="margin-left:15px; font-size:13px; color:#666; text-decoration:none;">Reset Filters</a>
        </form>

        {% if results is not none %}
            <hr>
            <h3>Results: {{ results|length }}</h3>
            {% if not results.empty %}
                <div style="overflow-x:auto;">
                    {{ results_html | safe }}
                </div>
            {% else %}
                <p style="color:#666;">No records found for the combination entered.</p>
            {% endif %}
        {% endif %}
    </div>

    <script>
    function handlePaste(e, startIdx) {
        const clipboardData = e.clipboardData || window.clipboardData;
        const pastedText = clipboardData.getData('Text');
        
        // Separa por tabulação ou múltiplos espaços (Excel)
        const values = pastedText.split(/\\t|\\n|\\r| {2,}/).map(v => v.trim()).filter(v => v !== "");

        if (values.length > 1) {
            e.preventDefault();
            values.forEach((val, i) => {
                const inputElement = document.getElementById('f_' + (startIdx + i));
                if (inputElement) {
                    inputElement.value = val;
                }
            });
        }
    }
    </script>
</body>
</html>
"""

@app.route('/toxo_genotypes', methods=['GET', 'POST'])
@app.route('/', methods=['GET', 'POST'])
def index():
    data, headers = get_data()
    if data is None:
        return f"Excel file not found: {FILE_PATH}"

    # Marcadores para filtro: Colunas B a L (índices 1 a 11 na nossa lista de colunas)
    # Como headers agora contém [A..M, O], os índices de B a L continuam de 1 a 11
    filter_cols = headers[1:12]
    
    # Gera combos dinâmicos baseados nos dados atuais da planilha
    unique_values = {}
    for col in filter_cols:
        unique_list = data[col].astype(str).str.strip().unique().tolist()
        unique_values[col] = sorted([v for v in unique_list if v != ''])
    
    results = None
    results_html = ""
    error = None

    if request.method == 'POST':
        user_filters = {col: request.form.get(col, '').strip() for col in filter_cols}
        
        # Validação: Exige preenchimento de todos os 11 campos
        if all(user_filters.values()):
            mask = pd.Series(True, index=data.index)
            for col, val in user_filters.items():
                mask &= (data[col].astype(str).str.strip() == val)
            
            results = data[mask]
            results_html = results.to_html(classes='table', index=False)
        else:
            error = "Please note: All 11 marker fields must be filled in to perform the search."

    return render_template_string(HTML_TEMPLATE, 
                                 filter_cols=filter_cols, 
                                 unique_values=unique_values, 
                                 results=results, 
                                 results_html=results_html,
                                 error=error)

if __name__ == "__main__":
    app.run()
