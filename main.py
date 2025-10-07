import pandas as pd
import webbrowser
import os
import json

# Data prep
try:
    df = pd.read_excel("C:\\Users\\SIMT\\Desktop\\Mesas\\usuarios_2025-10-07 (2).xlsx")
except FileNotFoundError:
    print("Error: No se encontr\u00f3 el archivo de Excel.")
    exit()

# --- Prepare data ---
chart_data = []
for index, row in df.iterrows():
    value = pd.to_numeric(row.get('Afinidades'), errors='coerce')
    chart_data.append({
        'name': str(row.get('Nombre Completo')),
        'value': 0 if pd.isna(value) else value
    })
chart_data_json = json.dumps(chart_data)

bar_chart_height = len(df) * 25

# --- Ranking Table ---
top_10 = df.sort_values(by="Afinidades", ascending=False).head(10)
ranking_html = "<h2>Ranking de Afinidades (Top 10)</h2>" + top_10[["Nombre Completo", "Afinidades"]].to_html(index=False, classes='table table-striped table-bordered table-hover').replace('<thead>', '<thead class="thead-dark">')

# --- Build HTML ---
html = []
html.append("<!DOCTYPE html><html><head>")
html.append("<title>Dashboard de Afinidades</title>")
html.append("<script src=\"https://cdn.jsdelivr.net/npm/chart.js\"></script>")
html.append("<script src=\"https://cdn.jsdelivr.net/npm/chartjs-chart-treemap@2.3.0/dist/chartjs-chart-treemap.min.js\"></script>")
html.append("<script src=\"https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.0.0\"></script>")
html.append("<link rel=\"stylesheet\" href=\"https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css\">")
html.append("<style>")
html.append("body { padding: 20px; }")
html.append(".container { margin-top: 20px; }")
html.append(".chart-container { position: relative; margin: auto; width: 80vw; }")
html.append(".treemap-container { height: 70vh; }")
html.append(".ranking-container { max-width: 600px; margin: auto; font-size: 0.9rem; }")
html.append(".ranking-container table { text-align: center; }")
html.append("h2 { text-align: center; margin-top: 40px; }")
html.append("</style></head><body><div class=\"container\">")
html.append("<h1>Jefes de Mesa zona Sur</h1>")
html.append('<div class="form-group mt-4"><label for="searchInput"><strong>Agentes</strong></label><input type="text" id="searchInput" class="form-control" placeholder="Buscar agente..."></div>')
html.append("<div class=\"mt-5 ranking-container\">")
html.append(ranking_html)
html.append("</div>")
html.append('<div class="mt-3 text-center"><h5>Adriana Thuman</h5><p>Cantidad de afinidad: 107</p></div>')
html.append("<h2>Mapa de Afinidades</h2>")
html.append("<div class=\"chart-container treemap-container mt-4\">")
html.append("<canvas id=\"treemapChart\"></canvas>")
html.append("</div>")
html.append("<h2>Detalle de Afinidades</h2>")
html.append(f'<div style="height: 800px; overflow-y: auto; border: 1px solid #ccc;"><div id="barChartContainer" class="chart-container mt-4" style="height:{bar_chart_height}px;">')
html.append("<canvas id=\"barChart\"></canvas>")
html.append("</div></div>")
html.append("</div><script>")

# --- JS Generation ---
js = []
js.append("Chart.register(ChartDataLabels);")
js.append(f"const originalData = {chart_data_json};")

# Treemap JS
js.append("var treemapCtx = document.getElementById('treemapChart').getContext('2d');")
js.append("var treemapChart = new Chart(treemapCtx, {")
js.append("    type: 'treemap',")
js.append("    data: { datasets: [{")
js.append("        label: 'Afinidades', tree: originalData, key: 'value', groups: ['name'],")
js.append("        backgroundColor: 'rgba(54, 162, 235, 0.5)', borderColor: 'rgba(54, 162, 235, 1)', borderWidth: 1,")
js.append("            labels: {")
js.append("                display: true,")
js.append("                color: 'black',")
js.append("                font: { size: 10, weight: 'normal' },")
js.append("                align: 'center',")
js.append("                position: 'center',")
js.append("                formatter: (ctx) => {")
js.append("                    if (!ctx.raw) { return ''; }")
js.append("                    const w = ctx.raw.w;")
js.append("                    const h = ctx.raw.h;")
js.append("                    if (w > 100 && h > 50) {")
js.append("                        const nameParts = ctx.raw.g.split(' ');")
js.append("                        if (nameParts.length > 1) {")
js.append("                            const firstName = nameParts[0];")
js.append("                            const restOfName = nameParts.slice(1).join(' ');")
js.append("                            return [firstName, restOfName];")
js.append("                        }")
js.append("                        return ctx.raw.g;")
js.append("                    }")
js.append("                    return '';")
js.append("                }")
js.append("            }")
js.append("    }]}, ")
js.append("    options: {")
js.append("        maintainAspectRatio: false,")
js.append("        plugins: { legend: { display: false }, datalabels: { display: false }, tooltip: { callbacks: { label: function(ctx) { const i = ctx.raw; return i ? `${i.g}: ${i.v}` : ''; } } } }")
js.append("    }")
js.append("});")

# Bar Chart JS
js.append("var barCtx = document.getElementById('barChart').getContext('2d');")
js.append("var barChart = new Chart(barCtx, {")
js.append("    type: 'bar',")
js.append("    data: { labels: originalData.map(d => d.name), datasets: [{")
js.append("        label: 'Afinidades', data: originalData.map(d => d.value), backgroundColor: 'rgba(255, 99, 132, 0.2)',")
js.append("        borderColor: 'rgba(255, 99, 132, 1)', borderWidth: 1")
js.append("    }]}, ")
js.append("    options: {")
js.append("        indexAxis: 'y',")
js.append("        maintainAspectRatio: false,")
js.append("        scales: { x: { beginAtZero: true } },")
js.append("        plugins: { datalabels: { anchor: 'end', align: 'end', formatter: Math.round, font: { weight: 'bold' } } }")
js.append("    }")
js.append("});")

# Search Logic
js.append("const searchInput = document.getElementById('searchInput');")
js.append("const barChartContainer = document.getElementById('barChartContainer');")
js.append("searchInput.addEventListener('keyup', (e) => {")
js.append("    const searchValue = e.target.value.toLowerCase();")
js.append("    const filteredData = originalData.filter(item => item.name.toLowerCase().includes(searchValue));")
js.append("    treemapChart.data.datasets[0].tree = filteredData;")
js.append("    treemapChart.update();")
js.append("    barChart.data.labels = filteredData.map(d => d.name);")
js.append("    barChart.data.datasets[0].data = filteredData.map(d => d.value);")
js.append("    const newBarHeight = filteredData.length * 25;")
js.append("    barChartContainer.style.height = `${newBarHeight}px`;")
js.append("    barChart.update();")
js.append("});")

html.append('\n'.join(js))
html.append("</script></body></html>")

final_html = '\n'.join(html)

# Write file
file_path = "C:\\Users\\SIMT\\Desktop\\Mesas\\index.html"
with open(file_path, "w", encoding="utf-8") as f:
    f.write(final_html)

print("File generated.")
webbrowser.open('file://' + os.path.realpath(file_path))