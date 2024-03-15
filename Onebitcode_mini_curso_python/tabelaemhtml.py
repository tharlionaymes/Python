import plotly.graph_objs as go
from plotly.subplots import make_subplots
import pandas as pd

# Carregar os dados do arquivo Excel
df = pd.read_excel("data/pivot_table.xlsx", sheet_name="Relatorio")

# Criar figura com plotly express
fig = make_subplots(rows=1, cols=1)
for column in df.columns[1:]:
    fig.add_trace(go.Bar(x=df.iloc[:, 0], y=df[column], name=column))

# Atualizar o layout do gráfico
fig.update_layout(
    title="Vendas Por Fabricantes",
    xaxis_title="Fabricantes",
    yaxis_title="Vendas",
    barmode="stack",
)

# Salvar a figura como um arquivo HTML
fig.write_html("chart.html")
