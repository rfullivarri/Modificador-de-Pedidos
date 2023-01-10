import dash
import openpyxl
import plotly.graph_objs as go
import plotly.offline
from dash import dcc, html
from dash.dependencies import ClientsideFunction, Input, Output

workbook = openpyxl.load_workbook('data.xlsx')
worksheet = workbook.active

# Creamos una lista con los datos del Excel
data = []
for row in worksheet.iter_rows():
  data.append(row[0].value)

# Importamos dash y creamos la aplicación
import dash
import dash_html_components as html

app = dash.Dash()

# Creamos la lista desplegable con la información del Excel
app.layout = html.Div([
    html.Label('Selecciona una opción:'),
    html.Div(
        dcc.Dropdown(
            id='my-dropdown',
            options=[{'label': item, 'value': item} for item in data],
            value=None,
            clearable=True,
            style={'width': '50%'}
        )
    )
])

# Agregamos la funcionalidad de búsqueda a la lista desplegable
app.clientside_callback(
    ClientsideFunction(
        namespace='clientside',
        function_name='search'
    ),
    Output('my-dropdown', 'options'),
    [Input('my-dropdown', 'search_value')]
)

def search(search_value, options):
  if search_value:
    return [option for option in options if search_value in option['label']]
  else:
    return options

# Iniciamos la aplicación
if __name__ == '__main__':
  app.run_server()