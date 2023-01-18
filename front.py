import dash
import openpyxl
import pandas as pd
from dash import dcc, html
from dash.dependencies import Input, Output
from openpyxl import load_workbook

# Cargamos el archivo Excel con los datos de los clientes y productos
df=pd.read_excel(r'C:\Users\ramferna\Documents\PruebaPowerApps.xlsx', sheet_name="Hoja1")
df.columns = ["CodProd","Producto" ,"Concat_Prod"]

df2=pd.read_excel(r'C:\Users\ramferna\Documents\PruebaPowerApps.xlsx', sheet_name="Hoja2")
df2.columns = ["Cliente", "RS", "Concat_Client"]

wb =load_workbook(r"C:\Users\ramferna\OneDrive - Anheuser-Busch InBev\MOPI232.xlsx")
mopi= wb["Hoja1"]
len_product=len(mopi['C':'C'])
len_client=len(mopi['A':'A'])


# Creamos una lista con los datos de los clientes
clientes = df2['Cliente'].dropna().unique().tolist()
#print(clientes,"/n")
# Creamos una lista con los datos de los productos
productos = df['CodProd'].dropna().unique().tolist()
#print(productos)


# Importamos dash y creamos la aplicación
external_stylesheets = [r'C:\Users\ramferna\Documents\GitHub\Python_Testing\Modificacion de Pedidos\styles.css']
app = dash.Dash(__name__, external_stylesheets=external_stylesheets)



# Creamos la lista desplegable de los clientes
dropdown_clientes = dcc.Dropdown(
            id='dropdown-clientes',
            options=[{'label': df2.loc[df2['Cliente'] == i, 'Concat_Client'].item(), 'value': i} for i in clientes],
            #options=[{'label': i, 'value': i} for i in clientes],
            value=None,
            clearable=True,
            style={'width': '50%'}
        )

# Creamos la lista desplegable de los productos
dropdown_productos = dcc.Dropdown(
            id='dropdown-productos',
            options=[{'label': df.loc[df['CodProd'] == i, 'Concat_Prod'].item(), 'value': i} for i in productos],
            #options=[{'label': i, 'value': i} for i in productos],
            value=None,
            clearable=True,
            style={'width': '50%'}
        )

slider = dcc.Slider(
            id='slider-quantity',
            min=0,
            max=50,
            step=1,
            value=1,
            marks={i: str(i) for i in range(0, 51, 5)},
            className="slider",
            updatemode='drag'
        )


# Creamos la lista desplegable con la información del Excel

app.layout = html.Div(children=[
    html.Label('Selecciona una opción:',
                        style={'display': 'flex', 'justifyContent': 'center','margin': '100px 0px 0px 0px'}),

    html.Div(style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center'},
        children=[
            html.H3('Selecciona tu cliente',
                        style={'display': 'flex',
                                'flexDirection': 'column', 'alignItems': 'center',
                                'margin': '5px 3px 0px 10px',
                                'font-size': '100%','text-align': 'center',
                                'font-family': 'Arial, sans-serif','color':'#122E50'}),
            dropdown_clientes]),


    html.Div(style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center','margin': '0px 0px 0px 0px'},
        children=[
            html.H3('Selecciona tu producto',
                        style={'vertical-align': 'middle','display': 'flex',
                                'flexDirection': 'column', 'alignItems': 'center',
                                'margin': '5px 3px 0px 10px',
                                'font-size': '100%','text-align': 'center',
                                'font-family': 'Arial, sans-serif','color':'#122E50'}),
            dropdown_productos]),


    html.Div(style={'display': 'flex', 'justifyContent': 'center', 'alignItems': 'center'}, 
        children=[
                  html.Label('Selecciona una cantidad:',style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center'}),
                  dcc.Input(id="quantity-input", type="number", value=1, style={'width': '5%'}),
    html.Div(style={'width': '50%'}, 
        children=[slider]),
                  ]),
    html.Div(style={'display': 'flex', 'flexDirection': 'column', 'alignItems': 'center'}, 
        children=[
    html.Button('Enviar',id='enviar-button')]),
    html.Div(id='output-clientes'),
    html.Div(id='output-productos')
])


@app.callback(
    Output("quantity-input", "value"),
    [Input("slider-quantity", "value")],
)
def update_quantity_input(quantity):
    return quantity



@app.callback(
    Output("enviar-button", "n_clicks"),
    Input("enviar-button", "n_clicks"))

def reset_n_clicks(n_clicks):
    if n_clicks > 1:
        return 0
    return n_clicks

# Asignamos las funciones a los elementos de salida del callback
@app.callback(
    [Output('output-clientes', 'children'),
     Output('output-productos', 'children'),
     Output('output-quantity', 'children')],

    [Input('dropdown-clientes', 'value'),
     Input('dropdown-productos', 'value'),
     Input('slider-quantity', 'value'),
     Input('enviar-button', 'n_clicks')
     ])



def update_outputs(client_value, product_value, quantity_value, n_clicks):

    if client_value and product_value and n_clicks > 0:
        mopi["C"+str(len_product+1)].value= product_value
        mopi["A"+str(len_client+1)].value= client_value
        mopi["D"+str(len_client+1)].value= quantity_value
        wb.save(r"C:\Users\ramferna\OneDrive - Anheuser-Busch InBev\MOPI232.xlsx")




# Iniciamos la aplicación
if __name__ == '__main__':
  app.run_server(debug=True, port=8090)