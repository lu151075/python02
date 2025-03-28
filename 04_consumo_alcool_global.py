from flask import Flask, request, render_template_string
import pandas as pd
import sqlite3 
import plotly.express as px
import plotly.io as pio
 
# configura o plotly para criar os arquivos no navegador por padrão
pio.renderers.default = "browser"


#carregar o drinkis.csv
df = pd.read_csv("C:/Users/noturno/Desktop/Luciano/Sistema/drinks.csv")
#df_avengers = pd.read_csv("avengers.csv")


# cria o banco de dados em sql e popular com os dados do arquivo csv
conn = sqlite3.connect("C:/Users/noturno/Desktop/Luciano/Sistema/consumo_alcool.db")
df.to_sql("drinks", conn, if_exists="replace", index=False )
#df_avengers.to_sql("Vingadores", conn, if_exists="replace", index=False)
conn.commit()
conn.close()


#  iniciar o Flask
app = Flask(__name__)

html_template = '''
<h1>Dashboard - Consumo de Alcool<h/1>
    <ul>
        <li> <a href="/grafico1"> Top 10 paises com maior consumo de alcool </a> </li>
        <li> <a href="/grafico2"> Media de consumo por tipo de bebida </a> </li>
        <li> <a href="/grafico3"> Consumo total por região </a> </li>
        <li> <a href="/grafico4"> Comparativo entre os tipos de bebidas </a> </li>
        <li> <a href="/pais?nome=Brazil"> Insigt por pais (ex: Brazil) </a> </li>
</ul>
'''
# rota inicial com o links para os graficos
@app.route('/')
def index():
    return render_template_string(html_template)

@app.route('/grafico1')
def grafico():
    conn = sqlite3("C:/Users/noturno/Desktop/Luciano/Sistema/consumo_alcool.db")
    df = pd.read_sql_query("""
        SELECT country, total_litres_of_pure_alcohol
        FROM drinks
        ORDER BY total_litres_of_pure_alcohol DESC
        LIMIT 10
    """ , conn)
    conn.close()
    fig = px.bar(
        df, 
        x = "country",
        Y = "total_litres_of_pure_alcohol",
        title="Top 10 paises com maior consumo de Alcool"
    )
    return fig.to_html()
@app.route('/grafico2')
def grafico2():
    conn = sqlite3.connect("C:/Users/noturno/Desktop/Luciano/Sistema/consumo_alcool.db")
    df.read_sql_query("SELECT AVG(beer_servings) AS cerveja, AVG(spirit_servings) AS destilados, AVG(wine_servings) AS vinhos FROM drinks", conn)
    conn.close()
    df_melted = df.melt(var_name="BEBIDAS", value_name="Média de Porções")
    fig = px.bar(df_melted, x="BEBIDAS", y="Média de Porções", title= "Média de consumo global por tipo")
    #return fig.to_html()
    return df_melted.to_html()


#iniciar o servidor flask
if __name__ == "__main__":
    app.run(debug=True)

