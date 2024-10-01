import flask
from flask import render_template, request
import openpyxl

app = flask.Flask(__name__)

@app.route('/buscar', methods=['GET', 'POST'])
def buscar():
    resultado = "" 
    if request.method == 'POST':
        caminho = "VLANSPGB.xlsx"
        
        pon = request.form['pon'].upper()  
        bairro = 'MESSEJANA'  
        workbook = openpyxl.load_workbook(caminho)

        work = workbook[bairro]
        for row in work.iter_rows(values_only=True):
            if row[0] == pon:
                resultado = ', '.join(map(str, row))
                break  
        if not resultado:
            resultado = "PON não encontrado."

    return render_template('index.html', resultado=resultado)

if __name__ == '__main__':
    app.run(debug=True)
