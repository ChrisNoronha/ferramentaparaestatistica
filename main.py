from flask import Flask, request, render_template, send_file
from openpyxl import Workbook
import math
import os

from ajustar_amplitude_classes import ajusta_amplitude, ajusta_classes

app = Flask(__name__)

def ordenar_contar_e_calcular_amplitude(numeros):
    numeros_ordenados = sorted(numeros)
    total_numeros = len(numeros)
    amplitude = numeros_ordenados[-1] - numeros_ordenados[0]
    quantidade_classes = round(math.sqrt(total_numeros))
    amplitude_classe = round((amplitude / quantidade_classes))
    amplitude_classe = ajusta_amplitude(amplitude_classe)
    quantidade_classes = ajusta_classes(quantidade_classes)
    limites_classes = []
    pontos_medios_classes = []
    limite_inferior = numeros_ordenados[0]

    for _ in range(quantidade_classes):
        limite_superior = limite_inferior + amplitude_classe
        ponto_medio = (limite_inferior + limite_superior) / 2
        limites_classes.append((limite_inferior, limite_superior))
        pontos_medios_classes.append(ponto_medio)
        limite_inferior = limite_superior

    return numeros_ordenados, total_numeros, amplitude, quantidade_classes, amplitude_classe, limites_classes, pontos_medios_classes

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        numeros_texto = request.form['numeros']
        numeros = list(map(int, numeros_texto.split()))
        resultados = ordenar_contar_e_calcular_amplitude(numeros)
        numeros_ordenados, total_numeros, amplitude, quantidade_classes, amplitude_classe, limites_classes, pontos_medios_classes = resultados

        contagem_por_classe = []
        for limite_inf, limite_sup in limites_classes:
            contagem_classe = sum(limite_inf <= numero < limite_sup for numero in numeros_ordenados)
            contagem_por_classe.append(contagem_classe)

        frequencia_acumulada = 0
        frequencia_acumulada_por_classe = []
        for contagem in contagem_por_classe:
            frequencia_acumulada += contagem
            frequencia_acumulada_por_classe.append(frequencia_acumulada)

        frequencia_por_ponto_medio = [contagem * ponto_medio for contagem, ponto_medio in zip(contagem_por_classe, pontos_medios_classes)]

        media = sum(frequencia_por_ponto_medio) / frequencia_acumulada

        diferencas_quadradas = [(ponto_medio - media) ** 2 * contagem for ponto_medio, contagem in zip(pontos_medios_classes, contagem_por_classe)]
        variancia = sum(diferencas_quadradas) / frequencia_acumulada
        desvio_padrao = math.sqrt(variancia)

        index_moda = contagem_por_classe.index(max(contagem_por_classe))
        moda = pontos_medios_classes[index_moda]

        media = round(media, 5)
        variancia = round(variancia, 5)
        desvio_padrao = round(desvio_padrao, 5)
        moda = round(moda, 5)

        if total_numeros % 2 == 1:
            mediana_index = total_numeros // 2
            mediana = numeros_ordenados[mediana_index]
        else:
            mediana_index_1 = total_numeros // 2 - 1
            mediana_index_2 = total_numeros // 2
            mediana = (numeros_ordenados[mediana_index_1] + numeros_ordenados[mediana_index_2]) / 2

        mediana = round(mediana, 5)

        wb = Workbook()
        ws = wb.active
        ws.title = "Estatísticas"
        ws.append(["Classe", "Limite", "P Médio", "Frequência", "Frequência Ac", "Freq * PM", "PM -Media", "(Freq * PM)²", "(Freq * PM)²*fi"])
        soma_freq_pm = sum(frequencia_por_ponto_medio)
        media_menos_ponto_medio_quadrado_por_freq = [(ponto_medio - media) ** 2 * contagem for contagem, ponto_medio in zip(contagem_por_classe, pontos_medios_classes)]
        soma_media_pm_quad_freq = sum(media_menos_ponto_medio_quadrado_por_freq)

        for i, (limite, ponto_medio, frequencia, frequencia_acumulada, freq_pm) in enumerate(zip(limites_classes, pontos_medios_classes, contagem_por_classe, frequencia_acumulada_por_classe, frequencia_por_ponto_medio), start=2):
            ws.cell(row=i, column=1, value=f"{i - 1}")
            ws.cell(row=i, column=2, value=f"{limite[0]} - {limite[1]}")
            ws.cell(row=i, column=3, value=ponto_medio)
            ws.cell(row=i, column=4, value=frequencia)
            ws.cell(row=i, column=5, value=frequencia_acumulada)
            ws.cell(row=i, column=6, value=freq_pm)
            ws.cell(row=i, column=7, value=ponto_medio - media)
            ws.cell(row=i, column=8, value=(ponto_medio - media) ** 2)
            ws.cell(row=i, column=9, value=((ponto_medio - media) ** 2) * frequencia)

        ws.cell(row=i + 1, column=4, value=frequencia_acumulada)
        ws.cell(row=i + 1, column=6, value=soma_freq_pm)
        ws.cell(row=i + 1, column=9, value=soma_media_pm_quad_freq)

        ws.cell(row=1, column=11, value="media")
        ws.cell(row=2, column=11, value=media)
        ws.cell(row=3, column=11, value="Variancia")
        ws.cell(row=4, column=11, value=variancia)
        ws.cell(row=5, column=11, value="Desvio Padrao")
        ws.cell(row=6, column=11, value=desvio_padrao)
        ws.cell(row=7, column=11, value="moda")
        ws.cell(row=8, column=11, value=moda)
        ws.cell(row=9, column=11, value="mediana")
        ws.cell(row=10, column=11, value=mediana)

        output_filename = 'classes.xlsx'
        wb.save(output_filename)

        return render_template('results.html', 
                               numeros_ordenados=numeros_ordenados, 
                               total_numeros=total_numeros, 
                               amplitude=amplitude, 
                               quantidade_classes=quantidade_classes, 
                               amplitude_classe=amplitude_classe, 
                               limites_classes=limites_classes, 
                               pontos_medios_classes=pontos_medios_classes, 
                               contagem_por_classe=contagem_por_classe, 
                               frequencia_acumulada_por_classe=frequencia_acumulada_por_classe, 
                               frequencia_por_ponto_medio=frequencia_por_ponto_medio, 
                               media=media, 
                               variancia=variancia, 
                               desvio_padrao=desvio_padrao, 
                               moda=moda, 
                               mediana=mediana, 
                               output_filename=output_filename, 
                               enumerate=enumerate, zip=zip)
    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
