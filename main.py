import PySimpleGUI as sg
from functions import buscarLigas, buscarTimeLigas, buscarTime, adicionarPlanilha
import pandas as pd

json_data = {
    'query': {
        'termo': [],
        'atividade_principal': [],
        'natureza_juridica': [],
        'uf': [],
        'municipio': [],
        'situacao_cadastral': 'ATIVA',
        'cep': [],
        'ddd': [],
    },
    'range_query': {
        'data_abertura': {
            'lte': '',
            'gte': '',
        },
        'capital_social': {
            'lte': None,
            'gte': None,
        },
    },
    'extras': {
        'somente_mei': True,
        'excluir_mei': False,
        'com_email': False,
        'incluir_atividade_secundaria': False,
        'com_contato_telefonico': False,
        'somente_fixo': True,
        'somente_celular': False,
        'somente_matriz': False,
        'somente_filial': False,
    },
    'page': 1,
}

Times=[]
timedaliga = []
Ligas=[]
media_gols_liga = ""


# print(Ligas[0][20])


sg.theme('DarkBlue13')


layout = [
    [sg.Text("Pegar dados soccerstats"), sg.Push(), sg.Text("Planilha:"), sg.Input(default_text="Resultados.xlsx", key="nome_planilha", justification="r", size=(15, 1))],
    [sg.Text("Filtros:")],
    [sg.Text("Ligas:"), sg.Combo(Ligas, key="Ligas", enable_events=True, size=(44, 1)), sg.Button("Carregar Ligas", size=(15, 1))],
    [sg.Text("Time A: "), sg.Combo(Times, key="TimeA",size=(15,1) ), sg.Text("Time B:"), sg.Combo(Times, key="TimeB", size=(15,1)), sg.Button("Pegar Resultados", size=(15, 1))],

    [sg.Button("Enviar para planilha"), sg.Button("Cancelar")],

    [sg.StatusBar("",key="-STAT-", size=(20, 1), auto_size_text=True, justification="c")]


]

janela = sg.Window("BOT SOCCERSTATS", layout)

while True:
    evento, valores = janela.read()
    if evento == sg.WIN_CLOSED or evento == "Cancelar":
        break

    if evento == "Carregar Ligas":
        janela['-STAT-'].update("Buscando Ligas, aguarde...") 
        janela.Refresh()
        Ligas = buscarLigas()
        janela['Ligas'].update(value="Selecione uma Liga", values=Ligas[0])
        janela['-STAT-'].update("Ligas carregadas") 
        

    if evento == "Ligas":
        janela['-STAT-'].update("Buscando Times, aguarde...") 
        janela.Refresh()
        indexliga = Ligas[0].index(valores['Ligas'])
        urlliga = Ligas[1][indexliga]
        timedaliga = buscarTimeLigas(urlliga)
        janela['TimeA'].update(value=timedaliga[0][0], values=timedaliga[0])
        janela['TimeB'].update(value=timedaliga[0][0], values=timedaliga[0])
        janela['-STAT-'].update("Times carregados") 



    if evento == "Pegar Resultados":
        try:
            if valores['Ligas'] == "Selecione uma Liga":
                sg.Popup("Selecione uma liga!")
            else:
                janela['-STAT-'].update("Buscando resultados, aguarde...") 
                janela.Refresh()
                indexTimeA = timedaliga[0].index(valores['TimeA'])
                urlTime = timedaliga[1][indexTimeA]
                dataframe_timeA = buscarTime(urlTime)

                indexTimeB = timedaliga[0].index(valores['TimeB'])
                urlTime = timedaliga[1][indexTimeB]
                dataframe_timeB = buscarTime(urlTime)
                janela['-STAT-'].update("Resultados carregados")
        except IndexError:
            sg.Popup("Selecione a Liga.")
            janela['-STAT-'].update("Selecione a Liga...") 



    if evento == "Enviar para planilha":
        try:
            media_gols_liga = timedaliga[2]
            janela['-STAT-'].update("Enviando para planilha, aguarde...") 
            resposta = adicionarPlanilha(dataframe_timeA, dataframe_timeB, valores['nome_planilha'], media_gols_liga)
            if resposta == 13:
                raise PermissionError
            janela['-STAT-'].update("Planilha salva!") 
        except NameError:
            sg.Popup("Faça a busca primeiro.")
            janela['-STAT-'].update("Faça a busca primeiro...")
        except IndexError:
            sg.Popup("Verifique se preenche tudo")
            janela['-STAT-'].update("Verifique se preenche tudo...")
        except PermissionError:
            sg.Popup("Feche a planilha")
            janela['-STAT-'].update("Feche a planilha...")







janela.close()