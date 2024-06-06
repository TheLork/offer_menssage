import openpyxl
from urllib.parse import quote
import webbrowser
import time
import pyautogui

workbook = openpyxl.load_workbook('./mensagem_oferta/planilha_leads.xlsx')
pagina_clientes = workbook['Página1']

webbrowser.open('https://web.whatsapp.com')
time.sleep(10)

for linha in pagina_clientes.iter_rows(min_row=2):

    telefone = linha[0].value
    

    mensagem = '''🚨*ATENÇÃO - ÚLTIMA OPORTUNIDADE*🚨

Olá! Estou passando para dizer que esta é a sua *ÚLTIMA OPORTUNIDADE* para garantir o *melhor custo benefício* para a segurança do seu veículo por um valor *IMPERDÍVEL*.


*Você terá direito a:*
✅ Localização em tempo real do seu veículo;
✅ Bloqueio remoto do veículo;
✅ Alertas de ignição ligada e desligada no seu celular;
✅ Central 0800 24h em caso de roubo/furto;
✅ Guincho;
✅ Chaveiro;
✅ Socorro elétrico/mecânico;
✅ Socorro para troca de pneu;

E muitas outras funcionalidades pelo valor de apenas *R$64,90* mensais!

Não espere ser tarde demais, nos chame AGORA mesmo para garantir sua contratação.

SOMENTE ATÉ DIA 30/05'''
    
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        time.sleep(10)
        seta = pyautogui.locateCenterOnScreen('botao whatsapp.png')
        time.sleep(10)
        pyautogui.click(seta[0],seta[1])
        time.sleep(5)
        pyautogui.hotkey('ctrl','w')
        time.sleep(5)
    except:
        print(f'Não foi possível envviar mensagem para {telefone}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
                arquivo.write(f'{telefone}')