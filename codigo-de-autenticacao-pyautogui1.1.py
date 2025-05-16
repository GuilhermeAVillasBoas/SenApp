import pyautogui as pa
import time
import tkinter as tk
from tkinter import messagebox
import winsound
import pygetwindow as gw
import os
import subprocess

# Nome da janela que deseja trazer para frente
janela_sge = "TOTVS Linha RM - Serviços  Alias: CorporeRM | 3-SENAI"  
janela_mecLogin = "[MEC - SISTEC - v.4255 ] - Google Chrome"
janela_mec = "[MEC - SisTec] - Google Chrome"

pa.PAUSE = 0  # Define um pequeno atraso entre as ações

#FUNÇÕES
def ativar_janela_chrome(titulos_parciais):
    """
    Procura janelas com qualquer um dos títulos especificados, ativa e maximiza.
    Retorna True se encontrar e ativar alguma janela.
    """
    for titulo_parcial in titulos_parciais:
        janelas = gw.getWindowsWithTitle(titulo_parcial)
        for janela in janelas:
            if "Chrome" in janela.title:
                try:
                    if janela.isMinimized:
                        janela.restore()
                    janela.activate()
                    janela.maximize()
                    print(f"Janela encontrada: {janela.title} - foi ativada e maximizada.")
                    return True
                except Exception as e:
                    print(f"Erro ao ativar janela: {e}")
    return False

def bring_or_open_window_fullscreen(window_title, program_path):
    # Verifica se a janela já está aberta
    windows = gw.getWindowsWithTitle(window_title)
    
    if windows:
        window = windows[0]
        # Se não estiver maximizada, então traz pro topo e maximiza
        if not window.isMaximized:
            window.restore()
            window.maximize()
            window.activate()
        else:
            # Apenas ativa, se já estiver em tela cheia
            window.activate()
    else:
        # Se não estiver aberta, inicia o programa
        process = subprocess.Popen(program_path, shell=True)
        time.sleep(3)  # Tempo para garantir que a janela seja criada
        
        windows = gw.getWindowsWithTitle(window_title)
        if windows:
            window = windows[0]
            window.restore()
            window.maximize()
            window.activate()
        mostrar_mensagem("Atenção", "Aguarde e faça login.", erro=False)

def abrir_url_em_nova_janela_se_necessario(urls_prioridade, titulos_parciais):
    """
    Ativa uma janela se já estiver aberta com qualquer título.
    Se nenhuma estiver aberta, abre a primeira URL da lista em nova janela maximizada.
    """
    if ativar_janela_chrome(titulos_parciais):
        print("Uma das janelas já estava aberta. Foi maximizada.")
        return

    chrome_path = r"C:\Users\GBOAS\AppData\Local\Google\Chrome\Application\chrome.exe"

    if os.path.exists(chrome_path):
        try:
            subprocess.Popen([
                chrome_path,
                "--new-window",
                "--start-maximized",
                urls_prioridade[0]
            ])
            print(f"Abrindo nova janela com: {urls_prioridade[0]}")
            # Espera a janela abrir para depois tentar ativar
            time.sleep(3)
            ativar_janela_chrome(titulos_parciais)
        except Exception as e:
            print(f"Erro ao abrir o Chrome: {e}")
    else:
        print("Chrome não encontrado no caminho especificado.")
    """
    Ativa a janela do Chrome se uma com os títulos estiver aberta.
    Se nenhuma estiver aberta, abre a primeira URL da lista em nova janela maximizada.
    """
    if ativar_janela_chrome(titulos_parciais):
        print("Uma das janelas já estava aberta. Apenas foi maximizada.")
        return

    chrome_path = r"C:\Users\GBOAS\AppData\Local\Google\Chrome\Application\chrome.exe"

    if os.path.exists(chrome_path):
        try:
            subprocess.Popen([
                chrome_path,
                "--new-window",
                "--start-maximized",
                urls_prioridade[0]
            ])
            print(f"Abrindo nova janela com: {urls_prioridade[0]}")
            time.sleep(2)
            ativar_janela_chrome(titulos_parciais)
        except Exception as e:
            print(f"Erro ao abrir o Chrome: {e}")
    else:
        print("Chrome não encontrado no caminho especificado.")
    """
    Ativa uma janela existente do Chrome se uma com os títulos for encontrada.
    Se nenhuma estiver aberta, abre apenas a primeira URL em nova janela maximizada.
    """
    if ativar_janela_chrome(titulos_parciais):
        print("Janela existente ativada e maximizada.")
        return

    chrome_path = r"C:\Users\GBOAS\AppData\Local\Google\Chrome\Application\chrome.exe"

    if os.path.exists(chrome_path):
        try:
            subprocess.Popen([
                chrome_path,
                "--new-window",
                "--start-maximized",
                urls_prioridade[0]  # Só a primeira URL será aberta
            ])
            print(f"Nova janela aberta com: {urls_prioridade[0]}")
            # opcional: aguarda e tenta maximizar a nova janela (caso não abra já maximizada)
            time.sleep(2)
            ativar_janela_chrome(titulos_parciais)
        except Exception as e:
            print(f"Erro ao abrir o Chrome: {e}")
    else:
        print("Chrome não encontrado no caminho especificado.")

def confirmar_processo(titulo, mensagem):
    """Exibe um pop-up fixo de confirmação com som e retorna a opção escolhida pelo usuário."""
    winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)  # Som de alerta do Windows

    top = tk.Tk()
    top.withdraw()  # Oculta a janela principal
    top.attributes("-topmost", True)  # Mantém a caixa de diálogo no topo
    
    resposta = messagebox.askyesno(titulo, mensagem)
    
    return resposta  # Retorna True para 'Sim', False para 'Não'

def mostrar_mensagem(titulo, mensagem, erro=False):
    """Exibe um pop-up informativo ou de erro."""
    top = tk.Tk()
    top.withdraw()  # Oculta a janela principal
    top.attributes("-topmost", True)  # Mantém a caixa de diálogo no topo
    if erro:
        messagebox.showerror(titulo, mensagem)  # Mensagem de erro
    else:
        messagebox.showinfo("Error", "Erro no sistema, repita o processo")  # Mensagem informativa



try:

    bring_or_open_window_fullscreen(janela_sge, "C:\Totvs\RM.NET\RM.exe")
    # Espera a tela carregar antes de clicar
    time.sleep(2)

    # Abrir Documentos do Aluno
    pa.click(854, 294)  # Click no aluno
    time.sleep(2)
    pa.click(736, 235)  # Click em "Documentos"
    time.sleep(2)
    # Copiar CPF do aluno
    pa.click(650, 302)
    pa.hotkey('ctrl', 'a')
    pa.hotkey('ctrl', 'c')
    time.sleep(1)

    # Abrir Campo Complementar do Curso
    pa.click(808, 205)  # Abrir Anexos
    time.sleep(1)
    pa.click(809, 542)  # Cursos e Habilitações

    if not confirmar_processo("Selecione o curso", "Deseja continuar?"):# Exibir pop-up fixo de confirmação
          mostrar_mensagem("Processo Cancelado", "O usuário cancelou o processo.", erro=True)
          exit()  # Encerra o programa
    #pa.doubleClick(776,403)#Abrir Curso Confeitaria
    #time.sleep(3)
    #pa.click(1167, 344)  # Campo Complementar
    #time.sleep(1)

    # Selecionar Perfil de Acesso
    #pa.click(851, 1057)  # Abrir Google
    #time.sleep(1)
    #pa.click(161, 23)  # Selecionar Aba
    #time.sleep(2)
    abrir_url_em_nova_janela_se_necessario([janela_mecLogin], [janela_mecLogin, janela_mec])
    pa.click(1821,156) #Alterar Perfil
    time.sleep(1)
    pa.click(1061, 480)  # Clicar no Campo de Seleção
    time.sleep(1)
    pa.click(1058, 554)  # Selecionar o campo
    time.sleep(1)
    pa.click(779, 505)  # Confirmar seleção
    time.sleep(4)

    # Cadastrar Individual
    pa.click(313, 157)  # Ciclo de Matrícula
    time.sleep(2)
    pa.click(56, 261)  # Aluno
    time.sleep(1)
    pa.click(91, 280)  # Cadastrar Individual
    time.sleep(1)
    pa.click(508, 256)  # Selecionar área de pesquisa
    pa.hotkey('ctrl', 'v')  # Colar CPF

    if not confirmar_processo("Confirme o CAPTCHA"):# Exibir pop-up fixo de confirmação
          mostrar_mensagem("Processo Cancelado", "O usuário cancelou o processo.", erro=True)
          exit()  # Encerra o programa

    pa.click(797,365)#Confirmar
    time.sleep(2)
    pa.click(1863,981)#Avançar
    time.sleep(2)
    
#Seleção da Turma Site Gov
    pa.click(701,314)#Barra de Pesquisa de turma
    time.sleep(1)

    pa.press('t')
    for _ in range(49):
        pa.press("down")
    pa.press("enter")
    time.sleep(1)

    pa.click(282,477)#Selecionar Turma
    time.sleep(1)
    pa.click(457,223)#Mês de Ocorrenia
    time.sleep(1)
    pa.click(280,333)#Selecionar Mês
    time.sleep(1)
    pa.click(620,222)#Dados de Matrícula
    time.sleep(2)
    pa.click(665,335)#Selecionar Situação
    time.sleep(1)
    pa.click(620,379)#Gratuito
    time.sleep(1)

    if not confirmar_processo():# Exibir pop-up fixo de confirmação
          mostrar_mensagem("Processo Cancelado", "O usuário cancelou o processo.", erro=True)
          exit()  # Encerra o programa

    pa.click(1874,977)#Salvar
    time.sleep(1)
    pa.click(868,545)#Confirmar
    time.sleep(1)

    #Pesquisar Aluno
    pa.click(112,385)#Página de Pesquisa
    time.sleep(1)
    pa.click(522,253)#Campo de Pesquisa
    time.sleep(1)
    pa.hotkey('ctrl', 'v')#Colar CPF
    time.sleep(1)
    pa.click(304,353)#Pesquisar
    time.sleep(1)
    
    #Registrar Conclusão
    pa.click(979,462)#Ação
    time.sleep(1)
    pa.click(952,461)#Alterar Status
    time.sleep(2)
    pa.click(768,530)#Campo de Seleção
    time.sleep(1)
    pa.click(620,671)#Selecionar Registrar Conclusão
    time.sleep(1)
    pa.click(275,644)#Selecionar Aluno
    time.sleep(1)
    pa.click(737,557)#Selecionar Mes de Conclusão
    time.sleep(1)
    pa.click(735,604)#mes
    time.sleep(1)
    if not confirmar_processo():# Exibir pop-up fixo de confirmação
          mostrar_mensagem("Processo Cancelado", "O usuário cancelou o processo.", erro=True)
          exit()  # Encerra o programa
    pa.click(330,981)#Salvar
    time.sleep(1)
    pa.click(883,569)#Confirmar
    time.sleep(1)
    pa.click(980,547)#Confirmar denovo
    time.sleep(1)

    #Alterar Perfil (Gestor Autenticador)
    pa.click(1821,156) #Alterar Perfil
    time.sleep(1)
    pa.click(1061, 480)  # Clicar no Campo de Seleção
    time.sleep(1)
    pa.click(1058, 649)  # Selecionar o campo
    time.sleep(1)
    pa.click(779, 505)  # Confirmar seleção
    time.sleep(4)

    pa.click(216,155)#Ciclo de Matrícula
    time.sleep(3)
    pa.click(97,242)#Validar Diploma de Curso Técnico(Pasta)
    time.sleep(2)
    pa.click(103,261)#Validar Diploma de Curso Técnico
    time.sleep(2)
    pa.click(535,354)#Selecionar área de pesquisa
    time.sleep(1)
    pa.hotkey('ctrl', 'v')#Colar CPF
    time.sleep(1)
    pa.click(313,417)#Pesquisar
    time.sleep(2)
    pa.click(272,531)#Selecionar Aluno
    time.sleep(1)
    pa.click(320,978)#Validar
    time.sleep(1)
    pa.click(882,587)#Confirmar
    time.sleep(1)
    pa.click(980,550)#Confirmar denovo
    time.sleep(1)
    
    #Consultar Diploma
    pa.click(160,279)#Consultar Diploma
    time.sleep(2)
    pa.click(535,354)#Selecionar área de pesquisa
    time.sleep(1)
    pa.hotkey('ctrl', 'v')#Colar CPF
    time.sleep(1)
    pa.click(313,417)#Pesquisar
    time.sleep(2)
    pa.tripleClick(1803,515)#Selecionar Código
    time.sleep(1)
    pa.hotkey('ctrl', 'c')#Copiar Código
    time.sleep(1)
    pa.click(1816,158)#Alterar Perfil
    time.sleep(1)
    pa.click(872,1053)#Abrir SGE1
    time.sleep(1)
    pa.click(974,945)#Abrir SGE2
    time.sleep(1)
    pa.click(1318,644)#Scroll
    time.sleep(1)
    pa.doubleClick(865,548)#Campo do Código
    time.sleep(1)
    pa.hotkey('ctrl', 'v')#Colar Código
    time.sleep(1)
    pa.click(1281,773)#Salvar
    time.sleep(1)

    mostrar_mensagem("Sucesso", "✅ Processo concluído com sucesso!")

except Exception as e:
    mostrar_mensagem("Erro", f"❌ Erro durante a automação:\n{e}", erro=True)
