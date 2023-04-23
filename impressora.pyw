import PySimpleGUI as sg
import os
import sys


if sys.platform == "win32":
    import ctypes

    ctypes.windll.kernel32.SetConsoleMode(ctypes.windll.kernel32.GetStdHandle(-11), 7)

# Obtém o diretório onde o script Python está sendo executado
dir_path = os.path.dirname(os.path.realpath(sys.argv[0]))

impressoras = [
    {
        "nome": "Impressora 1",
        "ip": "192.168.1.100",
        "caminho_driver": os.path.join(dir_path, "hplzp1006.inf"),
    },
    {
        "nome": "Impressora 2",
        "ip": "192.168.1.101",
        "caminho_driver": os.path.join(dir_path, "cnmlbp6030_32.inf"),
    },
    {
        "nome": "Impressora 3",
        "ip": "192.168.1.102",
        "caminho_driver": os.path.join(dir_path, "brhl2130_32.inf"),
    },
]
nomes_impressoras = [imp["nome"] for imp in impressoras]
layout = [
    [sg.Text("Selecione a impressora para instalar:")],
    [
        sg.Combo(
            nomes_impressoras, default_value=nomes_impressoras[0], key="-IMPRESSORA-"
        )
    ],
    [sg.Button("Instalar")],
]

janela = sg.Window("Instalação de Impressoras").Layout(layout)
while True:
    evento, valores = janela.Read()

    if evento == sg.WINDOW_CLOSED:
        break

    if evento == "Instalar":
        indice_impressora_selecionada = nomes_impressoras.index(valores["-IMPRESSORA-"])
        impressora_selecionada = impressoras[indice_impressora_selecionada]
        print(f"Impressora selecionada: {impressora_selecionada}")

        # Execute o comando para instalar a impressora no cmd
        cmd = f'rundll32 printui.dll,PrintUIEntry /if /b "{impressora_selecionada["nome"]}" /c "{impressora_selecionada["ip"]}" /m "{impressora_selecionada["caminho_driver"]}" /f'
        print(f"Comando a ser executado: {cmd}")
        sg.Popup(
            f"Instalando impressora {impressora_selecionada['nome']}..."
        )  # exibe uma mensagem para o usuário
        try:
            retorno = os.system(cmd)
            if retorno == 0:
                sg.Popup("Impressora instalada com sucesso!")
            else:
                sg.Popup("Erro ao instalar a impressora!")
        except:
            sg.Popup("Erro ao executar o comando!")
