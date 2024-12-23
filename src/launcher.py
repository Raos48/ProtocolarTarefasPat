import os
import sys
import requests
import subprocess
import json
import time

LAUNCHER_VERSION = "1.0.0"
MAIN_LOCAL_VERSION = "1.0.0"

VERSION_JSON_URL = "https://raw.githubusercontent.com/Raos48/ProtocolarTarefasPat/main/updates/version.json"

# Nome do executável principal
MAIN_EXE_NAME = "main_v1.0.1.exe"

def check_for_update():
    """Verifica se há nova versão do main.exe"""
    print(f"[Launcher] Versão local do main.exe: {MAIN_LOCAL_VERSION}")
    try:
        resp = requests.get(VERSION_JSON_URL, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            latest_ver = data.get("version")
            file_url = data.get("file_url")
            if latest_ver and file_url:
                if latest_ver > MAIN_LOCAL_VERSION:
                    print(f"[Launcher] Nova versão disponível: {latest_ver}")
                    choice = input("Deseja atualizar agora? (S/N): ")
                    if choice.upper() == "S":
                        return (latest_ver, file_url)
                else:
                    print("[Launcher] Você já está na versão mais recente do main.exe.")
            else:
                print("[Launcher] JSON de versão não contém os campos esperados.")
        else:
            print(f"[Launcher] Falha ao checar updates. HTTP {resp.status_code}")
    except Exception as e:
        print("[Launcher] Erro ao verificar updates:", e)

    return (None, None)  # se não houver update ou deu erro

def download_and_replace(file_url, new_version):
    """
    Baixa o novo executável e substitui o main.exe em etapas.
    Retorna True se tudo der certo, False em caso de falha.
    """
    temp_exe = f"update_main_{new_version}.exe"

    # Etapa 1: Confirmar download
    choice = input("[Launcher] Confirma DOWNLOAD do novo arquivo? (S/N): ")
    if choice.upper() != "S":
        print("[Launcher] Usuário não confirmou download. Cancelando.")
        return False

    # Realizar o download
    try:
        print("[Launcher] Baixando nova versão do main.exe...")
        r = requests.get(file_url, stream=True)
        if r.status_code == 200:
            with open(temp_exe, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
            print("[Launcher] Download concluído.")
        else:
            print("[Launcher] Falha no download. HTTP status:", r.status_code)
            return False
    except Exception as e:
        print("[Launcher] Erro no processo de download:", e)
        return False

    # Etapa 2: Confirmar renomear o arquivo antigo
    if os.path.exists(MAIN_EXE_NAME):
        choice = input("[Launcher] Deseja RENOMEAR o arquivo antigo? (S/N): ")
        if choice.upper() == "S":
            backup_name = f"old_main_{MAIN_LOCAL_VERSION}.exe"
            try:
                os.rename(MAIN_EXE_NAME, backup_name)
                print(f"[Launcher] Arquivo antigo renomeado para {backup_name}")
            except Exception as e:
                print("[Launcher] Erro ao renomear arquivo antigo:", e)
                return False
        else:
            print("[Launcher] Usuário não confirmou renomear. O arquivo antigo permanecerá.")

    # Etapa 3: Confirmar substituir pelo novo arquivo
    choice = input("[Launcher] Deseja SUBSTITUIR (renomear) o novo arquivo para main.exe? (S/N): ")
    if choice.upper() == "S":
        try:
            os.rename(temp_exe, MAIN_EXE_NAME)
            print(f"[Launcher] Novo main.exe atualizado para versão {new_version}.")
        except Exception as e:
            print("[Launcher] Erro ao renomear novo arquivo:", e)
            return False
    else:
        print("[Launcher] Usuário não confirmou substituir. O novo arquivo permanecerá como", temp_exe)
        return False

    return True

def run_main_exe():
    """ Executa o main.exe, perguntando antes se deseja executar """
    choice = input("[Launcher] Deseja EXECUTAR o programa atualizado? (S/N): ")
    if choice.upper() == "S":
        print("[Launcher] Iniciando main.exe...")
        try:
            # subprocess.Popen ou subprocess.run
            subprocess.Popen([MAIN_EXE_NAME])
        except Exception as e:
            print("[Launcher] Falha ao iniciar main.exe:", e)
            input("Pressione Enter para sair...")
    else:
        print("[Launcher] Usuário não confirmou execução. Saindo...")

def main():
    print("[Launcher] Iniciando launcher...")
    # 1) Verifica se há nova versão
    new_ver, file_url = check_for_update()
    if new_ver and file_url:
        updated = download_and_replace(file_url, new_ver)
        if updated:
            print("[Launcher] main.exe atualizado com sucesso.")
        else:
            print("[Launcher] Ocorreu falha em alguma etapa de atualização.")
            input("Pressione Enter para sair...")
            return
    # Se não houver nova versão ou se falhar, apenas executamos o main atual (opcional)
    run_main_exe()

    print("[Launcher] Encerrando launcher.")
    #input("Pressione Enter para sair...")  # se quiser manter janela aberta

if __name__ == "__main__":
    main()
