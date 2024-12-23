import time
import warnings
import chromedriver_autoinstaller
import requests
import urllib3
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from datetime import datetime, timedelta, timezone
import os
import sys
import configparser
import json
from tqdm import tqdm
import openpyxl
import subprocess


# Ignorar warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)
urllib3.disable_warnings()
requests.packages.urllib3.disable_warnings()


def main():
    def find_excel_file(filename):
        possible_paths = [
            os.path.join(os.path.dirname(sys.executable), filename),
            os.path.join(os.getcwd(), filename)
        ]
        for path in possible_paths:
            if os.path.exists(path):
                return path
        return None

    # --------------------------------------------------------------------------
    # 1) Verificação do arquivo Excel
    # --------------------------------------------------------------------------
    excel_filename = "protocolo_pat.xlsx"
    excel_path = find_excel_file(excel_filename)

    if excel_path:
        print(f"Arquivo Excel encontrado em: {excel_path}")
        try:
            df = pd.read_excel(excel_path)
            print("Arquivo Excel carregado com sucesso.")
        except Exception as e:
            print(f"Erro ao carregar o arquivo Excel: {str(e)}")
            input("Pressione Enter para sair...")
            sys.exit(1)
    else:
        print(f"Erro: O arquivo Excel '{excel_filename}' não foi encontrado.")
        print("Certifique-se de que o arquivo está no mesmo diretório que o executável.")
        input("Pressione Enter para sair...")
        sys.exit(1)

    print("Verificando versão Chromedriver...")
    chromedriver_autoinstaller.install()
    print("Checagem Chromedriver Finalizada...")

    # --------------------------------------------------------------------------
    # 2) Obter token se necessário
    # --------------------------------------------------------------------------
    def obter_token():
        print("Obter Token..")
        warnings.filterwarnings("ignore", category=DeprecationWarning)
        urllib3.disable_warnings()
        requests.packages.urllib3.disable_warnings()
        chromedriver_autoinstaller.install()

        config = configparser.ConfigParser()
        config.read('config.ini')
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Carregar headers do arquivo
        headers_file_path = os.path.join(os.getcwd(), config['FILES']['headers_file_path'])
        try:
            with open(headers_file_path, 'r') as file:
                headers = json.loads(file.read())
            print("Headers carregados com sucesso.")
        except Exception as e:
            print(f"Erro ao carregar headers: {e}")
            headers = {}

        # Fazer requisição à API
        url = config['API_URLS']['checagem_pat_url']
        try:
            print("Realizando requisição para checagem status do PAT Online/Offline")
            response = requests.get(url, verify=False, headers=headers)
            if response.status_code != 200:
                print(f"Erro na requisição. Código de status: {response.status_code, response.text}")
                print(f"API Offline - Data e hora: {current_time}")

                # Iniciar o navegador
                options = webdriver.ChromeOptions()
                for option in config['SELENIUM_SETTINGS']['chrome_options'].split(','):
                    options.add_argument(option)

                service = Service(executable_path=chromedriver_autoinstaller.install())
                driver = webdriver.Chrome(options=options, service=service)
                driver.maximize_window()
                driver.get("https://atendimento.inss.gov.br/")

                tempo_total = 5
                with tqdm(total=100, desc="Aguardando 10 seg para Estabelecimento do PAT..", unit="%", position=0) as pbar:
                    for i in range(101):
                        pbar.update(1)
                        time.sleep(tempo_total / 100)

                print("Aguardando conclusão do procedimento de Login")
                wait = WebDriverWait(driver, 120)
                wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[2]/div/header/div[1]/span")))

                for i in range(10, 0, -1):
                    print(f"Aguardando: {i} segundos")
                    time.sleep(1)
                print("Tempo encerrado!")

                js_script = "return localStorage.getItem('ifs_auth');"
                js_script2 = "return localStorage.getItem('srv_auth');"
                resultado = driver.execute_script(js_script)
                token = driver.execute_script(js_script2)

                if not token:
                    raise Exception("Não foi possível obter o token de acesso.")
                if not resultado:
                    raise Exception("Não foi possível obter o token de acesso.")

                headers = {
                    'Authorization': 'Bearer ' + resultado,
                    'Content-type': 'application/json',
                    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                                  'AppleWebKit/537.36 (KHTML, like Gecko) '
                                  'Chrome/111.0.0.0 Safari/537.36',
                    'tokenservidor': token,
                    'Accept': 'application/json'
                }

                driver.quit()
                with open(headers_file_path, 'w') as file:
                    file.write(json.dumps(headers))
            else:
                pass

        except Exception as e:
            print(f"Erro de conexão: {e}")

    # Chamar a função para garantir que os tokens estejam atualizados
    obter_token()

    # --------------------------------------------------------------------------
    # 3) Configurações do script principal
    # --------------------------------------------------------------------------
    config = configparser.ConfigParser()
    config.read('config.ini')
    CPF = config['API_URLS']['cpf']
    url_cadastro_requerimento = "https://atendimento.inss.gov.br/apis/requerimentosPortalApi/requerimento/cadastro/requerimento"

    # Carrega novamente os headers do arquivo
    headers_file_path = os.path.join(os.getcwd(), config['FILES']['headers_file_path'])
    try:
        with open(headers_file_path, 'r') as file:
            headers = json.loads(file.read())
        print("Headers para requisição POST carregados com sucesso.")
    except Exception as e:
        print(f"Erro ao carregar headers: {e}")
        headers = {}

    print("Iniciando automação...")
    print("Por favor, faça login quando o navegador abrir, caso não tenha feito...")

    # Reabrir planilha (para edição) caso queira manipular
    workbook = openpyxl.load_workbook(excel_filename)
    worksheet = workbook.active

    # --------------------------------------------------------------------------
    # 4) Definir funções auxiliares
    # --------------------------------------------------------------------------
    def obter_servidor_responsavel(cod_servico, headers):
        url_responsavel = (
            "https://atendimento.inss.gov.br/"
            f"apis/requerimentosPortalApi/requerimento/cadastro/servico/{cod_servico}/responsavel/inicial"
        )
        resp = requests.get(url_responsavel, headers=headers, verify=False)
        if resp.status_code == 200:
            lista_resp = resp.json()
            if isinstance(lista_resp, list) and len(lista_resp) > 0:
                return lista_resp[0].get("id")  # ex: 11509
            else:
                raise Exception("Não foi possível encontrar 'id' do responsável no JSON retornado.")
        else:
            raise Exception(
                f"Erro ao obter responsável (status: {resp.status_code}). Retorno: {resp.text}"
            )

    def obter_local_atendimento(cod_servico, cidadao, headers):
        url_local = (
            "https://atendimento.inss.gov.br/"
            f"apis/requerimentosPortalApi/requerimento/cadastro/servico/{cod_servico}/cidadao/{cidadao}/local/atendimento?void"
        )
        resp = requests.get(url_local, headers=headers, verify=False)
        if resp.status_code == 200:
            json_local = resp.json()
            id_local = json_local["local"]["id"]
            data_disponibilidade_original = json_local["vaga"]["data"]
            dt_local = datetime.fromisoformat(data_disponibilidade_original)
            dt_utc = dt_local.astimezone(timezone.utc)
            data_disponibilidade = dt_utc.isoformat(timespec="milliseconds").replace("+00:00", "Z")
            return id_local, data_disponibilidade
        else:
            raise Exception(
                f"Erro ao obter local de atendimento (status: {resp.status_code}). "
                f"Retorno: {resp.text}"
            )

    def enviar_comentario(headers, protocolo, despacho):
        url_comentario = "https://atendimento.inss.gov.br/apis/comentariosApi/comentarios"
        data_comentario = {
            "conteudo": despacho,
            "tarefa": {"protocolo": protocolo},
            "sistemaParceiro": {"sistema": "PAT"},
            "canal": "MODULO_TAREFAS",
            "origem": "COMENTARIO",
        }
        try:
            resp = requests.post(url_comentario, json=data_comentario, headers=headers, verify=False)
            if resp.status_code == 201:
                print("Comentário enviado com sucesso.")
                return True
            else:
                print("Erro no envio do comentário.")
                print("Status Code:", resp.status_code)
                print("Resposta:", resp.text)
                return False
        except requests.exceptions.RequestException as e:
            print("Exceção ao enviar comentário:", e)
            return None

    # --------------------------------------------------------------------------
    # 5) Loop principal
    # --------------------------------------------------------------------------
    linha = 2
    max_tentativas = 10

    while True:
        print("=================================================")
        print(f"Executando linha: {linha}")
        siape = worksheet.cell(row=linha, column=1).value
        cod = worksheet.cell(row=linha, column=2).value
        despacho = worksheet.cell(row=linha, column=3).value
        status_atual = worksheet.cell(row=linha, column=4).value

        if cod is None:
            # Se chegou numa linha vazia, paramos
            break

        if status_atual is not None:
            # Se já há status, pula esta linha
            linha += 1
            continue

        # Exemplo: data de disponibilidade (hoje + 5 dias)
        data_disponibilidade = (datetime.now() + timedelta(days=5)).strftime("%Y-%m-%dT03:00:00.000Z")

        # 1) Obter servidor responsável
        try:
            responsavel_id = obter_servidor_responsavel(cod, headers)
            print(f"ID do Responsável obtido: {responsavel_id}")
        except Exception as err:
            print("[ERRO] Falha ao obter servidoresResponsaveis:", err)
            input("Pressione Enter para sair...")
            sys.exit(1)

        # 2) Obter local de atendimento e dataDisponibilidade
        try:
            id_local_atendimento, data_disponibilidade = obter_local_atendimento(cod, CPF, headers)
            print(f"idLocalAtendimento obtido: {id_local_atendimento}")
            print(f"dataDisponibilidade obtida: {data_disponibilidade}")
        except Exception as err:
            print("[ERRO] Falha ao obter idLocalAtendimento:", err)
            input("Pressione Enter para sair...")
            sys.exit(1)

        # Montar payload
        payload = {
            "idServico": int(cod) if pd.notnull(cod) else None,
            "origemSolicitacao": "INTRANET",
            "servidoresResponsaveis": [{"id": responsavel_id}],
            "idLocalAtendimento": id_local_atendimento,
            "dataDisponibilidade": data_disponibilidade,
            "requerente": {},
            "houveAtualizacaoDadosCadastrais": False,
            "camposAdicionais": [],
            "anexos": [],
            "vaga": {
                "dataDisponivel": ""
            }
        }

        # 3) POST para criar requerimento
        tentativas = 0
        url = "https://atendimento.inss.gov.br/apis/requerimentosPortalApi/requerimento/cadastro/requerimento"

        while True:
            tentativas += 1
            try:
                response = requests.post(url, headers=headers, json=payload, verify=False)
                print(f"[POST] Tentativa {tentativas} - Status Code: {response.status_code}")

                if response.status_code == 200:
                    print("[OK] Tarefa executada com sucesso.")
                    try:
                        resp_json = response.json()
                        protocolo_req = resp_json["answer"]["protocoloRequerimento"]
                        worksheet.cell(row=linha, column=4).value = protocolo_req
                        print("Protocolo gerado:", protocolo_req)

                        # Enviar despacho como comentário (opcional)
                        if despacho and isinstance(despacho, str) and despacho.strip():
                            ok_comentario = enviar_comentario(headers, protocolo_req, despacho)
                            if ok_comentario is True:
                                worksheet.cell(row=linha, column=5).value = "Despacho enviado."
                            elif ok_comentario is False:
                                worksheet.cell(row=linha, column=5).value = "Erro no envio do Despacho."
                            else:
                                worksheet.cell(row=linha, column=5).value = "Exceção no envio do Despacho."
                        else:
                            print("Não há despacho para enviar.")
                    except Exception as e:
                        print(f"Erro ao extrair protocolo e enviar Despacho: {e}")
                        worksheet.cell(row=linha, column=5).value = f"Exceção: {str(e)}"
                    break

                elif response.status_code in [403, 406]:
                    if tentativas < max_tentativas:
                        print(f"[ERRO] Status {response.status_code}. Tentativa {tentativas} de {max_tentativas}. Aguardando 30 segundos...")
                        time.sleep(30)
                        continue
                    else:
                        raise Exception(f"Excesso de tentativas (status: {response.status_code}).")

                else:
                    print(f"[ERRO] Falha não relacionada a 403 ou 406. Status code: {response.status_code}")
                    break

            except Exception as e:
                worksheet.cell(row=linha, column=4).value = f"Exceção: {str(e)}"
                print("[ERRO] Exceção ao enviar requisição:", e)
                break

        workbook.save(excel_filename)
        linha += 1

    # Salvar novamente ao final
    workbook.save(excel_filename)
    print("Processamento finalizado, planilha salva.")
    print("Processo finalizado.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("[FATAL] Ocorreu um erro não tratado:")
        print(e)
        input("Pressione Enter para sair...")
