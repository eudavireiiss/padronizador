from flask import Flask, render_template, request, jsonify
import subprocess
import threading
import pandas as pd
import os

app = Flask(__name__)

CAMINHO_BASE_PRODUTOS = os.path.join(os.path.dirname(__file__), "base_produtos.xlsx")


# Carrega a base de produtos ao iniciar o programa
def carregar_base_produtos():
    try:
        # Lê as colunas 1 (código) e 2 (nome) da planilha "Planilha1"
        base_produtos = pd.read_excel(CAMINHO_BASE_PRODUTOS, sheet_name="Planilha1", usecols=[0, 1], header=None, dtype=str)
        base_produtos.columns = ["COD", "NOME"]  # Renomeia as colunas
        base_produtos["COD"] = base_produtos["COD"].str.strip().str.lstrip('0')  # Remove espaços e zeros à esquerda

        # Cria um dicionário para mapear códigos para nomes
        mapeamento_codigos = dict(zip(base_produtos["COD"], base_produtos["NOME"]))
        print(f"Base de produtos carregada com {len(mapeamento_codigos)} itens.")  # Log para depuração
        return mapeamento_codigos
    except Exception as e:
        print(f"Erro ao carregar a base de produtos: {e}")
        return {}

# Carrega a base de produtos ao iniciar o programa
BASE_PRODUTOS = carregar_base_produtos()
MAPEAMENTO_CODIGOS = BASE_PRODUTOS  # Reutiliza a mesma base


@app.route("/buscar_produtos", methods=["POST"])
def buscar_produtos():
    texto = request.form["texto"]
    print(f"Texto recebido: {texto}")  # Log para depuração

    # Extrai os códigos (parte antes do primeiro delimitador: '-' ou ';')
    codigos = []
    for linha in texto.strip().split('\n'):
        # Remove espaços e zeros à esquerda, e pega a parte antes do primeiro delimitador
        codigo = linha.replace(';', '-').split('-')[0].strip().lstrip('0')
        codigos.append(codigo)
    print(f"Códigos extraídos: {codigos}")  # Log para depuração

    try:
        # Busca os nomes dos produtos com base nos códigos fornecidos
        resultado = []
        for codigo in codigos:
            nome = MAPEAMENTO_CODIGOS.get(codigo, "Código não encontrado")  # Retorna o nome ou uma mensagem de erro
            resultado.append(f"{codigo} - {nome}")

        # Retorna o resultado como uma string separada por quebras de linha
        return "\n".join(resultado)
    except Exception as e:
        print(f"Erro ao buscar produtos: {e}")
        return "Erro ao buscar produtos. Verifique o arquivo de base de produtos."

@app.route("/")
def home():
    return render_template("index.html")


@app.route("/padronizar", methods=["POST"])
@app.route("/padronizar", methods=["POST"])
def padronizar():
    tipo = request.form["tipo"]
    texto_original = request.form["texto"]
    formacao_diferente = request.form.get("formacao_diferente") == "true"
    vitrine_padrao = request.form.get("vitrine_padrao") == "true"

    texto_sem_letras = remover_letras(texto_original)
    validacao, erro = validar_codigos(texto_sem_letras)  # Já filtra linhas vazias aqui

    if not validacao:
        return erro

    if tipo == "Mateus Mais":
        texto_padronizado = padronizar_mateus_mais(texto_sem_letras, vitrine_padrao)
    else:
        texto_padronizado = padronizar_gm_core(texto_sem_letras, formacao_diferente)  # Corrigi um typo aqui ("padronizado")

    return texto_padronizado

# Remove letras do texto
def remover_letras(texto):
    return ''.join([c for c in texto if not c.isalpha()])


# Valida se os códigos estão na base de produtos
def validar_codigos(texto):
    linhas = [linha.strip() for linha in texto.split('\n') if linha.strip()]  # Filtra linhas vazias
    codigos_invalidos = []

    for i, linha in enumerate(linhas, start=1):
        partes = linha.replace(';', '-').split('-')
        codigo = partes[0].strip().lstrip('0')  # Pega o código (parte antes do primeiro '-')
        
        if not codigo:  # Se o código for vazio (linha como "-1,00")
            codigos_invalidos.append((i, linha))
        elif codigo not in BASE_PRODUTOS:
            codigos_invalidos.append((i, codigo))

    if codigos_invalidos:
        mensagem_erro = "Erro: Os seguintes códigos não foram encontrados na base de produtos:\n"
        for linha, codigo in codigos_invalidos:
            mensagem_erro += f"- Linha {linha}: '{codigo}'\n"
        return False, mensagem_erro.strip()
    
    return True, None


# Padroniza o valor (trunca após a segunda casa decimal)
def padronizar_valor(valor):
    if ',' in valor:
        partes = valor.split(',')
        if len(partes[1]) > 2:
            partes[1] = partes[1][:2]
        return ','.join(partes)
    return valor

# Padroniza o texto para o formato "Mateus Mais"
def padronizar_mateus_mais(texto, vitrine_padrao):
    linhas = [linha.strip() for linha in texto.split('\n') if linha.strip()]  # Filtra linhas vazias
    resultado = []
    valor_padrao = "100000" if vitrine_padrao else "10000"

    for i, linha in enumerate(linhas, start=1):
        partes = linha.replace(';', '-').split('-')
        if len(partes) == 3:
            valor = padronizar_valor(partes[2])
            resultado.append(f"SIM\t{partes[0]}\t{partes[1]}\t{valor_padrao}\t{valor_padrao}\t{valor}")
        elif len(partes) == 2:
            valor = padronizar_valor(partes[1])
            resultado.append(f"SIM\t{partes[0]}\t\t{valor_padrao}\t{valor_padrao}\t{valor}")
        else:
            return f"Erro: Formato inválido na linha {i} ('{linha}'). Use 'COD-QTD-VALOR' ou 'COD-VALOR'."
    return "\n".join(resultado) + "\n" * 30


# Padroniza o texto para o formato "gm_core"
def padronizar_gm_core(texto, formacao_diferente):
    linhas = [linha.strip() for linha in texto.split('\n') if linha.strip()]  # Filtra linhas vazias
    resultado = []
    
    for i, linha in enumerate(linhas, start=1):
        partes = linha.replace(';', '-').split('-')
        if len(partes) == 3:
            valor = padronizar_valor(partes[2])
            if formacao_diferente:
                resultado.append(f"{partes[0]}\t{partes[1]}\t{valor}\t{valor}\t{valor}\t{valor}\t{valor}")
            else:
                resultado.append(f"{partes[0]}\t{partes[1]}\t0\t{valor}\t0\t0\t0")
        elif len(partes) == 2:
            valor = padronizar_valor(partes[1])
            if formacao_diferente:
                resultado.append(f"{partes[0]}\t{partes[1]}\t{valor}\t{valor}\t{valor}\t{valor}\t{valor}")
            else:
                resultado.append(f"{partes[0]}\t{partes[1]}\t0\t{valor}\t0\t0\t0")
        else:
            return f"Erro: Formato inválido na linha {i} ('{linha}'). Use 'COD-QTD-VALOR' ou 'COD-VALOR'."
    resultado += ["0\t0\t0\t0\t0\t0\t0"] * 60
    return "\n".join(resultado)


# Inicia o servidor Flask
def run_server():
    app.run(debug=False, use_reloader=False)


if __name__ == "__main__":
    threading.Thread(target=run_server).start()
    subprocess.Popen(['cmd', '/c', 'exit'], creationflags=subprocess.CREATE_NEW_CONSOLE)