from flask import Flask, render_template, request, jsonify, redirect, url_for
import sqlite3
import os
from werkzeug.utils import secure_filename

import shutil

from PyPDF2 import PdfReader # type: ignore
import pandas as pd # type: ignore
from openpyxl import load_workbook # type: ignore
from datetime import datetime, date, timedelta

app = Flask(__name__)

app.config['UPLOAD_FOLDER'] = 'static/uploads'
app.config['DATABASE'] = 'colaborador.db'

# Cria a pasta de upload se não existir
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Função para buscar dados
def get_dados():
    conn = sqlite3.connect('meubanco.db')
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Saurus")
    dados = cursor.fetchall()
    conn.close()
    return dados

@app.route('/')
def index():
    dados = get_dados()
    nomes_colunas = {
        'mov_nNf': 'Pedido',
        'mov_descResumido': 'Status',
        'mov_dhEmi': 'Finalizado',
        'dest_xNome': 'Cliente',
        'dest_fone': 'Contato',
        'tot_qtdItens': 'Mix cliente',
        'tot_qCom': 'Itens',
        'tot_vNF': 'Valor',
        'pdf': 'PDF',
        'status': 'Status',
        'categoria': 'Categoria',
        'vendedor': 'Vendedor',
        'Data': 'Data',
        'Informações': 'Informações',
    }

    hoje = datetime.today().date()
    fim = hoje + timedelta(days=365)
    datas_validas = gerar_datas_validas_com_dia_semana(hoje, fim)

    return render_template('index.html', dados=dados, nomes_colunas=nomes_colunas, datas_validas=datas_validas)

# Rota para atualizar telefone
@app.route('/atualizar-telefone', methods=['POST'])
def atualizar_telefone():
    dados = request.get_json()
    id_mov = dados.get('id')
    coluna = dados.get('coluna')
    novo_valor = dados.get('valor')

    if not (id_mov and coluna and novo_valor):
        return jsonify({'erro': 'Dados insuficientes'}), 400

    try:
        conn = sqlite3.connect('meubanco.db')
        cursor = conn.cursor()
        cursor.execute(f"UPDATE Saurus SET {coluna} = ? WHERE mov_idMov = ?", (novo_valor, id_mov))
        conn.commit()
        conn.close()
    except Exception as e:
        return jsonify({'erro': str(e)}), 500

    return jsonify({'status': 'sucesso'})

# Rota exemplo para o botão "Iniciar"
@app.route('/iniciar-status', methods=['POST'])
def iniciar_status():
    data = request.get_json()
    id_mov = data.get('id')

    try:
        conn = sqlite3.connect('meubanco.db')
        cursor = conn.cursor()
        cursor.execute("UPDATE Saurus SET status = 'Iniciado' WHERE mov_idMov = ?", (id_mov,))
        conn.commit()
        conn.close()
        return jsonify({'status': 'ok'})
    except Exception as e:
        return jsonify({'erro': str(e)}), 500

@app.route('/atualizar-categoria', methods=['POST'])
def atualizar_categoria():
    dados = request.get_json()
    id_mov = dados.get('id')
    coluna = dados.get('coluna')
    novo_valor = dados.get('valor')

    if not (id_mov and coluna and novo_valor):
        return jsonify({'erro': 'Dados inválidos'}), 400

    try:
        conn = sqlite3.connect('meubanco.db')
        cursor = conn.cursor()
        cursor.execute(f"UPDATE Saurus SET {coluna} = ? WHERE mov_idMov = ?", (novo_valor, id_mov))
        conn.commit()
        conn.close()
        return jsonify({'status': 'sucesso'})
    except Exception as e:
        return jsonify({'erro': str(e)}), 500

# Cria o banco de dados e a tabela se não existir
def init_db():
    with sqlite3.connect(app.config['DATABASE']) as conn:
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS colaboradores (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL,
                cargo TEXT NOT NULL,
                categoria TEXT NOT NULL,
                imagem TEXT NOT NULL
            )
        ''')
        conn.commit()

# Página principal colaboradores
@app.route('/colaborador', methods=['GET', 'POST'])
def cadastrar_colaborador():
    init_db()  # Garante que o banco e tabela existam

    if request.method == 'POST':
        nome = request.form['nome']
        cargo = request.form['cargo']
        categoria = request.form['categoria']
        imagem = request.files['imagem']

        if imagem:
            filename = secure_filename(imagem.filename)
            imagem_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            imagem.save(imagem_path)

            with sqlite3.connect(app.config['DATABASE']) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO colaboradores (nome, cargo, categoria, imagem)
                    VALUES (?, ?, ?, ?)
                ''', (nome, cargo, categoria, filename))
                conn.commit()

        return redirect(url_for('index'))

    # Buscar os colaboradores por categoria
    with sqlite3.connect(app.config['DATABASE']) as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM colaboradores WHERE categoria = 'conferente'")
        conferentes = cursor.fetchall()
        cursor.execute("SELECT * FROM colaboradores WHERE categoria = 'gerente'")
        gerentes = cursor.fetchall()
        cursor.execute("SELECT * FROM colaboradores WHERE categoria = 'encarregado'")
        encarregados = cursor.fetchall()
        cursor.execute("SELECT * FROM colaboradores WHERE categoria = 'separador'")
        separadores = cursor.fetchall()
        cursor.execute("SELECT * FROM colaboradores WHERE categoria = 'vendedor'")
        vendedores = cursor.fetchall()

    return render_template('colaboradores.html',
                           conferentes=conferentes,
                           gerentes=gerentes,
                           encarregados=encarregados,
                           Separador=separadores,
                           vendedor=vendedores)

@app.route('/atualizar-informacoes', methods=['POST'])
def atualizar_informacoes():
    dados = request.get_json()
    id_mov = dados.get('id')
    coluna = dados.get('coluna')
    novo_valor = dados.get('valor')

    if not (id_mov and coluna):
        return jsonify({'erro': 'Dados inválidos'}), 400

    if coluna == 'Data' and novo_valor:
        try:
            dt = datetime.strptime(novo_valor, '%Y-%m-%d')
            dias_semana = ['Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo']
            nome_dia = dias_semana[dt.weekday()]  # weekday() = 0 (segunda) até 6 (domingo)
            novo_valor = dt.strftime('%d/%m/%Y') + f" - {nome_dia}"
        except Exception as e:
            print('Erro conversão data:', e)
            # Se erro, mantém o valor enviado

    try:
        conn = sqlite3.connect('meubanco.db')
        cursor = conn.cursor()
        cursor.execute(f"UPDATE Saurus SET {coluna} = ? WHERE mov_idMov = ?", (novo_valor, id_mov))
        conn.commit()
        conn.close()
        return jsonify({'status': 'sucesso'})
    except Exception as e:
        return jsonify({'erro': str(e)}), 500

@app.route('/analise')
def analise():
    dias_semana = [
        'Segunda-feira', 'Terça-feira', 'Quarta-feira',
        'Quinta-feira', 'Sexta-feira', 'Sábado', 'Domingo'
    ]

    hoje = date.today()
    inicio_semana = hoje - timedelta(days=hoje.weekday())

    # gera lista de dias da semana (com data formatada e data ISO)
    dias_correntes = []
    for i in range(7):
        dia_data = inicio_semana + timedelta(days=i)
        dias_correntes.append({
            "nome": dias_semana[i],
            "data": dia_data.strftime("%d/%m/%Y"),
            "data_iso": dia_data.strftime("%Y-%m-%d")
        })

    dia_selecionado = request.args.get('dia')
    print("\n===============================")
    print("DIA SELECIONADO (original):", dia_selecionado)

    with sqlite3.connect("meubanco.db") as conn:
        cursor = conn.cursor()

        if dia_selecionado:
            try:
                # Converte de YYYY-MM-DD
                data_dt = datetime.strptime(dia_selecionado.strip(), "%Y-%m-%d").date()
                nome_dia = dias_semana[data_dt.weekday()]
                data_banco_prefixo = f"{data_dt.strftime('%d/%m/%Y')} - {nome_dia}"

            except ValueError:
                # Se vier formatado de outro jeito
                data_banco_prefixo = dia_selecionado.strip()

            print("FILTRO NO BANCO (prefixo):", data_banco_prefixo)

            cursor.execute(
                "SELECT * FROM Saurus WHERE Data LIKE ?",
                (f"%{data_banco_prefixo}%",)
            )
        else:
            cursor.execute("SELECT * FROM Saurus")

        resultados = cursor.fetchall()
        colunas = [desc[0] for desc in cursor.description]

    dados = [dict(zip(colunas, linha)) for linha in resultados]

    print("REGISTROS ENCONTRADOS:", len(dados))

    destino_pasta = os.path.join("static", "analise")
    os.makedirs(destino_pasta, exist_ok=True)

    # REALIZO A LIMPEZA DA TABELA DO BANCO ANTES DE FAZER A ANALISE
    # Conecta ao banco
    with sqlite3.connect("logistica.db") as conn:
        cursor = conn.cursor()

        # Lista todas as tabelas do banco
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tabelas = cursor.fetchall()

        for tabela in tabelas:
            nome_tabela = tabela[0]
            cursor.execute(f"DELETE FROM {nome_tabela};")
            print(f"Todos os registros da tabela '{nome_tabela}' foram apagados.")

        conn.commit()

    # --- Limpar arquivos existentes na pasta ---
    for arquivo in os.listdir(destino_pasta):
        caminho_arquivo = os.path.join(destino_pasta, arquivo)
        if os.path.isfile(caminho_arquivo):
            os.remove(caminho_arquivo)
    print("Arquivos existentes na pasta 'analise' foram apagados.")

    for registro in dados:
        caminho_pdf = registro.get('pdf')  # aqui pdf_nome contém o caminho retornado pelo banco
        if caminho_pdf:
            # Extrai só o nome do arquivo para colocar na pasta de destino
            nome_arquivo = os.path.basename(caminho_pdf)
            arquivo_destino = os.path.join(destino_pasta, nome_arquivo)

            if os.path.exists(caminho_pdf):
                shutil.copy2(caminho_pdf, arquivo_destino)
                print("Arquivo PDF copiado:", arquivo_destino)
            else:
                print("Arquivo PDF não encontrado*****************************************:", caminho_pdf)
        else:
            print("Nenhum arquivo PDF associado ao registro.")

    # Inicializa as variáveis vazias
    #dados1 = []
    #colunas1 = []

    if dia_selecionado:
        try:
            # Converte de YYYY-MM-DD
            data_dt = datetime.strptime(dia_selecionado.strip(), "%Y-%m-%d").date()
            nome_dia = dias_semana[data_dt.weekday()]
            data_banco_prefixo = f"{data_dt.strftime('%d/%m/%Y')} - {nome_dia}"
            
#_______________________________________ANALISE CRUA______________________________________________

            origem = "C://Users/anglo/OneDrive/Área de Trabalho/Seleniun Saurus/static/analise"

            #Users\anglo\OneDrive\Área de Trabalho\Seleniun Saurus\static\pdfjs

            for caminho, subpasta, arquivos in os.walk(origem):
                for nome in arquivos:
                    #print(nome)
                    quantidade_arquivos = len(arquivos)

                # QUANDO ENCONTRO DADOS REALIZO O PROCESSAMENTO
                if len(dados) > 0:
                
                    for i in range(quantidade_arquivos):
                        arquivo = arquivos[i]

                        caminho_arquivo = f"C://Users/anglo/OneDrive/Área de Trabalho/Seleniun Saurus/static/analise/{arquivo}"
                        
                            
                        with open(caminho_arquivo, "rb") as arquivo:
                            print("********************************LEITURA ARQUIVO********************************",arquivo)
                            leitor_pdf = PdfReader(arquivo)
                            texto = ""
                            for pagina in leitor_pdf.pages:
                                texto += pagina.extract_text()

                    # Separar as folhas
                            pagina = texto.split(" ")
                            numero_pagina = pagina[20]
                            folha = numero_pagina[2:3]

                            if folha == '1':
                                (lerUmaPagina(texto))

                            if folha == '2':
                                (lerDuasPaginas(texto))

                            if folha == '3':
                                (lerTresPaginas(texto))
                
                else:    
                    print("NENHUM REGISTRO ENCONTRADO")

                # Conectar ao banco de dados existente
                conn = sqlite3.connect("logistica.db")

                # Ler a tabela 'logistica' e carregar em um DataFrame
                df = pd.read_sql_query("SELECT * FROM logistica", conn)

                # Fechar a conexão
                conn.close()

                # Converter a coluna 'quantidade' para número
                # - errors='coerce' transforma valores não numéricos em NaN
                # - Se houver vírgulas decimais, substituímos por ponto antes
                df["quantidade"] = (
                    df["quantidade"]
                    .astype(str)                # garante que tudo é string
                    .str.replace(",", ".", regex=False)  # troca vírgula por ponto
                )
                df["quantidade"] = pd.to_numeric(df["quantidade"], errors="coerce")

                # (Opcional) remover linhas onde a conversão falhou
                df = df.dropna(subset=["quantidade"])

                # Agrupar por 'descricao' e somar as quantidades
                df_grouped = df.groupby("descricao", as_index=False)["quantidade"].sum()

                # Ordenar do maior para o menor
                df_sorted = df_grouped.sort_values(by="quantidade", ascending=False)

                # Exibir resultado
                print(df_sorted)

                # --- Criar um novo banco de dados e salvar o DataFrame ---
                novo_banco = "logistica_resumo.db"
                conn_novo = sqlite3.connect(novo_banco)

                # Salvar DataFrame como tabela 'logistica_resumo' no novo banco
                df_sorted.to_sql("logistica_resumo", conn_novo, if_exists="replace", index=False)


                # Fechar conexão
                conn_novo.close()

                print(f"✅ Dados salvos com sucesso no novo banco '{novo_banco}' na tabela 'logistica_resumo'")
                
                # Conecta ao banco
                with sqlite3.connect("logistica_resumo.db") as conn:
                    cursor = conn.cursor()
                    cursor.execute("SELECT descricao, quantidade FROM logistica_resumo")  # nome da tabela = produtos
                    resultados = cursor.fetchall()
                    colunas1 = [desc[0] for desc in cursor.description]

                # Transformar em lista de dicionários
                dados1 = [dict(zip(colunas1, linha)) for linha in resultados]

        except ValueError:
            # Se vier formatado de outro jeito
            data_banco_prefixo = dia_selecionado.strip()
    

    if dia_selecionado == 'semana':

        hoje = date.today()
        inicio_semana = hoje - timedelta(days=hoje.weekday())  # segunda-feira
        fim_semana = inicio_semana + timedelta(days=6)         # domingo

        
        if dia_selecionado == "semana":
            datas_semana = [(inicio_semana + timedelta(days=i)).strftime("%d/%m/%Y") for i in range(7)]

            print("Data semana", datas_semana)

            

            with sqlite3.connect("meubanco.db") as conn:
                cursor = conn.cursor()

                filtros = " OR ".join([f"Data LIKE '%{d}%'" for d in datas_semana])
                query = f"SELECT * FROM Saurus WHERE {filtros}"

                #################################################

                print("Query executada:", query)

                cursor.execute(query)
                resultados = cursor.fetchall()

                # pegar nomes das colunas
                colunas = [desc[0] for desc in cursor.description]

                # transformar em lista de dicionários
                dados = [dict(zip(colunas, linha)) for linha in resultados]

            # imprimir os resultados
            if dados:
                for registro in dados:
                    print(registro)
            else:
                print("Nenhum registro encontrado para a semana.")

            ############LIMPANDO A PASTA###########################

            # Caminho da pasta
            pasta_analise_semana = "C://Users/anglo/OneDrive/Área de Trabalho/Seleniun Saurus/static/analise semana"

            # Verifica se a pasta existe
            if os.path.exists(pasta_analise_semana):
                # Percorre todos os arquivos na pasta
                for arquivo in os.listdir(pasta_analise_semana):
                    caminho_arquivo = os.path.join(pasta_analise_semana, arquivo)
                    
                    # Apaga apenas se for arquivo
                    if os.path.isfile(caminho_arquivo):
                        os.remove(caminho_arquivo)
                        print(f"Arquivo apagado: {arquivo}")
            else:
                print("A pasta 'analise semana' não existe.")


            # caminho da pasta origem (onde os PDFs estão atualmente)
            origem = "C://Users/anglo/OneDrive/Área de Trabalho/Seleniun Saurus/static/pdfjs"

            # caminho da nova pasta
            destino_semana = "C://Users/anglo/OneDrive/Área de Trabalho/Seleniun Saurus/static/analise semana"

            # cria a pasta se não existir
            os.makedirs(destino_semana, exist_ok=True)

            # supondo que 'dados' contém os registros com o caminho do PDF no campo 'pdf'
            for registro in dados:
                caminho_pdf = registro.get('pdf')
                if caminho_pdf and os.path.exists(caminho_pdf):
                    nome_arquivo = os.path.basename(caminho_pdf)
                    caminho_destino = os.path.join(destino_semana, nome_arquivo)
                    shutil.copy2(caminho_pdf, caminho_destino)  # move o arquivo
                    print(f"Arquivo movido: {caminho_destino}")
                else:
                    print(f"Arquivo não encontrado ou registro sem PDF: {caminho_pdf}")

        


        else:
            print("ok")


        origem = "C://Users/anglo/OneDrive/Área de Trabalho/Seleniun Saurus/static/analise semana"

        #Users\anglo\OneDrive\Área de Trabalho\Seleniun Saurus\static\pdfjs

        for caminho, subpasta, arquivos in os.walk(origem):
            for nome in arquivos:
                #print(nome)
                quantidade_arquivos = len(arquivos)

                # QUANDO ENCONTRO DADOS REALIZO O PROCESSAMENTO
        if len(dados) > 0:
                
            for i in range(quantidade_arquivos):
                arquivo = arquivos[i]

                caminho_arquivo = f"C://Users/anglo/OneDrive/Área de Trabalho/Seleniun Saurus/static/analise semana/{arquivo}"
                        
                            
                with open(caminho_arquivo, "rb") as arquivo:
                    print("********************************LEITURA ARQUIVO********************************",arquivo)
                    leitor_pdf = PdfReader(arquivo)
                    texto = ""
                    for pagina in leitor_pdf.pages:
                        texto += pagina.extract_text()

                    # Separar as folhas
                pagina = texto.split(" ")
                numero_pagina = pagina[20]
                folha = numero_pagina[2:3]

                if folha == '1':
                    (lerUmaPagina(texto))

                if folha == '2':
                    (lerDuasPaginas(texto))

                if folha == '3':
                    (lerTresPaginas(texto))
                
    else:    
        print("NENHUM REGISTRO ENCONTRADO")

    # Conectar ao banco de dados existente
    conn = sqlite3.connect("logistica.db")

    # Ler a tabela 'logistica' e carregar em um DataFrame
    df = pd.read_sql_query("SELECT * FROM logistica", conn)

    # Fechar a conexão
    conn.close()

    # Converter a coluna 'quantidade' para número
    # - errors='coerce' transforma valores não numéricos em NaN
    # - Se houver vírgulas decimais, substituímos por ponto antes
    df["quantidade"] = (
    df["quantidade"]
        .astype(str)                # garante que tudo é string
        .str.replace(",", ".", regex=False)  # troca vírgula por ponto
    )
    df["quantidade"] = pd.to_numeric(df["quantidade"], errors="coerce")

    # (Opcional) remover linhas onde a conversão falhou
    df = df.dropna(subset=["quantidade"])

    # Agrupar por 'descricao' e somar as quantidades
    df_grouped = df.groupby("descricao", as_index=False)["quantidade"].sum()

    # Ordenar do maior para o menor
    df_sorted = df_grouped.sort_values(by="quantidade", ascending=False)

    # Exibir resultado
    print(df_sorted)

    # --- Criar um novo banco de dados e salvar o DataFrame ---
    novo_banco = "logistica_resumo.db"
    conn_novo = sqlite3.connect(novo_banco)

    # Salvar DataFrame como tabela 'logistica_resumo' no novo banco
    df_sorted.to_sql("logistica_resumo", conn_novo, if_exists="replace", index=False)


    # Fechar conexão
    conn_novo.close()

    print(f"✅ Dados salvos com sucesso no novo banco '{novo_banco}' na tabela 'logistica_resumo'")
                
    # Conecta ao banco
    with sqlite3.connect("logistica_resumo.db") as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT descricao, quantidade FROM logistica_resumo")  # nome da tabela = produtos
        resultados = cursor.fetchall()
        colunas1 = [desc[0] for desc in cursor.description]

    # Transformar em lista de dicionários
    dados1 = [dict(zip(colunas1, linha)) for linha in resultados]



    return render_template(
        "analise.html",
        dias=dias_correntes,
        dia_selecionado=dia_selecionado,
        dados=dados,
        nomes_colunas={c: c for c in colunas},

        dados1=dados1,
        nomes_colunas1=colunas1
    )

@app.template_filter('formatar_para_date_input')
def formatar_para_date_input(data_str):
    # Exemplo de data no banco: "15/10/2025 - Quarta-feira"
    try:
        # Pega só a parte da data (antes do " - ")
        data_limpa = data_str.split(' - ')[0]
        dt = datetime.strptime(data_limpa, '%d/%m/%Y')
        return dt.strftime('%Y-%m-%d')
    except Exception:
        return ''

@app.route('/atualizar-data', methods=['POST'])
def atualizar_data():
    dados = request.get_json()
    id_mov = dados.get('id')
    coluna = dados.get('coluna')
    valor = dados.get('valor')

    if not (id_mov and coluna and valor):
        return jsonify({'erro': 'Dados inválidos'}), 400

    try:
        conn = sqlite3.connect('meubanco.db')
        cursor = conn.cursor()
        cursor.execute(f"UPDATE Saurus SET {coluna} = ? WHERE mov_idMov = ?", (valor, id_mov))
        conn.commit()
        conn.close()
        return jsonify({'status': 'sucesso'})
    except Exception as e:
        return jsonify({'erro': str(e)}), 500
    
def gerar_datas_validas_com_dia_semana(inicio, fim):
    dias = ['Domingo', 'Segunda-feira', 'Terça-feira', 'Quarta-feira', 'Quinta-feira', 'Sexta-feira', 'Sábado']
    delta = fim - inicio
    datas = []
    for i in range(delta.days + 1):
        dia = inicio + timedelta(days=i)
        dia_semana = dias[(dia.weekday() + 1) % 7]
        texto = dia.strftime('%d/%m/%Y') + ' - ' + dia_semana
        datas.append(texto)
    return datas



#FUNCÕES DE LEITURA DOS ARQUIVOS DO PDF
def lerUmaPagina(scanner):
    inicio = scanner.find("CFOP") + 5
    fim = scanner.find("Totais")-1

    
    pedido = scanner[inicio:fim].replace("1,5L","1.5L")

    tamanho_lista = len(pedido.split("\n"))

    descricao = pedido.split("\n")

    produto = descricao[0].split(" ")

    yasmin = "YASMIN"
    allan = "ALLAN"
    leia = "LEIA"

    if yasmin in produto:

        for i in range(tamanho_lista):
            produto = descricao[i].split(" ")
            invertida = produto[::-1]
            quantidade = invertida[8]
            med = invertida[9]
            desc = descricao[i].split(' ')
            t_desc = len(desc)-11
            desc_prod = desc[1:t_desc]

            string = ''
            for elemento in desc_prod:
                string += elemento + ' '
            
            cod = (desc[0])

            # Conectar (ou criar) o banco de dados
            conn = sqlite3.connect("logistica.db")
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS logistica (
                cod TEXT,
                descricao TEXT,
                unidade TEXT,
                quantidade INTEGER
            )
            """)

            # Inserir dados
            cursor.execute("""
            INSERT INTO logistica (cod, descricao, unidade, quantidade)
            VALUES (?, ?, ?, ?)
            """, (cod, string, med, quantidade))

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()
    
    elif allan in produto:

        for i in range(tamanho_lista):
            produto = descricao[i].split(" ")
            invertida = produto[::-1]
            quantidade = invertida[8]
            med = invertida[9]
            desc = descricao[i].split(' ')
            t_desc = len(desc)-11
            desc_prod = desc[1:t_desc]

            string = ''
            for elemento in desc_prod:
                string += elemento + ' '
            
            cod = (desc[0])

            # Conectar (ou criar) o banco de dados
            conn = sqlite3.connect("logistica.db")
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS logistica (
                cod TEXT,
                descricao TEXT,
                unidade TEXT,
                quantidade INTEGER
            )
            """)

            # Inserir dados
            cursor.execute("""
            INSERT INTO logistica (cod, descricao, unidade, quantidade)
            VALUES (?, ?, ?, ?)
            """, (cod, string, med, quantidade))

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()
    
    elif leia in produto:

        for i in range(tamanho_lista):
            produto = descricao[i].split(" ")
            invertida = produto[::-1]
            quantidade = invertida[8]
            med = invertida[9]
            desc = descricao[i].split(' ')
            t_desc = len(desc)-11
            desc_prod = desc[1:t_desc]

            string = ''
            for elemento in desc_prod:
                string += elemento + ' '
            
            cod = (desc[0])
            
            # Conectar (ou criar) o banco de dados
            conn = sqlite3.connect("logistica.db")
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS logistica (
                cod TEXT,
                descricao TEXT,
                unidade TEXT,
                quantidade INTEGER
            )
            """)

            # Inserir dados
            cursor.execute("""
            INSERT INTO logistica (cod, descricao, unidade, quantidade)
            VALUES (?, ?, ?, ?)
            """, (cod, string, med, quantidade))

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()

    
    else:
  
        for i in range(tamanho_lista):
            produto = descricao[i].split(" ")
            invertida = produto[::-1]
            quantidade = invertida[7]
            med = invertida[8]
            desc = descricao[i].split(' ')
            t_desc = len(desc)-10
            desc_prod = desc[1:t_desc]        

            string = ''
            for elemento in desc_prod:
                string += elemento + ' '
            
            cod = (desc[0])

            # Conectar (ou criar) o banco de dados
            conn = sqlite3.connect("logistica.db")
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS logistica (
                cod TEXT,
                descricao TEXT,
                unidade TEXT,
                quantidade INTEGER
            )
            """)

            # Inserir dados
            cursor.execute("""
            INSERT INTO logistica (cod, descricao, unidade, quantidade)
            VALUES (?, ?, ?, ?)
            """, (cod, string, med, quantidade))

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()
        
def lerDuasPaginas(scanner):
    inicio = scanner.find("CFOP") + 5
    fim = scanner.find("Totais")

    inicio1 = scanner.find("CONSUMIDORFinalizado")

    pedido = scanner[inicio:fim].replace("1,5L","1.5L")
    corte_p01 = pedido.find("Venda")-4
   
    pedido01 = pedido[0:corte_p01]
    
    palavra_dititacao = ("CONSUMIDORDigitação")
    palavra_finalizado = ("CONSUMIDORFinalizado")
    palavra_finalizado_0 = ("LTDAFinalizado")

    if palavra_dititacao in pedido:
        corte_inicio_p02 = pedido.find("CONSUMIDORDigitação")+19
        corte_fim_p02 = len(pedido)-1
        pedido02 = pedido[corte_inicio_p02:corte_fim_p02]

    if palavra_finalizado in pedido:
        corte_inicio_p02 = pedido.find("CONSUMIDORFinalizado")+20
        corte_fim_p02 = len(pedido)-1
        pedido02 = pedido[corte_inicio_p02:corte_fim_p02]
    
    if palavra_finalizado_0 in pedido:
        corte_inicio_p02 = pedido.find("LTDAFinalizado")+14
        corte_fim_p02 = len(pedido)-1
        pedido02 = pedido[corte_inicio_p02:corte_fim_p02]


    pedido02.split("\n")

    pedido03 = pedido01+pedido02

    tamanho_lista = len(pedido03.split("\n"))
    descricao = pedido03.split("\n")
    produto = descricao[0].split(" ")

    larissa = "LARISSA"
    allan = "ALLAN" 
    leia = "LEIA"

    if larissa in produto:
      
        for i in range(tamanho_lista):
            produto = descricao[i].split(" ")
            invertida = produto[::-1]
            quantidade = invertida[8]
            med = invertida[9]
            desc = descricao[i].split(' ')
            t_desc = len(desc)-11
            desc_prod = desc[1:t_desc]

            string = ''
            for elemento in desc_prod:
                string += elemento + ' '
            
            cod = (desc[0])

            # Conectar (ou criar) o banco de dados
            conn = sqlite3.connect("logistica.db")
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS logistica (
                cod TEXT,
                descricao TEXT,
                unidade TEXT,
                quantidade INTEGER
            )
            """)

            # Inserir dados
            cursor.execute("""
            INSERT INTO logistica (cod, descricao, unidade, quantidade)
            VALUES (?, ?, ?, ?)
            """, (cod, string, med, quantidade))

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()

    elif allan in produto:

        for i in range(tamanho_lista):
            produto = descricao[i].split(" ")
            invertida = produto[::-1]
            quantidade = invertida[8]
            med = invertida[9]
            desc = descricao[i].split(' ')
            t_desc = len(desc)-11
            desc_prod = desc[1:t_desc]

            string = ''
            for elemento in desc_prod:
                string += elemento + ' '
            
            cod = (desc[0])

            # Conectar (ou criar) o banco de dados
            conn = sqlite3.connect("logistica.db")
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS logistica (
                cod TEXT,
                descricao TEXT,
                unidade TEXT,
                quantidade INTEGER
            )
            """)

            # Inserir dados
            cursor.execute("""
            INSERT INTO logistica (cod, descricao, unidade, quantidade)
            VALUES (?, ?, ?, ?)
            """, (cod, string, med, quantidade))

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()
    
    elif leia in produto:

        for i in range(tamanho_lista):
            produto = descricao[i].split(" ")
            invertida = produto[::-1]
            quantidade = invertida[8]
            med = invertida[9]
            desc = descricao[i].split(' ')
            t_desc = len(desc)-11
            desc_prod = desc[1:t_desc]

            string = ''
            for elemento in desc_prod:
                string += elemento + ' '
            
            cod = (desc[0])
            
            # Conectar (ou criar) o banco de dados
            conn = sqlite3.connect("logistica.db")
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS logistica (
                cod TEXT,
                descricao TEXT,
                unidade TEXT,
                quantidade INTEGER
            )
            """)

            # Inserir dados
            cursor.execute("""
            INSERT INTO logistica (cod, descricao, unidade, quantidade)
            VALUES (?, ?, ?, ?)
            """, (cod, string, med, quantidade))

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()
 
 
    else:

        for i in range(tamanho_lista):
            produto = descricao[i].split(" ")
            invertida = produto[::-1]
            quantidade = invertida[7]
            med = invertida[8]

            desc = descricao[i].split(' ')
    
            t_desc = len(desc)-10

            desc_prod = (desc[1:t_desc])
            
            string = ''
            for elemento in desc_prod:
                string += elemento + ' '
            
            cod = (desc[0])
            pedido_df = pd.DataFrame(descricao)
            #print("Codigo:",cod,"Produto:",string,"Med:",med,"Quantidade:",quantidade)

            # Conectar (ou criar) o banco de dados
            conn = sqlite3.connect("logistica.db")
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS logistica (
                cod TEXT,
                descricao TEXT,
                unidade TEXT,
                quantidade INTEGER
            )
            """)

            # Inserir dados
            cursor.execute("""
            INSERT INTO logistica (cod, descricao, unidade, quantidade)
            VALUES (?, ?, ?, ?)
            """, (cod, string, med, quantidade))

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()
    
def lerTresPaginas(scanner):
    inicio = scanner.find("CFOP") + 5
    fim = scanner.find("Totais")

    inicio1 = scanner.find("CONSUMIDORFinalizado")

    pedido = scanner[inicio:fim].replace("1,5L","1.5L")

    lista = pedido.split("\n")

    lista = pedido.find("Nº")

    lista = pedido.find("Venda")

    lista = pedido

    corte = pedido[1627:2104]

    corte_p01 = pedido.find("Venda")-4

    pedido01 = pedido[0:corte_p01]

    palavra_dititacao = ("CONSUMIDORDigitação")
    palavra_finalizado = ("CONSUMIDORFinalizado")

    if palavra_dititacao in pedido:
        corte_inicio_p02 = pedido.find("CONSUMIDORDigitação")+20
        corte_fim_p02 = pedido.find("3/3")-190
        pedido02 = pedido[corte_inicio_p02:corte_fim_p02]

    if palavra_finalizado in pedido:
        corte_inicio_p02 = pedido.find("CONSUMIDORFinalizado")+20
        corte_fim_p02 = pedido.find("3/3")-190
        pedido02 = pedido[corte_inicio_p02:corte_fim_p02]

    corte_inicio_p03 = pedido.find("3/3")+299

    corte_fim_p03 = len(pedido)-1

    pedido03 = pedido[corte_inicio_p03:corte_fim_p03]
    
    pedido04 = pedido01+pedido02+pedido03

    tamanho_lista = len(pedido03.split("\n"))

    descricao = pedido04.split("\n")

    produto = descricao[0].split(" ")

    yasmin = "YASMIN"
    allan = "ALLAN"
    leia = "LEIA"

    if yasmin in produto:
        for i in range(tamanho_lista):
            produto = descricao[i].split(" ")
            invertida = produto[::-1]
            quantidade = invertida[8]
            med = invertida[9]
            desc = descricao[i].split(' ')
            t_desc = len(desc)-11
            desc_prod = desc[1:t_desc]

            string = ''
            for elemento in desc_prod:
                string += elemento + ' '
            
            cod = (desc[0])

            # Conectar (ou criar) o banco de dados
            conn = sqlite3.connect("logistica.db")
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS logistica (
                cod TEXT,
                descricao TEXT,
                unidade TEXT,
                quantidade INTEGER
            )
            """)

            # Inserir dados
            cursor.execute("""
            INSERT INTO logistica (cod, descricao, unidade, quantidade)
            VALUES (?, ?, ?, ?)
            """, (cod, string, med, quantidade))

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()


    elif allan in produto:

        for i in range(tamanho_lista):
            produto = descricao[i].split(" ")
            invertida = produto[::-1]
            quantidade = invertida[8]
            med = invertida[9]
            desc = descricao[i].split(' ')
            t_desc = len(desc)-11
            desc_prod = desc[1:t_desc]

            string = ''
            for elemento in desc_prod:
                string += elemento + ' '
            
            cod = (desc[0])

            # Conectar (ou criar) o banco de dados
            conn = sqlite3.connect("logistica.db")
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS logistica (
                cod TEXT,
                descricao TEXT,
                unidade TEXT,
                quantidade INTEGER
            )
            """)

            # Inserir dados
            cursor.execute("""
            INSERT INTO logistica (cod, descricao, unidade, quantidade)
            VALUES (?, ?, ?, ?)
            """, (cod, string, med, quantidade))

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()
    
    elif leia in produto:

        for i in range(tamanho_lista):
            produto = descricao[i].split(" ")
            invertida = produto[::-1]
            quantidade = invertida[8]
            med = invertida[9]
            desc = descricao[i].split(' ')
            t_desc = len(desc)-11
            desc_prod = desc[1:t_desc]

            string = ''
            for elemento in desc_prod:
                string += elemento + ' '
            
            cod = (desc[0])
            
            # Conectar (ou criar) o banco de dados
            conn = sqlite3.connect("logistica.db")
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS logistica (
                cod TEXT,
                descricao TEXT,
                unidade TEXT,
                quantidade INTEGER
            )
            """)

            # Inserir dados
            cursor.execute("""
            INSERT INTO logistica (cod, descricao, unidade, quantidade)
            VALUES (?, ?, ?, ?)
            """, (cod, string, med, quantidade))

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()
 
    else:

        for i in range(tamanho_lista):
            produto = descricao[i].split(" ")
            invertida = produto[::-1]
            quantidade = invertida[7]
            med = invertida[8]

            desc = descricao[i].split(' ')
    
            t_desc = len(desc)-10

            desc_prod = (desc[1:t_desc])
            
            string = ''
            for elemento in desc_prod:
                string += elemento + ' '
            
            cod = (desc[0])
            pedido_df = pd.DataFrame(descricao)
            #print("Codigo:",cod,"Produto:",string,"Med:",med,"Quantidade:",quantidade)

            # Conectar (ou criar) o banco de dados
            conn = sqlite3.connect("logistica.db")
            cursor = conn.cursor()

            # Criar tabela se não existir
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS logistica (
                cod TEXT,
                descricao TEXT,
                unidade TEXT,
                quantidade INTEGER
            )
            """)

            # Inserir dados
            cursor.execute("""
            INSERT INTO logistica (cod, descricao, unidade, quantidade)
            VALUES (?, ?, ?, ?)
            """, (cod, string, med, quantidade))

            # Confirmar e fechar conexão
            conn.commit()
            conn.close()


if __name__ == '__main__':
    app.run(debug=True)
