# -*- coding: utf-8 -*-
import csv
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

# --- Funções de Leitura de Arquivos com Fallback de Codificação ---

def ler_arquivo_com_fallback(caminho):
    """
    Tenta ler o arquivo de forma segura, com várias codificações.
    """
    if not Path(caminho).exists():
        raise FileNotFoundError(f"O arquivo não foi encontrado no caminho: {caminho}")

    for enc in ['utf-8-sig', 'utf-8', 'windows-1252', 'latin-1']:
        try:
            return Path(caminho).read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(f"Não foi possível ler o arquivo {caminho} com codificações conhecidas.")

# --- Funções de Carregamento de Dados ---

def carregar_dados_gerais(caminho_csv):
    """
    Carrega dados gerais de um CSV.
    Assume o formato <CHAVE>;<VALOR>.
    """
    dados_gerais = {}
    try:
        print(f"DEBUG: Carregando dados gerais de: {caminho_csv}")
        conteudo = ler_arquivo_com_fallback(caminho_csv)
        leitor = csv.reader(conteudo.splitlines(), delimiter=';')
    
        try:
            primeira_linha = next(leitor)
            if primeira_linha and primeira_linha[0].strip() == "<PONTO>":
                raise ValueError("Você selecionou o arquivo de pontos para os dados gerais. Por favor, selecione o arquivo de dados gerais (<CHAVE>;<VALOR>).")
        except StopIteration:
            raise ValueError("O arquivo CSV de dados gerais está vazio.")

        linhas_restantes = [primeira_linha] + list(leitor)

        for i, linha in enumerate(linhas_restantes):
            if not linha:
                continue
            
            try:
                if len(linha) >= 2 and re.match(r"<[A-Z_]+>", linha[0].strip()):
                    chave = linha[0].strip()
                    valor = linha[1].strip()
                    dados_gerais[chave] = valor
                    print(f"DEBUG: Linha {i+1} - Chave: {chave}, Valor: {valor}")
            except IndexError:
                raise ValueError(f"Linha malformada no arquivo de dados gerais, linha {i+1}: {linha}. Esperado '<CHAVE>;<VALOR>'.")
        
        if not dados_gerais:
            raise ValueError("Nenhum dado geral foi encontrado. Verifique se o formato é '<CHAVE>;<VALOR>'.")
        
        print(f"DEBUG: {len(dados_gerais)} dados gerais carregados com sucesso.")
    
    except Exception as e:
        raise ValueError(f"Erro ao carregar dados gerais do CSV. Detalhes: {e}")
    
    return dados_gerais

def carregar_pontos(caminho_csv):
    """
    Carrega a lista de pontos de um CSV.
    Assume que os dados começam após a linha de cabeçalho '<PONTO>;...'.
    """
    pontos = []
    lendo_pontos = False
    try:
        print(f"DEBUG: Carregando pontos de: {caminho_csv}")
        conteudo = ler_arquivo_com_fallback(caminho_csv)
        
        linhas = [linha.strip() for linha in conteudo.splitlines() if linha.strip()]
        
        cabecalho = []
        for i, linha in enumerate(linhas):
            campos = [campo.strip().strip('"') for campo in linha.split(';')]
            
            while campos and not campos[-1]:
                campos.pop()

            if campos and campos[0] == "<PONTO>":
                lendo_pontos = True
                cabecalho = [col.strip().strip('"') for col in campos]
                print(f"DEBUG: Cabeçalho encontrado: {cabecalho}")
                continue
            
            if lendo_pontos:
                if not any(campos):
                    break
                
                try:
                    num_campos_esperados = len(cabecalho)
                    if len(campos) < num_campos_esperados:
                        campos.extend([''] * (num_campos_esperados - len(campos)))
                    elif len(campos) > num_campos_esperados:
                        campos = campos[:num_campos_esperados]

                    ponto = dict(zip(cabecalho, [val.strip() for val in campos]))
                    pontos.append(ponto)
                    print(f"DEBUG: Ponto {ponto.get('<PONTO>', '')} carregado.")
                except IndexError:
                    raise ValueError(f"Linha malformada no arquivo de pontos, linha {i+1}: {linha}. Verifique o número de colunas.")

        if not pontos:
            raise ValueError("Nenhum ponto foi encontrado. Verifique se o cabeçalho '<PONTO>' existe.")
            
        print(f"DEBUG: {len(pontos)} pontos carregados com sucesso.")
    except Exception as e:
        raise ValueError(f"Erro ao carregar os pontos do CSV. Detalhes: {e}")
    
    return pontos
    
def normalizar_chaves_rtf(texto_rtf, chaves):
    """
    Remove formatação RTF das chaves como <IMOVEL> que podem estar salvas como <\f0 I\f1 M\f2 O...>
    """
    for chave in chaves:
        # Constrói padrão que aceita qualquer coisa entre os caracteres da chave
        padrao = ''
        for letra in chave:
            padrao += re.escape(letra) + r'(?:\\[a-z]+\d* ?|[\s{}])*?'
        padrao_regex = re.compile(padrao, flags=re.IGNORECASE)
        texto_rtf = padrao_regex.sub(chave, texto_rtf)
    return texto_rtf


def processar_rtf(modelo_rtf, dados_gerais, pontos, saida_rtf):
    """
    Processa o modelo RTF, preenche com os dados gerais e os dados de pontos,
    e salva o resultado em um arquivo RTF.
    """
    try:
        print(f"DEBUG: Iniciando processamento do RTF: {modelo_rtf}")
        
        with open(modelo_rtf, 'r', encoding='windows-1252', errors='ignore') as f:
            texto_modelo = f.read()

        if not texto_modelo:
            raise ValueError("O arquivo modelo RTF está vazio ou não pôde ser lido.")
        
        # Lista de todas as chaves a serem substituídas
        chaves = list(dados_gerais.keys()) + list(pontos[0].keys()) + ['<***>']
        texto_pronto = normalizar_chaves_rtf(texto_modelo, chaves)
        
        # 1. Substitui apenas os dados gerais
        print("\nDEBUG: Substituindo dados gerais...")
        for chave, valor in dados_gerais.items():
            texto_pronto = texto_pronto.replace(chave, str(valor))

        # 2. Encontra o bloco de repetição e o texto de fechamento
        padrao_bloco = r"(?s)<\*\*\*>\s*(.*?)\s*<\*\*\*>."
        match_bloco = re.search(padrao_bloco, texto_pronto)

        if not match_bloco or len(pontos) < 2:
            print("\nDEBUG: Bloco de repetição não encontrado ou número de pontos insuficiente. Processando como arquivo simples.")
            with open(saida_rtf, 'w', encoding='windows-1252') as f:
                f.write(texto_pronto)
            return

        bloco_base = match_bloco.group(1).replace('<***>', '')
        blocos_gerados = []
        print(f"\nDEBUG: Encontrado bloco base de repetição. Gerando {len(pontos) - 1} blocos...")

        # 3. Monta os blocos ponto a ponto
        for i, ponto_atual in enumerate(pontos[:-1]):
            print(f"DEBUG: Processando ponto {ponto_atual.get('<PONTO>')}. Confrontando com {pontos[i+1].get('<PONTO>')}")

            bloco_formatado = bloco_base

            # Substitui os dados do ponto atual
            for chave, valor in ponto_atual.items():
                bloco_formatado = bloco_formatado.replace(chave, str(valor))

            # Substitui os dados do próximo ponto (encadeamento do memorial)
            for chave in ["<PONTO>", "<UTMX>", "<UTMY>", "<CONFRONTANTE>", "<AZIMUTE>", "<DISTANCIA>", "<RUMO>"]:
                if chave in bloco_formatado:
                    bloco_formatado = bloco_formatado.replace(chave, pontos[i+1].get(chave, ""), 1)

            blocos_gerados.append(bloco_formatado)
        
        texto_depois = texto_pronto[match_bloco.end():]
        
        # 4. Trata o fechamento do perímetro
        if pontos:
            ultimo_ponto = pontos[-1]
            primeiro_ponto = pontos[0]
            fechamento_texto = (
                f"deste segue confrontando com a propriedade de {ultimo_ponto.get('<CONFRONTANTE>', '')}, "
                f"com azimute de {ultimo_ponto.get('<AZIMUTE>', '')} por uma distância de {ultimo_ponto.get('<DISTANCIA>', '')}m, "
                f"até o ponto {primeiro_ponto.get('<PONTO>', '')}, onde teve inicio essa descrição."
            )

            print("DEBUG: Gerando texto de fechamento do perímetro.")
            texto_depois = texto_depois.replace(
                "deste segue confrontando com a propriedade de <CONFRONTANTE>, com azimute de <AZIMUTE> por uma  distância de <DISTANCIA>m, até o ponto <PONTO>, onde teve inicio essa descrição.",
                fechamento_texto
            )
            # Substitui variáveis de fechamento
            for chave, valor in ultimo_ponto.items():
                texto_depois = texto_depois.replace(chave, str(valor))
            texto_depois = texto_depois.replace("<PONTO>", primeiro_ponto.get('<PONTO>', ''))

        # 5. Monta o texto final
        texto_final = texto_pronto[:match_bloco.start()] + "".join(blocos_gerados) + texto_depois
        
        with open(saida_rtf, 'w', encoding='windows-1252', errors='ignore') as f:
            f.write(texto_final)
        print(f"DEBUG: Arquivo final salvo em: {saida_rtf}")

    except Exception as e:
        raise Exception(f"Erro ao processar o RTF. Detalhes: {e}")


def selecionar_arquivos_e_processar():
    """
    Função principal que gerencia a seleção de arquivos e o processamento.
    Esta função usa tkinter e só funcionará em um ambiente com GUI.
    """
    root = tk.Tk()
    root.withdraw()

    try:
        modelo_rtf = filedialog.askopenfilename(title="1. Selecione o RTF Modelo", filetypes=[("RTF Files", "*.rtf")])
        if not modelo_rtf:
            messagebox.showinfo("Aviso", "Seleção cancelada.")
            return

        csv_pontos = filedialog.askopenfilename(title="2. Selecione o CSV com os pontos", filetypes=[("CSV Files", "*.csv")])
        if not csv_pontos:
            messagebox.showinfo("Aviso", "Seleção cancelada.")
            return

        csv_gerais = filedialog.askopenfilename(title="3. Selecione o CSV com os dados gerais", filetypes=[("CSV Files", "*.csv")])
        if not csv_gerais:
            messagebox.showinfo("Aviso", "Seleção cancelada.")
            return

        saida_rtf = filedialog.asksaveasfilename(title="4. Salvar arquivo final RTF", defaultextension=".rtf", filetypes=[("RTF Files", "*.rtf")])
        if not saida_rtf:
            messagebox.showinfo("Aviso", "Seleção cancelada.")
            return

        print("--- Iniciando carregamento dos arquivos ---")
        dados_gerais = carregar_dados_gerais(csv_gerais)
        pontos = carregar_pontos(csv_pontos)
        if "<AREAHE>" in dados_gerais and "<AREAM2>" not in dados_gerais:
            dados_gerais["<AREAM2>"] = dados_gerais["<AREAHE>"]      
        print("\n--- Verificação de dados carregados ---")
        if dados_gerais:
            print("Dados gerais carregados com sucesso.")
            print(f"  Número de itens: {len(dados_gerais)}")
        else:
            raise ValueError("Não foi possível carregar os dados gerais.")
        
        if pontos:
            print("Pontos carregados com sucesso.")
            print(f"  Número de pontos: {len(pontos)}")
        else:
            raise ValueError("Não foi possível carregar os pontos.")
        
        print("\n--- Iniciando o processamento do RTF ---")
        processar_rtf(modelo_rtf, dados_gerais, pontos, saida_rtf)
        messagebox.showinfo("Sucesso", f"Memorial gerado com sucesso:\n{saida_rtf}")
    except Exception as e:
        print(f"ERRO: {e}")
        messagebox.showerror("Erro", f"Erro ao gerar arquivo:\n{e}")
        

if __name__ == "__main__":
    selecionar_arquivos_e_processar()
