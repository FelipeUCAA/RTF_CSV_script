# -*- coding: utf-8 -*-
import csv
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import argparse
import sys
# O 'os' foi removido, pois não há mais manipulação de arquivos temporários

# O suporte a DOC e pywin32 foi removido, deixando apenas a MockDocument como fallback.
class MockDocument:
    def __init__(self, path=None): pass
    def save(self, path): pass
    
# Adicionado para suporte a arquivos DOCX (requer a biblioteca python-docx instalada)
try:
    from docx import Document
    from docx.shared import Inches
    # Importações adicionais para a lógica de reconstrução
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run
except ImportError:
    Document = MockDocument
    print("AVISO: 'python-docx' não está instalado. O processamento de arquivos DOCX não funcionará.")


# --- Funções de Leitura de Arquivos com Fallback de Codificação (inalteradas) ---
def ler_arquivo_com_fallback(caminho):
    if not Path(caminho).exists():
        raise FileNotFoundError(f"O arquivo não foi encontrado no caminho: {caminho}")

    for enc in ['utf-8-sig', 'utf-8', 'windows-1252', 'latin-1']:
        try:
            return Path(caminho).read_text(encoding=enc)
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(f"Não foi possível ler o arquivo {caminho} com codificações conhecidas.")

# --- Funções de Carregamento de Dados (inalteradas) ---
def carregar_dados_gerais(caminho_csv):
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
    
    if len(pontos) > 1 and pontos[0].get('<PONTO>') == pontos[-1].get('<PONTO>'):
        print("DEBUG: Ponto final duplicado encontrado. Removendo da lista de pontos para o loop de repetição.")
        pontos.pop()

    return pontos
    
def normalizar_chaves_rtf(texto_rtf, chaves):
    for chave in chaves:
        padrao = ''
        for letra in chave:
            padrao += re.escape(letra) + r'(?:\\[a-z]+\d* ?|[\s{}])*?'
        padrao_regex = re.compile(padrao, flags=re.IGNORECASE)
        texto_rtf = padrao_regex.sub(chave, texto_rtf)
    return texto_rtf

def replace_placeholder_in_paragraph(paragraph, placeholder, replacement):
    if placeholder not in paragraph.text:
        return

    text_before = ''
    start_run_index = -1
    start_offset = -1
    
    current_char_count = 0
    for i, run in enumerate(paragraph.runs):
        run_text = run.text
        run_len = len(run_text)
        
        if current_char_count + run_len > paragraph.text.find(placeholder) and start_run_index == -1:
            start_run_index = i
            start_pos_in_para = paragraph.text.find(placeholder)
            start_offset = start_pos_in_para - current_char_count
            break
            
        current_char_count += run_len

    if start_run_index == -1:
        return

    end_run_index = start_run_index
    end_offset = -1
    char_count = current_char_count
    
    placeholder_end_pos = paragraph.text.find(placeholder) + len(placeholder)
    
    for i in range(start_run_index, len(paragraph.runs)):
        run = paragraph.runs[i]
        run_len = len(run.text)
        
        if char_count + run_len >= placeholder_end_pos:
            end_run_index = i
            end_offset = placeholder_end_pos - char_count
            break
            
        char_count += run_len
    
    if end_run_index == -1:
        return
        
    start_run = paragraph.runs[start_run_index]
    end_run = paragraph.runs[end_run_index]
    
    prefix_text = start_run.text[:start_offset]
    suffix_text = end_run.text[end_offset:]
    
    if start_run_index == end_run_index:
        start_run.text = prefix_text + str(replacement) + suffix_text
    else:
        start_run.text = prefix_text + str(replacement)
        
        for i in range(start_run_index + 1, end_run_index):
            paragraph.runs[i].clear()
            
        end_run.text = suffix_text
        
        for i in range(len(paragraph.runs) - 1, -1, -1):
            if not paragraph.runs[i].text:
                p = paragraph._element
                r = paragraph.runs[i]._element
                p.remove(r)


def substituir_texto_em_docx(documento, dados_substituicao):
    def substituir_em_estrutura(estrutura):
        for paragrafo in estrutura.paragraphs:
            for chave, valor in dados_substituicao.items():
                replace_placeholder_in_paragraph(paragrafo, chave, str(valor))
    
    substituir_em_estrutura(documento)

    for tabela in documento.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                substituir_em_estrutura(celula)
                    
# Função converter_doc_para_docx removida, pois o suporte a .doc foi descontinuado.
    
def processar_modelo(caminho_modelo, dados_gerais, pontos, saida_caminho):
    """
    Processa o modelo (RTF ou DOCX), preenche com os dados,
    e salva o resultado. Roteador de extensão.
    """
    caminho_modelo_path = Path(caminho_modelo)
    extensao = caminho_modelo_path.suffix.lower()
    
    try:
        if extensao == '.rtf':
            # --- Lógica Original para RTF (Baseada em Manipulação de String) ---
            return processar_rtf_string(caminho_modelo, dados_gerais, pontos, saida_caminho)
        
        elif extensao == '.docx':
            # --- Lógica para DOCX ---
            
            if Document is MockDocument:
                raise Exception("A biblioteca 'python-docx' não foi encontrada. Instale-a para processar arquivos DOCX (pip install python-docx).")

            print(f"DEBUG: Iniciando processamento do DOCX: {caminho_modelo}")
            
            # 1. Carrega o documento
            documento = Document(caminho_modelo)
            
            # 2. Substitui apenas os dados gerais (em todo o documento)
            print("\nDEBUG: Substituindo dados gerais...")
            # Usa a nova função de substituição que preserva o negrito/formatação
            substituir_texto_em_docx(documento, dados_gerais)

            # 3. Encontra e processa o bloco de repetição
            padrao_bloco = r"(?s)<\*\*\*>(.*?)<\*\*\*>"
            
            # 3a. Encontra o parágrafo marcador de INÍCIO
            paragrafo_bloco_inicio = None
            start_index = -1
            end_index = -1
            
            # Encontra o marcador de início
            for i, p in enumerate(documento.paragraphs):
                if "<***>" in p.text:
                    start_index = i
                    paragrafo_bloco_inicio = p
                    break
            
            # Encontra o marcador de FIM
            if start_index != -1:
                for i in range(start_index, len(documento.paragraphs)):
                    p = documento.paragraphs[i]
                    if p.text.count('<***>') >= 2 or (i > start_index and '<***>' in p.text):
                        end_index = i
                        break

            # VERIFICAÇÃO DE BLOCO E PONTOS
            if start_index == -1 or end_index == -1 or len(pontos) < 2:
                # ... (lógica de fallback inalterada)
                print("\nDEBUG: Bloco de repetição não encontrado ou número de pontos insuficiente. Processando como arquivo simples.")
                if paragrafo_bloco_inicio:
                    paragrafo_bloco_inicio._element.getparent().remove(paragrafo_bloco_inicio._element)
                documento.save(saida_caminho)
                print(f"DEBUG: Arquivo final salvo em: {saida_caminho}")
                return
                
            # 3b. Extrair o Bloco Base de repetição
            
            # Compile o texto do bloco abrangendo P_start até P_end
            block_text_raw = ""
            paragraphs_to_remove = []
            
            for i in range(start_index, end_index + 1):
                p = documento.paragraphs[i]
                block_text_raw += p.text + "\n"
                paragraphs_to_remove.append(p)
                
            # Agora, aplica o regex no texto multi-parágrafo para extrair o conteúdo exato.
            match_bloco = re.search(padrao_bloco, block_text_raw)
            
            # Se o match falhar aqui, o modelo está malformado.
            if not match_bloco:
                raise Exception("Erro fatal ao extrair o conteúdo do bloco de repetição. Verifique se há dois marcadores '<***>' no modelo.")
            
            bloco_base_string = match_bloco.group(1).replace('<***>', '').strip()
            
            # 3c. Substitui variáveis do primeiro ponto (M01) em qualquer lugar
            primeiro_ponto = pontos[0]
            # A abertura já está substituída pelo Passo 2.
            substituir_texto_em_docx(documento, primeiro_ponto)

            # --- GERAÇÃO DOS PARÁGRAFOS DE TRANSIÇÃO (Lógica de String Original) ---
            print(f"\nDEBUG: Encontrado bloco base de repetição. Gerando {len(pontos) - 1} blocos...")
            
            blocos_gerados = []
            ultimo_confrontante = None

            for idx in range(len(pontos) - 1):
                ponto_atual = pontos[idx]
                proximo_ponto = pontos[idx + 1]
                numero_confronto = idx + 1

                bloco_formatado = bloco_base_string[:]

                # Dados básicos
                dados_paragrafo = ponto_atual.copy()
                dados_paragrafo.update({
                    '<PONTO>': proximo_ponto.get('<PONTO>', ''),
                    '<UTMX>': proximo_ponto.get('<UTMY>', ''),
                    '<UTMY>': proximo_ponto.get('<UTMX>', ''),
                })

                # --- TRATAMENTO: confrontante repetido ---
                confrontante_atual = ponto_atual.get('<CONFRONTANTE>', '').strip()
                if confrontante_atual and confrontante_atual == ultimo_confrontante:
                    dados_paragrafo['<CONFRONTANTE>'] = "o mesmo"
                else:
                    dados_paragrafo['<CONFRONTANTE>'] = confrontante_atual
                    ultimo_confrontante = confrontante_atual
                # -----------------------------------------
                
                # --- INÍCIO DA CORREÇÃO MANUAL DE RECONSTRUÇÃO (Geração da String) ---
                # Substitui primeiro o <PONTO> por um token para que a variável real não seja substituída
                # ainda pelo loop de substituição de strings abaixo.
                token_ponto = "@@PONTO_NEGRITO@@"
                bloco_formatado_sem_ponto = bloco_formatado.replace('<PONTO>', token_ponto)
                
                # Substitui placeholders que NÃO são o token
                for chave, valor in dados_paragrafo.items():
                    if chave == '<PONTO>': continue # Pula o placeholder de PONTO
                    bloco_formatado_sem_ponto = bloco_formatado_sem_ponto.replace(chave, str(valor))

                if '<CONFRO>' in bloco_formatado_sem_ponto:
                    bloco_formatado_sem_ponto = bloco_formatado_sem_ponto.replace("<CONFRO>", f"Confronto {numero_confronto}")

                # Guarda a string sem o PONTO real, mas com o token.
                blocos_gerados.append(bloco_formatado_sem_ponto)
            
            # 4. Inserção dos Blocos e Limpeza
            print("DEBUG: Inserindo blocos gerados e limpando marcadores.")
            
            paragrafo_ref = paragraphs_to_remove[0] if paragraphs_to_remove else paragrafo_bloco_inicio
            estilo_referencia = paragrafo_ref.style
            
            # Adiciona os blocos gerados
            for idx, bloco_texto_sem_ponto in enumerate(blocos_gerados):
                if bloco_texto_sem_ponto.strip():
                    
                    # Para o loop de repetição, o ponto de destino é sempre o (idx + 1)
                    proximo_ponto_nome = pontos[idx + 1].get('<PONTO>', '')
                    
                    novo_paragrafo = paragrafo_bloco_inicio.insert_paragraph_before('') # Inicia vazio
                    novo_paragrafo.style = estilo_referencia

                    # Divide o texto no token
                    partes = bloco_texto_sem_ponto.split(token_ponto)
                    
                    # 1. Adiciona a parte antes do <PONTO>
                    if len(partes) > 0:
                        novo_paragrafo.add_run(partes[0])

                    # 2. Adiciona o NOME do PONTO de DESTINO em NEGRITO
                    if len(partes) > 1:
                        novo_paragrafo.add_run(proximo_ponto_nome).bold = True
                        
                        # 3. Adiciona a parte depois
                        novo_paragrafo.add_run(partes[1])
            
            # 4c. Remove todos os parágrafos que continham o bloco de repetição original
            for p_remove in paragraphs_to_remove:
                p_remove._element.getparent().remove(p_remove._element)

            # --- FIM DA CORREÇÃO MANUAL DE RECONSTRUÇÃO (Passo 4) ---

            # --- PASSO 5 REMOVIDO: Não gera mais o parágrafo de fechamento "Confronto X: ..." ---
            # O código termina aqui, salvando o documento após o último bloco de repetição.
            
            # 6. Salva o documento
            documento.save(saida_caminho)
            print(f"DEBUG: Arquivo final salvo em: {saida_caminho}")
            return

        else:
            # Mensagem de erro atualizada
            raise ValueError(f"Extensão de arquivo não suportada: {extensao}. Suportadas: .rtf, .docx")
            
    finally:
        # A remoção de arquivos temporários não é mais necessária
        pass 


def processar_rtf_string(modelo_rtf, dados_gerais, pontos, saida_rtf):
    # ... (função inalterada)
    try:
        print(f"DEBUG: Iniciando processamento do RTF (string-based): {modelo_rtf}")
        
        with open(modelo_rtf, 'r', encoding='windows-1252', errors='ignore') as f:
            texto_modelo = f.read()

        if not texto_modelo:
            raise ValueError("O arquivo modelo RTF está vazio ou não pôde ser lido.")
        
        chaves = list(dados_gerais.keys()) + list(pontos[0].keys()) + ['<***>']
        texto_pronto = normalizar_chaves_rtf(texto_modelo, chaves)
        
        # 1. Substitui apenas os dados gerais
        print("\nDEBUG: Substituindo dados gerais...")
        for chave, valor in dados_gerais.items():
            texto_pronto = texto_pronto.replace(chave, str(valor))

        # 2. Encontra o bloco de repetição e o texto de fechamento
        padrao_bloco = r"(?s)<\*\*\*>(.*?)<\*\*\*>"
        match_bloco = re.search(padrao_bloco, texto_pronto)

        if not match_bloco or len(pontos) < 2:
            print("\nDEBUG: Bloco de repetição não encontrado ou número de pontos insuficiente. Processando como arquivo simples.")
            with open(saida_rtf, 'w', encoding='windows-1252') as f:
                f.write(texto_pronto)
            return

        # Separar o texto em partes
        texto_antes = texto_pronto[:match_bloco.start()]
        bloco_base = match_bloco.group(1).replace('<***>', '')
        texto_depois = texto_pronto[match_bloco.end():]
        blocos_gerados = []
        
        # --- ABERTURA: substitui variáveis do primeiro ponto (M01) no texto antes do bloco ---
        first_point = pontos[0]
        for key, value in first_point.items():
            if key in texto_antes:
                texto_antes = texto_antes.replace(key, str(value))
        # --------------------------------------------------------------------------------------

        # --- GERAÇÃO DOS PARÁGRAFOS DE TRANSIÇÃO ---
        print(f"\nDEBUG: Encontrado bloco base de repetição. Gerando {len(pontos) - 1} blocos...")
        
        last_confrontante = None

        for idx in range(len(pontos) - 1):
            current_point = pontos[idx]
            next_point = pontos[idx + 1]
            confronto_num = idx + 1

            formatted_block = bloco_base[:]

            # Dados básicos
            paragraph_data = current_point.copy()
            paragraph_data.update({
                '<PONTO>': next_point.get('<PONTO>', ''),
                '<UTMX>': next_point.get('<UTMY>', ''),
                '<UTMY>': next_point.get('<UTMX>', ''),
            })

            # Tratamento: confrontante repetido
            current_confrontante = current_point.get('<CONFRONTANTE>', '').strip()
            if current_confrontante and current_confrontante == last_confrontante:
                paragraph_data['<CONFRONTANTE>'] = "o mesmo"
            else:
                paragraph_data['<CONFRONTANTE>'] = current_confrontante
                last_confrontante = current_confrontante

            # Substitui placeholders
            for key, value in paragraph_data.items():
                formatted_block = formatted_block.replace(key, str(value))

            if '<CONFRO>' in formatted_block:
                formatted_block = formatted_block.replace("<CONFRO>", f"Confronto {confronto_num}")

            blocos_gerados.append(formatted_block)


        # --- FECHAMENTO: último ponto voltando para M01 (mantido para RTF) ---
        last_point = pontos[-1]
        first_point = pontos[0]
        confronto_num = len(pontos)

        closing_text = (
            f"Confronto {confronto_num}: deste segue confrontando com a propriedade de {last_point.get('<CONFRONTANTE>', '')}, "
            f"com azimute de {last_point.get('<AZIMUTE>', '')} por uma distância de {last_point.get('<DISTANCIA>', '')}m, "
            f"até o ponto {first_point.get('<PONTO>', '')}, onde teve início essa descrição."
        )

        print("DEBUG: Gerando texto de fechamento do perímetro.")
        closing_pattern = r"Confronto\s*\d+:\s*deste\s*segue.*?essa\s*descrição\."
        texto_depois_final, subs_feitas = re.subn(
            closing_pattern,
            closing_text,
            texto_depois,
            flags=re.DOTALL | re.IGNORECASE
        )

        if subs_feitas == 0:
            print("DEBUG: Nenhuma frase de fechamento encontrada no modelo. Fechamento será adicionado ao final.")
            texto_depois_final = texto_depois.strip() + " " + closing_text
        
        texto_depois_final = texto_depois_final.replace('<PONTO>', first_point.get('<PONTO>', ''))
        for key, value in last_point.items():
              texto_depois_final = texto_depois_final.replace(key, str(value))
              
        # 5. Monta o texto final
        texto_final = texto_antes + "".join(blocos_gerados) + texto_depois_final
        
        # 6. Salva o arquivo RTF
        with open(saida_rtf, 'w', encoding='windows-1252', errors='ignore') as f:
            f.write(texto_final)
        print(f"DEBUG: Arquivo final salvo em: {saida_rtf}")

    except Exception as e:
        raise Exception(f"Erro ao processar o RTF. Detalhes: {e}")


def processar_rtf(modelo_rtf, dados_gerais, pontos, saida_rtf):
    """
    Função wrapper que mantém a assinatura original e chama o roteador de extensão.
    """
    return processar_modelo(modelo_rtf, dados_gerais, pontos, saida_rtf)


def selecionar_arquivos_e_processar():
    # ... (função inalterada)
    root = tk.Tk()
    root.withdraw()

    try:
        file_types = [
            ("Modelos de Documento", "*.rtf *.docx"),
            ("RTF Files", "*.rtf"),
            ("DOCX Files", "*.docx")
        ]
        
        modelo_caminho = filedialog.askopenfilename(title="1. Selecione o Modelo (RTF ou DOCX)", filetypes=file_types)
        if not modelo_caminho:
            messagebox.showinfo("Aviso", "Seleção cancelada.")
            return
        
        extensao_modelo = Path(modelo_caminho).suffix.lower()
        if extensao_modelo not in ('.rtf', '.docx'):
            raise ValueError("Extensão de modelo não reconhecida. Por favor, use .rtf ou .docx.")
        
        output_file_types = [
            ("Arquivo Final", f"*{extensao_modelo}"),
            ("RTF Files", "*.rtf"),
            ("DOCX Files", "*.docx")
        ]
        
        csv_pontos = filedialog.askopenfilename(title="2. Selecione o CSV com os pontos", filetypes=[("CSV Files", "*.csv")])
        if not csv_pontos:
            messagebox.showinfo("Aviso", "Seleção cancelada.")
            return

        csv_gerais = filedialog.askopenfilename(title="3. Selecione o CSV com os dados gerais", filetypes=[("CSV Files", "*.csv")])
        if not csv_gerais:
            messagebox.showinfo("Aviso", "Seleção cancelada.")
            return

        saida_caminho = filedialog.asksaveasfilename(
            title="4. Salvar arquivo final", 
            defaultextension=extensao_modelo, 
            filetypes=output_file_types
        )
        if not saida_caminho:
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
            print(f"   Número de itens: {len(dados_gerais)}")
        else:
            raise ValueError("Não foi possível carregar os dados gerais.")
        
        if pontos:
            print("Pontos carregados com sucesso.")
            print(f"   Número de pontos: {len(pontos)}")
        else:
            raise ValueError("Não foi possível carregar os pontos.")
        
        print(f"\n--- Iniciando o processamento do Modelo ({extensao_modelo}) ---")
        # Chamada da função unificada
        processar_modelo(modelo_caminho, dados_gerais, pontos, saida_caminho)
        messagebox.showinfo("Sucesso", f"Memorial gerado com sucesso:\n{saida_caminho}")
    except Exception as e:
        print(f"ERRO: {e}")
        messagebox.showerror("Erro", f"Erro ao gerar arquivo:\n{e}")

def main():
    # ... (função inalterada)
    parser = argparse.ArgumentParser(description="Gerador de Memorial Descritivo DOCX/RTF")
    # Ajuda e descrição atualizadas
    parser.add_argument("modelo", nargs="?", help="Caminho do Modelo (RTF ou DOCX)")
    parser.add_argument("base", nargs="?", help="Caminho base (sem extensão, ex: D:\\TESTE-MOD)")
    parser.add_argument("--csv_pontos", help="Caminho do CSV com pontos")
    parser.add_argument("--csv_gerais", help="Caminho do CSV com dados gerais")
    parser.add_argument("--saida", help="Caminho para salvar o arquivo final")
    args = parser.parse_args()

    try:
        # --- MODO SIMPLIFICADO ---
        if args.modelo and args.base and not (args.csv_pontos or args.csv_gerais or args.saida):
            modelo_path = Path(args.modelo)
            extensao = modelo_path.suffix.upper() # Captura a extensão do modelo
            
            csv_gerais = args.base + "1.CSV"
            csv_pontos = args.base + "2.CSV"
            saida_caminho = args.base + extensao # Usa a mesma extensão do modelo para a saída

            print(f"--- Arquivos detectados ---")
            print(f"Modelo : {args.modelo}")
            print(f"CSV Gerais : {csv_gerais}")
            print(f"CSV Pontos : {csv_pontos}")
            print(f"Saída : {saida_caminho}")

            dados_gerais = carregar_dados_gerais(csv_gerais)
            pontos = carregar_pontos(csv_pontos)
            if "<AREAHE>" in dados_gerais and "<AREAM2>" not in dados_gerais:
                dados_gerais["<AREAM2>"] = dados_gerais["<AREAHE>"]
            
            # Chamada da função unificada
            processar_modelo(args.modelo, dados_gerais, pontos, saida_caminho)
            print(f"\n✅ Memorial gerado com sucesso: {saida_caminho}")

        # --- MODO COMPLETO (flags antigas) ---
        elif args.modelo and args.csv_pontos and args.csv_gerais and args.saida:
            dados_gerais = carregar_dados_gerais(args.csv_gerais)
            pontos = carregar_pontos(args.csv_pontos)
            if "<AREAHE>" in dados_gerais and "<AREAM2>" not in dados_gerais:
                dados_gerais["<AREAM2>"] = dados_gerais["<AREAHE>"]
            
            # Chamada da função unificada
            processar_modelo(args.modelo, dados_gerais, pontos, args.saida)
            print(f"\n✅ Memorial gerado com sucesso: {args.saida}")

        # --- MODO GUI (fallback) ---
        else:
            selecionar_arquivos_e_processar()

    except Exception as e:
        print(f"ERRO FATAL: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()