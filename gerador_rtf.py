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
    # print("AVISO: 'python-docx' não está instalado. O processamento de arquivos DOCX não funcionará.")


# --- Funções de Leitura de Arquivos com Fallback de Codificação ---
def ler_arquivo_com_fallback(caminho):
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
    dados_gerais = {}
    try:
        print(f"DEBUG: Carregando dados gerais de: {caminho_csv}")
        conteudo = ler_arquivo_com_fallback(caminho_csv)
        # Usando o módulo csv para garantir a separação correta por ponto e vírgula
        leitor = csv.reader(conteudo.splitlines(), delimiter=';')
    
        linhas_restantes = list(leitor)

        for i, linha in enumerate(linhas_restantes):
            if not linha:
                continue
            
            try:
                # O arquivo de dados gerais deve ter <CHAVE>;<VALOR>
                if len(linha) >= 2 and re.match(r"<[A-Z_]+>", linha[0].strip()):
                    chave = linha[0].strip()
                    valor = linha[1].strip()
                    
                    if chave == "<PONTO>":
                         raise ValueError("Você selecionou o arquivo de pontos para os dados gerais. Por favor, selecione o arquivo de dados gerais (<CHAVE>;<VALOR>).")

                    dados_gerais[chave] = valor
                    # print(f"DEBUG: Linha {i+1} - Chave: {chave}, Valor: {valor}")
                # Ignora linhas que não são chaves válidas
                
            except IndexError:
                raise ValueError(f"Linha malformada no arquivo de dados gerais, linha {i+1}: {linha}. Esperado '<CHAVE>;<VALOR>'.")
        
        if not dados_gerais:
            # Não lança erro se o arquivo estiver vazio, mas se não encontrar dados válidos.
            pass
        
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
            # Usando split(';') sem o módulo csv porque o formato <PONTO> é mais maleável
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
                    if num_campos_esperados == 0:
                        raise ValueError("Cabeçalho '<PONTO>' está vazio ou malformado.")

                    if len(campos) < num_campos_esperados:
                        campos.extend([''] * (num_campos_esperados - len(campos)))
                    elif len(campos) > num_campos_esperados:
                        campos = campos[:num_campos_esperados]

                    ponto = dict(zip(cabecalho, [val.strip() for val in campos]))
                    pontos.append(ponto)
                    # print(f"DEBUG: Ponto {ponto.get('<PONTO>', '')} carregado.")
                except IndexError:
                    raise ValueError(f"Linha malformada no arquivo de pontos, linha {i+1}: {linha}. Verifique o número de colunas.")

        if not pontos:
            raise ValueError("Nenhum ponto foi encontrado. Verifique se o cabeçalho '<PONTO>' existe.")
            
        print(f"DEBUG: {len(pontos)} pontos carregados com sucesso.")
    except Exception as e:
        raise ValueError(f"Erro ao carregar os pontos do CSV. Detalhes: {e}")
    
    # Lógica para evitar ponto final duplicado (se o último for igual ao primeiro)
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
        run_len = len(run.text)
        
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
        
        # Clear intermediate runs
        for i in range(start_run_index + 1, end_run_index):
            paragraph.runs[i].clear()
            
        end_run.text = suffix_text
        
        # Remove empty runs elements from the underlying XML
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
                    
def processar_modelo(caminho_modelo, dados_gerais, pontos, saida_caminho, ignorar_confrontante_repetido=False):
    """
    Processa o modelo (RTF ou DOCX), preenche com os dados,
    e salva o resultado. Roteador de extensão.
    """
    caminho_modelo_path = Path(caminho_modelo)
    extensao = caminho_modelo_path.suffix.lower()
    
    try:
        if extensao == '.rtf':
            # Passa a flag para a função RTF
            return processar_rtf_string(caminho_modelo, dados_gerais, pontos, saida_caminho, ignorar_confrontante_repetido)
        
        elif extensao == '.docx':
            # --- Lógica para DOCX ---
            
            if Document is MockDocument:
                raise Exception("A biblioteca 'python-docx' não foi encontrada. Instale-a para processar arquivos DOCX (pip install python-docx).")

            print(f"DEBUG: Iniciando processamento do DOCX: {caminho_modelo}")
            
            # 1. Carrega o documento
            documento = Document(caminho_modelo)
            
            # 2. Substitui apenas os dados gerais (em todo o documento)
            print("\nDEBUG: Substituindo dados gerais...")
            substituir_texto_em_docx(documento, dados_gerais)

            # 3. Encontra e processa o bloco de repetição
            
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
                print("\nDEBUG: Bloco de repetição não encontrado ou número de pontos insuficiente. Processando como arquivo simples.")
                if paragrafo_bloco_inicio:
                    paragrafo_bloco_inicio.text = paragrafo_bloco_inicio.text.replace("<***>", "")
                documento.save(saida_caminho)
                print(f"DEBUG: Arquivo final salvo em: {saida_caminho}")
                return
                
            # 3b. Extrair o Bloco Base de repetição
            
            block_text_raw = ""
            paragraphs_to_remove = []
            
            for i in range(start_index, end_index + 1):
                p = documento.paragraphs[i]
                block_text_raw += p.text + "\n"
                paragraphs_to_remove.append(p)
                
            padrao_bloco = r"(?s)<\*\*\*>(.*?)<\*\*\*>"
            match_bloco = re.search(padrao_bloco, block_text_raw)
            
            if not match_bloco:
                raise Exception("Erro fatal ao extrair o conteúdo do bloco de repetição. Verifique se há dois marcadores '<***>' no modelo.")
            
            bloco_base_string = match_bloco.group(1).replace('<***>', '').strip()
            
            # 3c. Substitui variáveis do primeiro ponto (M01) em qualquer lugar
            primeiro_ponto = pontos[0]
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

                # --- TRATAMENTO: confrontante repetido (AGORA CONDICIONAL) ---
                if not ignorar_confrontante_repetido:
                    confrontante_atual = ponto_atual.get('<CONFRONTANTE>', '').strip()
                    if confrontante_atual and confrontante_atual == ultimo_confrontante:
                        dados_paragrafo['<CONFRONTANTE>'] = "o mesmo"
                    else:
                        dados_paragrafo['<CONFRONTANTE>'] = confrontante_atual
                        ultimo_confrontante = confrontante_atual
                else:
                    dados_paragrafo['<CONFRONTANTE>'] = ponto_atual.get('<CONFRONTANTE>', '')

                # --- INÍCIO DA CORREÇÃO MANUAL DE RECONSTRUÇÃO (Geração da String) ---
                token_ponto = "@@PONTO_NEGRITO@@"
                bloco_formatado_sem_ponto = bloco_formatado.replace('<PONTO>', token_ponto)
                
                for chave, valor in dados_paragrafo.items():
                    if chave == '<PONTO>': continue 
                    bloco_formatado_sem_ponto = bloco_formatado_sem_ponto.replace(chave, str(valor))

                if '<CONFRO>' in bloco_formatado_sem_ponto:
                    bloco_formatado_sem_ponto = bloco_formatado_sem_ponto.replace("<CONFRO>", f"Confronto {numero_confronto}")

                blocos_gerados.append(bloco_formatado_sem_ponto)
            
            # 4. Inserção dos Blocos e Limpeza
            paragrafo_ref = paragraphs_to_remove[0] if paragraphs_to_remove else paragrafo_bloco_inicio
            estilo_referencia = paragrafo_ref.style
            
            run_ref = None
            for r in paragrafo_ref.runs:
                if r.text.strip():
                    run_ref = r
                    break
            if run_ref is None and paragrafo_ref.runs:
                 run_ref = paragrafo_ref.runs[-1]
            
            font_ref = "Arial" 
            size_ref = run_ref.font.size if run_ref and run_ref.font.size else None
            bold_ref = run_ref.bold if run_ref else False
            
            paragrafo_ancora = documento.paragraphs[end_index] if end_index < len(documento.paragraphs) else documento.paragraphs[-1]
            
            for idx, bloco_texto_sem_ponto in enumerate(blocos_gerados):
                if bloco_texto_sem_ponto.strip():
                    
                    proximo_ponto_nome = pontos[idx + 1].get('<PONTO>', '')
                    
                    novo_paragrafo = paragrafo_ancora.insert_paragraph_before('') 
                    novo_paragrafo.style = estilo_referencia

                    def apply_base_format(r):
                        if font_ref: r.font.name = font_ref
                        if size_ref: r.font.size = size_ref
                        if bold_ref: r.bold = True 
                        return r
                    
                    partes = bloco_texto_sem_ponto.split(token_ponto)
                    
                    if len(partes) > 0:
                        run = novo_paragrafo.add_run(partes[0])
                        apply_base_format(run)

                    if len(partes) > 1:
                        run = novo_paragrafo.add_run(proximo_ponto_nome)
                        run.bold = True
                        if font_ref: run.font.name = font_ref
                        if size_ref: run.font.size = size_ref
                        
                        run = novo_paragrafo.add_run(partes[1])
                        apply_base_format(run)
            
            for p_remove in paragraphs_to_remove:
                if p_remove._element.getparent() is not None:
                     p_remove._element.getparent().remove(p_remove._element)


            # 5. Salva o documento
            documento.save(saida_caminho)
            print(f"DEBUG: Arquivo final salvo em: {saida_caminho}")
            return

        else:
            raise ValueError(f"Extensão de arquivo não suportada: {extensao}. Suportadas: .rtf, .docx")
            
    finally:
        pass 


def processar_rtf_string(modelo_rtf, dados_gerais, pontos, saida_rtf, ignorar_confrontante_repetido=False):
    """
    Lógica ORIGINAL para processar modelos RTF (manipulação de string).
    A lógica de confrontante repetido foi adicionada aqui.
    """
    try:
        print(f"DEBUG: Iniciando processamento do RTF (string-based): {modelo_rtf}")
        
        with open(modelo_rtf, 'r', encoding='windows-1252', errors='ignore') as f:
            texto_modelo = f.read()

        if not texto_modelo:
            raise ValueError("O arquivo modelo RTF está vazio ou não pôde ser lido.")
        
        chaves = list(dados_gerais.keys()) + list(pontos[0].keys()) + ['<***>', '<CONFRO>']
        texto_pronto = normalizar_chaves_rtf(texto_modelo, chaves)
        
        # 1. Substitui apenas os dados gerais
        print("\nDEBUG: Substituindo dados gerais...")
        for chave, valor in dados_gerais.items():
            texto_pronto = texto_pronto.replace(chave, str(valor))

        # 2. Encontra o bloco de repetição
        padrao_bloco = r"(?s)<\*\*\*>(.*?)<\*\*\*>"
        match_bloco = re.search(padrao_bloco, texto_pronto)

        if not match_bloco or len(pontos) < 2:
            print("\nDEBUG: Bloco de repetição não encontrado ou número de pontos insuficiente. Processando como arquivo simples.")
            with open(saida_rtf, 'w', encoding='windows-1252') as f:
                texto_pronto = texto_pronto.replace('<***>', '')
                f.write(texto_pronto)
            print(f"DEBUG: Arquivo final salvo em: {saida_rtf}")
            return

        # Separar o texto em partes
        texto_antes = texto_pronto[:match_bloco.start()]
        bloco_base = match_bloco.group(1).replace('<***>', '')
        texto_depois = texto_pronto[match_bloco.end():]
        blocos_gerados = []
        
        # --- ABERTURA: substitui variáveis do primeiro ponto (M01) no texto antes do bloco ---
        first_point = pontos[0]
        for key, value in first_point.items():
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
                '<PONTO>': next_point.get('<PONTO>', ''), # PONTO de DESTINO
                '<UTMX>': next_point.get('<UTMY>', ''), # Troca de UTMX/UTMY foi mantida
                '<UTMY>': next_point.get('<UTMX>', ''), # Troca de UTMX/UTMY foi mantida
            })

            # --- TRATAMENTO: confrontante repetido (AGORA CONDICIONAL) ---
            if not ignorar_confrontante_repetido:
                current_confrontante = current_point.get('<CONFRONTANTE>', '').strip()
                if current_confrontante and current_confrontante == last_confrontante:
                    paragraph_data['<CONFRONTANTE>'] = "o mesmo"
                else:
                    paragraph_data['<CONFRONTANTE>'] = current_confrontante
                    last_confrontante = current_confrontante
            else:
                paragraph_data['<CONFRONTANTE>'] = current_point.get('<CONFRONTANTE>', '')

            # Substitui placeholders
            for key, value in paragraph_data.items():
                formatted_block = formatted_block.replace(key, str(value))

            if '<CONFRO>' in formatted_block:
                formatted_block = formatted_block.replace("<CONFRO>", f"Confronto {confronto_num}")

            blocos_gerados.append(formatted_block)


        # --- FECHAMENTO: último ponto voltando para M01 ---
        # Usa o penúltimo ponto do CSV (pontos[-2]) para o último confronto
        last_point_of_perimeter = pontos[-2]
        first_point = pontos[0]
        confronto_num = len(pontos) - 1 # Número total de confrontos

        closing_text = (
            f"Confronto {confronto_num}: deste segue confrontando com a propriedade de {last_point_of_perimeter.get('<CONFRONTANTE>', '')}, "
            f"com azimute de {last_point_of_perimeter.get('<AZIMUTE>', '')} por uma distância de {last_point_of_perimeter.get('<DISTANCIA>', '')}m, "
            f"até o ponto {first_point.get('<PONTO>', '')}, onde teve início essa descrição."
        )

        print(f"DEBUG: Gerando texto de fechamento do perímetro ({confronto_num}º confronto).")
        
        closing_pattern = r"Confronto\s*\d+:\s*deste\s*segue.*?essa\s*descrição\."
        
        texto_depois_final, subs_feitas = re.subn(
            closing_pattern,
            closing_text,
            texto_depois,
            flags=re.DOTALL | re.IGNORECASE
        )

        if subs_feitas == 0:
            print("DEBUG: Nenhuma frase de fechamento padrão encontrada. Tentando substituição de chaves no texto restante.")
            texto_depois_final = texto_depois
        
        # Substitui chaves restantes no texto final (Azimute/Distancia/Confrontante do último segmento e o Ponto inicial)
        texto_depois_final = texto_depois_final.replace('<PONTO>', first_point.get('<PONTO>', ''))
        
        for key, value in last_point_of_perimeter.items():
             texto_depois_final = texto_depois_final.replace(key, str(value))
             
        # 5. Monta o texto final
        texto_final = texto_antes + "".join(blocos_gerados) + texto_depois_final
        
        # 6. Salva o arquivo RTF
        with open(saida_rtf, 'w', encoding='windows-1252') as f:
            f.write(texto_final)
        print(f"DEBUG: Arquivo final salvo em: {saida_rtf}")

    except Exception as e:
        raise Exception(f"Erro ao processar o RTF. Detalhes: {e}")


def selecionar_arquivos_e_processar():
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
        if not dados_gerais:
             print("Aviso: Nenhum dado geral carregado.")
        
        if pontos:
            print(f"Pontos carregados com sucesso: {len(pontos)}.")
        else:
            raise ValueError("Não foi possível carregar os pontos.")
        
        print(f"\n--- Iniciando o processamento do Modelo ({extensao_modelo}) ---")
        # Chamada da função unificada (padrão para GUI é não ignorar a repetição)
        processar_modelo(modelo_caminho, dados_gerais, pontos, saida_caminho, ignorar_confrontante_repetido=False)
        messagebox.showinfo("Sucesso", f"Memorial gerado com sucesso:\n{saida_caminho}")
    except Exception as e:
        # Apenas mostra o erro na tela (GUI)
        print(f"ERRO: {e}")
        messagebox.showerror("Erro", f"Erro ao gerar arquivo:\n{e}")

def main():
    parser = argparse.ArgumentParser(description="Gerador de Memorial Descritivo DOCX/RTF")
    # ALTERAÇÃO: Argumento simplificado para --x
    parser.add_argument("--x", action="store_true", 
                        help="Desativa a substituição de confrontantes repetidos por 'o mesmo'.")
                        
    # Argumentos originais
    parser.add_argument("modelo", nargs="?", help="Caminho do Modelo (RTF ou DOCX)")
    parser.add_argument("base", nargs="?", help="Caminho base (sem extensão, ex: D:\\TESTE-MOD)")
    parser.add_argument("--csv_pontos", help="Caminho do CSV com pontos")
    parser.add_argument("--csv_gerais", help="Caminho do CSV com dados gerais")
    parser.add_argument("--saida", help="Caminho para salvar o arquivo final")
    args = parser.parse_args()
    
    # Mapeia a nova flag para o parâmetro da função:
    ignorar_rep = args.x
    
    # Adicionando um aviso se a biblioteca docx não estiver disponível
    if args.modelo and Path(args.modelo).suffix.lower() == '.docx' and Document is MockDocument:
         print("AVISO: 'python-docx' não está instalado. O processamento de arquivos DOCX NÃO funcionará.", file=sys.stderr)

    try:
        # --- MODO SIMPLIFICADO ---
        if args.modelo and args.base and not (args.csv_pontos or args.csv_gerais or args.saida):
            modelo_path = Path(args.modelo)
            # Usa .lower() para garantir consistência do nome de saída
            extensao = modelo_path.suffix.lower() 
            
            if not extensao: extensao = ".rtf" 
            
            # Formato esperado para o modo simplificado
            csv_gerais = args.base + "1.CSV"
            csv_pontos = args.base + "2.CSV"
            saida_caminho = args.base + extensao 

            print(f"--- MODO SIMPLIFICADO DETECTADO ({'Ignorar Repetição: SIM' if ignorar_rep else 'Ignorar Repetição: NÃO'}) ---")
            print(f"Modelo : {args.modelo}")
            print(f"CSV Gerais : {csv_gerais}")
            print(f"CSV Pontos : {csv_pontos}")
            print(f"Saída Esperada : {saida_caminho}")

            dados_gerais = carregar_dados_gerais(csv_gerais)
            pontos = carregar_pontos(csv_pontos)
            
            if "<AREAHE>" in dados_gerais and "<AREAM2>" not in dados_gerais:
                dados_gerais["<AREAM2>"] = dados_gerais["<AREAHE>"]
            
            # Chamada da função unificada (passando a flag)
            processar_modelo(args.modelo, dados_gerais, pontos, saida_caminho, ignorar_confrontante_repetido=ignorar_rep)
            print(f"\n✅ Memorial gerado com sucesso: {saida_caminho}")

        # --- MODO COMPLETO (flags antigas) ---
        elif args.modelo and args.csv_pontos and args.csv_gerais and args.saida:
            print(f"--- MODO COMPLETO DETECTADO ({'Ignorar Repetição: SIM' if ignorar_rep else 'Ignorar Repetição: NÃO'}) ---")
            
            dados_gerais = carregar_dados_gerais(args.csv_gerais)
            pontos = carregar_pontos(args.csv_pontos)
            if "<AREAHE>" in dados_gerais and "<AREAM2>" not in dados_gerais:
                dados_gerais["<AREAM2>"] = dados_gerais["<AREAHE>"]
            
            # Chamada da função unificada (passando a flag)
            processar_modelo(args.modelo, dados_gerais, pontos, args.saida, ignorar_confrontante_repetido=ignorar_rep)
            print(f"\n✅ Memorial gerado com sucesso: {args.saida}")

        # --- MODO GUI (fallback) ---
        else:
            print("--- MODO GRÁFICO (GUI) DETECTADO ---")
            selecionar_arquivos_e_processar()

    except Exception as e:
        print(f"ERRO FATAL: Ocorreu um erro durante o processamento. Detalhes: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
