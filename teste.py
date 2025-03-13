import csv
import os
import pandas as pd
import re
from datetime import datetime, date
from difflib import SequenceMatcher

def formatar_data(data_str):
    """
    Converte uma data no formato 'DD/MM' para 'DD de mês de 2025'
    
    Args:
        data_str: String de data no formato 'DD/MM'
        
    Returns:
        String formatada como 'DD de mês de 2025'
    """
    try:
        partes = data_str.split('/')
        dia = int(partes[0])
        mes = int(partes[1])
        
        # Dicionário para converter número do mês para nome
        meses = {
            1: 'janeiro',
            2: 'fevereiro',
            3: 'março',
            4: 'abril',
            5: 'maio',
            6: 'junho',
            7: 'julho',
            8: 'agosto',
            9: 'setembro',
            10: 'outubro',
            11: 'novembro',
            12: 'dezembro'
        }
        
        # Formar a data no formato solicitado
        return f"{dia} de {meses[mes]} de 2025"
    except:
        # Se houver algum erro na conversão, retorna a data original
        return data_str


def obter_nome_curso_para_certificado(curso_original):
    """
    Retorna o nome do curso formatado para o certificado
    
    Args:
        curso_original: Nome original do curso
        
    Returns:
        Nome do curso formatado para o certificado
    """
    curso_lower = curso_original.lower()
    
    if curso_lower == "pc gamer":
        return "Montagem e configuração de computadores de alto desempenho - PC Gamer"
    elif curso_lower == "m. celular" or curso_lower == "manutenção de celulares":
        return "Manutenção de Celulares"
    elif curso_lower == "robótica" or curso_lower == "robótica, programação":
        return "Robótica, Programação"
    elif curso_lower == "dev. jogos":
        return "Desenvolvimento de Jogos"
    else:
        return curso_original


def normalizar_nome(nome):
    """
    Normaliza um nome para comparação, removendo espaços extras, maiúsculas, acentos, etc.
    
    Args:
        nome: O nome a ser normalizado
        
    Returns:
        O nome normalizado
    """
    # Garantir que o nome seja uma string
    if nome is None:
        return ""
    
    # Converter para string caso seja outro tipo
    if not isinstance(nome, str):
        try:
            nome = str(nome)
        except:
            return ""
    
    # Remover espaços extras, converter para minúsculas
    nome_norm = nome.lower().strip()
    
    # Remover termos especiais como "(prof)"
    nome_norm = re.sub(r'\([^)]*\)', '', nome_norm).strip()
    
    # Remover acentos
    acentos = {
        'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a', 'ä': 'a',
        'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
        'í': 'i', 'ì': 'i', 'î': 'i', 'ï': 'i',
        'ó': 'o', 'ò': 'o', 'ô': 'o', 'õ': 'o', 'ö': 'o',
        'ú': 'u', 'ù': 'u', 'û': 'u', 'ü': 'u',
        'ç': 'c', 'ñ': 'n'
    }
    
    for acento, sem_acento in acentos.items():
        nome_norm = nome_norm.replace(acento, sem_acento)
    
    # Remover espaços múltiplos
    nome_norm = re.sub(r'\s+', ' ', nome_norm)
    
    return nome_norm


def capitalizar_nome(nome):
    """
    Capitaliza as primeiras letras de cada palavra do nome
    
    Args:
        nome: O nome a ser capitalizado
        
    Returns:
        O nome com as primeiras letras maiúsculas
    """
    # Garantir que o nome seja uma string
    if nome is None:
        return ""
    
    # Converter para string caso seja outro tipo
    if not isinstance(nome, str):
        try:
            nome = str(nome)
        except:
            return ""
    
    if not nome:
        return ""
    
    # Tratar palavras que não devem ser capitalizadas
    palavras_minusculas = ['de', 'da', 'do', 'dos', 'das', 'e']
    
    # Dividir o nome em palavras e capitalizar cada uma
    palavras = nome.lower().split()
    resultado = []
    
    for i, palavra in enumerate(palavras):
        # Se não for a primeira palavra e for uma palavra minúscula
        if i > 0 and palavra in palavras_minusculas:
            resultado.append(palavra)
        else:
            resultado.append(palavra.capitalize())
    
    return ' '.join(resultado)


def calcular_similaridade(nome1, nome2):
    """
    Calcula a similaridade entre dois nomes usando SequenceMatcher
    
    Args:
        nome1: Primeiro nome
        nome2: Segundo nome
        
    Returns:
        Valor de similaridade entre 0 e 1
    """
    return SequenceMatcher(None, nome1, nome2).ratio()


def nomes_similares(nome1, nome2, limiar=0.6):
    """
    Verifica se dois nomes são similares, considerando um limiar de similaridade.
    
    Args:
        nome1: Primeiro nome
        nome2: Segundo nome
        limiar: Limiar de similaridade (0.0 a 1.0)
        
    Returns:
        True se os nomes forem similares, False caso contrário
    """
    nome1_norm = normalizar_nome(nome1)
    nome2_norm = normalizar_nome(nome2)
    
    # Primeiro critério: igualdade após normalização
    if nome1_norm == nome2_norm:
        return True
    
    # Calcular similaridade direta
    similaridade_direta = calcular_similaridade(nome1_norm, nome2_norm)
    if similaridade_direta >= limiar:
        return True
    
    # Segundo critério: um contém o outro
    if nome1_norm in nome2_norm or nome2_norm in nome1_norm:
        return True
    
    # Terceiro critério: similaridade de palavras
    palavras1 = set(nome1_norm.split())
    palavras2 = set(nome2_norm.split())
    
    # Ignorar palavras muito comuns que podem causar falsos positivos
    palavras_comuns = {'de', 'da', 'do', 'dos', 'das', 'e', 'a', 'o', 'as', 'os'}
    palavras1 = palavras1 - palavras_comuns
    palavras2 = palavras2 - palavras_comuns
    
    # Se um dos conjuntos estiver vazio após remover palavras comuns, usar os originais
    if not palavras1 or not palavras2:
        palavras1 = set(nome1_norm.split())
        palavras2 = set(nome2_norm.split())
    
    # Calcular similaridade de Jaccard (proporção de palavras comuns)
    intersecao = palavras1.intersection(palavras2)
    uniao = palavras1.union(palavras2)
    
    if len(uniao) == 0:
        return False
    
    similaridade = len(intersecao) / len(uniao)
    
    # Verificar similaridade mínima
    if similaridade >= limiar:
        return True
    
    # Quarto critério: correspondência de partes importantes do nome
    # Verificar se sobrenomes importantes coincidem
    sobrenomes1 = [p for p in palavras1 if len(p) > 3 and p not in {'jose', 'maria', 'joao', 'pedro', 'ana'}]
    sobrenomes2 = [p for p in palavras2 if len(p) > 3 and p not in {'jose', 'maria', 'joao', 'pedro', 'ana'}]
    
    if sobrenomes1 and sobrenomes2:
        # Se pelo menos 50% dos sobrenomes importantes coincidem
        sobrenomes_comuns = set(sobrenomes1).intersection(set(sobrenomes2))
        if sobrenomes_comuns:
            proporcao1 = len(sobrenomes_comuns) / len(sobrenomes1)
            proporcao2 = len(sobrenomes_comuns) / len(sobrenomes2)
            if proporcao1 >= 0.5 or proporcao2 >= 0.5:
                return True
    
    return False


def mapear_curso_inscricao(curso_frequencia):
    """
    Mapeia o curso da frequência para o formato correspondente na inscrição
    
    Args:
        curso_frequencia: Nome do curso no arquivo de frequência
        
    Returns:
        Nome do curso no formato do arquivo de inscrição
    """
    mapeamento = {
        "PC Gamer": "Montagem e Configuração de Computadores de Alto Desempenho (PC GAMER)",
        "Robótica": "Robótica",
        "Robótica, Programação": "Robótica",
        "Manutenção de Celulares": "Manutenção de Celulares",
        "M. Celular": "Manutenção de Celulares",
        "Dev. Jogos": "Desenvolvimento de Jogos"
    }
    
    return mapeamento.get(curso_frequencia, curso_frequencia)


def cursos_correspondentes(curso_freq, curso_insc):
    """
    Verifica se o curso da frequência corresponde ao curso da inscrição
    
    Args:
        curso_freq: Nome do curso no arquivo de frequência
        curso_insc: Nome do curso no arquivo de inscrição
        
    Returns:
        True se os cursos corresponderem, False caso contrário
    """
    # Normalizar os cursos para comparação
    curso_freq_norm = normalizar_nome(curso_freq)
    curso_insc_norm = normalizar_nome(curso_insc)
    
    # Verificar correspondência direta
    if curso_freq_norm == curso_insc_norm:
        return True
    
    # Verificar correspondências específicas
    if "pcgamer" in curso_freq_norm and "pcgamer" in curso_insc_norm:
        return True
    if "robotica" in curso_freq_norm and "robotica" in curso_insc_norm:
        return True
    if ("celular" in curso_freq_norm or "m.celular" == curso_freq_norm) and "celular" in curso_insc_norm:
        return True
    if ("celular" in curso_insc_norm or "m.celular" == curso_insc_norm) and "celular" in curso_freq_norm:
        return True
    if ("dev.jogos" == curso_freq_norm or "desenvolvimentodejogos" in curso_freq_norm) and ("desenvolvimentodejogos" in curso_insc_norm or "dev.jogos" == curso_insc_norm):
        return True
    if ("robotica,programacao" in curso_freq_norm or "roboticaprogramacao" in curso_freq_norm) and "robotica" in curso_insc_norm:
        return True
    
    # Correspondências com abreviações específicas
    if curso_freq_norm == "m.celular" and "manutencao" in curso_insc_norm:
        return True
    if "manutencaocelular" in curso_freq_norm and curso_insc_norm == "m.celular":
        return True
    if curso_freq_norm == "dev.jogos" and "jogos" in curso_insc_norm:
        return True
    
    # Verificar usando o mapeamento
    curso_mapeado = mapear_curso_inscricao(curso_freq)
    if normalizar_nome(curso_mapeado) == curso_insc_norm:
        return True
    
    return False


def processar_frequencia_pandas(arquivo_frequencia):
    """
    Processa o arquivo de frequência usando pandas e retorna a lista de alunos aprovados.
    
    Um aluno é aprovado se tiver pelo menos 60% de presença (P ou FJ) nos dias válidos.
    Dias válidos são aqueles cujos valores são P, F ou FJ. Outros valores são ignorados.
    
    Args:
        arquivo_frequencia: Caminho para o arquivo CSV de frequência
        
    Returns:
        Tupla contendo: (lista de alunos aprovados, lista de alunos reprovados, total de alunos)
    """
    # Ler o arquivo CSV com pandas
    try:
        df = pd.read_csv(arquivo_frequencia)
    except Exception as e:
        print(f"Erro ao ler o arquivo de frequência: {str(e)}")
        # Tentar com codificação diferente
        try:
            df = pd.read_csv(arquivo_frequencia, encoding='latin1')
        except Exception as e2:
            print(f"Segundo erro ao ler arquivo: {str(e2)}")
            raise
    
    # Total de alunos
    total_alunos = len(df)
    
    # Identificar as colunas de data (últimas 5 colunas)
    colunas_data = df.columns[-5:]
    data_conclusao = colunas_data[-1]  # Última coluna de data
    
    # Listas para armazenar os resultados
    alunos_aprovados = []
    alunos_reprovados = []
    
    # Processar cada linha
    for idx, row in df.iterrows():
        # Converter valores para string para evitar problemas de tipo
        nome_aluno = str(row['ALUNOS']) if pd.notna(row['ALUNOS']) else ""
        curso = str(row['CURSO']) if pd.notna(row['CURSO']) else ""
        
        # Obter os valores de presença das últimas 5 colunas e converter para string
        valores_presenca = []
        for col in colunas_data:
            valor = row[col]
            if pd.isna(valor):
                valor = ""
            else:
                valor = str(valor).upper()
            valores_presenca.append(valor)
        
        # Contar apenas os dias válidos (com valores P, F ou FJ)
        dias_validos = [valor for valor in valores_presenca if valor in ['P', 'F', 'FJ']]
        total_dias_validos = len(dias_validos)
        
        # Se não houver dias válidos, o aluno é reprovado
        if total_dias_validos == 0:
            alunos_reprovados.append({
                'NOME': nome_aluno,
                'CURSO': curso
            })
            continue
            
        # Contar as presenças (P ou FJ)
        dias_presentes = sum(1 for valor in dias_validos if valor in ['P', 'FJ'])
        
        # Calcular a porcentagem de presença
        porcentagem_presenca = (dias_presentes / total_dias_validos) * 100
        
        # Verificar se o aluno foi aprovado (pelo menos 60% de presença)
        if porcentagem_presenca >= 60:
            # Converter a data de conclusão para o formato desejado
            data_formatada = formatar_data(data_conclusao)
            
            # Obter o nome do curso formatado para o certificado
            curso_certificado = obter_nome_curso_para_certificado(curso)
            
            # Normalizar o nome para capitalização correta
            nome_capitalizado = capitalizar_nome(nome_aluno)
            
            alunos_aprovados.append({
                'NOME': nome_capitalizado,
                'NOME_ORIGINAL': nome_aluno,
                'CURSO': curso_certificado,
                'CURSO_ORIGINAL': curso,
                'DATA_CONCLUSAO': data_formatada
            })
        else:
            alunos_reprovados.append({
                'NOME': nome_aluno,
                'CURSO': curso
            })
    
    return alunos_aprovados, alunos_reprovados, total_alunos


def encontrar_melhor_correspondencia(nome_aluno, inscritos, limiar=0.65):
    """
    Encontra a melhor correspondência para um nome de aluno entre os inscritos
    
    Args:
        nome_aluno: Nome do aluno na frequência
        inscritos: Lista de dicionários com inscrições
        limiar: Limiar mínimo de similaridade para considerar correspondência
        
    Returns:
        O índice da melhor correspondência ou None se não encontrar
    """
    nome_aluno_norm = normalizar_nome(nome_aluno)
    melhor_idx = None
    melhor_similaridade = 0
    
    for idx, inscrito in enumerate(inscritos):
        nome_inscrito = inscrito['Nome completo'] if 'Nome completo' in inscrito else ''
        nome_inscrito_norm = normalizar_nome(nome_inscrito)
        
        # Verificar similaridade
        similaridade = calcular_similaridade(nome_aluno_norm, nome_inscrito_norm)
        
        # Verificar se é melhor que a anterior e atende ao limiar
        if similaridade > melhor_similaridade and similaridade >= limiar:
            melhor_similaridade = similaridade
            melhor_idx = idx
        
        # Se encontrar correspondência exata, retornar imediatamente
        if similaridade > 0.9:
            return idx
    
    return melhor_idx


def processar_inscricao_pandas(arquivo_inscricao, alunos_aprovados, arquivo_saida):
    """
    Processa o arquivo de inscrição usando pandas e gera o arquivo final com todos os dados solicitados.
    
    Args:
        arquivo_inscricao: Caminho para o arquivo CSV de inscrição
        alunos_aprovados: Lista de alunos aprovados por frequência
        arquivo_saida: Caminho para o arquivo CSV de saída
    """
    try:
        # Ler o arquivo de inscrição com pandas
        try:
            df_inscricao = pd.read_csv(arquivo_inscricao)
        except Exception as e:
            print(f"Erro ao ler o arquivo de inscrição: {str(e)}")
            # Tentar com codificação diferente
            try:
                df_inscricao = pd.read_csv(arquivo_inscricao, encoding='latin1')
            except Exception as e2:
                print(f"Segundo erro ao ler arquivo: {str(e2)}")
                raise
        
        # Converter DataFrame para lista de dicionários para facilitar o processamento
        inscritos = df_inscricao.to_dict('records')
        
        # Campos que serão incluídos no arquivo final
        campos_finais = [
            'DATA_ADESAO', 'ESTADO', 'ESCOLA', 'NOME', 'CURSO', 'TELEFONE', 
            'EMAIL', 'CPF', 'DIA', 'MES', 'ANO', 'IDADE', 'COR_PELE', 
            'SEXO', 'SERIE_ESCOLAR', 'DATA_CONCLUSAO'
        ]
        
        # Mapeamento dos campos do arquivo de inscrição para os campos finais
        mapeamento_campos = {
            'Carimbo de data/hora': 'DATA_ADESAO',
            'Estado': 'ESTADO',
            'Escreva o nome da Unidade de Ensino que você frequenta atualmente:': 'ESCOLA',
            'Nome completo': 'NOME',
            'Telefone celular do ALUNO (que use WHATSAPP)': 'TELEFONE',
            'E-mail': 'EMAIL',
            'CPF': 'CPF',
            'Data de Nascimento (DIA)': 'DIA',
            'Data de Nascimento (MÊS)': 'MES',
            'Data de Nascimento (ANO)': 'ANO',
            'Idade': 'IDADE',
            'Cor da pele / Raça / Etnia': 'COR_PELE',
            'Sexo': 'SEXO',
            'Qual nível ou série escolar você está frequentando atualmente?': 'SERIE_ESCOLAR'
        }
        
        # Lista para armazenar os dados finais
        dados_finais = []
        
        # Dicionário para verificar se já incluímos um aluno+curso específico
        alunos_incluidos = {}
        
        # Contador de correspondências encontradas
        correspondencias_encontradas = 0
        correspondencias_por_curso = {}
        
        # Para cada aluno aprovado, encontrar sua inscrição correspondente
        for aluno in alunos_aprovados:
            nome_aluno = aluno['NOME_ORIGINAL']
            nome_normalizado = aluno['NOME']
            curso_frequencia = aluno['CURSO_ORIGINAL']
            curso_certificado = aluno['CURSO']
            data_conclusao = aluno['DATA_CONCLUSAO']
            
            # Registrar estatísticas por curso
            if curso_frequencia not in correspondencias_por_curso:
                correspondencias_por_curso[curso_frequencia] = {'total': 0, 'encontrados': 0}
            correspondencias_por_curso[curso_frequencia]['total'] += 1
            
            # Garantir que nome_aluno seja uma string
            if not isinstance(nome_aluno, str):
                nome_aluno = str(nome_aluno)
            
            # Chave única para este aluno+curso
            chave_unica = f"{normalizar_nome(nome_aluno)}|{curso_frequencia}"
            
            # Verificar se já incluímos este aluno+curso
            if chave_unica in alunos_incluidos:
                continue
            
            # Encontrar a melhor correspondência para este aluno
            inscritos_filtrados = []
            for inscrito in inscritos:
                nome_inscrito = inscrito.get('Nome completo', '')
                curso_inscrito = inscrito.get('Para qual curso você quer se inscrever?', '')
                
                # Converter para string se necessário
                if not isinstance(nome_inscrito, str):
                    try:
                        nome_inscrito = str(nome_inscrito)
                    except:
                        nome_inscrito = ""
                
                if not isinstance(curso_inscrito, str):
                    try:
                        curso_inscrito = str(curso_inscrito)
                    except:
                        curso_inscrito = ""
                
                # Verificar se o nome e curso correspondem
                try:
                    nome_similar = nomes_similares(nome_aluno, nome_inscrito)
                    curso_correspondente = cursos_correspondentes(curso_frequencia, curso_inscrito)
                    
                    if nome_similar and curso_correspondente:
                        inscritos_filtrados.append(inscrito)
                except Exception as e:
                    print(f"Erro ao comparar nomes: {nome_aluno} vs {nome_inscrito}")
                    print(f"Erro: {str(e)}")
                    continue
            
            # Se encontramos correspondências, ordenar por data (mais recente primeiro)
            inscricao_aluno = None
            if inscritos_filtrados:
                # Ordenar por data/hora de inscrição (se disponível)
                try:
                    inscritos_filtrados.sort(
                        key=lambda x: str(x.get('Carimbo de data/hora', '')),
                        reverse=True
                    )
                except Exception as e:
                    print(f"Erro ao ordenar inscrições: {str(e)}")
                
                inscricao_aluno = inscritos_filtrados[0]
            else:
                # Se não encontrou correspondência específica para o curso, buscar a melhor correspondência pelo nome
                try:
                    idx_melhor = encontrar_melhor_correspondencia(nome_aluno, inscritos)
                    if idx_melhor is not None:
                        inscricao_aluno = inscritos[idx_melhor]
                except Exception as e:
                    print(f"Erro ao encontrar melhor correspondência: {str(e)}")
            
            # Se encontrou uma inscrição para o aluno
            if inscricao_aluno:
                # Marcar como incluído
                alunos_incluidos[chave_unica] = True
                correspondencias_encontradas += 1
                correspondencias_por_curso[curso_frequencia]['encontrados'] += 1
                
                # Criar um novo dicionário com os campos finais
                aluno_final = {}
                
                # Garantir que todos os campos estejam inicializados
                for campo in campos_finais:
                    aluno_final[campo] = ""
                
                # Mapear os campos do arquivo de inscrição para os campos finais
                for campo_inscricao, campo_final in mapeamento_campos.items():
                    # Verificar se o campo existe no arquivo de inscrição
                    if campo_inscricao in inscricao_aluno and pd.notna(inscricao_aluno[campo_inscricao]):
                        # Armazenar o valor, fazendo conversões necessárias
                        valor = inscricao_aluno[campo_inscricao]
                        
                        # Converter para string se não for
                        if not isinstance(valor, str):
                            try:
                                valor = str(valor)
                            except:
                                valor = ""
                        
                        # Tratar o valor conforme o campo
                        if campo_final == 'NOME':
                            # Usar o nome normalizado da frequência
                            valor = nome_normalizado
                        elif campo_final == 'ESTADO' and valor:
                            # Normalizar sigla de estado
                            valor = str(valor).strip().upper()
                        elif campo_final in ['DIA', 'ANO', 'IDADE'] and pd.notna(valor):
                            # Garantir que seja um inteiro
                            try:
                                valor = int(float(valor))
                            except:
                                pass
                        
                        aluno_final[campo_final] = valor
                
                # Adicionar a data de conclusão e o curso formatado do aluno aprovado
                aluno_final['DATA_CONCLUSAO'] = data_conclusao
                aluno_final['CURSO'] = curso_certificado
                
                # Garantir que o nome seja capitalizado corretamente
                if aluno_final['NOME']:
                    aluno_final['NOME'] = capitalizar_nome(aluno_final['NOME'])
                
                dados_finais.append(aluno_final)
        
        # Converter para DataFrame para facilitar a escrita
        df_final = pd.DataFrame(dados_finais)
        
        # Garantir que todas as colunas estejam presentes no DataFrame
        for campo in campos_finais:
            if campo not in df_final.columns:
                df_final[campo] = ""
        
        # Reordenar as colunas conforme os campos finais
        df_final = df_final[campos_finais]
        
        # Escrever os dados finais no arquivo de saída
        try:
            df_final.to_csv(arquivo_saida, index=False, encoding='utf-8')
            print(f"Arquivo '{arquivo_saida}' gerado com sucesso!")
            print(f"Total de alunos no arquivo final: {len(df_final)} de {len(alunos_aprovados)} aprovados ({(len(df_final)/len(alunos_aprovados)*100):.1f}%)")
        except Exception as e:
            print(f"Erro ao escrever arquivo: {str(e)}")
            import traceback
            traceback.print_exc()
        
        # Verificar quais alunos aprovados não foram incluídos no arquivo final
        alunos_nao_incluidos = []
        for aluno in alunos_aprovados:
            nome_aluno = aluno['NOME_ORIGINAL']
            curso = aluno['CURSO_ORIGINAL']
            
            # Garantir que nome_aluno seja uma string
            if not isinstance(nome_aluno, str):
                nome_aluno = str(nome_aluno)
                
            chave = f"{normalizar_nome(nome_aluno)}|{curso}"
            if chave not in alunos_incluidos:
                alunos_nao_incluidos.append(f"{nome_aluno} - {curso}")
        
        if alunos_nao_incluidos:
            print(f"\nAlunos aprovados mas não incluídos (por falta de inscrição ou inscrição para curso incorreto): {len(alunos_nao_incluidos)}")
            print("-" * 80)
            print("| {:<4} | {:<40} | {:<20} |".format("Nº", "Nome do Aluno", "Curso"))
            print("|" + "-"*6 + "|" + "-"*42 + "|" + "-"*22 + "|")
            
            for i, aluno_info in enumerate(alunos_nao_incluidos, 1):
                partes = aluno_info.split(" - ", 1)
                nome = partes[0] if len(partes) > 0 else ""
                curso = partes[1] if len(partes) > 1 else ""
                print("| {:<4} | {:<40} | {:<20} |".format(i, nome[:40], curso[:20]))
        
        # Exibir estatísticas por curso
        print("\n" + "-"*80)
        print("ESTATÍSTICAS POR CURSO")
        print("-"*80)
        print("| {:<30} | {:<10} | {:<12} | {:<10} |".format("Curso", "Aprovados", "Certificados", "Taxa (%)"))
        print("|" + "-"*32 + "|" + "-"*12 + "|" + "-"*14 + "|" + "-"*12 + "|")
        
        for curso, stats in sorted(correspondencias_por_curso.items()):
            total = stats['total']
            encontrados = stats['encontrados']
            taxa = (encontrados / total * 100) if total > 0 else 0
            print("| {:<30} | {:<10} | {:<12} | {:<10.1f} |".format(
                curso[:30], total, encontrados, taxa
            ))
        
        print("\n" + "-"*80)
        
    except Exception as e:
        print(f"Erro geral em processar_inscricao_pandas: {str(e)}")
        import traceback
        traceback.print_exc()


def ordenar_por_nome(aluno):
    """
    Função para ordenação segura por nome, tratando diferentes tipos de dados
    
    Args:
        aluno: Dicionário com informações do aluno
        
    Returns:
        Nome do aluno convertido para string para ordenação segura
    """
    nome = aluno.get('NOME', '')
    if nome is None:
        return ""
    return str(nome).lower()


def main():
    arquivo_frequencia = 'frequencia.csv'
    arquivo_inscricao = 'inscricao.csv'
    arquivo_saida = 'certificados_final.csv'
    
    if not os.path.exists(arquivo_frequencia):
        print(f"Erro: O arquivo '{arquivo_frequencia}' não foi encontrado.")
        return
        
    if not os.path.exists(arquivo_inscricao):
        print(f"Erro: O arquivo '{arquivo_inscricao}' não foi encontrado.")
        return
    
    try:
        # Processar o arquivo de frequência para obter os alunos aprovados e reprovados
        print("\n" + "="*80)
        print("PROCESSAMENTO DE FREQUÊNCIA E INSCRIÇÕES PARA CERTIFICADOS")
        print("="*80)
        
        print("\nProcessando arquivo de frequência...")
        alunos_aprovados, alunos_reprovados, total_alunos = processar_frequencia_pandas(arquivo_frequencia)
        
        # Exibir informações sobre o total de alunos e aprovações/reprovações
        print("\n" + "-"*80)
        print("RESUMO DE FREQUÊNCIA")
        print("-"*80)
        print(f"Total de Alunos na Chamada: {total_alunos}")
        print(f"Total de alunos aprovados por frequência: {len(alunos_aprovados)} ({(len(alunos_aprovados)/total_alunos*100):.1f}%)")
        print(f"Total de alunos reprovados por falta: {len(alunos_reprovados)} ({(len(alunos_reprovados)/total_alunos*100):.1f}%)")
        
        # Listar os alunos reprovados por falta de forma mais detalhada
        if alunos_reprovados:
            print("\n" + "-"*80)
            print("DETALHAMENTO DE ALUNOS REPROVADOS POR FALTA")
            print("-"*80)
            print("| {:<4} | {:<40} | {:<20} |".format("Nº", "Nome do Aluno", "Curso"))
            print("|" + "-"*6 + "|" + "-"*42 + "|" + "-"*22 + "|")
            
            for i, aluno in enumerate(sorted(alunos_reprovados, key=ordenar_por_nome), 1):
                nome = str(aluno['NOME'])
                curso = str(aluno['CURSO'])
                print("| {:<4} | {:<40} | {:<20} |".format(i, nome[:40], curso[:20]))
        
        # Processar o arquivo de inscrição e gerar o arquivo final
        print("\n" + "-"*80)
        print("PROCESSAMENTO DE INSCRIÇÕES")
        print("-"*80)
        print("Processando arquivo de inscrição e gerando certificados...")
        processar_inscricao_pandas(arquivo_inscricao, alunos_aprovados, arquivo_saida)
        
        # Exibir estatísticas finais
        print("\n" + "-"*80)
        print("RESUMO FINAL")
        print("-"*80)
        print(f"Total de alunos na chamada: {total_alunos}")
        print(f"Total de alunos aprovados: {len(alunos_aprovados)}")
        print(f"Total de alunos reprovados: {len(alunos_reprovados)}")
        
    except Exception as e:
        print(f"\nErro ao processar os arquivos: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()