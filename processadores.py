import pandas as pd
import os
from datetime import datetime
import re

from utils import (
    normalizar_nome, capitalizar_nome, formatar_data, 
    obter_nome_curso_para_certificado, nomes_similares, 
    cursos_correspondentes, encontrar_melhor_correspondencia,
    calcular_similaridade
)
from relatorios import gerar_planilhas_por_curso, salvar_relatório_nao_incluidos, exibir_estatisticas_por_curso

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
    
    # Registrar anomalias
    dados_anomalos = []
    
    # Processar cada linha em lotes para melhorar desempenho com tabelas grandes
    BATCH_SIZE = 500
    
    for i in range(0, len(df), BATCH_SIZE):
        batch = df.iloc[i:i+BATCH_SIZE]
        
        # Processar cada linha do lote
        for idx, row in batch.iterrows():
            # Converter valores para string para evitar problemas de tipo
            nome_aluno = str(row['ALUNOS']) if pd.notna(row['ALUNOS']) else ""
            curso = str(row['CURSO']) if pd.notna(row['CURSO']) else ""
            
            # Verificar dados inconsistentes ou anomalias
            if not nome_aluno or not curso:
                dados_anomalos.append({
                    'indice': idx, 
                    'tipo': 'Dados faltantes',
                    'detalhes': f"Nome: '{nome_aluno}', Curso: '{curso}'",
                })
                continue
            
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
            
            # Resumo da presença para relatórios
            resumo_presenca = ", ".join([f"{col}: {val}" for col, val in zip(colunas_data, valores_presenca)])
            
            # Se não houver dias válidos, o aluno é reprovado com motivo específico
            if total_dias_validos == 0:
                alunos_reprovados.append({
                    'NOME': nome_aluno,
                    'CURSO': curso,
                    'MOTIVO': 'Sem dias válidos de presença',
                    'DETALHES_PRESENCA': resumo_presenca
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
                    'DATA_CONCLUSAO': data_formatada,
                    'DETALHES_PRESENCA': resumo_presenca
                    # Removido: 'PORCENTAGEM_PRESENCA': f"{porcentagem_presenca:.1f}%"
                })
            else:
                alunos_reprovados.append({
                    'NOME': nome_aluno,
                    'CURSO': curso,
                    'MOTIVO': f'Presença insuficiente ({porcentagem_presenca:.1f}%)',
                    'DETALHES_PRESENCA': resumo_presenca
                })
    
    # Se encontrou anomalias, registrar em arquivo separado
    if dados_anomalos:
        df_anomalias = pd.DataFrame(dados_anomalos)
        df_anomalias.to_csv('anomalias_frequencia.csv', index=False, encoding='utf-8')
        print(f"Atenção: {len(dados_anomalos)} anomalias encontradas no arquivo de frequência. Detalhes salvos em 'anomalias_frequencia.csv'")
    
    return alunos_aprovados, alunos_reprovados, total_alunos


def processar_inscricao_pandas(arquivo_inscricao, alunos_aprovados, arquivo_saida, gerar_por_curso=True):
    """
    Processa o arquivo de inscrição usando pandas e gera o arquivo final com todos os dados solicitados.
    Também gera relatórios adicionais e planilhas por curso.
    
    Args:
        arquivo_inscricao: Caminho para o arquivo CSV de inscrição
        alunos_aprovados: Lista de alunos aprovados por frequência
        arquivo_saida: Caminho para o arquivo CSV de saída
        gerar_por_curso: Se True, gera arquivos separados por curso
        
    Returns:
        Lista de dicionários com os dados dos alunos incluídos no arquivo final
    """
    try:
        # Criar registro de log
        log_entries = []
        def log(mensagem):
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            entry = f"[{timestamp}] {mensagem}"
            log_entries.append(entry)
            return entry
        
        log("Iniciando processamento de inscrições")
        
        # Ler o arquivo de inscrição com pandas
        try:
            df_inscricao = pd.read_csv(arquivo_inscricao)
            log(f"Arquivo de inscrição lido com sucesso. {len(df_inscricao)} inscrições encontradas.")
        except Exception as e:
            log(f"Erro ao ler o arquivo de inscrição: {str(e)}")
            # Tentar com codificação diferente
            try:
                df_inscricao = pd.read_csv(arquivo_inscricao, encoding='latin1')
                log(f"Arquivo lido com codificação alternativa. {len(df_inscricao)} inscrições encontradas.")
            except Exception as e2:
                log(f"Segundo erro ao ler arquivo: {str(e2)}")
                raise
        
        # Verificar valores ausentes nas colunas importantes
        colunas_essenciais = ['Nome completo', 'Para qual curso você quer se inscrever?']
        for coluna in colunas_essenciais:
            if coluna in df_inscricao.columns:
                ausentes = df_inscricao[coluna].isna().sum()
                if ausentes > 0:
                    log(f"AVISO: {ausentes} valores ausentes na coluna '{coluna}'")
        
        # Converter DataFrame para lista de dicionários para facilitar o processamento
        inscritos = df_inscricao.to_dict('records')
        log(f"Dados convertidos para processamento. Iniciando correspondência de alunos.")
        
        # Campos que serão incluídos no arquivo final
        campos_finais = [
            'DATA_ADESAO', 'ESTADO', 'ESCOLA', 'NOME', 'CURSO', 'TELEFONE', 
            'EMAIL', 'CPF', 'DIA', 'MES', 'ANO', 'IDADE', 'COR_PELE', 
            'SEXO', 'SERIE_ESCOLAR', 'DATA_CONCLUSAO'
            # Removido: 'PORCENTAGEM_PRESENCA'
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
        
        # Lista para registrar casos de correspondência duvidosa
        correspondencias_duvidosas = []
        
        # Processar em lotes para performance
        BATCH_SIZE = 500
        log(f"Processando alunos aprovados em lotes de {BATCH_SIZE}")
        
        for i in range(0, len(alunos_aprovados), BATCH_SIZE):
            batch = alunos_aprovados[i:i+BATCH_SIZE]
            log(f"Processando lote {i//BATCH_SIZE + 1} ({len(batch)} alunos)")
            
            # Para cada aluno aprovado no lote, encontrar sua inscrição correspondente
            for aluno in batch:
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
                    log(f"Aluno já processado: {nome_aluno} - {curso_frequencia}")
                    continue
                
                # Encontrar a melhor correspondência para este aluno
                inscritos_filtrados = []
                melhor_similaridade = 0
                
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
                        
                        # Calcular similaridade para registro de casos duvidosos
                        similaridade = calcular_similaridade(
                            normalizar_nome(nome_aluno), 
                            normalizar_nome(nome_inscrito)
                        )
                        
                        if similaridade > melhor_similaridade:
                            melhor_similaridade = similaridade
                        
                        if nome_similar and curso_correspondente:
                            inscritos_filtrados.append({
                                'inscricao': inscrito,
                                'similaridade': similaridade
                            })
                            
                            # Registrar casos de similaridade baixa mas aceitável
                            if 0.6 <= similaridade < 0.8:
                                correspondencias_duvidosas.append({
                                    'nome_frequencia': nome_aluno,
                                    'nome_inscricao': nome_inscrito,
                                    'curso': curso_frequencia,
                                    'similaridade': similaridade
                                })
                    except Exception as e:
                        log(f"Erro ao comparar nomes: {nome_aluno} vs {nome_inscrito} - {str(e)}")
                        continue
                
                # Se encontramos correspondências, ordenar por similaridade (melhor primeiro)
                inscricao_aluno = None
                if inscritos_filtrados:
                    # Ordenar por similaridade e depois por data (mais recente primeiro)
                    try:
                        inscritos_filtrados.sort(
                            key=lambda x: (
                                x['similaridade'], 
                                str(x['inscricao'].get('Carimbo de data/hora', ''))
                            ),
                            reverse=True
                        )
                    except Exception as e:
                        log(f"Erro ao ordenar inscrições: {str(e)}")
                    
                    inscricao_aluno = inscritos_filtrados[0]['inscricao']
                    log(f"Correspondência encontrada para {nome_aluno} - {curso_frequencia} (similaridade: {inscritos_filtrados[0]['similaridade']:.2f})")
                else:
                    # Se não encontrou correspondência específica para o curso, buscar a melhor correspondência pelo nome
                    try:
                        idx_melhor = encontrar_melhor_correspondencia(nome_aluno, inscritos)
                        if idx_melhor is not None:
                            inscricao_aluno = inscritos[idx_melhor]
                            log(f"Correspondência parcial para {nome_aluno} - {curso_frequencia} (melhor similaridade: {melhor_similaridade:.2f})")
                    except Exception as e:
                        log(f"Erro ao encontrar melhor correspondência: {str(e)}")
                
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
                    
                    # Adicionar a data de conclusão e curso formatado
                    aluno_final['DATA_CONCLUSAO'] = data_conclusao
                    aluno_final['CURSO'] = curso_certificado
                    
                    # Garantir que o nome seja capitalizado corretamente
                    if aluno_final['NOME']:
                        aluno_final['NOME'] = capitalizar_nome(aluno_final['NOME'])
                    
                    dados_finais.append(aluno_final)
                else:
                    log(f"Não foi encontrada inscrição para: {nome_aluno} - {curso_frequencia}")
        
        # Salvar log de processamento
        with open("processamento_log.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(log_entries))
        log(f"Log de processamento salvo em 'processamento_log.txt'")
        
        # Salvar correspondências duvidosas
        if correspondencias_duvidosas:
            df_duvidosas = pd.DataFrame(correspondencias_duvidosas)
            df_duvidosas.to_csv('correspondencias_duvidosas.csv', index=False, encoding='utf-8')
            log(f"Correspondências duvidosas salvas em 'correspondencias_duvidosas.csv' ({len(correspondencias_duvidosas)} registros)")
        
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
            log(f"Arquivo final gerado com {len(df_final)} alunos")
        except Exception as e:
            error_msg = f"Erro ao escrever arquivo: {str(e)}"
            print(error_msg)
            log(error_msg)
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
        
        # Gerar relatório de alunos não incluídos
        if alunos_nao_incluidos:
            salvar_relatório_nao_incluidos(alunos_nao_incluidos)
            
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
        exibir_estatisticas_por_curso(correspondencias_por_curso)
        
        # Gerar planilhas separadas por curso se solicitado
        if gerar_por_curso and dados_finais:
            gerar_planilhas_por_curso(dados_finais)
        
        print("\n" + "-"*80)
        
        return dados_finais
    except Exception as e:
        print(f"Erro geral em processar_inscricao_pandas: {str(e)}")
        import traceback
        traceback.print_exc()
        return []  # Retornar lista vazia em caso de erro