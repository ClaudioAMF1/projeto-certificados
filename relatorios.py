# Arquivo: relatorios.py
# Modificar a função para incluir todos os alunos não certificados

import os
import pandas as pd

def gerar_planilhas_por_curso(dados_finais, diretorio_saida="certificados_por_curso"):
    """
    Gera planilhas separadas para cada curso
    
    Args:
        dados_finais: Lista de dicionários com dados dos alunos aprovados
        diretorio_saida: Diretório onde serão salvos os arquivos
    """
    # Criar diretório se não existir
    if not os.path.exists(diretorio_saida):
        os.makedirs(diretorio_saida)
    
    # Agrupar alunos por curso
    alunos_por_curso = {}
    for aluno in dados_finais:
        curso = aluno['CURSO']
        if curso not in alunos_por_curso:
            alunos_por_curso[curso] = []
        alunos_por_curso[curso].append(aluno)
    
    # Gerar uma planilha para cada curso
    for curso, alunos in alunos_por_curso.items():
        # Criar nome de arquivo seguro (remover caracteres especiais)
        nome_arquivo = ''.join(c if c.isalnum() else '_' for c in curso)
        caminho_arquivo = os.path.join(diretorio_saida, f"{nome_arquivo}.csv")
        
        # Criar DataFrame e salvar
        df_curso = pd.DataFrame(alunos)
        df_curso.to_csv(caminho_arquivo, index=False, encoding='utf-8')
        print(f"Arquivo para curso '{curso}' gerado: {caminho_arquivo} ({len(alunos)} alunos)")


def salvar_relatório_reprovados(alunos_reprovados, arquivo_saida="alunos_reprovados.csv"):
    """
    Salva um relatório detalhado dos alunos reprovados
    
    Args:
        alunos_reprovados: Lista de dicionários com dados dos alunos reprovados
        arquivo_saida: Caminho para o arquivo CSV de saída
    """
    # Preparar dados para o relatório
    dados_relatorio = []
    for aluno in alunos_reprovados:
        nome = str(aluno['NOME'])
        curso = str(aluno['CURSO'])
        motivo = "Presença insuficiente" if 'MOTIVO' not in aluno else aluno['MOTIVO']
        
        # Coletar detalhes de presença se disponíveis
        presenca_info = ""
        if 'DETALHES_PRESENCA' in aluno:
            presenca_info = aluno['DETALHES_PRESENCA']
        
        dados_relatorio.append({
            'NOME': nome,
            'CURSO': curso,
            'MOTIVO_REPROVACAO': motivo,
            'DETALHES': presenca_info
        })
    
    # Salvar relatório
    df_reprovados = pd.DataFrame(dados_relatorio)
    df_reprovados.to_csv(arquivo_saida, index=False, encoding='utf-8')
    print(f"Relatório de reprovados gerado: {arquivo_saida} ({len(dados_relatorio)} alunos)")
    
    # Retornar os dados para uso posterior
    return dados_relatorio


def salvar_relatório_nao_incluidos(alunos_nao_incluidos, arquivo_saida="alunos_nao_incluidos.csv"):
    """
    Salva um relatório dos alunos aprovados por frequência mas não incluídos no arquivo final
    
    Args:
        alunos_nao_incluidos: Lista de strings com informações dos alunos não incluídos
        arquivo_saida: Caminho para o arquivo CSV de saída
    """
    # Preparar dados para o relatório
    dados_relatorio = []
    for aluno_info in alunos_nao_incluidos:
        partes = aluno_info.split(" - ", 1)
        nome = partes[0] if len(partes) > 0 else ""
        curso = partes[1] if len(partes) > 1 else ""
        
        dados_relatorio.append({
            'NOME': nome,
            'CURSO': curso,
            'MOTIVO': "Inscrição não encontrada ou curso não correspondente"
        })
    
    # Salvar relatório
    df_nao_incluidos = pd.DataFrame(dados_relatorio)
    df_nao_incluidos.to_csv(arquivo_saida, index=False, encoding='utf-8')
    print(f"Relatório de alunos não incluídos gerado: {arquivo_saida} ({len(dados_relatorio)} alunos)")
    
    # Retornar os dados para uso posterior
    return dados_relatorio


def salvar_relatorio_todos_nao_certificados(alunos_reprovados, alunos_nao_incluidos, arquivo_saida="alunos_nao_certificados.csv"):
    """
    Salva um relatório consolidado de todos os alunos que não receberam certificados, 
    independente do motivo (reprovados por frequência ou sem inscrição correspondente)
    
    Args:
        alunos_reprovados: Dados dos alunos reprovados por frequência
        alunos_nao_incluidos: Dados dos alunos sem inscrição correspondente
        arquivo_saida: Caminho para o arquivo CSV de saída
    """
    # Preparar dados para o relatório consolidado
    dados_consolidados = []
    
    # Adicionar alunos reprovados por frequência
    for aluno in alunos_reprovados:
        dados_consolidados.append({
            'NOME': aluno['NOME'],
            'CURSO': aluno['CURSO'],
            'MOTIVO': aluno['MOTIVO_REPROVACAO'],
            'DETALHES': aluno.get('DETALHES', '')
        })
    
    # Adicionar alunos sem inscrição correspondente
    for aluno in alunos_nao_incluidos:
        dados_consolidados.append({
            'NOME': aluno['NOME'],
            'CURSO': aluno['CURSO'],
            'MOTIVO': aluno['MOTIVO'],
            'DETALHES': ''
        })
    
    # Salvar relatório consolidado
    df_consolidado = pd.DataFrame(dados_consolidados)
    df_consolidado.to_csv(arquivo_saida, index=False, encoding='utf-8')
    print(f"\nRELATÓRIO CONSOLIDADO DE ALUNOS NÃO CERTIFICADOS")
    print("-" * 80)
    print(f"Total de alunos sem certificado: {len(dados_consolidados)}")
    print(f"Relatório consolidado gerado: {arquivo_saida}")
    
    # Listar todos os alunos
    print("\nLista de todos os alunos sem certificado:")
    print("| {:<4} | {:<40} | {:<20} | {:<30} |".format("Nº", "Nome do Aluno", "Curso", "Motivo"))
    print("|" + "-"*6 + "|" + "-"*42 + "|" + "-"*22 + "|" + "-"*32 + "|")
    
    for i, aluno in enumerate(sorted(dados_consolidados, key=lambda x: x['NOME']), 1):
        nome = aluno['NOME'][:40]
        curso = aluno['CURSO'][:20]
        motivo = aluno['MOTIVO'][:30]
        print("| {:<4} | {:<40} | {:<20} | {:<30} |".format(i, nome, curso, motivo))
    
    return dados_consolidados


def exibir_estatisticas_por_curso(correspondencias_por_curso):
    """
    Exibe estatísticas detalhadas por curso
    
    Args:
        correspondencias_por_curso: Dicionário com estatísticas por curso
    """
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