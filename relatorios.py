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