import os
import traceback
from datetime import datetime
import pandas as pd

from processadores import processar_frequencia_pandas, processar_inscricao_pandas
from utils import ordenar_por_nome

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
        dados_finais = processar_inscricao_pandas(arquivo_inscricao, alunos_aprovados, arquivo_saida, gerar_por_curso=True)
        
        # Exibir estatísticas finais
        print("\n" + "-"*80)
        print("RESUMO FINAL")
        print("-"*80)
        print(f"Total de alunos na chamada: {total_alunos}")
        print(f"Total de alunos aprovados: {len(alunos_aprovados)}")
        print(f"Total de alunos reprovados: {len(alunos_reprovados)}")
        print(f"Total de certificados gerados: {len(dados_finais)}")
        
        # Identificar e exibir alunos que foram aprovados na frequência mas não foram incluídos no arquivo final
        alunos_sem_inscricao = []
        
        # Criar lista de nomes incluídos no certificado final (normalizado para comparação)
        nomes_incluidos = [aluno['NOME'].lower().strip() for aluno in dados_finais]
        
        # Verificar cada aluno aprovado que não foi para o arquivo final
        for aluno in alunos_aprovados:
            nome = aluno['NOME']
            nome_normalizado = nome.lower().strip()
            curso = aluno['CURSO_ORIGINAL']
            
            if nome_normalizado not in nomes_incluidos:
                alunos_sem_inscricao.append({
                    'NOME': nome,
                    'CURSO': curso
                })
        
        # Mostrar tabela de alunos sem inscrição
        if alunos_sem_inscricao:
            print("\n" + "-"*80)
            print(f"ALUNOS APROVADOS POR FREQUÊNCIA, MAS SEM INSCRIÇÃO CORRESPONDENTE (≥ 80%): {len(alunos_sem_inscricao)}")
            print("-"*80)
            print("| {:<4} | {:<40} | {:<20} |".format("Nº", "Nome do Aluno", "Curso"))
            print("|" + "-"*6 + "|" + "-"*42 + "|" + "-"*22 + "|")
            
            for i, aluno in enumerate(sorted(alunos_sem_inscricao, key=ordenar_por_nome), 1):
                nome = str(aluno['NOME'])
                curso = str(aluno['CURSO'])
                print("| {:<4} | {:<40} | {:<20} |".format(i, nome[:40], curso[:20]))
        
        # Verificar se existem alunos no arquivo de correspondências limítrofes
        if os.path.exists("correspondencias_limitrofes.csv"):
            try:
                df_limitrofes = pd.read_csv("correspondencias_limitrofes.csv")
                if not df_limitrofes.empty:
                    print("\n" + "-"*80)
                    print(f"ALUNOS COM CORRESPONDÊNCIA ENTRE 70% E 80%: {len(df_limitrofes)}")
                    print("-"*80)
                    print("| {:<4} | {:<25} | {:<25} | {:<10} |".format("Nº", "Nome na Frequência", "Nome na Inscrição", "Similaridade"))
                    print("|" + "-"*6 + "|" + "-"*27 + "|" + "-"*27 + "|" + "-"*12 + "|")
                    
                    for i, row in df_limitrofes.iterrows():
                        nome_freq = str(row['nome_frequencia'])
                        nome_insc = str(row['nome_inscricao'])
                        similaridade = float(row['similaridade'])
                        print("| {:<4} | {:<25} | {:<25} | {:<10.2%} |".format(
                            i+1, nome_freq[:25], nome_insc[:25], similaridade
                        ))
            except Exception as e:
                print(f"Erro ao ler o arquivo 'correspondencias_limitrofes.csv': {str(e)}")
        
        print("\n" + "-"*80)
        print("NOTA: Alunos com similaridade menor que 80% foram excluídos do arquivo final")
        print("Correspondências dos alunos incluídos foram salvas em 'correspondencias_incluidos.csv'")
        print("Correspondências entre 70% e 80% foram salvas em 'correspondencias_limitrofes.csv' para análise manual")
        print("-"*80)
        
    except Exception as e:
        print(f"\nErro ao processar os arquivos: {str(e)}")
        traceback.print_exc()
        
        # Registrar o erro em um arquivo de log
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        with open(f"erro_{timestamp}.log", "w", encoding="utf-8") as f:
            f.write(f"Erro ao processar os arquivos: {str(e)}\n\n")
            traceback.print_exc(file=f)
        print(f"Detalhes do erro foram salvos em 'erro_{timestamp}.log'")


if __name__ == "__main__":
    main()