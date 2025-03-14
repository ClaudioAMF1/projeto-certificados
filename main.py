import os
import traceback
from datetime import datetime

from processadores import processar_frequencia_pandas, processar_inscricao_pandas
from relatorios import salvar_relatório_reprovados, salvar_relatório_nao_incluidos, salvar_relatorio_todos_nao_certificados
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
            
            # Salvar relatório de reprovados e armazenar o resultado
            dados_reprovados = salvar_relatório_reprovados(alunos_reprovados)
        else:
            dados_reprovados = []
        
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
        
        # Verificar quais alunos aprovados não foram incluídos no arquivo final
        alunos_nao_incluidos = []
        for aluno in alunos_aprovados:
            nome_aluno = aluno['NOME_ORIGINAL']
            curso = aluno['CURSO_ORIGINAL']
            
            # Garantir que nome_aluno seja uma string
            if not isinstance(nome_aluno, str):
                nome_aluno = str(nome_aluno)
                
            # Verificamos se o aluno foi incluído nos certificados finais
            incluido = False
            for cert in dados_finais:
                if cert['NOME'] == aluno['NOME']:
                    incluido = True
                    break
                    
            if not incluido:
                alunos_nao_incluidos.append(f"{nome_aluno} - {curso}")
        
        # Gerar relatório de alunos não incluídos
        if alunos_nao_incluidos:
            # Salvar relatório de alunos não incluídos e armazenar o resultado
            dados_nao_incluidos = salvar_relatório_nao_incluidos(alunos_nao_incluidos)
            
            print(f"\nAlunos aprovados mas não incluídos (por falta de inscrição ou inscrição para curso incorreto): {len(alunos_nao_incluidos)}")
            print("-" * 80)
            print("| {:<4} | {:<40} | {:<20} |".format("Nº", "Nome do Aluno", "Curso"))
            print("|" + "-"*6 + "|" + "-"*42 + "|" + "-"*22 + "|")
            
            for i, aluno_info in enumerate(alunos_nao_incluidos, 1):
                partes = aluno_info.split(" - ", 1)
                nome = partes[0] if len(partes) > 0 else ""
                curso = partes[1] if len(partes) > 1 else ""
                print("| {:<4} | {:<40} | {:<20} |".format(i, nome[:40], curso[:20]))
        else:
            dados_nao_incluidos = []
        
        # Criar relatório consolidado de todos os alunos que não receberam certificados
        salvar_relatorio_todos_nao_certificados(dados_reprovados, dados_nao_incluidos)
        
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