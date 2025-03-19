"""
Script para testar o sistema modificado de processamento de certificados
"""
import os
import subprocess
import pandas as pd

def executar_codigo():
    """Executa o script main.py"""
    print("\n=== EXECUTANDO O SCRIPT PRINCIPAL ===\n")
    subprocess.run(["python", "main.py"], check=True)

def verificar_resultados():
    """Verifica os resultados gerados"""
    print("\n=== VERIFICANDO RESULTADOS ===\n")
    
    # Verificar arquivo de certificados final
    if os.path.exists("certificados_final.csv"):
        try:
            df = pd.read_csv("certificados_final.csv")
            print(f"Arquivo 'certificados_final.csv' gerado com sucesso.")
            print(f"Total de alunos no arquivo: {len(df)}")
            
            # Mostrar distribuição por curso
            print("\nDistribuição por curso:")
            cursos = df["CURSO"].value_counts()
            for curso, quantidade in cursos.items():
                print(f"- {curso}: {quantidade} alunos")
        except Exception as e:
            print(f"Erro ao ler o arquivo 'certificados_final.csv': {str(e)}")
    else:
        print("Arquivo 'certificados_final.csv' não foi gerado.")
    
    # Verificar arquivo de correspondências dos alunos incluídos (>=80%)
    if os.path.exists("correspondencias_incluidos.csv"):
        try:
            df = pd.read_csv("correspondencias_incluidos.csv")
            print(f"\nArquivo 'correspondencias_incluidos.csv' gerado com sucesso.")
            print(f"Total de correspondências: {len(df)}")
            
            # Analisar similaridades
            similaridades = df["similaridade"].describe()
            print("\nEstatísticas de similaridade (incluídos):")
            print(f"- Mínima: {similaridades['min']:.4f}")
            print(f"- Média: {similaridades['mean']:.4f}")
            print(f"- Máxima: {similaridades['max']:.4f}")
            
            # Verificar se todos têm similaridade >= 0.8
            abaixo_limiar = df[df["similaridade"] < 0.8]
            if len(abaixo_limiar) > 0:
                print(f"\nATENÇÃO: Encontrados {len(abaixo_limiar)} registros com similaridade < 0.8")
                print("Exemplos:")
                for _, row in abaixo_limiar.head(3).iterrows():
                    print(f"  {row['nome_frequencia']} ↔ {row['nome_inscricao']}: {row['similaridade']}")
            else:
                print("\nTodos os registros têm similaridade >= 0.8, conforme esperado.")
        except Exception as e:
            print(f"Erro ao ler o arquivo 'correspondencias_incluidos.csv': {str(e)}")
    else:
        print("\nArquivo 'correspondencias_incluidos.csv' não foi gerado.")
    
    # Verificar arquivo de correspondências limítrofes (entre 70% e 80%)
    if os.path.exists("correspondencias_limitrofes.csv"):
        try:
            df = pd.read_csv("correspondencias_limitrofes.csv")
            print(f"\nArquivo 'correspondencias_limitrofes.csv' gerado com sucesso.")
            print(f"Total de correspondências limítrofes: {len(df)}")
            
            # Analisar similaridades
            if len(df) > 0:
                similaridades = df["similaridade"].describe()
                print("\nEstatísticas de similaridade (limítrofes):")
                print(f"- Mínima: {similaridades['min']:.4f}")
                print(f"- Média: {similaridades['mean']:.4f}")
                print(f"- Máxima: {similaridades['max']:.4f}")
                
                # Verificar se todos estão na faixa esperada (0.7 - 0.8)
                fora_faixa = df[(df["similaridade"] < 0.7) | (df["similaridade"] >= 0.8)]
                if len(fora_faixa) > 0:
                    print(f"\nATENÇÃO: Encontrados {len(fora_faixa)} registros fora da faixa 0.7-0.8")
                else:
                    print("\nTodos os registros limítrofes estão na faixa 0.7-0.8, conforme esperado.")
                
                # Mostrar exemplos de correspondências limítrofes
                print("\nExemplos de correspondências limítrofes:")
                for _, row in df.head(3).iterrows():
                    print(f"  {row['nome_frequencia']} ↔ {row['nome_inscricao']}: {row['similaridade']}")
            else:
                print("Não foram encontrados casos limítrofes (similaridade entre 0.7 e 0.8).")
        except Exception as e:
            print(f"Erro ao ler o arquivo 'correspondencias_limitrofes.csv': {str(e)}")
    else:
        print("\nArquivo 'correspondencias_limitrofes.csv' não foi gerado.")
    
    # Verificar se os arquivos que devem ser removidos não foram gerados
    arquivos_removidos = ["alunos_reprovados.csv", "alunos_nao_incluidos.csv", "alunos_nao_certificados.csv", "todas_correspondencias.csv"]
    for arquivo in arquivos_removidos:
        if os.path.exists(arquivo):
            print(f"\nATENÇÃO: O arquivo '{arquivo}' ainda foi gerado e deveria ter sido removido.")
        else:
            print(f"\nArquivo '{arquivo}' removido com sucesso.")
    
    # Verificar diretório de certificados por curso
    diretorio_curso = "certificados_por_curso"
    if os.path.exists(diretorio_curso) and os.path.isdir(diretorio_curso):
        arquivos = os.listdir(diretorio_curso)
        print(f"\nDiretório '{diretorio_curso}' gerado com {len(arquivos)} arquivos.")
    else:
        print(f"\nDiretório '{diretorio_curso}' não foi gerado.")

if __name__ == "__main__":
    print("Iniciando teste do sistema de processamento de certificados...\n")
    try:
        executar_codigo()
        verificar_resultados()
        print("\n=== TESTE CONCLUÍDO ===")
    except Exception as e:
        print(f"\nERRO NO TESTE: {str(e)}")