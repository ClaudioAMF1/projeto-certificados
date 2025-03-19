"""
Script para testar a nova função de correspondência de nomes especificamente para os casos problemáticos
"""
from utils import normalizar_nome, calcular_similaridade, verificar_subconjunto_nomes, nomes_similares

def testar_correspondencia(nome_frequencia, nome_inscricao):
    """
    Testa a correspondência entre dois nomes usando diferentes métodos e exibe os resultados
    """
    nome_freq_norm = normalizar_nome(nome_frequencia)
    nome_insc_norm = normalizar_nome(nome_inscricao)
    
    print(f"\nTeste de correspondência para os nomes:")
    print(f"Nome na frequência: '{nome_frequencia}'")
    print(f"Nome na inscrição:  '{nome_inscricao}'")
    print("-" * 60)
    
    # Similaridade direta
    similaridade_direta = calcular_similaridade(nome_freq_norm, nome_insc_norm)
    print(f"Similaridade direta: {similaridade_direta:.2%}")
    
    # Verificação de subconjunto
    e_subconjunto, similaridade_subconjunto = verificar_subconjunto_nomes(nome_freq_norm, nome_insc_norm)
    subconjunto_status = "SIM" if e_subconjunto else "NÃO"
    print(f"É subconjunto: {subconjunto_status}, Similaridade subconjunto: {similaridade_subconjunto:.2%}")
    
    # Melhor similaridade
    melhor_similaridade = max(similaridade_direta, similaridade_subconjunto if e_subconjunto else 0)
    print(f"Melhor similaridade: {melhor_similaridade:.2%}")
    
    # Resultado final da função nomes_similares
    resultado = nomes_similares(nome_frequencia, nome_inscricao)
    resultado_texto = "CORRESPONDÊNCIA ENCONTRADA" if resultado else "CORRESPONDÊNCIA NÃO ENCONTRADA"
    print(f"Resultado final: {resultado_texto}")
    
    return melhor_similaridade

if __name__ == "__main__":
    print("=" * 60)
    print("TESTE DE CORRESPONDÊNCIA DE NOMES")
    print("=" * 60)
    
    # Casos específicos mencionados pelo usuário
    casos_teste = [
        ("Brenda Raiane", "Brenda Raiane Agradem da Silva"),
        ("Rafael Lopes", "Rafael Lopes Farias"),
        ("Lavinia Vasconcelos", "Lavinia Eduarda Vargas Vasconcelos"),
        # Outros casos para validação
        ("Maria Silva", "Maria da Silva Santos"),
        ("João Pedro", "João Pedro Almeida"),
        ("Ana Carolina", "Ana Carolina Ferreira Mendes"),
        # Casos que não devem corresponder
        ("Carlos Eduardo", "Eduardo Carlos Ferreira"),
        ("Pedro Silva", "Paulo Silveira")
    ]
    
    resultados = []
    
    # Testar cada caso
    for nome_freq, nome_insc in casos_teste:
        similaridade = testar_correspondencia(nome_freq, nome_insc)
        resultados.append((nome_freq, nome_insc, similaridade))
    
    # Resumo dos resultados
    print("\n" + "=" * 60)
    print("RESUMO DOS RESULTADOS")
    print("=" * 60)
    print(f"{'Nome Frequência':<20} | {'Nome Inscrição':<30} | {'Similaridade':<15}")
    print("-" * 70)
    
    for nome_freq, nome_insc, similaridade in resultados:
        status = "✓" if similaridade >= 0.8 else "✗"
        print(f"{nome_freq:<20} | {nome_insc:<30} | {similaridade:.2%} {status}")