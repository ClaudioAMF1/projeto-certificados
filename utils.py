import re
from difflib import SequenceMatcher
import pandas as pd

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
        nome_inscrito = inscrito.get('Nome completo', '')
        if not isinstance(nome_inscrito, str):
            try:
                nome_inscrito = str(nome_inscrito)
            except:
                nome_inscrito = ""
                
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