"""
Módulo responsável por escanear pastas e identificar aquelas com validade
vencida há 3 meses ou mais.

IMPORTANTE: este módulo é SOMENTE LEITURA. Em nenhum momento cria, altera,
move ou exclui arquivos/pastas — apenas lê nomes e navega na árvore de
diretórios.

Versão com leitura em paralelo: como o gargalo em uma pasta de rede é a
LATÊNCIA de cada chamada (não o processamento), várias pastas são lidas ao
mesmo tempo em vez de uma de cada vez. Isso costuma reduzir bastante o tempo
total em compartilhamentos grandes.
"""

import os
import re
import calendar
import threading
from datetime import date
from concurrent.futures import ThreadPoolExecutor

# Regex que localiza datas no formato MM.AA ou MM.AAAA dentro do nome da pasta.
# (?<!\d) e (?!\d) garantem que não pegamos pedaços de números maiores.
# Exemplos que casam: 08.26 | 02.09 | 11.2026
PADRAO_DATA = re.compile(r'(?<!\d)(\d{2})\.(\d{4}|\d{2})(?!\d)')

# Quantidade de meses de tolerância antes de considerar a pasta "vencida".
MESES_LIMITE_VENCIMENTO = 3

# Quantas pastas são lidas ao mesmo tempo. Como o gargalo é a rede (e não a
# CPU), esse número pode ser bem maior que a quantidade de núcleos da
# máquina. Ajuste para cima/baixo conforme o comportamento do servidor de
# arquivos (números muito altos podem sobrecarregar um servidor de rede
# lento ou levar a mais erros de "muitas conexões simultâneas").
MAX_LEITURAS_PARALELAS = 16


def pasta_e_candidata(nome_pasta: str) -> bool:
    """True se o nome da pasta começa com 'val' (ignorando maiúsculas/minúsculas)."""
    return nome_pasta.strip().lower().startswith('val')


def extrair_validade(nome_pasta: str):
    """
    Encontra a ÚLTIMA ocorrência de MM.AA ou MM.AAAA no nome da pasta e
    retorna a data correspondente ao último dia daquele mês/ano.

    Retorna None se nenhuma data válida for encontrada.
    """
    ocorrencias = PADRAO_DATA.findall(nome_pasta)
    if not ocorrencias:
        return None

    mes_str, ano_str = ocorrencias[-1]  # última ocorrência encontrada

    mes = int(mes_str)
    ano = int(ano_str)

    # Ano com 2 dígitos -> assume-se 20XX (ex: 26 -> 2026)
    if ano < 100:
        ano += 2000

    if mes < 1 or mes > 12:
        return None

    ultimo_dia_do_mes = calendar.monthrange(ano, mes)[1]
    return date(ano, mes, ultimo_dia_do_mes)


def meses_entre(data_referencia: date, data_validade: date) -> int:
    """
    Quantidade de meses completos já passados desde o fim do mês de validade.
    Considera corretamente a virada de ano (ex: validade 11/2025 e hoje
    fevereiro/2026 = 3 meses).
    """
    return (
        (data_referencia.year - data_validade.year) * 12
        + (data_referencia.month - data_validade.month)
    )


def escanear_pastas(caminho_base: str, data_referencia: date = None):
    """
    Percorre recursivamente 'caminho_base' e retorna a tupla (dados, erro).

    - 'dados' é um dicionário: {'resultados': [...], 'avisos': [...]}
    - 'erro' é uma string de erro, ou None se tudo correu bem.

    Nenhum arquivo ou pasta é alterado, movido ou excluído — apenas leitura.
    """
    if data_referencia is None:
        data_referencia = date.today()

    try:
        if not os.path.exists(caminho_base):
            return None, f'Caminho não encontrado ou inacessível: {caminho_base}'

        if not os.path.isdir(caminho_base):
            return None, f'O caminho informado não é uma pasta: {caminho_base}'
    except OSError as erro:
        return None, f'Não foi possível acessar o caminho informado: {erro}'

    resultados = []
    avisos = []
    lock = threading.Lock()  # protege 'resultados' e 'avisos' entre as threads

    def listar_subpastas(caminho):
        """Lê apenas os NOMES das subpastas diretas de 'caminho' (somente leitura)."""
        try:
            with os.scandir(caminho) as it:
                return [
                    entrada.name for entrada in it
                    if entrada.is_dir(follow_symlinks=False)
                ]
        except OSError as erro:
            with lock:
                avisos.append(f'Sem permissão/acesso para ler: {caminho}')
            return []

    def processar_pasta(caminho):
        """
        Lê as subpastas diretas de 'caminho'. Para cada uma:
          - se for candidata ("Val...") e estiver vencida, registra no resultado;
          - senão, devolve o caminho para ser lido na próxima rodada.
        Não desce dentro de pastas "Val..." (evita varredura desnecessária).
        """
        proximas = []
        for nome in listar_subpastas(caminho):
            if pasta_e_candidata(nome):
                data_validade = extrair_validade(nome)
                if data_validade is not None:
                    meses_vencida = meses_entre(data_referencia, data_validade)
                    if meses_vencida >= MESES_LIMITE_VENCIMENTO:
                        with lock:
                            resultados.append({
                                'nome': nome,
                                'caminho': os.path.join(caminho, nome),
                                'validade': data_validade.strftime('%m/%Y'),
                                'meses_vencida': meses_vencida,
                            })
                # já classificada como "Val...": não entra dentro dela.
            else:
                proximas.append(os.path.join(caminho, nome))
        return proximas

    # Varredura "por nível" (largura), com cada nível lido em paralelo:
    # primeiro lê a pasta raiz, depois lê TODAS as subpastas encontradas ao
    # mesmo tempo (até MAX_LEITURAS_PARALELAS por vez), depois o próximo
    # nível, e assim por diante — em vez de esperar uma pasta terminar para
    # só então começar a próxima.
    try:
        with ThreadPoolExecutor(max_workers=MAX_LEITURAS_PARALELAS) as executor:
            nivel_atual = [caminho_base]
            while nivel_atual:
                resultados_do_nivel = executor.map(processar_pasta, nivel_atual)
                proximo_nivel = []
                for subpastas_para_ler in resultados_do_nivel:
                    proximo_nivel.extend(subpastas_para_ler)
                nivel_atual = proximo_nivel
    except OSError as erro:
        return None, f'Erro durante a leitura das pastas: {erro}'

    # Mostra primeiro as mais vencidas
    resultados.sort(key=lambda item: item['meses_vencida'], reverse=True)

    return {'resultados': resultados, 'avisos': avisos}, None
