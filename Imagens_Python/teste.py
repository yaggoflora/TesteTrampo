# -*- coding: utf-8 -*-
"""
=================================================================================================
 SCRIPT: comparar_imagens_excel.py  (porte do PowerShell v3 para Python 3.12, 100% stdlib)
 OBJETIVO:
     Comparar uma imagem de referencia (arquivo local) com as imagens flutuantes presentes na
     coluna alvo de uma aba especifica de um arquivo Excel (.xlsx), classificando cada uma como:
         - "Identica"          -> bytes exatamente iguais (hash SHA256 identico)
         - "Muito semelhante"  -> visualmente quase igual, mas nao byte-a-byte igual
         - "Semelhante"        -> visualmente parecida, acima do limiar minimo
         - (nada)               -> abaixo do limiar minimo -> NAO aparece no resultado

 -------------------------------------------------------------------------------------------
 POR QUE FUNCIONA SEM INSTALAR NENHUMA BIBLIOTECA (nem Pillow, nem numpy):
 -------------------------------------------------------------------------------------------
 A parte de ler o .xlsx como ZIP e achar as imagens flutuantes usa apenas:
     zipfile, re, hashlib -> 100% biblioteca padrao do Python.

 A parte de DECODIFICAR PNG/JPEG e redimensionar para comparar pixels normalmente exigiria
 uma lib externa (Pillow). Como isso nao pode ser instalado no PC corporativo, este script
 usa "ctypes" para chamar diretamente o "gdiplus.dll" -- a mesma biblioteca nativa do Windows
 que o .NET "System.Drawing" usa por baixo dos panos no seu script PowerShell original.
 O gdiplus.dll ja vem em QUALQUER instalacao do Windows (nao e uma lib Python, e um DLL do
 sistema operacional), entao chama-lo via ctypes nao conta como "instalar biblioteca".

 Resultado: mesma logica, mesmos limiares, mesmo metodo de comparacao (grade reduzida NxN +
 diferenca media de cor) do script PowerShell, rodando em Python puro.

 REQUISITOS:
     - Windows (o gdiplus.dll so existe no Windows).
     - Python 3.12.4 (ou qualquer 3.x) -- sem pip install de nada.
     - Excel NAO precisa estar instalado.

 COMO USAR:
     1) Ajuste as variaveis na secao "CONFIGURACOES DO USUARIO" logo abaixo.
     2) Rode: python comparar_imagens_excel.py
=================================================================================================
"""

import ctypes
import hashlib
import os
import re
import sys
import tempfile
import zipfile

if sys.platform != "win32":
    sys.exit("Este script depende do gdiplus.dll e so roda no Windows.")

# =================================================================================================
# CONFIGURACOES DO USUARIO -- AJUSTE SOMENTE ESTA SECAO
# =================================================================================================

CAMINHO_ARQUIVO_EXCEL = r"C:\Caminho\Para\Arquivo.xlsx"
NOME_ABA = "Planilha1"
CAMINHO_IMAGEM_REFERENCIA = r"C:\Caminho\Para\ImagemReferencia.png"
COLUNA_ALVO = 2  # Coluna B = 2, Coluna A = 1, Coluna C = 3, etc.

LIMIAR_MUITO_SEMELHANTE = 95   # percentual minimo (%) para ser "Muito semelhante"
LIMIAR_SEMELHANTE = 80         # percentual minimo (%) para ser "Semelhante"
RESOLUCAO_GRADE_COMPARACAO = 16  # grade NxN usada na comparacao visual (mesmo valor do PS1)

# =================================================================================================
# GDI+ via ctypes (nenhuma instalacao necessaria -- gdiplus.dll ja vem no Windows)
# =================================================================================================

gdiplus = ctypes.WinDLL("gdiplus.dll")

PIXEL_FORMAT_32BPP_ARGB = 2498570
IMAGE_LOCK_MODE_READ = 1
INTERPOLATION_HIGH_QUALITY_BICUBIC = 7


class GdiplusStartupInput(ctypes.Structure):
    _fields_ = [
        ("GdiplusVersion", ctypes.c_uint32),
        ("DebugEventCallback", ctypes.c_void_p),
        ("SuppressBackgroundThread", ctypes.c_int),
        ("SuppressExternalCodecs", ctypes.c_int),
    ]


class GpRect(ctypes.Structure):
    _fields_ = [("X", ctypes.c_int), ("Y", ctypes.c_int),
                ("Width", ctypes.c_int), ("Height", ctypes.c_int)]


class BitmapData(ctypes.Structure):
    _fields_ = [
        ("Width", ctypes.c_uint),
        ("Height", ctypes.c_uint),
        ("Stride", ctypes.c_int),
        ("PixelFormat", ctypes.c_int),
        ("Scan0", ctypes.c_void_p),
        ("Reserved", ctypes.c_uint),
    ]


_token = ctypes.c_ulong()
_startup_input = GdiplusStartupInput(1, None, 0, 0)
if gdiplus.GdiplusStartup(ctypes.byref(_token), ctypes.byref(_startup_input), None) != 0:
    raise RuntimeError("Falha ao iniciar o GDI+ (GdiplusStartup).")


def _checar(status: int, mensagem: str) -> None:
    if status != 0:
        raise RuntimeError(f"{mensagem} (codigo de erro GDI+: {status})")


def carregar_imagem_de_bytes(dados: bytes):
    """Grava os bytes num arquivo temporario e carrega via GDI+ (suporta PNG, JPEG, BMP, GIF)."""
    sufixo = ".png" if dados[:8] == b"\x89PNG\r\n\x1a\n" else ".jpg"
    tmp_path = None
    try:
        fd, tmp_path = tempfile.mkstemp(suffix=sufixo)
        with os.fdopen(fd, "wb") as f:
            f.write(dados)
        ptr_imagem = ctypes.c_void_p()
        status = gdiplus.GdipLoadImageFromFile(tmp_path, ctypes.byref(ptr_imagem))
        _checar(status, "Falha ao carregar imagem via GDI+ (formato invalido/corrompido?)")
        return ptr_imagem
    finally:
        if tmp_path:
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def liberar_imagem(ptr_imagem) -> None:
    gdiplus.GdipDisposeImage(ptr_imagem)


def reduzir_para_grade(ptr_imagem_origem, tamanho: int):
    """Redesenha a imagem original numa miniatura NxN (32bppArgb) e devolve os bytes + stride."""
    ptr_bitmap = ctypes.c_void_p()
    status = gdiplus.GdipCreateBitmapFromScan0(
        tamanho, tamanho, 0, PIXEL_FORMAT_32BPP_ARGB, None, ctypes.byref(ptr_bitmap)
    )
    _checar(status, "Falha ao criar bitmap de destino")

    ptr_graphics = ctypes.c_void_p()
    status = gdiplus.GdipGetImageGraphicsContext(ptr_bitmap, ctypes.byref(ptr_graphics))
    _checar(status, "Falha ao obter contexto grafico")

    gdiplus.GdipSetInterpolationMode(ptr_graphics, INTERPOLATION_HIGH_QUALITY_BICUBIC)

    status = gdiplus.GdipDrawImageRectI(ptr_graphics, ptr_imagem_origem, 0, 0, tamanho, tamanho)
    _checar(status, "Falha ao desenhar imagem reduzida")
    gdiplus.GdipDeleteGraphics(ptr_graphics)

    rect = GpRect(0, 0, tamanho, tamanho)
    dados_bitmap = BitmapData()
    status = gdiplus.GdipBitmapLockBits(
        ptr_bitmap, ctypes.byref(rect), IMAGE_LOCK_MODE_READ,
        PIXEL_FORMAT_32BPP_ARGB, ctypes.byref(dados_bitmap)
    )
    _checar(status, "Falha ao bloquear bits do bitmap")

    tamanho_buffer = dados_bitmap.Stride * tamanho
    buffer = ctypes.string_at(dados_bitmap.Scan0, tamanho_buffer)
    stride = dados_bitmap.Stride

    gdiplus.GdipBitmapUnlockBits(ptr_bitmap, ctypes.byref(dados_bitmap))
    gdiplus.GdipDisposeImage(ptr_bitmap)

    return buffer, stride


def calcular_percentual_similaridade(ptr_img1, ptr_img2, tamanho: int = 16) -> float:
    buffer1, stride1 = reduzir_para_grade(ptr_img1, tamanho)
    buffer2, stride2 = reduzir_para_grade(ptr_img2, tamanho)

    soma_diferencas = 0.0
    for y in range(tamanho):
        for x in range(tamanho):
            i1 = y * stride1 + x * 4  # formato BGRA
            i2 = y * stride2 + x * 4
            b1, g1, r1 = buffer1[i1], buffer1[i1 + 1], buffer1[i1 + 2]
            b2, g2, r2 = buffer2[i2], buffer2[i2 + 1], buffer2[i2 + 2]
            soma_diferencas += (abs(r1 - r2) + abs(g1 - g2) + abs(b1 - b2)) / 3.0

    total_pontos = tamanho * tamanho
    diferenca_media = soma_diferencas / total_pontos  # 0 (identico) a 255 (totalmente oposto)
    percentual = 100 - (diferenca_media / 255 * 100)
    return round(percentual, 2)


# =================================================================================================
# LEITURA DO .xlsx COMO ZIP (workbook.xml -> sheet -> drawing.xml -> media)
# =================================================================================================

def numero_para_letra_coluna(numero: int) -> str:
    letras = ""
    n = numero
    while n > 0:
        n, resto = divmod(n - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def resolver_caminho_zip(diretorio_base: str, caminho_relativo: str) -> str:
    if caminho_relativo.startswith("/"):
        return caminho_relativo.lstrip("/")
    pilha = [p for p in diretorio_base.split("/") if p]
    for parte in caminho_relativo.split("/"):
        if parte == "..":
            if pilha:
                pilha.pop()
        elif parte in ("", "."):
            continue
        else:
            pilha.append(parte)
    return "/".join(pilha)


def mapa_relationships(zf: zipfile.ZipFile, caminho_rels: str) -> dict:
    mapa = {}
    try:
        texto = zf.read(caminho_rels).decode("utf-8")
    except KeyError:
        return mapa
    for tag in re.findall(r"<Relationship\b[^>]*/>", texto):
        m_id = re.search(r'Id="([^"]*)"', tag)
        m_target = re.search(r'Target="([^"]*)"', tag)
        if m_id and m_target:
            mapa[m_id.group(1)] = m_target.group(1)
    return mapa


def localizar_imagens_da_coluna(zf: zipfile.ZipFile, nome_aba: str, coluna_alvo: int):
    workbook_xml = zf.read("xl/workbook.xml").decode("utf-8")

    rid_aba = None
    for tag in re.findall(r"<sheet\b[^>]*/>", workbook_xml):
        m_nome = re.search(r'name="([^"]*)"', tag)
        m_rid = re.search(r'r:id="([^"]*)"', tag)
        if m_nome and m_nome.group(1) == nome_aba and m_rid:
            rid_aba = m_rid.group(1)
            break
    if not rid_aba:
        raise RuntimeError(f"A aba '{nome_aba}' nao foi encontrada em xl/workbook.xml.")

    mapa_wb_rels = mapa_relationships(zf, "xl/_rels/workbook.xml.rels")
    if rid_aba not in mapa_wb_rels:
        raise RuntimeError(f"Nao foi possivel resolver o caminho interno da aba (rId {rid_aba}).")
    caminho_sheet = resolver_caminho_zip("xl", mapa_wb_rels[rid_aba])

    sheet_xml = zf.read(caminho_sheet).decode("utf-8")
    m_drawing = re.search(r'<drawing\b[^>]*r:id="(rId\d+)"', sheet_xml)
    if not m_drawing:
        print(f"      A aba '{nome_aba}' nao possui nenhum desenho associado.")
        return []

    idx = caminho_sheet.rfind("/")
    dir_sheet = caminho_sheet[:idx]
    nome_sheet = caminho_sheet[idx + 1:]
    caminho_sheet_rels = f"{dir_sheet}/_rels/{nome_sheet}.rels"

    mapa_sheet_rels = mapa_relationships(zf, caminho_sheet_rels)
    rid_drawing = m_drawing.group(1)
    if rid_drawing not in mapa_sheet_rels:
        return []
    caminho_drawing = resolver_caminho_zip(dir_sheet, mapa_sheet_rels[rid_drawing])

    idx_d = caminho_drawing.rfind("/")
    dir_drawing = caminho_drawing[:idx_d]
    nome_drawing = caminho_drawing[idx_d + 1:]
    caminho_drawing_rels = f"{dir_drawing}/_rels/{nome_drawing}.rels"
    mapa_drawing_rels = mapa_relationships(zf, caminho_drawing_rels)

    drawing_xml = zf.read(caminho_drawing).decode("utf-8")

    blocos = re.findall(r"<xdr:twoCellAnchor.*?</xdr:twoCellAnchor>", drawing_xml, re.S)
    blocos += re.findall(r"<xdr:oneCellAnchor.*?</xdr:oneCellAnchor>", drawing_xml, re.S)

    candidatos = []
    for bloco in blocos:
        if "<xdr:pic>" not in bloco:
            continue
        m_from = re.search(r"<xdr:from>(.*?)</xdr:from>", bloco, re.S)
        if not m_from:
            continue
        m_col = re.search(r"<xdr:col>(\d+)</xdr:col>", m_from.group(1))
        m_row = re.search(r"<xdr:row>(\d+)</xdr:row>", m_from.group(1))
        if not (m_col and m_row):
            continue

        coluna_real = int(m_col.group(1)) + 1  # xdr:col e zero-based
        linha_real = int(m_row.group(1)) + 1
        if coluna_real != coluna_alvo:
            continue

        m_embed = re.search(r'r:embed="(rId\d+)"', bloco)
        if not m_embed:
            continue
        rid_imagem = m_embed.group(1)
        if rid_imagem not in mapa_drawing_rels:
            continue

        caminho_imagem = resolver_caminho_zip(dir_drawing, mapa_drawing_rels[rid_imagem])
        celula = f"{numero_para_letra_coluna(coluna_real)}{linha_real}"
        candidatos.append((celula, caminho_imagem))

    return candidatos


# =================================================================================================
# MAIN
# =================================================================================================

def main() -> None:
    print("==================================================================")
    print(" Iniciando comparacao de imagens - Excel x Imagem de Referencia")
    print("==================================================================")

    if not os.path.isfile(CAMINHO_ARQUIVO_EXCEL):
        sys.exit(f"Arquivo Excel nao encontrado em: {CAMINHO_ARQUIVO_EXCEL}")
    if not os.path.isfile(CAMINHO_IMAGEM_REFERENCIA):
        sys.exit(f"Imagem de referencia nao encontrada em: {CAMINHO_IMAGEM_REFERENCIA}")

    print("\n[1/4] Carregando imagem de referencia...")
    with open(CAMINHO_IMAGEM_REFERENCIA, "rb") as f:
        bytes_referencia = f.read()
    hash_referencia = hashlib.sha256(bytes_referencia).hexdigest().upper()
    ptr_referencia = carregar_imagem_de_bytes(bytes_referencia)
    print(f"      Hash SHA256 : {hash_referencia}")

    resultados = []

    print(f"\n[2/4] Localizando a aba '{NOME_ABA}' dentro do arquivo...")
    with zipfile.ZipFile(CAMINHO_ARQUIVO_EXCEL) as zf:
        candidatos = localizar_imagens_da_coluna(zf, NOME_ABA, COLUNA_ALVO)
        print(f"      Imagens flutuantes encontradas na coluna alvo: {len(candidatos)}")

        total = len(candidatos)
        for i, (celula, caminho_imagem) in enumerate(candidatos, 1):
            if i % 20 == 0 or i == total:
                print(f"      Processando {i}/{total} ({round(i / total * 100, 1)}%)")

            try:
                dados_candidato = zf.read(caminho_imagem)
            except KeyError:
                continue

            hash_candidato = hashlib.sha256(dados_candidato).hexdigest().upper()
            if hash_candidato == hash_referencia:
                resultados.append((celula, "Identica", 100.0))
                continue

            try:
                ptr_candidato = carregar_imagem_de_bytes(dados_candidato)
            except RuntimeError:
                continue  # arquivo nao e uma imagem valida (equivalente ao catch do PS1)

            try:
                percentual = calcular_percentual_similaridade(
                    ptr_referencia, ptr_candidato, RESOLUCAO_GRADE_COMPARACAO
                )
            finally:
                liberar_imagem(ptr_candidato)

            if percentual >= LIMIAR_MUITO_SEMELHANTE:
                resultados.append((celula, "Muito semelhante", percentual))
            elif percentual >= LIMIAR_SEMELHANTE:
                resultados.append((celula, "Semelhante", percentual))
            # abaixo do limiar minimo -> nao adiciona nada (regra de negocio)

    print("\n[3/4] Liberando recursos...")
    liberar_imagem(ptr_referencia)
    gdiplus.GdiplusShutdown(_token)

    print("\n[4/4] Resultado final:")
    print("\n==================================================================")
    print(" RESULTADO DA COMPARACAO")
    print("==================================================================")

    if not resultados:
        print(f"Nenhuma imagem identica ou semelhante (>= {LIMIAR_SEMELHANTE}%) foi encontrada.")
    else:
        resultados.sort(key=lambda r: r[0])
        print(f"{'Celula':<10}{'Status':<20}{'Similaridade':<12}")
        print("-" * 42)
        for celula, status, pct in resultados:
            print(f"{celula:<10}{status:<20}{pct}%")

    print(f"\nProcesso concluido. Total de correspondencias encontradas: {len(resultados)}")


if __name__ == "__main__":
    main()