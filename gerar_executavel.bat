@echo off
REM ============================================================================
REM Verificador de Validade de Pastas — gerar_executavel.bat
REM
REM Este script NAO e para o time usar no dia a dia. E para RODAR UMA VEZ,
REM na sua maquina (a que ja tem Python), para gerar um arquivo .exe pronto.
REM
REM Depois de pronto, esse .exe pode ser copiado para a maquina de qualquer
REM colega — mesmo quem NAO tem Python, VSCode ou qualquer coisa de
REM programacao instalada. E so copiar o arquivo e dar dois cliques.
REM ============================================================================

setlocal enabledelayedexpansion
title Gerando o executavel do Verificador de Validade

set "PASTA_DO_PROGRAMA=%~dp0"
cd /d "%PASTA_DO_PROGRAMA%"

REM --- 0. Cria (só na primeira vez) um atalho com ícone vermelho de
REM        engrenagem, diferente do ícone de "Iniciar", pra ficar claro que
REM        este é o arquivo que GERA o executável (não o que abre o
REM        programa no dia a dia). Clicar direto neste .bat continua
REM        funcionando normalmente.
if not exist "%PASTA_DO_PROGRAMA%Gerar Executavel.lnk" (
    powershell -NoProfile -WindowStyle Hidden -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut('%PASTA_DO_PROGRAMA%Gerar Executavel.lnk'); $sc.TargetPath = '%PASTA_DO_PROGRAMA%gerar_executavel.bat'; $sc.WorkingDirectory = '%PASTA_DO_PROGRAMA%'; $sc.IconLocation = '%PASTA_DO_PROGRAMA%icones\gerar_executavel.ico'; $sc.Description = 'Gera o executavel (.exe) do Verificador de Validade - uso pontual'; $sc.Save()" >nul 2>nul
    if exist "%PASTA_DO_PROGRAMA%Gerar Executavel.lnk" (
        echo(
        echo   Foi criado o atalho "Gerar Executavel" ^(icone vermelho de
        echo   engrenagem^) nesta pasta. Da proxima vez, pode usar ele em
        echo   vez deste arquivo .bat.
        echo(
    )
)

REM Importante: NAO usamos "where python" aqui pelo mesmo motivo do
REM iniciar.bat — em máquina sem Python instalado, o Windows costuma ter um
REM "atalho fantasma" python.exe (alias da Microsoft Store) que o "where"
REM encontraria e acharia, errado, que o Python está instalado. Rodamos
REM "python --version" de verdade e conferimos o código 9009, que aparece
REM tanto quando não existe python.exe nenhum quanto quando só existe esse
REM alias fantasma.
python --version >nul 2>nul
set "PYTHON_ENCONTRADO=%errorlevel%"
if "%PYTHON_ENCONTRADO%"=="9009" (
    echo(
    echo   Este script precisa do Python instalado NESTA maquina para gerar
    echo   o executavel ^(so aqui, uma vez^). Instale em
    echo   https://www.python.org/downloads/ marcando "Add Python to PATH".
    echo(
    echo   Atencao: se o Windows tentar abrir a Microsoft Store sozinho, e
    echo   sinal de que existe um "atalho" do Windows para o Python que nao
    echo   e o Python de verdade. Para desativa-lo, va em Configuracoes ^>
    echo   Aplicativos ^> Configuracoes avancadas do aplicativo ^> Aliases
    echo   de execucao do aplicativo, e desligue "python.exe" e
    echo   "python3.exe". Depois instale o Python pelo link acima.
    echo(
    pause
    exit /b 1
)

echo(
echo   Preparando o ambiente...
echo(

set "PYTHON_EXE=python"
set "USER_FLAG=--user"

if not exist "venv\Scripts\python.exe" (
    python -m venv venv
)

if exist "venv\Scripts\python.exe" (
    REM venv criado com sucesso: usamos o python de dentro dele. Dentro de
    REM um venv o "--user" nao existe (o pip ate recusa), entao tiramos essa
    REM flag nesse caminho.
    set "PYTHON_EXE=venv\Scripts\python.exe"
    set "USER_FLAG="
    call venv\Scripts\activate.bat
) else (
    REM Em algumas maquinas corporativas o "python -m venv" falha
    REM silenciosamente (antivirus, pasta protegida, permissao restrita).
    REM Nesse caso seguimos com o Python global da maquina, instalando os
    REM pacotes so para o usuario atual (--user), em vez de travar aqui.
    echo(
    echo   Nao foi possivel criar o ambiente isolado ^(venv^) nesta maquina.
    echo   Seguindo com o Python global, instalando os pacotes so para o
    echo   seu usuario.
    echo(
)

REM Importante: chamamos sempre "%PYTHON_EXE% -m pip" / "%PYTHON_EXE% -m
REM PyInstaller" em vez de "pip" / "pyinstaller" soltos. Rodar o comando
REM solto depende desses .exe estarem numa pasta que o Windows conhece
REM (o PATH) — e em maquina corporativa, sem venv, eles costumam ir parar
REM numa pasta pessoal que NAO esta no PATH, causando o erro "nao e
REM reconhecido". Chamando via "python -m", nao importa onde o pacote foi
REM instalado: o proprio Python sabe achar.
"%PYTHON_EXE%" -m pip install --upgrade pip --quiet %USER_FLAG%
"%PYTHON_EXE%" -m pip install -r requirements.txt --quiet %USER_FLAG%
"%PYTHON_EXE%" -m pip install pyinstaller --quiet %USER_FLAG%
if errorlevel 1 (
    echo(
    echo   Nao foi possivel instalar o PyInstaller. Verifique a conexao com
    echo   a internet e tente de novo.
    echo(
    pause
    exit /b 1
)

echo(
echo   Gerando o executavel — isso pode levar alguns minutos na primeira vez.
echo(

REM --onefile:  tudo (Python + Flask + o programa) vira UM arquivo .exe so,
REM             o mais facil de copiar e distribuir para o time.
REM --add-data: inclui a tela (template) e o visual (static) dentro do .exe —
REM             sem isso, o programa abriria em branco, sem CSS nem JS.
REM             No Windows o separador entre origem e destino e ";".
REM --noupx:    NAO compacta o .exe com UPX. Compactar deixaria o arquivo
REM             menor, mas e exatamente o padrao que antivirus corporativos
REM             (comuns em bancos) mais associam a malware, gerando falso
REM             positivo. Sem UPX, o .exe fica maior mas passa mais tranquilo.
REM --icon:     da ao .exe final o icone vermelho de "play" (o mesmo do
REM             atalho de Iniciar) — diferente de um .bat, um .exe aceita
REM             icone embutido de verdade, entao quem receber o arquivo ja
REM             ve um icone vermelho reconhecivel em vez do icone padrao
REM             de executavel do Windows.
"%PYTHON_EXE%" -m PyInstaller --onefile --noupx --name VerificadorDeValidade ^
    --icon "icones\iniciar.ico" ^
    --add-data "template;template" ^
    --add-data "static;static" ^
    app.py

if not exist "dist\VerificadorDeValidade.exe" (
    echo(
    echo   Algo deu errado e o executavel nao foi gerado. Role a tela para
    echo   cima para ver a mensagem de erro do PyInstaller.
    echo(
    pause
    exit /b 1
)

echo(
echo   ==========================================================
echo    Pronto! O executavel esta em:
echo(
echo      dist\VerificadorDeValidade.exe
echo(
echo    Copie SOMENTE esse arquivo para a maquina de quem vai usar.
echo    La, a pessoa so precisa dar dois cliques — nao precisa
echo    instalar Python, nem nada de programacao.
echo(
echo    O navegador abre sozinho. Para fechar o programa, e so
echo    fechar a janela preta que aparece junto.
echo   ==========================================================
echo(
pause
