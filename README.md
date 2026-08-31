# Verificador de Validade de Pastas

Sistema web simples (HTML + CSS + Python/Flask) que verifica, em uma pasta de
rede, quais subpastas cujo nome começa com **"Val"** têm uma data de validade
**vencida há 3 meses ou mais**.

Este sistema é **somente leitura**: em nenhum momento cria, altera, move ou
exclui arquivos ou pastas. Ele apenas lê a estrutura de diretórios.

## Como a busca funciona (importante)

O sistema entra na pasta de rede informada e **percorre automaticamente
todos os níveis de subpastas** dentro dela — não é preciso que as pastas
"Val..." estejam logo no primeiro nível. Por exemplo, essa estrutura funciona
normalmente:

```
\\servidor\rede\
├── Filial_SP\
│   ├── Val_03.24        ← encontrada
│   └── Val_09.26
├── Filial_RJ\
│   ├── Val_ATE_01.25     ← encontrada
│   └── DocumentosGerais\
└── Filial_MG\
    └── Sub_Nivel_2\
        └── Val_05.20    ← encontrada mesmo em um 3º nível
```

Assim que uma pasta "Val..." é identificada, o sistema não entra no conteúdo
dela (não é necessário e evita varredura desnecessária em pastas de rede
grandes).

A leitura das pastas é feita **em paralelo** (várias pastas lidas ao mesmo
tempo, não uma de cada vez), já que em uma pasta de rede o tempo gasto é
principalmente de latência de rede, não de processamento. Isso reduz bastante
o tempo total em compartilhamentos com muitas pastas. O número de leituras
simultâneas pode ser ajustado em `scanner.py`, na constante
`MAX_LEITURAS_PARALELAS` (padrão: 16).

## Regras usadas para identificar as pastas

1. O nome da pasta precisa começar com `Val` (não diferencia maiúsculas de
   minúsculas: `val`, `VAL`, `Val` etc. são aceitos).
2. A validade é a **última ocorrência** de uma data no formato `MM.AA` ou
   `MM.AAAA` encontrada no nome da pasta. Exemplos:
   - `Val_08.26` → validade 08/2026
   - `Val_02.09.26_ATE_11.26` → validade 11/2026 (última ocorrência)
   - `Val_ATE_02.27` → validade 02/2027
3. A pasta só aparece no resultado se a validade estiver vencida há **3 meses
   ou mais**, contando corretamente a virada de ano (ex: validade 11/2025,
   hoje fevereiro/2026 → 3 meses vencida).

## Estrutura do projeto

```
verificador-validade/
├── app.py              # Rotas web e da API (Flask)
├── scanner.py          # Lógica de varredura das pastas (sem dependências web)
├── requirements.txt    # Dependências Python
├── README.md
├── templates/
│   └── index.html      # Página principal
└── static/
    ├── style.css        # Estilo visual
    └── script.js        # Chamada da API e exibição dos resultados
```

## Como executar localmente

### 1. Pré-requisitos
- Python 3.9 ou superior instalado.

### 2. Criar um ambiente virtual (recomendado)

```bash
python -m venv venv
```

Ativar o ambiente:
- **Windows:** `venv\Scripts\activate`
- **Linux/macOS:** `source venv/bin/activate`

### 3. Instalar as dependências

```bash
pip install -r requirements.txt
```

### 4. Executar o sistema

```bash
python app.py
```

### 5. Acessar no navegador

Abra: [http://127.0.0.1:5000](http://127.0.0.1:5000)

Digite o caminho da pasta de rede que deseja verificar, por exemplo:
- Windows: `\\servidor\compartilhamento\pasta`
- Linux/macOS (montagem de rede): `/mnt/rede/pasta`

Clique em **Buscar**. O sistema mostrará uma tabela com nome, validade, meses
vencida e caminho completo de cada pasta encontrada. Os resultados sempre
aparecem na tela; o botão **Salvar log** é opcional e gera, além disso, um
arquivo `.txt` com o mesmo conteúdo, e o botão **Exportar para Excel** gera
uma planilha `.xlsx`.

## Observações importantes

- **Segurança:** por padrão, o servidor roda apenas em `127.0.0.1` (acessível
  só na própria máquina). Só altere para `0.0.0.0` em `app.py` se realmente
  precisar acessar de outros computadores da rede, e com cautela — essa
  ferramenta permite listar caminhos de pastas informados por quem a usa.
- **Erros de acesso:** se alguma subpasta não puder ser lida (permissão
  negada, por exemplo), o sistema continua a varredura nas demais pastas e
  mostra um aviso na tela, em vez de interromper tudo.
- **Caminho inexistente:** se o caminho informado não existir ou não for uma
  pasta, uma mensagem de erro clara é exibida.
