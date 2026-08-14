# Bradesco · Modal/Card Console — HTML/CSS/JS puro

Aplicação **100% no navegador** para o time de CRM Bradesco (Publicação e Orquestração)
buscar **e agora também editar** imagens duplicadas ou semelhantes em planilhas Excel —
sem servidor, sem instalar nada, sem precisar ficar abrindo e fechando o Excel toda hora.

## Novidades desta versão

- **Pesquisa por pasta**: além de buscar 1 imagem, dá pra selecionar uma pasta inteira
  do computador — o sistema procura cada foto da pasta na planilha, igual "abrir pasta"
  no VS Code.
- **Ordenação por urgência**: os resultados aparecem sempre com **VENCIDA** e **VENCE HOJE**
  no topo, depois os que vencem em breve, depois o resto.
- **Edição direta na planilha, sem abrir o Excel**: clique no lápis ✏️ de qualquer card
  (ou vá na nova aba **Planilha**) pra trocar a imagem, a descrição, o vencimento ou
  **qualquer outra coluna** daquela linha.
- **Aba "Planilha"**: uma tabela com todas as linhas encontradas, pra você conferir tudo
  visualmente sem precisar reabrir o arquivo original — só abre o Excel de novo depois
  de exportar, pra conferência final.
- **Backup automático**: toda vez que você exporta, o sistema salva uma cópia dos dados
  originais **numa aba nova dentro do próprio arquivo** (ex: `Backup_Desconsiderados_143022`)
  antes de aplicar qualquer edição.
- **Guia "Como usar"**: passo a passo direto na tela pra quem não tem tanta intimidade
  com tecnologia.

## Arquivos

```
standalone/
├── index.html   # abra este no navegador
├── styles.css
├── app.js
└── README.md
```

## Como usar

1. **Baixe** os 4 arquivos (`index.html`, `styles.css`, `app.js`, `README.md`) e coloque na mesma pasta
2. **Duplo clique** em `index.html` (ou arraste pro navegador)
3. Arraste a **planilha `.xlsx`** para a primeira caixa
4. Escolha **"Imagem única"** (1 foto) ou **"Pasta de imagens"** (várias fotos de uma vez) e selecione
5. (Opcional) Ajuste os **Parâmetros** (aba, colunas, similaridade mínima, etc.)
6. Clique em **EXECUTAR** — o terminal ao lado mostra tudo em tempo real
7. Os resultados aparecem ordenados por urgência (vencidas e "vence hoje" primeiro)
8. **Precisa corrigir algo?** Clique no lápis ✏️ do card (ou vá na aba **Planilha**) e altere
   a imagem, a descrição, o vencimento ou qualquer outra coluna daquela linha
9. Clique em **"Salvar alterações no Excel"** — o sistema baixa um novo `.xlsx` já atualizado,
   com um backup automático dos dados originais numa aba separada
10. **Abra o arquivo baixado no Excel** pra conferir antes de enviar pra alguém
11. Na aba **Histórico** você acompanha as últimas 100 execuções (salvas no `localStorage` do navegador)

> A tela tem um guia "Como usar" passo a passo, pensado pra quem não tem tanta intimidade
> com tecnologia — é só clicar pra abrir.

> Precisa de internet apenas na **primeira** abertura, para carregar as bibliotecas
> `SheetJS` e `JSZip` via CDN. Depois, funciona offline.
> Nenhum arquivo sai da sua máquina — todo o processamento é feito no navegador.

## Editando a planilha pelo navegador (sem abrir/fechar o Excel)

Sim, dá pra editar direto — é o mesmo princípio da leitura (JSZip manipulando o `.xlsx`
como um zip), só que na direção contrária:

- **Texto/descrição/qualquer coluna**: o app localiza a célula (`<c r="B57">`) dentro do
  XML da aba e substitui o conteúdo por uma string inline (`inlineStr`), sem mexer no
  restante da planilha. Funciona mesmo se a coluna que você quer editar não é nem a de
  descrição nem a de vencimento — dá pra adicionar quantas colunas quiser no modal de edição.
- **Vencimento (data)**: convertido pro formato serial que o Excel usa internamente
  (dias desde 30/12/1899), então continua sendo reconhecido como data de verdade.
- **Célula mesclada**: o app resolve automaticamente qual é a célula "mestre" do
  intervalo mesclado e escreve nela (é assim que o Excel espera).
- **Imagem**: como a célula da imagem já tem uma âncora (linha/coluna) e uma referência
  pro arquivo de mídia dentro do `.xlsx` (`xl/media/imageN.png`), o app simplesmente
  **substitui os bytes desse arquivo** pela nova imagem (recodificada no mesmo formato).
  Não precisa mexer em fórmulas de desenho (`drawings`) porque a posição já existe.
- **Backup**: antes de aplicar qualquer edição, o app duplica a aba inteira (dados +
  referência das imagens) para uma aba nova chamada `Backup_<aba>_<hora>`, registrada
  corretamente no `workbook.xml`, nos relacionamentos e no `[Content_Types].xml` — o
  Excel abre normalmente e mostra a aba de backup na lista de abas.

Todas as edições ficam **em memória, no navegador**, até você clicar em "Salvar
alterações no Excel". Nada é enviado pra nenhum servidor. Se fechar a aba antes de
exportar, as edições se perdem — então sempre exporte antes de sair.

### Limitações atuais (importante saber)

- Uma célula que continha **fórmula** vira valor fixo depois de editada (perde a fórmula).
- O backup duplica os **dados** e a **referência às imagens** da aba, mas não faz uma
  cópia física separada do arquivo `.xlsx` inteiro — tudo fica dentro do mesmo arquivo
  exportado, em abas diferentes.
- Arquivos `.xlsm` (com macro) perdem a macro ao serem reexportados — a biblioteca usada
  não reescreve VBA. Se sua planilha tiver macro, use apenas para leitura/consulta.

## Como funciona (por baixo dos panos)

- **SheetJS** lê os valores da planilha (descrição, vencimento) já resolvendo células mescladas
- **JSZip** desempacota o `.xlsx` (que é um zip) para extrair as imagens embutidas em `xl/media/` e ler os `drawings` que mapeiam cada imagem à sua célula
- **pHash em JavaScript**: reduz cada imagem para 32×32 tons de cinza → DCT 2D → pega o quadrante 8×8 de baixa frequência → gera um hash de 64 bits
- **Comparação por distância de Hamming**: se `distância == 0` → IDÊNTICA; se `similaridade >= mínima` → SEMELHANTE
- **Nitidez** usa o mesmo kernel do PIL (`FIND_EDGES`) aplicado via `canvas` e retorna a variância das bordas
- **Vencimento** usa `Date` do JS para classificar VENCIDA / VENCE HOJE / VENCE EM N DIAS / OK
- **Histórico** vai pro `localStorage` (até 100 execuções, ~5MB)
- **CSV** é gerado no cliente com `Blob` + `download` (`;` como separador, BOM UTF-8 pro Excel BR)

## Parâmetros

| Campo | Padrão | Descrição |
|---|---|---|
| Nome da aba | `Desconsiderados` | Aba da planilha onde estão os cards |
| Col. Descrição | `B` | Coluna da descrição da peça |
| Col. Vencimento | `F` | Coluna da data de vencimento |
| Similaridade mínima (%) | `75` | Match "semelhante" acima desse limiar |
| Dias de alerta | `15,7,3,0` | Alerta quando faltam N dias |
| Limiar de nitidez | `500` | Variância mínima de bordas para "boa nitidez" |

## Compatibilidade

Testado em Chrome, Edge e Firefox atuais.
Para planilhas gigantes (3000+ imagens) pode levar alguns minutos — o navegador
processa em segundo plano com "yields" para não travar a UI.
A seleção de **pasta** usa o atributo `webkitdirectory`, suportado no Chrome, Edge e
Firefox atuais (não funciona no Safari mobile).

## Ideias para os próximos passos (ainda não implementado)

Depois que essa leva de funcionalidades estiver validada com o time, os próximos
candidatos naturais são:
- **Monitoramento programado**: alerta automático quando alguma peça "vencer amanhã"
  sem precisar rodar a pesquisa manualmente.
- **Relatório em PDF**: gerar um PDF pronto pra imprimir/enviar com o resumo da execução.
- **Edição em lote**: aplicar a mesma alteração (ex: nova data de vencimento) pra várias
  linhas de uma vez, em vez de uma por uma.

## Diferenças em relação ao script Python original

O algoritmo é o mesmo (pHash 32×32 → DCT → 8×8 → mediana), mas o resultado numérico
pode variar em 1-2 bits em relação à `imagehash` do Python porque as bibliotecas
de redimensionamento e DCT têm implementações ligeiramente diferentes. A tolerância
de "similaridade mínima" (padrão 75%) absorve essa variação sem problemas.

## Créditos

Feito pro Iago · CRM Bradesco · Publicação e Orquestração.
Baseado no script Python original.
