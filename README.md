# Bradesco · Modal/Card Console — HTML/CSS/JS puro

Aplicação **100% no navegador** para o time de CRM Bradesco (Publicação e Orquestração)
buscar imagens duplicadas ou semelhantes em planilhas Excel — sem servidor, sem instalar nada.

## Arquivos

```
standalone/
├── index.html   # abra este no navegador
├── styles.css
├── app.js
└── README.md
```

## Como usar

1. **Baixe** os 3 arquivos (`index.html`, `styles.css`, `app.js`) e coloque na mesma pasta
2. **Duplo clique** em `index.html` (ou arraste pro navegador)
3. Arraste a **planilha `.xlsx`** e a **imagem de pesquisa** para as áreas de upload
4. (Opcional) Ajuste os **Parâmetros** (aba, colunas, similaridade mínima, etc.)
5. Clique em **EXECUTAR** — o terminal ao lado mostra tudo em tempo real
6. No final, veja o resumo à esquerda e clique em **Exportar CSV**
7. Na aba **Histórico** você acompanha as últimas 100 execuções (salvas no `localStorage` do navegador)

> Precisa de internet apenas na **primeira** abertura, para carregar as bibliotecas
> `SheetJS` e `JSZip` via CDN. Depois, funciona offline.
> Nenhum arquivo sai da sua máquina — todo o processamento é feito no navegador.

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

## Diferenças em relação ao script Python original

O algoritmo é o mesmo (pHash 32×32 → DCT → 8×8 → mediana), mas o resultado numérico
pode variar em 1-2 bits em relação à `imagehash` do Python porque as bibliotecas
de redimensionamento e DCT têm implementações ligeiramente diferentes. A tolerância
de "similaridade mínima" (padrão 75%) absorve essa variação sem problemas.

## Créditos

Feito pro Iago · CRM Bradesco · Publicação e Orquestração.
Baseado no script Python original.
