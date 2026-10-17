# Agregador de Dados para Declaração de Imposto de Renda (Excel)

Projeto do desafio da DIO: uma ferramenta feita inteiramente no Excel para reunir, validar e consolidar as informações necessárias para a declaração de imposto de renda, com menu de navegação, validações automáticas e links rápidos.

## Objetivo

Centralizar em um único arquivo os dados que normalmente ficam espalhados em informes, recibos e extratos, reduzindo erros de digitação e facilitando a conferência antes de preencher a declaração oficial.

## Estrutura da planilha

| Aba | Função |
|---|---|
| **Menu** | Painel de navegação com links para todas as abas |
| **Contribuinte** | Nome, CPF, ano-base, ano do exercício e dependentes |
| **Rendimentos** | Fontes pagadoras, tipo de rendimento, valores e IR retido |
| **Deducoes** | Despesas dedutíveis por categoria, com data e valor |
| **Bens_Direitos** | Patrimônio em 31/12 do ano anterior e do ano-base, com variação |
| **Resumo** | Totais por categoria e indicadores consolidados |
| **Links** | Atalhos para sites oficiais (Receita Federal, gov.br, e-CAC, consulta de CPF) |
| **Listas** | Opções que alimentam as listas suspensas |

> Capturas de tela: adicione na pasta `/images` e referencie aqui, por exemplo `![Menu](images/menu.png)`.

## Recursos do Excel utilizados

### Navegação
- **Hiperlinks internos** no Menu para cada aba e o link "← Voltar ao Menu" em todas as páginas.
- **Hiperlinks externos** na aba Links.
- **Painéis congelados** nas abas de lançamento para manter o cabeçalho visível.

### Validação de dados
- **Lista suspensa** para tipo de rendimento, categoria de dedução, tipo de bem e situação do documento, alimentadas por **intervalos nomeados** (`lst_tipos_rend`, `lst_cats_ded`, `lst_tipos_bem`, `lst_situacao`).
- **Personalizada** para CPF (11 dígitos numéricos) e CPF/CNPJ (11 ou 14 dígitos numéricos):
  `=AND(LEN(B4)=11,ISNUMBER(VALUE(B4)))`
- **Número inteiro** para ano-base (2020 a 2035) e dependentes (0 a 20).
- **Decimal** para valores monetários (não aceita negativos).
- **Data** para o pagamento das deduções, que só aceita datas dentro do ano-base informado na aba Contribuinte:
  `=DATE(Contribuinte!$B$5,1,1)` até `=DATE(Contribuinte!$B$5,12,31)`
- Mensagens de erro personalizadas em cada regra.

### Fórmulas
- `SOMASES` (`SUMIFS`) e `CONT.SES` (`COUNTIFS`) para totais e quantidade de lançamentos por categoria no Resumo.
- `SE` e `E` para a variação patrimonial: `=IF(AND(C4="",D4=""),"",D4-C4)`.
- `DATA` para montar datas a partir do ano-base e `MÁXIMO` para evitar base de cálculo negativa.
- Ano do exercício calculado automaticamente: `=B5+1`.

### Formatação
- **Formatação condicional** destaca em laranja linhas com valor preenchido, mas sem descrição ou categoria.
- Convenção de cores: células amarelas com texto azul são campos de entrada; células cinza ou com texto verde são fórmulas.
- Formato monetário em reais e datas no padrão dd/mm/aaaa.

## Como usar

1. Abra a aba **Contribuinte** e preencha seus dados (o ano-base libera as validações de data).
2. Registre cada fonte pagadora em **Rendimentos**.
3. Registre despesas dedutíveis em **Deducoes**, usando a coluna de situação para acompanhar quais comprovantes você já tem.
4. Informe seu patrimônio em **Bens_Direitos**.
5. Consulte o **Resumo** para conferir totais antes de preencher o programa oficial.

As linhas de exemplo (linha 4 de cada aba de lançamento e os dados de Contribuinte) devem ser apagadas ou substituídas.

## Limitações e decisões

- A "base de cálculo simplificada" do Resumo é apenas organizacional: considera como tributáveis os três primeiros tipos de rendimento e subtrai o total de deduções. Não aplica tabela progressiva, limites de dedução nem desconto simplificado. O cálculo oficial é feito no programa da Receita Federal.
- Os links da aba Links devem ser conferidos de tempos em tempos, pois os portais do governo mudam.
- Cada aba de lançamento comporta 100 linhas; para mais, estenda o intervalo das validações e das fórmulas do Resumo.
- Foi usado `SOMASES` em vez de funções mais novas (como `XLOOKUP`) para manter compatibilidade com versões antigas do Excel.

## O que aprendi

*(Escreva aqui, com suas palavras: quais aulas foram mais úteis, quais dificuldades teve, o que faria diferente e o que pretende melhorar, como relatórios por mês ou proteção de células com senha.)*

## Arquivos do repositório

- `agregador_imposto_de_renda.xlsx` — a planilha
- `README.md` — esta documentação
- `images/` — capturas de tela
