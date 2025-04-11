# Interpolação de Preços

## Visão Geral

Este sistema foi desenvolvido para processar planilhas Excel contendo registros de preços associados a datas, preenchendo automaticamente os valores ausentes através de interpolação linear. O programa identifica os preços faltantes e calcula novos valores com base nos registros de datas anteriores e posteriores, garantindo uma série temporal de preços completa sem a necessidade de inserção manual.

## Funcionalidades

- Carregamento de planilhas Excel contendo dados de preços
- Identificação automática de valores faltantes
- Interpolação linear baseada em datas para preenchimento de preços
- Formatação dos valores no padrão brasileiro de moeda (R$ 0,00)
- Geração de uma nova planilha com todos os registros preenchidos
- Ajuste automático da largura das colunas para melhor visualização

## Requisitos

- Python 3.6 ou superior
- Bibliotecas: pandas, numpy, openpyxl

Para instalar as dependências necessárias:

```bash
pip install pandas numpy openpyxl
```

## Como Usar

1. Execute o script principal:

```bash
python interpolacao_precos.py
```

Ou apenas execute no notebook

2. Quando solicitado, forneça o caminho completo para o arquivo Excel:

```
Digite o caminho do arquivo Excel: caminho/para/seu/arquivo.xlsx
```

3. O programa processará o arquivo e gerará um novo com o sufixo "_preenchido" adicionado ao nome original.

## Formato da Planilha

A planilha de entrada deve conter as seguintes colunas:

- `data do preco`: Data associada ao registro de preço
- `preco produto`: Valor do preço (pode estar em formato de texto com "R$")

Outras colunas podem estar presentes e serão preservadas no arquivo de saída.

## Algoritmo de Interpolação

O sistema utiliza um algoritmo de interpolação linear para preencher os valores faltantes:

1. Localiza o registro válido mais próximo antes da data com preço faltante
2. Localiza o registro válido mais próximo depois da data com preço faltante
3. Calcula o valor proporcional com base na distância temporal entre estes pontos
4. Se as datas coincidirem, usa a média aritmética dos preços

## Limitações

- O sistema requer pelo menos dois registros com preços válidos para realizar a interpolação
- Não é possível interpolar valores no início ou final da série temporal (apenas no meio)
- O programa assume que a relação entre as datas e os preços é linear

## Saída

O programa gera:

1. Um arquivo Excel contendo todos os dados originais com os preços faltantes preenchidos
2. Informações no console sobre o processo de preenchimento (quantidade de valores preenchidos)

## Tratamento de Erros

O sistema inclui tratamento para:
- Arquivos Excel inexistentes ou corrompidos
- Formatos de data inválidos
- Valores de preço em diferentes formatos
- Ausência das colunas necessárias
