# Plano de Higienização de Vendas da Porsche

> Mapa de dados de colunas

## Colunas de entrada canônicas

| Coluna | Obrigatório | Descrição | Exemplos de valores brutos |
|---|---:|---|---|
| `sale_id` | sim | ID único do registro de venda | `1`, `86` |
| `data_da_venda` | sim | Data bruta da venda | `2024-13-05`, `02/05/24`, `15 de julho de 2024` |
| `nome_do_cliente` | sim | Nome do cliente | `John Smith`, `Mary Johnson` |
| `porsche_model` | sim | Modelo e versão do Porsche | `911 Carrera`, `Macan Elétrico` |
| `model_year` | sim | Ano do modelo do veículo | `2023`, `20-20`, `dois mil e vinte e dois` |
| `preço_de_venda` | sim | Preço de venda em formatos USD desorganizados | `$ 985.000,00`, `188 mil USD` |
| `vehicle_mileage` | sim | Quilometragem do veículo em formatos confusos de milhas/km | `12.500 milhas`, `15.200 km`, `novo` |
| `método_de_pagamento` | sim | Método de pagamento | `Cartão de crédito`, `transferência bancária`, `criptomoeda` |
| `cidade` | sim | Cidade | `Nova York`, `San Jose` |
| `estado` | sim | Nome ou abreviação de estado dos EUA | `ny`, `Califórnia` |
| `vendedor` | sim | Nome do vendedor | `William Clark`, `SARAH lee` |
| `status_da_entrega` | sim | Status da entrega | `Entregue!!!`, `em trânsito`, `ENTREGADO` |

## Colunas de saída higienizadas

| Origem | Saída | Formato |
|---|---|---|
| `data_da_venda` | `DataDaVendaSanitizada` | `AAAA-MM-DD` ou `INVÁLIDO` |
| `porsche_model` | `PorscheModelSanitized` | Etiqueta canônica do modelo/acabamento da Porsche |
| `model_year` | `ModelYearSanitized` | Ano com quatro dígitos ou `INVALID` |
| `preço_de_venda` | `Preço_de_vendaSanitizado` | Valor em USD como string decimal com duas casas decimais |
| `vehicle_mileage` | `VehicleMileageSanitized` | Milhas inteiras |
| `método_de_pagamento` | `Método_de_pagamentoSanitizado` | Rótulo controlado |
| `cidade` | `CidadeSanitizada` | Cidade com inicial maiúscula |
| `estado` | `EstadoSanitizado` | Código de duas letras do USPS ou `INVÁLIDO` |
| `status_de_entrega` | `Status_de_entrega_higienizado` | Etiqueta controlada |

Regras de higienização

### Datas

Normalizar para o formato ISO `AAAA-MM-DD`.

Aceitar formatos comuns:

- `AAAA-MM-DD`
- `AAAA/MM/DD`
- `AAAA.MM.DD`
- `MM/DD/AAAA`
- `MM/DD/AA`
- `MM-DD-AA`
- `Mês DD, AAAA`
- `Seg DD de AAAA`

Datas de calendário inválidas, como `2024-13-05`, `2024-02-30`, `31 de abril de 2024` e `2027-06-40`, devem ser marcadas como `INVÁLIDAS`.

### Modelos Porsche

Normalizar os nomes dos modelos para o formato de título canônico. Preservar caracteres de aparar válidos, como:

- `911 Carrera`
- `911 Carrera S`
- `911 Carrera GTS`
- `911 Turbo`
- `911 Turbo S`
- `911 GT3`
- `911 GT3 RS`
- `911 Dakar`
- `911 Targa 4`
- `911 Targa 4S`
- `718 Cayman`
- `718 Cayman S`
- `718 Cayman GT4 RS`
- `718 Boxster`
- `718 Boxster GTS`
- `718 Spyder RS`
- `Pimenta`
- `Cayenne S`
- `Cayenne Coupé`
- `Cayenne E-Hybrid`
- `Cayenne Turbo`
- `Cayenne Turbo GT`
- `Macan`
- `Macan S`
- `Macan T`
- `Macan GTS`
- `Macan Electric`
- `Panamera`
- `Panamera 4`
- `Panamera 4S`
- `Panamera Turbo`
- `Panamera Turbo S`
- `Panamera 4 E-Hybrid`
- `Taycan`
- `Taycan 4S`
- `Taycan GTS`
- `Taycan Turbo`
- `Taycan Turbo S`
- `Taycan Cross Turismo`

Modelos desconhecidos devem ser escritos com a primeira letra de cada palavra em maiúscula, em vez de omitidos.

### Ano do modelo

Normalizar para um ano de quatro dígitos. Converter formas de texto comuns como "vinte e vinte e quatro" e formatos como "20-24" ou "20 24". Valores fora do período de "1990" a "2035" devem se tornar "INVÁLIDOS".

### Preço de venda

Converter para um valor decimal em USD sem símbolos, vírgulas ou texto.

Exemplos:

- `$ 985.000,00` -> `985.000,00`
- `870000 USD` -> `870000,00`
- `$645 mil` -> `645.000,00`
- `188 mil USD` -> `188000,00`
- `oitenta e dois mil USD` -> `82000,00`
- `duzentos mil dólares americanos` -> `200000,00`

### Quilometragem do veículo

Normalizar para milhas inteiras.

Exemplos:

- `12.500 milhas` -> `12.500`
- `Milhas: 15.200` -> `15200`
- `zero milhas`, `novo`, `0 mi`, `carro novo` -> `0`
- `doze mil milhas` -> `12000`

Se um valor for explicitamente rotulado como `KM`, converta quilômetros em milhas usando `1 km = 0,621371 milhas`, arredondado para o número inteiro mais próximo.

### Método de pagamento

Normalizar para um dos seguintes valores:

- `Cartão de Crédito`
- `Cartão de Débito`
- `Transferência bancária`
- `Transferência Bancária`
- `Financiamento`
- `Arrendamento`
- `Dinheiro`
- `Pagamento ACH`
- `Pagamento em Criptomoedas`

Métodos de pagamento desconhecidos devem ser escritos com a primeira letra de cada palavra em maiúscula.

### Cidade

Normalizar os nomes das cidades para maiúsculas apenas no início de cada palavra e preservar a pontuação conhecida, como "St. Louis".

### Estado

Normalizar os estados dos EUA para códigos de duas letras do USPS. Converter nomes e abreviações, sem distinção entre maiúsculas e minúsculas.

Exemplos:

- `Califórnia` -> `CA`
- `california` -> `CA`
- `ca` -> `CA`
- `Nova Iorque` -> `NY`

Estados desconhecidos devem se tornar `INVÁLIDOS`.

### Status da entrega

Normalizar para um dos seguintes valores:

- `Entregue`
- `Pendente`
- `Em trânsito`
- `Cancelado`
- `Aguardando Entrega`
- `Aguardando Recolhimento`
- `Aguardando aprovação`
- `Em análise`
- `Enviado`
- `Aguardando revisão`

Corrija erros de pontuação, maiúsculas e minúsculas, hifenização e o erro ortográfico comum `DELIVERD`.

## Verificações de qualidade

Após a normalização, verifique:

- Nenhuma coluna original foi removida.
- As colunas higienizadas aparecem imediatamente após suas colunas de origem.
- As datas estão no formato ISO ou são `INVÁLIDAS`.
- Os valores de estado são códigos USPS de duas letras ou `INVÁLIDO`.
- A quilometragem é expressa em milhas inteiras.
- O preço é numérico, com duas casas decimais.
- Não são introduzidos novos valores em branco nas colunas higienizadas; em vez disso, use `INVALID` ou valores padrão normalizados.

## Manipulação inválida

Use `INVALID` quando um valor não puder ser normalizado com segurança. Nunca deixe um campo sanitizado em branco.