# Viva Cuidar — Pendência Diária

Painel Streamlit para acompanhamento diário de locações, vencimentos e criticidade operacional.

## Controle de acesso por unidade

O painel possui usuários separados por loja. Cada usuário de loja visualiza somente os dados da própria unidade, inclusive KPIs, tabelas, gráficos, histórico e exportações.

Usuários configurados:

- `USERANONI` — administrador / acesso geral.
- `VCVIVA` — Vila Clementino.
- `GESTORA` — Cuidar Bem.
- `VIVAINDEPENDENCIA` — Independência.
- `LIMEIRAVIVA` — Limeira Nova.
- `RIO2VIVA` — Rio Claro 2.
- `BARAOVIVA` — Viva Barão.
- `MATRIZVIVA` — Viva Rio Claro Matriz.
- `SCVIVA` — Viva São Carlos.

As senhas não ficam gravadas em texto puro no código; apenas hashes SHA-256 são utilizados para validação.

### Permissões

- **USERANONI**: visualiza todas as unidades, pode usar o filtro de unidade, publicar a planilha diária e baixar a base completa.
- **Usuários das lojas**: visualizam somente sua unidade e não têm acesso ao upload da base completa nem ao download da base geral.

## O que o site mostra

- Login protegido.
- Upload diário da planilha Excel pelo administrador.
- % de ocupação do parque.
- Quantidade e % dos locados em atraso.
- Atrasos de 24h, 48h e 72h+.
- Atraso acumulado ≥24h, ≥48h e ≥72h.
- Vencimentos de hoje e dos próximos 5 dias.
- Gráficos de ocupação, criticidade, vencimentos e grupos com atraso.
- Filtros por grupo, situação e busca por cliente/contrato/equipamento.
- Filtro de unidade disponível somente para o administrador.
- Fila de casos críticos e tabelas operacionais.
- Resumo por grupo.
- Histórico respeitando a permissão da unidade logada.
- Exportação da lista filtrada para Excel.

## Rodar no Windows

Abra o Anaconda Prompt ou CMD dentro da pasta do projeto e rode:

```bat
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Uso diário

1. O administrador entra com `USERANONI`.
2. Abre **Atualizar base do dia** no menu lateral.
3. Envia o novo arquivo `.xlsx`.
4. Clica em **Publicar atualização**.
5. Cada loja acessa o mesmo site com seu próprio usuário e enxerga somente seus dados.

A planilha precisa conter, no mínimo:

`LOCAL`, `Equipamento`, `Status`, `Nome Fantasia`, `Produto`, `Grupo`, `Contrato`, `DATA FINAL`.

## Nomes das unidades

O controle de acesso compara o campo `LOCAL` da planilha com as unidades permitidas. O sistema já aceita alguns nomes alternativos comuns, como `RC2`, `RC MATRIZ`, `BARAO` e versões sem acento.

Se uma loja entrar e aparecer a mensagem de que a unidade não foi encontrada, confira exatamente como o nome está escrito na coluna `LOCAL` da planilha e ajuste os aliases no bloco `USER_ACCOUNTS` do `app.py`.

## Privacidade

A base contém nomes de clientes, números de contrato e informações operacionais. Mantenha o repositório privado e compartilhe o endereço do painel somente com pessoas autorizadas.

## Aba exclusiva de Inadimplência

Somente o usuário administrador `USERANONI` consegue ver a aba **💰 Inadimplência** e publicar o arquivo `MonitorReceber.xlsx`.

A aba mostra:

- valor total vencido que continua em aberto;
- quantidade de títulos em aberto;
- clientes inadimplentes;
- ticket médio e maior atraso;
- cartões de **24h**, **48h** e **72h+**, com quantidade de casos e valor em aberto;
- filtros por unidade, faixa de atraso e busca por cliente/contrato/fatura/documento;
- inadimplência por unidade;
- inadimplência por faixa de atraso;
- maiores valores em aberto por cliente;
- valores em aberto por data de vencimento;
- tabela detalhada de títulos em aberto para baixa;
- exportação da inadimplência filtrada para Excel.

O saldo usa a coluna **Valor Vencido** do relatório MonitorReceber. Assim, quando existe pagamento parcial, o painel considera apenas o valor que continua vencido.

Para atualizar, entre como administrador e abra **💰 Atualizar inadimplência** no menu lateral.

## Inadimplência — origem, faturamento e períodos

A aba de inadimplência (somente administrador) também mostra:

- Origem do título: Venda, Locação, Intermediação ou Não identificado.
- Status de faturamento: Faturado ou Não faturado.
- Filtro de vencimento dos últimos 30, 60, 90 ou 120 dias.
- Radar 24h, 48h e 72h+.
- Gráfico de aging em 30 / 60 / 90 / 120 dias.

Regras usadas no MonitorReceber:
- Pedido preenchido > 0 = Venda.
- Contrato iniciado por INTERM = Intermediação.
- Demais contratos preenchidos = Locação.
- Sem número de Fatura e sem `DATA FAT` na Observação = Não faturado.

## Data manual da base

Somente o usuário administrador `USERANONI` vê a opção **🗓️ Data exibida da base** no menu lateral.

O administrador pode escolher manualmente a data e a hora que aparecem no selo **Base atualizada** do cabeçalho. A data escolhida é exibida para todos os usuários. Também existe a opção **Usar data automática** para voltar a mostrar a data/hora real do arquivo publicado.
