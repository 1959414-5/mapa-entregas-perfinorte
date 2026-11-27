# Faturamento e importação de NF-e

## Fluxo por perfil

- **ADMIN** administra os módulos, alterna entre Teste e Produção e tem acesso às funções administrativas.
- **OPERACIONAL** acompanha somente as viagens liberadas para a operação e executa as ações de parada e viagem.
- **FATURAMENTO** cria viagens em estado `RASCUNHO`, seleciona entregas, vincula uma ou várias NF-e por entrega e libera a viagem depois da conferência.

A viagem em rascunho não aparece para OPERACIONAL. A liberação altera o status para `PROGRAMADA`. O backend exige ao menos uma NF-e ativa para cada parada antes de liberar.

## Campos lidos do XML

O parser lê os elementos pelo nome local, sem depender do prefixo de namespace:

| Campo | Caminho |
|---|---|
| Chave de acesso | `infNFe@Id` (remove o prefixo `NFe`) |
| Número e série | `infNFe/ide/nNF` e `infNFe/ide/serie` |
| Emissão | `infNFe/ide/dhEmi` ou `dEmi` |
| Valor total | `infNFe/total/vNFTot`; alternativa `infNFe/total/ICMSTot/vNF` |
| Peso líquido | Soma de cada `infNFe/transp/vol/pesoL` |
| Emitente | CNPJ/CPF e `xNome` de `emit` |
| Destinatário | CNPJ/CPF, `xNome`, município e UF de `dest` |

O peso é somado uma vez por grupo `vol`, sem somar novamente o peso dos itens. O total da viagem é a soma das NF-e ativas vinculadas àquela viagem.

## Armazenamento

- `VIAGEM_NOTAS_FISCAIS`: metadados normalizados e vínculo entre NF, viagem e parada.
- `VIAGEM_NOTAS_FISCAIS_TESTE`: registros exclusivos do ambiente de teste.
- `PERFINORTE_NFE_XML` e `PERFINORTE_NFE_XML_TESTE`: arquivos XML originais no Drive.
- `CADASTRO_VIAGENS` recebe total de valor de NF e peso líquido em colunas próprias. A versão de teste usa abas próprias.

A chave de acesso de 44 dígitos impede que uma mesma NF-e ativa seja importada novamente no mesmo ambiente.

## Isolamento de teste

O ambiente de teste grava viagens, paradas, ocorrências, veículos e NF-e em abas/arquivos separados. As entregas da operação podem ser consultadas para seleção, mas a conclusão de parada em teste não chama a baixa da entrega real. A ação `concluirEntrega` fica bloqueada em teste.

## Secrets de autenticação no Cloudflare

O Worker atual mantém as credenciais legadas de OPERACIONAL em `APP_USER` e `APP_PASSWORD`. As novas contas são configuradas separadamente por secrets:

- `APP_ADMIN_USER`
- `APP_ADMIN_PASSWORD`
- `APP_FATURAMENTO_USER`
- `APP_FATURAMENTO_PASSWORD`

Não guardar senhas, tokens ou secrets no repositório.

## Publicação

Este documento acompanha a implementação na branch de desenvolvimento. Não publicar o código Apps Script ou substituir o Worker em produção até concluir os testes integrados e aprovar a implantação.
