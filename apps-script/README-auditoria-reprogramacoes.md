# Auditoria de prazos e reprogramações

Este diretório contém o patch complementar do Apps Script para a funcionalidade de auditoria do Mapa de Entregas.

## Aplicação

1. Abra o projeto Apps Script vinculado à planilha.
2. Em `MAPA_CFG`, altere `ULTIMA_COLUNA: 168` para `ULTIMA_COLUNA: 173`.
3. Em `doPost(e)`, adicione ao `switch (acao)`:

```js
case "concluirEntrega":
  return jsonOutput_(concluirEntrega(body.linha, body.motivo, body.observacao));
case "obterConfigMotivoAtraso":
  return jsonOutput_(obterConfigMotivoAtraso_());
case "salvarConfigMotivoAtraso":
  return jsonOutput_(salvarConfigMotivoAtraso_(body.ativo));
```

4. Na função `onEdit(e)`, depois da validação da aba e da linha, adicione o trecho abaixo. Mantenha os tratamentos existentes de ID/endereço:

```js
if (col === MAPA_CFG.COL_PREVISAO) {
  registrarEdicaoPrevisaoManual_(e);
  return;
}
if (col === MAPA_CFG.COL_STATUS &&
    String(e.range.getValue() || "").trim().toUpperCase() === "ENTREGUE") {
  registrarConclusaoManual_(sheet, row);
  return;
}
```

5. Cole o conteúdo de `patch-auditoria-reprogramacoes.gs` no final do projeto Apps Script.
6. Em `aplicarConfirmacao_(hist, registro)`, dentro do bloco `if (status === "ENTREGUE")`, adicione a chamada abaixo logo depois de gravar o status:
   ```js
   registrarConclusaoManual_(hist, linhaNum);
   ```
   Isso também registra o horário real quando uma confirmação do motorista é aprovada no mapa.
7. Execute `inicializarAuditoriaEntregas()` uma vez no editor, autorize os acessos e verifique se a aba `Historico_Reprogramacoes` foi criada.
8. Em **Implantar > Gerenciar implantações > Editar > Nova versão > Implantar**, publique a nova versão do Web App.
9. Atualize o site com Ctrl+F5 depois que a nova implantação estiver ativa.

## Colunas reservadas

| Coluna | Conteúdo |
|---|---|
| FM (169) | Data/hora real da conclusão |
| FN (170) | Motivo do atraso |
| FO (171) | Observação do motivo |
| FP (172) | Previsão original, preservada |
| FQ (173) | Identificador estável da entrega |

O log completo fica na aba `Historico_Reprogramacoes`, com previsão anterior, nova previsão e data/hora da alteração.

**Migração:** para linhas antigas, a previsão original começa com a data existente no momento da primeira inicialização. O sistema não tem como reconstruir reprogramações anteriores à implantação.

**Critério inicial de atraso:** compara a data real da conclusão com a data original. Como ainda não foram definidos horários-limite para manhã e tarde, o atraso começa no dia seguinte à previsão original.

**Registros que não são entregas:** linhas com `FA = AVISO` e `I` (código/ID da entrega) vazio são ignoradas na migração da auditoria e na lista retornada ao site. Elas não entram nos totais, atrasos, rankings ou gráficos e não geram reprogramações/conclusões na auditoria.


## Controle liga/desliga da justificativa de atraso

O Dashboard inclui um toggle **Motivos de atraso**. Ligado, exige a seleção de um dos 12 motivos ao concluir uma entrega fora da previsão original. Desligado, permite concluir sem abrir o modal e sem gravar motivo. A configuração fica em `ScriptProperties` e vale para o sistema inteiro; não depende do navegador. O padrão inicial é desligado. A exigência só passa a valer depois que alguém ativar o toggle.

Para ativar a integração do toggle:

1. No `doPost(e)`, inclua os dois cases `obterConfigMotivoAtraso` e `salvarConfigMotivoAtraso` mostrados acima.
2. No bloco de auditoria que já foi adicionado ao Apps Script, inclua a constante `CONFIG_EXIGIR_MOTIVO_ATRASO_` e as funções `obterConfigMotivoAtraso_()` e `salvarConfigMotivoAtraso_()` da versão atual de `patch-auditoria-reprogramacoes.gs`.
3. Substitua as funções `concluirEntrega()` e `atualizarCampoEntrega()` existentes pelas versões atuais do mesmo arquivo. Não cole o patch completo por cima sem remover a versão anterior: isso duplicaria declarações `const`.
4. Salve e publique uma nova versão do Web App.

O site lê a configuração ao autenticar e salva cada alteração diretamente no Apps Script. A preferência permanece após fechar o navegador ou entrar em outro computador.


## Arquivo consolidado do Apps Script

A versão integral e consolidada do projeto está em [Perfinorte-Mapa-Entregas.gs](Perfinorte-Mapa-Entregas.gs). Ela reúne o código principal, auditoria, filtros de avisos sem código e toggle de motivos de atraso (desativado por padrão).

**Para atualizar o projeto:** substitua o conteúdo do arquivo principal atual no editor do Apps Script pelo conteúdo desse arquivo. Não adicione o arquivo como um segundo arquivo com as mesmas funções, pois o projeto já possui declarações existentes.

**Propriedades do script necessárias:** antes de implantar, configure em Configurações do projeto > Propriedades do script:
- `GOOGLE_GEOCODING_KEY`: sua chave atual do Google Geocoding.
- `ORS_API_KEY`: sua chave atual do OpenRouteService.
- `MAPA_WORKER_SECRET`: mantenha o valor que já está configurado para autenticar as chamadas do Cloudflare Worker.

As chaves de API não ficam gravadas neste arquivo do Git. Se elas ainda estiverem apenas no código antigo, copie os valores para as Propriedades do script antes de substituir o arquivo. Não publique essas chaves no repositório.
