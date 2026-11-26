# Auditoria de prazos e reprogramações

Este diretório contém o patch complementar do Apps Script para a funcionalidade de auditoria do Mapa de Entregas.

## Aplicação

1. Abra o projeto Apps Script vinculado à planilha.
2. Em `MAPA_CFG`, altere `ULTIMA_COLUNA: 168` para `ULTIMA_COLUNA: 173`.
3. Em `doPost(e)`, adicione ao `switch (acao)`:

```js
case "concluirEntrega":
  return jsonOutput_(concluirEntrega(body.linha, body.motivo, body.observacao));
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
