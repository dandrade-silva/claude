## Principais comandos
1. **/model:** Selecionar o modelo que quer usar (haiku, sonnet, opus ou Fable) 
1. **/usage:** Permite visualizar com está o seu uso do claude
1. **/context:** Mostra quanto de contexto do claude code foi preenchido
1. **/clear:** Limpa histórico de conversas para reduzir o contexto
1. **/compact:** Resume tudo que foi feito até agora e reduz o contexto

## Organizacao do repositorio

- `requirements.txt` e `venv/` ficam na raiz e sao compartilhados por todas
  as automacoes.
- Cada automacao vive em `projetos/<nome>/`, com seu proprio script(s),
  dados, log e README explicando o uso especifico.

## Projetos

- [projetos/refresh_excel/](projetos/refresh_excel/README.md) — atualiza uma
  planilha Excel via RefreshAll (automacao COM).

