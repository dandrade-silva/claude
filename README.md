## Principais comandos
1. **/model:** Selecionar o modelo que quer usar (haiku, sonnet, opus ou Fable) 
1. **/usage:** Permite visualizar com está o seu uso do claude
1. **/context:** Mostra quanto de contexto do claude code foi preenchido
1. **/clear:** Limpa histórico de conversas para reduzir o contexto
1. **/compact:** Resume tudo que foi feito até agora e reduz o contexto

## Organizacao do repositorio

- Cada automacao vive em `projetos/<nome>/`, isolada e autossuficiente: seu
  proprio script(s), `requirements.txt`, dados, log e README com o uso
  especifico.
- Cada automacao tem seu proprio `venv/` local (nao versionado, coberto pelo
  `.gitignore` na raiz).

## Projetos

- [projetos/refresh_excel/](projetos/refresh_excel/README.md) — atualiza uma
  planilha Excel via RefreshAll (automacao COM).

