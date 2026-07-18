## refresh_excel.py

Script que atualiza os dados de uma planilha Excel via RefreshAll nativo do
Excel. Automatiza o Excel via COM (`pywin32`), aciona o `RefreshAll` e aguarda
a conclusao antes de salvar.

### Uso
1. Coloque o arquivo `.xlsx`/`.xlsm` na mesma pasta do script.
2. Ajuste a constante `EXCEL_FILENAME` no topo de `refresh_excel.py` para o
   nome real do arquivo.
3. Instale as dependencias (na raiz do repositorio): `pip install -r ../../requirements.txt`
4. Execute: `python refresh_excel.py`
5. Confira o log gerado em `refresh_excel.log` (mesma pasta).

Requer Windows com Microsoft Excel instalado.

### Observacao sobre agendamento (Task Scheduler)
O script ja roda "headless" (sem exibir o Excel nem popups), preparado para
uso futuro como tarefa agendada. Um ponto de atencao conhecido: automacao COM
do Excel sem sessao interativa pode ser instavel (limitacao da Microsoft para
automacao server-side do Office). Se isso ocorrer ao agendar, o ajuste
recomendado e configurar a tarefa como "Run only when user is logged on".
