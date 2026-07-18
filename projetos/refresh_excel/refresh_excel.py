"""Atualiza (RefreshAll) as conexoes de dados de uma planilha Excel.

Abre a planilha via automacao COM do Excel, dispara o RefreshAll, aguarda
a conclusao, salva e fecha o Excel.
"""

import logging
import sys
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

# --- Configuracao ---
BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILENAME = "feriados2026.xlsx"  # Ajustar para o nome real do arquivo
EXCEL_PATH = BASE_DIR / EXCEL_FILENAME
LOG_FILE = BASE_DIR / "refresh_excel.log"

log = logging.getLogger("refresh_excel")


def configure_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
        ],
    )


def validate_target_file(path: Path) -> None:
    if not path.exists():
        raise FileNotFoundError(f"Arquivo Excel nao encontrado: {path}")


def wait_for_refresh_completion(excel) -> None:
    try:
        excel.CalculateUntilAsyncQueriesDone()
    except Exception:
        log.warning("CalculateUntilAsyncQueriesDone indisponivel ou falhou.")


def refresh_workbook(path: Path) -> bool:
    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        log.info("Iniciando Excel (DispatchEx).")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False

        log.info("Abrindo planilha: %s", path)
        wb = excel.Workbooks.Open(
            str(path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
        )

        log.info("Disparando RefreshAll.")
        wb.RefreshAll()

        wait_for_refresh_completion(excel)

        log.info("Salvando planilha.")
        wb.Save()

        log.info("Atualizacao concluida com sucesso.")
        return True

    except pywintypes.com_error as e:
        log.error("Erro COM ao atualizar planilha: %s", e.args)
        return False
    except Exception:
        log.exception("Erro inesperado ao atualizar planilha.")
        return False
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                log.exception("Falha ao fechar a planilha.")
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                log.exception("Falha ao encerrar o Excel.")
        pythoncom.CoUninitialize()


def main() -> None:
    configure_logging()
    log.info("=== Iniciando script de atualizacao da planilha ===")

    try:
        validate_target_file(EXCEL_PATH)
    except FileNotFoundError as e:
        log.error(str(e))
        sys.exit(1)

    success = refresh_workbook(EXCEL_PATH)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
