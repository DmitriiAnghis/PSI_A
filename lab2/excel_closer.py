"""Закрытие книги, открытой в Microsoft Excel (только Windows, через pywin32)."""
import logging
import os

log = logging.getLogger("tickets")


def _same_file(a, b):
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def close_in_excel(path):
    """Сохраняет и закрывает книгу path, если она открыта в Excel.

    Возвращает True, если книга была найдена и закрыта. На других ОС или без
    pywin32 возвращает False, и программа показывает обычное окно «Файл занят».
    Другие открытые книги Excel не трогаются.
    """
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return False
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return False
    try:
        for book in list(excel.Workbooks):
            if _same_file(book.FullName, path):
                book.Close(SaveChanges=True)
                log.info("Книга %s закрыта в Excel перед записью", path)
                if excel.Workbooks.Count == 0:
                    excel.Quit()
                return True
    except Exception as exc:
        log.warning("Не удалось закрыть %s в Excel: %s", path, exc)
    return False
