# YM_matcher_mac — версия скрипта для запуска на macOS.
# Рядом со скриптом должны находиться:
#   - zvonko_in.xlsx
#   - ZVONKO Digital X Яндекс Музыка.xlsx
# Результат: zvonko_out_YM.xlsx и Нет в zvonko - YM.xlsx
# Установка зависимости: pip install openpyxl

from pathlib import Path
import re
import sys
from collections import defaultdict
from copy import copy

import openpyxl
from openpyxl.styles import Alignment

# При запуске как EXE ищем файлы рядом с EXE, а не во временной папке PyInstaller.
if getattr(sys, "frozen", False) and sys.platform == "darwin":
    # Внутри .app исполняемый файл лежит в Contents/MacOS; отчёты ищем рядом с .app.
    BASE_DIR = Path(sys.executable).resolve().parents[3]
elif getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
# Используем только файл с точным именем zvonko_in.xlsx.
INPUT_FILE = BASE_DIR / "zvonko_in.xlsx"

SOURCE_FILE = BASE_DIR / "ZVONKO Digital X Яндекс Музыка.xlsx"
OUTPUT_FILE = BASE_DIR / "zvonko_out_YM.xlsx"
MISSING_FILE = BASE_DIR / "Нет в zvonko - YM.xlsx"


def normalize(value):
    """Нормализует текст для сравнения, не меняя исходное значение в Excel."""
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().casefold()


def artist_tokens(value):
    """Выделяет отдельные имена из ячейки с несколькими артистами."""
    if value is None:
        return set()
    parts = re.split(r"\s*(?:,|;|\n|&|\bfeat\.?\b|\bft\.?\b|\s+x\s+)\s*", str(value), flags=re.IGNORECASE)
    return {normalize(part) for part in parts if normalize(part)}


def find_header(ws, required_headers, max_scan=50):
    """Находит строку заголовков по заданным названиям столбцов."""
    for row in range(1, min(ws.max_row or 1, max_scan) + 1):
        headers = {
            normalize(ws.cell(row, col).value): col
            for col in range(1, (ws.max_column or 1) + 1)
            if ws.cell(row, col).value is not None
        }
        if all(normalize(name) in headers for name in required_headers):
            return row, headers
    return None, {}


def split_values(value):
    """Возвращает отдельные значения и убирает служебный x/х в начале названия."""
    if value is None:
        return []
    result = []
    for line in str(value).splitlines():
        part = line.strip()
        if not part or part.casefold() == "не поддержали":
            continue
        # В исходном отчёте перед некоторыми названиями плейлистов стоит
        # служебный символ x (латиница) или х (кириллица). Убираем его только
        # в начале строки и только если за ним следует пробел.
        part = re.sub(r"^[xх]\s+", "", part, flags=re.IGNORECASE).strip()
        if part and part.casefold() != "не поддержали":
            result.append(part)
    return result


def main():
    for path in (INPUT_FILE, SOURCE_FILE):
        if not path.exists():
            raise FileNotFoundError(f"Не найден файл: {path}\nПоложите оба Excel-файла рядом со скриптом.")

    # Исходную книгу Яндекс Музыки открываем только для чтения.
    # Скрипт не заменяет формулы и не изменяет исходный файл.
    source_wb = openpyxl.load_workbook(SOURCE_FILE, data_only=True)
    playlists_by_artist_release = defaultdict(list)
    seen_by_artist_release = defaultdict(set)
    playlists_by_artist = defaultdict(list)
    seen_by_artist = defaultdict(set)
    source_artists = {}
    source_sheets = []

    # Обрабатываем строго один указанный лист — остальные листы игнорируются.
    target_sheet = "ZVONKO Digital X Яндекс Музыка"
    if target_sheet not in source_wb.sheetnames:
        raise ValueError(f"В файле Яндекс Музыки не найден лист «{target_sheet}».")
    source_ws = source_wb[target_sheet]
    # Берём столбцы F–L (6–12). Столбец E и знак «+» игнорируются.
    for source_ws in [source_ws]:
        header_row, headers = find_header(source_ws, ("Артист", "Название"))
        if header_row is None:
            continue
        artist_col = headers[normalize("Артист")]
        release_col = headers[normalize("Название")]
        source_sheets.append(source_ws.title)

        for row in range(header_row + 1, (source_ws.max_row or header_row) + 1):
            artist_value = source_ws.cell(row, artist_col).value
            artist = normalize(artist_value)
            release = normalize(source_ws.cell(row, release_col).value)
            if not artist:
                continue

            source_artists.setdefault(artist, str(artist_value).strip())
            tokens = artist_tokens(artist_value)
            for token in tokens:
                source_artists.setdefault(token, token)

            # Обрабатываем только F–L; столбец E полностью игнорируется.
            # Если в F, G или H есть любое непустое значение, подставляем
            # название соответствующего столбца вместо содержимого ячейки.
            # В I–L переносим фактические значения, удаляя служебный x/х в начале.
            row_values = []
            for col in range(6, 13):
                cell_value = source_ws.cell(row, col).value
                if col in (6, 7, 8):
                    if cell_value is not None and str(cell_value).strip():
                        header_value = source_ws.cell(header_row, col).value
                        if header_value is not None and str(header_value).strip():
                            row_values.append(str(header_value).strip())
                else:
                    row_values.extend(split_values(cell_value))

            # Удаляем дубли внутри строки, сохраняя исходный порядок.
            unique_values = []
            row_seen = set()
            for value in row_values:
                key = normalize(value)
                if key and key not in row_seen:
                    unique_values.append(value)
                    row_seen.add(key)

            if not unique_values:
                continue

            for value in unique_values:
                key = normalize(value)
                if key not in seen_by_artist[artist]:
                    playlists_by_artist[artist].append(value)
                    seen_by_artist[artist].add(key)
                if release:
                    # Сопоставляем строго по паре артист + название релиза.
                    for artist_key in ({artist} | tokens):
                        pair = (artist_key, release)
                        if key not in seen_by_artist_release[pair]:
                            playlists_by_artist_release[pair].append(value)
                            seen_by_artist_release[pair].add(key)

    if not source_sheets:
        raise ValueError("Не найдены листы с заголовками «Артист» и «Название» в файле Яндекс Музыки.")

    # Открываем zvonko_in и запоминаем заголовки ДО вставки нового столбца.
    input_wb = openpyxl.load_workbook(INPUT_FILE)
    input_meta = {}
    input_artists = set()
    for ws in input_wb.worksheets:
        header_row, headers = find_header(ws, ("Артист",))
        if header_row is None:
            # Совместимость с таблицей без заголовков: артист в столбце A.
            header_row, artist_col, release_col = 1, 1, None
        else:
            artist_col = headers.get(normalize("Артист")) or headers.get(normalize("Имя артиста"))
            release_col = headers.get(normalize("Релиз")) or headers.get(normalize("Название релиза"))
            if not artist_col:
                artist_col = headers.get(normalize("Имя артиста"), 1)
        # Если строка найдена по «Имя артиста», но не по «Артист», поддерживаем и этот вариант.
        if header_row == 1 and artist_col == 1:
            _, headers2 = find_header(ws, ("Имя артиста",))
            if headers2:
                artist_col = headers2[normalize("Имя артиста")]
                release_col = headers2.get(normalize("Релиз")) or headers2.get(normalize("Название релиза"))
        input_meta[ws.title] = (header_row, artist_col, release_col)
        for row in range(header_row + 1, (ws.max_row or header_row) + 1):
            value = ws.cell(row, artist_col).value
            if normalize(value):
                input_artists.add(normalize(value))
                input_artists.update(artist_tokens(value))

    # Вставляем столбец C «Плейлисты» на каждом листе zvonko_in.
    for ws in input_wb.worksheets:
        header_row, artist_col, release_col = input_meta[ws.title]
        ws.insert_cols(3, 1)
        # Ставим заголовок в строке заголовков, а также в C1 для привычного вида старого шаблона.
        header_cell = ws.cell(header_row, 3)
        header_cell.value = "Плейлисты"
        style_source = ws.cell(header_row, 2)
        if style_source.has_style:
            header_cell._style = copy(style_source._style)
        if style_source.number_format:
            header_cell.number_format = style_source.number_format

        if artist_col >= 3:
            artist_col += 1
        if release_col is not None and release_col >= 3:
            release_col += 1

        if release_col is None:
            print(f"Предупреждение: на листе «{ws.title}» не найден столбец «Релиз»/«Название релиза». Сопоставление по одному артисту не выполняется.")

        for row in range(header_row + 1, (ws.max_row or header_row) + 1):
            artist_value = ws.cell(row, artist_col).value
            artist = normalize(artist_value)
            release = normalize(ws.cell(row, release_col).value) if release_col else ""
            playlists = []
            if artist and release:
                playlists = list(playlists_by_artist_release.get((artist, release), []))
                if not playlists:
                    for token in artist_tokens(artist_value):
                        for value in playlists_by_artist_release.get((token, release), []):
                            if value not in playlists:
                                playlists.append(value)
            elif artist and not release and release_col is not None:
                # Если релиз не указан, используем совпадение по артисту.
                playlists = list(playlists_by_artist.get(artist, []))
            cell = ws.cell(row, 3)
            cell.value = "\n".join(playlists) if playlists else None
            base_alignment = copy(ws.cell(row, 2).alignment)
            cell.alignment = Alignment(horizontal=base_alignment.horizontal, vertical="top", wrap_text=True)

        ws.column_dimensions["C"].width = 42

    input_wb.save(OUTPUT_FILE)
    input_wb.close()
    print(f"Готово: {OUTPUT_FILE}")
    print("Для F, G и H при наличии любого значения подставляются названия столбцов; фактические плейлисты берутся из I–L. Столбец E игнорируется.")
    print("Если релиз заполнен, сопоставление выполняется только по точной паре «артист + релиз».")

    # Создаём отдельную книгу с исходными листами источника, содержащими артистов,
    # которых нет в zvonko_in. В ней сохраняются все исходные столбцы и значения.
    missing_wb = openpyxl.load_workbook(SOURCE_FILE, data_only=True)
    # В файл «Нет в zvonko - YM.xlsx» также включаем только целевой лист.
    for sheet_name in list(missing_wb.sheetnames):
        if sheet_name != target_sheet:
            del missing_wb[sheet_name]
    ws = missing_wb[target_sheet]
    header_row, headers = find_header(ws, ("Артист", "Название"))
    if header_row is None:
        raise ValueError(f"На листе «{target_sheet}» не найдены заголовки «Артист» и «Название».")
    artist_col = headers[normalize("Артист")]
    remove_rows = []
    for row in range(header_row + 1, (ws.max_row or header_row) + 1):
        artist_key = normalize(ws.cell(row, artist_col).value)
        if not artist_key or artist_key in input_artists or any(token in input_artists for token in artist_tokens(ws.cell(row, artist_col).value)):
            remove_rows.append(row)
    for row in reversed(remove_rows):
        ws.delete_rows(row, 1)
    if not missing_wb.sheetnames:
        missing_wb.create_sheet("Нет данных")
    missing_wb.save(MISSING_FILE)
    missing_wb.close()
    source_wb.close()
    print(f"Создан файл с артистами, которых нет в zvonko_in: {MISSING_FILE}")
    print(f"Листы источника, использованные для сопоставления: {', '.join(source_sheets)}")


if __name__ == "__main__":
    try:
        main()
        if sys.platform == "darwin" and getattr(sys, "frozen", False):
            import subprocess
            subprocess.run(["/usr/bin/osascript", "-e", 'display dialog "Обработка завершена.\n\nСозданы файлы:\nzvonko_out_YM.xlsx\nНет в zvonko - YM.xlsx" buttons {"OK"} default button "OK" with title "YM Matcher"'], check=False)
    except Exception as error:
        message = (f"Ошибка: {error}\n\nПроверьте, что рядом с YM_matcher.app находятся файлы zvonko_in.xlsx и ZVONKO Digital X Яндекс Музыка.xlsx.")
        try:
            (BASE_DIR / "YM_matcher_error.txt").write_text(message, encoding="utf-8")
        except Exception:
            pass
        if sys.platform == "darwin" and getattr(sys, "frozen", False):
            import subprocess
            escaped = message.replace('\\', '\\\\').replace('"', '\\"')
            subprocess.run(["/usr/bin/osascript", "-e", f'display dialog "{escaped}" buttons {{"OK"}} default button "OK" with title "YM Matcher — ошибка"'], check=False)
        else:
            print(message)
