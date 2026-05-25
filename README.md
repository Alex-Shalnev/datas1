import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import time
import sys

# ------------------- ПАКЕТНОЕ ЧТЕНИЕ ЛИСТА -------------------
def get_all_values_safe(worksheet, chunk_size=5000):
    """
    Читает лист по частям (по chunk_size строк) и возвращает все данные.
    Это предотвращает ошибки 503 при чтении больших листов.
    """
    all_data = []
    row = 1
    while True:
        # Определяем диапазон: от row до row+chunk_size-1
        end_row = row + chunk_size - 1
        # Получаем блок данных (столбцы до ZZ, можно расширить если нужно)
        range_name = f"A{row}:ZZ{end_row}"
        try:
            chunk = worksheet.get(range_name, value_render_option='UNFORMATTED_VALUE')
            if not chunk:
                break
            all_data.extend(chunk)
            if len(chunk) < chunk_size:
                # Последний блок, выходим
                break
            row += chunk_size
            time.sleep(0.5)  # небольшая пауза между блоками
        except Exception as e:
            print(f"Ошибка при чтении блока {range_name}: {e}")
            # Пробуем уменьшить размер блока
            chunk_size = max(1000, chunk_size // 2)
            print(f"Уменьшаем размер блока до {chunk_size} и повторяем...")
            continue
    # Удаляем возможные пустые строки в конце
    while all_data and all(len(r) == 0 for r in all_data[-1:]):
        all_data.pop()
    return all_data

# ------------------- ОСТАЛЬНЫЕ ФУНКЦИИ (БЕЗ ИЗМЕНЕНИЙ) -------------------
def retry_api_call(func, max_retries=5, base_delay=1):
    for attempt in range(max_retries):
        try:
            return func()
        except gspread.exceptions.APIError as e:
            if e.response.status_code in (503, 429) and attempt < max_retries - 1:
                delay = base_delay * (2 ** attempt)
                print(f"API ошибка {e.response.status_code}. Повтор через {delay} сек... (попытка {attempt + 1})")
                time.sleep(delay)
            else:
                raise
        except Exception as e:
            print(f"Неожиданная ошибка: {e}")
            raise
    raise Exception("Превышено максимальное количество попыток.")

def get_sheet_dimensions(worksheet):
    try:
        all_values = worksheet.get_all_values()
        current_rows = len(all_values)
        current_cols = max(len(row) for row in all_values) if all_values else 0
        return current_rows, current_cols
    except Exception as e:
        print(f"Ошибка получения размеров: {e}")
        return 1000, 20

def ensure_sheet_size(worksheet, required_rows, required_cols):
    try:
        current_rows, current_cols = get_sheet_dimensions(worksheet)
        print(f"Текущие размеры: {current_rows} строк, {current_cols} столбцов")
        print(f"Требуется: {required_rows} строк, {required_cols} столбцов")
        if required_rows > current_rows:
            rows_to_add = required_rows - current_rows + 1000
            print(f"Добавляем {rows_to_add} строк")
            retry_api_call(lambda: worksheet.add_rows(rows_to_add))
        if required_cols > current_cols:
            cols_to_add = required_cols - current_cols + 10
            print(f"Добавляем {cols_to_add} столбцов")
            retry_api_call(lambda: worksheet.add_cols(cols_to_add))
        return True
    except Exception as e:
        print(f"Ошибка при проверке размера листа: {e}")
        return False

def batch_update_sheet(worksheet, data, clear_first=True):
    if clear_first:
        retry_api_call(lambda: worksheet.clear())
        print("Лист очищен.")
    if not data:
        print("Нет данных для записи.")
        return
    headers = [data[0]]
    rows = data[1:]
    required_rows = len(data)
    required_cols = len(data[0]) if data else 0
    ensure_sheet_size(worksheet, required_rows, required_cols)
    chunk_size = 1000
    retry_api_call(lambda: worksheet.update(range_name='A1', values=headers))
    print("Заголовки записаны.")
    for i in range(0, len(rows), chunk_size):
        chunk = rows[i:i + chunk_size]
        start_row = 2 + i
        end_row = start_row + len(chunk) - 1
        range_name = f'A{start_row}'
        retry_api_call(lambda c=chunk, r=range_name: worksheet.update(range_name=r, values=c))
        print(f"Записано строк: {start_row} - {end_row}")
        time.sleep(1)

# ------------------- ОСНОВНАЯ ФУНКЦИЯ -------------------
def merge_data():
    start_time = datetime.now()
    print(f'Начало выполнения: {start_time}')

    # АВТОРИЗАЦИЯ (используем ваш новый JSON, если он есть)
    # Замените путь на актуальный
    json_path = r"C:\Users\a.shalnev\Desktop\Scripts\rollout-sheets-d86a3b31d340.json"
    # Если этого файла нет, вернитесь к старому:
    # json_path = r"C:\Users\a.shalnev\Desktop\Scripts\ap-orders-441610-bf8cb50a9e0a.json"
    
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(json_path, scope)
    client = gspread.authorize(creds)

    SPREADSHEET_ID = '1bYgFHuApF0v7dvLd8gX85kR3jCVe3rRkvCyBV3oVC7M'
    spreadsheet = retry_api_call(lambda: client.open_by_key(SPREADSHEET_ID))

    shop_sheet = retry_api_call(lambda: spreadsheet.worksheet('shop'))
    ar_tem_sheet = retry_api_call(lambda: spreadsheet.worksheet('ArTem'))
    coordinator_sheet = retry_api_call(lambda: spreadsheet.worksheet('Координатор'))

    print('Чтение листа shop (безопасное, по частям)...')
    shop_data = get_all_values_safe(shop_sheet, chunk_size=5000)
    print(f'  shop: {len(shop_data)} строк')
    
    print('Чтение листа ArTem...')
    ar_tem_data = get_all_values_safe(ar_tem_sheet, chunk_size=5000)
    print(f'  ArTem: {len(ar_tem_data)} строк')
    
    print('Чтение листа Координатор...')
    coordinator_data = get_all_values_safe(coordinator_sheet, chunk_size=5000)
    print(f'  Координатор: {len(coordinator_data)} строк')

    # --- Ключи для shop (столбец Q) ---
    print('Создание ключей для shop...')
    shop_keys = []
    for row in shop_data[1:]:
        partner_code = row[1].strip().lower() if len(row) > 1 else ''
        partner_name = row[4].strip().lower() if len(row) > 4 else ''
        shop_keys.append([f"{partner_code}_{partner_name}"])
    if shop_keys:
        retry_api_call(lambda: shop_sheet.update(range_name=f'Q2:Q{len(shop_keys)+1}', values=shop_keys))
        print(f"Ключи shop записаны в диапазон Q2:Q{len(shop_keys)+1}")

    # --- Ключи для ArTem (столбец AP) ---
    print('Создание ключей для ArTem...')
    ar_tem_keys = []
    for row in ar_tem_data[1:]:
        store_number = row[6].strip().lower() if len(row) > 6 else ''
        partner = row[1].strip().lower() if len(row) > 1 else ''
        ar_tem_keys.append([f"{store_number}_{partner}"])
    if ar_tem_keys:
        retry_api_call(lambda: ar_tem_sheet.update(range_name=f'AP2:AP{len(ar_tem_keys)+1}', values=ar_tem_keys))
        print(f"Ключи ArTem записаны в диапазон AP2:AP{len(ar_tem_keys)+1}")

    # --- Карты ---
    shop_map = {}
    for i, row in enumerate(shop_data[1:]):
        key = shop_keys[i][0] if i < len(shop_keys) else ''
        shop_map[key] = row

    coordinator_map = {}
    for row in coordinator_data[1:]:
        if len(row) < 5:
            continue
        territory = row[0].strip().lower()
        name = row[1].strip()
        partner = row[2].strip().lower()
        project = row[3].strip().lower()
        subject = row[4].strip().lower()
        if not partner:
            continue
        keys_to_try = []
        if project and subject:
            keys_to_try.append(f"{territory}_{partner}_{project}_{subject}")
        if project:
            keys_to_try.append(f"{territory}_{partner}_{project}")
        keys_to_try.append(f"{territory}_{partner}")
        for key in keys_to_try:
            if key not in coordinator_map:
                coordinator_map[key] = name

    # --- Объединение ---
    print('Объединение данных...')
    shop_header = shop_data[0] if shop_data else []
    artem_header = ar_tem_data[0] if ar_tem_data else []
    headers = artem_header + shop_header + [
        "Код_в_партнерской_сети",
        "Координатор отдела кофе",
        "Ключ ArTem",
        "Ключ Shop"
    ]
    merged_data = [headers]

    for i, row in enumerate(ar_tem_data[1:]):
        key_ar_tem = ar_tem_keys[i][0] if i < len(ar_tem_keys) else ''
        if not key_ar_tem or '_' not in key_ar_tem:
            shop_row = ['ошибка'] * len(shop_header)
            coordinator = ''
        else:
            shop_row = shop_map.get(key_ar_tem, ['нет данных'] * len(shop_header))
            region = row[2].strip().lower() if len(row) > 2 else ''
            partner = row[1].strip().lower() if len(row) > 1 else ''
            project = row[8].strip().lower() if len(row) > 8 else ''
            subject = row[13].strip().lower() if len(row) > 13 else ''
            coordinator = ''
            if partner:
                if project and subject:
                    coordinator = coordinator_map.get(f"{region}_{partner}_{project}_{subject}", '')
                if not coordinator and project:
                    coordinator = coordinator_map.get(f"{region}_{partner}_{project}", '')
                if not coordinator:
                    coordinator = coordinator_map.get(f"{region}_{partner}", '')
        full_row = row + shop_row + ['', coordinator, key_ar_tem, key_ar_tem]
        merged_data.append(full_row)

    print(f"Подготовлено строк для записи: {len(merged_data)}")

    # --- Запись в Datas ---
    print('Запись результатов в лист "Datas"...')
    try:
        datas_sheet = retry_api_call(lambda: spreadsheet.worksheet('Datas'))
    except gspread.exceptions.WorksheetNotFound:
        print('Лист "Datas" не найден. Создаём новый...')
        datas_sheet = retry_api_call(lambda: spreadsheet.add_worksheet(title='Datas', rows=20000, cols=len(headers)))
    batch_update_sheet(datas_sheet, merged_data, clear_first=True)

    end_time = datetime.now()
    execution_time = (end_time - start_time).total_seconds()
    print(f'Выполнение завершено за {execution_time:.2f} секунд')
    print(f'Обработано строк из ArTem: {len(ar_tem_data) - 1}')

if __name__ == '__main__':
    try:
        merge_data()
    except KeyboardInterrupt:
        print("\nСкрипт прерван пользователем.")
        sys.exit(1)
    except Exception as e:
        print(f"\nКритическая ошибка: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
