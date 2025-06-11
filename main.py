import os
import pandas as pd
import argparse
import psycopg

# Константы
HOST = "localhost"
PORT = "5432"
DEFAULT_DIRECTORY = r"C:\Users\eugen\OneDrive\Рабочий стол\sql_to_xlsx"


# Функция для получения следующего доступного имени файла
def get_next_filename(directory, base_filename="table", extension=".xlsx"):
    i = 0
    while os.path.exists(os.path.join(directory, f"{base_filename}{i}{extension}")):
        i += 1
    return f"{base_filename}{i}{extension}"


# Функция для получения данных из базы данных
def fetch_data_from_db(dbname, user, password, table_name):
    try:
        conn = psycopg.connect(
            dbname=dbname,
            user=user,
            password=password,
            host=HOST,
            port=PORT,
            options='-c client_encoding=UTF8'  # Указание кодировки
        )
        print("Подключение к базе данных успешно установлено.")

        # Формируем запрос для указанной таблицы
        query = f"SELECT * FROM {table_name}"

        with conn.cursor() as cursor:
            cursor.execute(query)
            columns = [desc[0] for desc in cursor.description]
            rows = cursor.fetchall()

            # Преобразуем данные в DataFrame
            df = pd.DataFrame(rows, columns=columns)

            # Проверяем и обрабатываем кодировку для строковых столбцов
            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].apply(lambda x: str(x).encode('utf-8', errors='ignore').decode('utf-8') if isinstance(x, str) else x)

        conn.close()

        return df
    except Exception as e:
        print(f"Ошибка при подключении или запросе: {e}")
        return None


# Основная логика программы
def main():
    # Создание парсера аргументов командной строки
    parser = argparse.ArgumentParser(description="Преобразование данных из SQL в Excel.")
    parser.add_argument("--filename", type=str, help="Название выходного файла (например, table0.xlsx)")
    parser.add_argument("--directory", type=str, help="Директория для сохранения файла")

    args = parser.parse_args()

    # Запрос информации для подключения к БД и имя таблицы
    dbname = input("Введите имя базы данных: ")
    user = input("Введите имя пользователя: ")
    password = input("Введите пароль: ")
    table_name = input("Введите имя таблицы, которую хотите экспортировать: ")

    # Дефолтные значения для директории и имени файла
    directory = args.directory or DEFAULT_DIRECTORY
    filename = args.filename or get_next_filename(directory)

    # Проверка существования директории
    if not os.path.exists(directory):
        print(f"Ошибка: директория {directory} не существует.")
        return

    # Извлечение данных из базы данных
    df = fetch_data_from_db(dbname, user, password, table_name)

    if df is not None:
        # Сохранение данных в файл Excel с кодировкой UTF-8
        file_path = os.path.join(directory, filename)
        df.to_excel(file_path, index=False, engine='openpyxl')
        print(f"Данные успешно сохранены в файл {file_path}")


if __name__ == "__main__":
    main()
