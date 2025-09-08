import database
import doc_generator
import json
import os

def select_conference(conferences):
    """Выбор конференции из списка."""
    if not conferences:
        print("Ошибка: в базе данных нет доступных конференций.")
        return None
    
    if len(conferences) == 1:
        print(f"Найдена одна конференция: {conferences[0]['title']}")
        return conferences[0]
    
    print("Доступные конференции:")
    for i, conf in enumerate(conferences, 1):
        print(f"{i}. {conf['title']}")
    
    while True:
        try:
            choice = input("Введите номер конференции для генерации документов: ")
            choice = int(choice)
            if 1 <= choice <= len(conferences):
                return conferences[choice - 1]
            else:
                print(f"Пожалуйста, выберите номер от 1 до {len(conferences)}.")
        except ValueError:
            print("Пожалуйста, введите корректный номер.")

def main():
    json_file_path = "conference_data.json"
    
    print("Запуск программы генерации документов конференции...")
    
    # Этап 1: Извлечение данных из базы и создание JSON
    print("Извлечение данных из базы данных...")
    try:
        database.create_conference_json(json_file_path)
        print(f"JSON файл успешно создан: {json_file_path}")
    except Exception as e:
        print(f"Ошибка при создании JSON файла: {e}")
        return

    # Этап 2: Чтение JSON и выбор конференции
    try:
        with open(json_file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
        conferences = data["conferences"]
    except Exception as e:
        print(f"Ошибка при чтении JSON файла: {e}")
        return

    selected_conference = select_conference(conferences)
    if not selected_conference:
        return

    # Этап 3: Создание папки для конференции
    conference_title = selected_conference["title"]
    output_dir = conference_title
    try:
        os.makedirs(output_dir, exist_ok=True)
        print(f"Создана папка для документов: {output_dir}")
    except Exception as e:
        print(f"Ошибка при создании папки {output_dir}: {e}")
        return

    # Этап 4: Генерация DOCX документов
    print(f"Генерация DOCX документов для конференции '{conference_title}'...")
    try:
        doc_generator.create_conference_docx(selected_conference, output_dir)
        print("Документы успешно созданы:")
        print(f"- {output_dir}/1_Программа_к43.docx")
        print(f"- {output_dir}/2_Отчет о проведении {conference_title}.docx")
        print(f"- {output_dir}/3_Список представляемых к публикации докладов.docx")
    except Exception as e:
        print(f"Ошибка при создании DOCX документов: {e}")
        return

    print("Программа успешно завершена.")

if __name__ == "__main__":
    main()
