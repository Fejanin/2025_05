import PyPDF2

def extract_text_from_pdf(pdf_path: str) -> str:
    try:
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            # Проверка на шифрование
            if reader.is_encrypted:
                print("Файл зашифрован. Попытка расшифровки без пароля...")
                if not reader.decrypt(''):
                    raise ValueError("Требуется пароль для расшифровки PDF")
            
            full_text = ''
            for page_num, page in enumerate(reader.pages):
                text = page.extract_text()
                if text:
                    full_text += f"--- Страница {page_num + 1} ---\n{text}\n"
            return full_text
    except FileNotFoundError:
        return f"Ошибка: файл '{pdf_path}' не найден."
    except Exception as e:
        return f"Произошла ошибка: {e}"

# Пример использования
pdf_file = 'Бердянск на 21,06,.pdf'
text = extract_text_from_pdf(pdf_file)
print(text)
