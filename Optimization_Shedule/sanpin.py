# Источник: СанПиН 2.4.2.2821-10 (Приложение 3)

SANPIN_DATA = {
    # МАТЕМАТИКА
    "Математика": {"difficulty": 10, "direction": "Математика"},
    "Алгебра": {"difficulty": 9, "direction": "Математика"},
    "Геометрия": {"difficulty": 8, "direction": "Математика"},
    "Вероятность и статистика": {"difficulty": 7, "direction": "Математика"},
    "Мат. вертикаль": {"difficulty": 8, "direction": "Математика"},

    # ФИЛОЛОГИЯ
    "Русский язык": {"difficulty": 9, "direction": "Филология"},
    "Литература": {"difficulty": 7, "direction": "Филология"},
    "Иностранный язык": {"difficulty": 7, "direction": "Филология"},
    "Английский язык": {"difficulty": 7, "direction": "Филология"},
    "Немецкий язык": {"difficulty": 7, "direction": "Филология"},
    "Французский язык": {"difficulty": 7, "direction": "Филология"},
    "Родной язык": {"difficulty": 6, "direction": "Филология"},

    # ЕСТЕСТВОЗНАНИЕ 
    "Физика": {"difficulty": 9, "direction": "Естествознание"},
    "Химия": {"difficulty": 8, "direction": "Естествознание"},
    "Биология": {"difficulty": 6, "direction": "Естествознание"},
    "География": {"difficulty": 5, "direction": "Естествознание"},
    "Астрономия": {"difficulty": 7, "direction": "Естествознание"},
    "Окружающий мир": {"difficulty": 4, "direction": "Естествознание"},

    # ОБЩЕСТВОЗНАНИЕ 
    "История": {"difficulty": 6, "direction": "Обществознание"},
    "Обществознание": {"difficulty": 6, "direction": "Обществознание"},
    "Право": {"difficulty": 6, "direction": "Обществознание"},
    "Экономика": {"difficulty": 6, "direction": "Обществознание"},

    # ТЕХНОЛОГИЯ И ИНФОРМАТИКА
    "Информатика": {"difficulty": 7, "direction": "Технология"},
    "Технология": {"difficulty": 4, "direction": "Технология"},
    "Черчение": {"difficulty": 5, "direction": "Технология"},

    # ФИЗКУЛЬТУРА И ОБЖ 
    "Физкультура": {"difficulty": 2, "direction": "Физкультура"},
    "ОБЖ": {"difficulty": 3, "direction": "Физкультура"},
    "НВП": {"difficulty": 3, "direction": "Физкультура"},

    # ИСКУССТВО 
    "ИЗО": {"difficulty": 3, "direction": "Искусство"},
    "Музыка": {"difficulty": 3, "direction": "Искусство"},
    "МХК": {"difficulty": 4, "direction": "Искусство"}
}

DEFAULT_SUBJECT = {"difficulty": 5, "direction": "Разное"}

def enrich_subject_data(subject_name):
    """
    Получить данные SanPin для предмета.
    Если точного совпадения нет, ищет частичное.
    """
    if not subject_name:
        return DEFAULT_SUBJECT.copy()
    
    clean_name = subject_name.strip()
    
    # 1. Точное совпадение
    if clean_name in SANPIN_DATA:
        return SANPIN_DATA[clean_name]
    
    # 2. Частичное совпадение
    for key, data in SANPIN_DATA.items():
        if key.lower() in clean_name.lower():
            return data
            
    return DEFAULT_SUBJECT.copy()
