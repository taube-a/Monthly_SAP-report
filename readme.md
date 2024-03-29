# Автоматизия составления отчёта (Python)

### Задача

В течение месяца изменяется файл задач:
+ Добавляются новые задачи;
+ Изменяется статус решения задачи (находятся на уточнении, уже в работе, закрыты);
+ Решаются старые задачи.

В конце месяца необходимо составить отчёт о состоянии задач. Все задачи в отчёте должны быть отсортированы в следующем порядке:
1. Критичность: самые приотритетные к решению - выше;
2. Дата дедлайна: задача с ближайшей датой дедлайна - выше;
3. Дата создания: если у двух задач дата дедлайна совпадает, то выше должна оказаться задача, созданная раньше. 

Также в очтёте должны быть следующие строки: 
* Общая статистика по всем задачам;
* Подытоги по критичности.

Строки итогов должны находиться над соотвествующими записями.

В конце отчёта должны находиться 2 строки с согласованием отчёта. 

Пример изначального файла - "example.xlsx"
____

**Библиотеки**: datetime, os, openpyxl, pandas

**Установка**:
1. Клонирование репозитория;

2. Создание и активация виртуального окружения:
```
MacOS:

python3 -m venv <venv_name>

source <venv_name>/bin/activate

python3 -m pip install --upgrade pip

Windows:

python -m venv <venv_name> 

source <venv_name>/Scripts/activate

python -m pip install --upgrade pip
```

3. Установить зависимости из файла requirements.txt:
```
pip install -r requirements.txt
```

4. Запустить код из папки, где находится изначальный файл для формирования отчёта.
