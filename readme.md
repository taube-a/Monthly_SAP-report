# Автоматизия составления отчёта
____

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
Пример ежемесячного отчёта - "Список проблем SAP на 28.03.2022 - SAP issues list on 28.03.2022"

____

ЯП реализации автоматизации: Python

Библиотеки: openpyxl, pandas, datetime, os, re
