# GetMoexData
get data from moex.com using selenium


![Status](https://github.com/elegantovich/GetMoexData/actions/workflows/main.yml/badge.svg)
## Description
Скрипт для обработки данных с интерфейса сервиса "Московская биржа".

### Tech
Python 3.10, Selenium 4.3, Pandas 1.3, Numpy 1.2, Lxml 4.9


### How to start a project:
Тестирование производилось на браузере Google chrome. Для старта необходимо узнать версию браузера.
```
chrome://settings/help
```
Скачайте и положите в папку с приложением драйвер, с версией поддерживаемой вашим браузером:
```
https://sites.google.com/chromium.org/driver/
```


Clone and move to local repository:
```
git clone https://github.com/Elegantovich/GetMoexData/
```
Create a virtual environment (win):
```
python -m venv venv
```
Activate a virtual environment:
```
source venv/Scripts/activate
```
Install dependencies from file requirements.txt:
```
python -m pip install --upgrade pip
```
```
pip install -r requirements.txt
```
Run the script
```
python main.py
```

## Checklist
ID| Option| Status |
| ------ | ------ | ------ |
| 1 | Открыть https://www.moex.com | Done |  |
| 2 | Перейти по следующим элементам: Меню -> Срочный рынок -> Индикативные курсы; | Done |
| 3 | В выпадающем списке выбрать валюты: USD/RUB | Done |
| 4 | Сформировать данные за предыдущий месяц | Done |
| 5 | Скопировать данные в Excel | Done |
| 6 | Повторить шаги для валют: JPY/RUB | Done |
| 7 | Скопировать данные в Excel | Done |  |
| 8 | Для каждой строки полученного файла поделить курс USD/RUB на JPY/RUB, полученное значение записать в ячейку (G) Результат | Done |
| 9 | Выровнять – автоширина | Done |
| 10 | Формат чисел – финансовый | Done |  
| 11 | Проверить, что автосумма в Excel распознаёт ячейки как числовой формат | Done |
| 12 | Направить итоговый файл отчета на почту Mik***v@Gr****.ru | Failed |
| 13 | В письме указать количество строк в Excel в правильном склонении | Done |


## Notes
- Драйвер должен находиться в папке с проектом.
- Результаты работы в виде ексель документа 'currency.xlsx'.
- По 12 пункту возникает ошибка. Пробовал неоднократно.

