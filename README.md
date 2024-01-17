# Taxapp

Веб-приложение на Django, с помощью которого пользователь может отправить Excel-файл с исходными данными и в ответ получить сформированный отчет, основанный на исходных данных.

## Установка

### Запуск приложения (на Windows)

- Клонировать репозиторий

```
git clone https://github.com/edgar1148/django-excel.git
```

- Установить и активировать виртуальное окружение

```
python -m venv venv
```

```
source venv/Scripts/activate
```

- Установить зависимости requirements.txt

```
pip install -r requirements.txt
```

- Запустить docker-compose

```
docker compose up
```

- Или зайти в директорию taxapp с файлом manage.py и выполнить команду

```
python manage.py runserver
```


### Возможности приложения:

#### Написано в соответствии с ТЗ из файла (example_data/test_task_text.md)

- После запуска сервера переходим по адресу 

```
http://0.0.0.0:8000/files/upload

```

- Нажимаем кнопку 'Загрузить файл'

```
Прикрепляем excel-файл с тестовыми данными (example_data/example_data.xlsx)

```

- Нажимаем кнопку 'Обработать'

```
Загружаем готовый .xlsx файл в нужную нам директорию

```


#### Технологии
- python==3.8.7
- Django==4.2.9
- openpyxl==3.1.2
- pandas==2.0.3

#### Автор
[Евгений Екишев - edgar1148](https://github.com/edgar1148)
