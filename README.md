# Scraping

Консольная утилита для скачивания и обработки данных с tektorg.ru

# Использование
1) Сделать поисковый запрос здесь https://www.tektorg.ru/procedures , дождаться его выполнения

2) Скопировать URL (содержимое адресной строки браузера)

3) В консоли зупустить скрипт, предварительно заменив путь и URL на корректные:
python /PATH/TO/scraping.py https://www.tektorg.ru/procedures?q=%D0%A3%D0%B7%D0%B5%D0%BB+%D1%83%D1%87%D0%B5%D1%82%D0%B0+%D0%BD%D0%B5%D1%84%D1%82%D0%B8

4) Данные появятся в папке workdir, путь к которой можно изменить в settings.json


# Установка
1) python3 https://www.python.org/downloads/

2) pip https://www.liquidweb.com/kb/install-pip-windows/

3) git https://git-scm.com/book/en/v2/Getting-Started-Installing-Git

4) tesseract https://tesseract-ocr.github.io/tessdoc/Home.html

5) Открыть консоль, перейти в папку, куда нужно скачать проект, например, так:

cd ~\Documents

6) Скачать проект:

git clone https://github.com/seregakol007/scraping.git

7) При необходимости изменить путь к tesseract в settings.py

8) Перейти в папку проекта:

cd scraping

9) Выполнить тестовый запрос

python scraping.py https://www.tektorg.ru/procedures?q=%D0%A3%D0%B7%D0%B5%D0%BB+%D1%83%D1%87%D0%B5%D1%82%D0%B0+%D0%BD%D0%B5%D1%84%D1%82%D0%B8
