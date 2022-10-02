# Course-project
## Telegram Bot Details_order
Этот бот создан для больших фабрик или небольших производств где используется промышленное оборудование для легкой промышленности. Персонал обслуживающий оборудование имеет на каждую модель Parts List запчастей, откуда и может взять номер детали.
Он удобен для быстрого нахождения информации по номеру детали, есть ли на складе в регионе, а также заказа этой детали у продавца. Так как компьютер или сервер может работать 24/7 можно в реальном времени на производственной смене после рабочих часов фирмы продавца, не тревожа менеджеров по телефону, узнать всю информацию и заказать сломанную деталь. 

## Создание бота
Создание бота происходит через специального бота **BotFather**. Когда вы создадите бота, **BotFather** даст вам его токен. Токен выглядит примерно  так: ***110201543:AAHdqTcvCH1vGWJxfSeofSAs0K5PALDsaw***. Именно с помощью токена вы сможете управлять ботом.
Имя бота выглядит как обычный юзернейм, но он должен заканчиваться на **"bot"**.
Для работы Вашего кода также необходимо импортировать некоторые библиотеки:
-	import **telebot**   # библиотека для разработки **telegram-ботов**
-	from **telebot** import **types**   # библиотека для создания кнопок
-	import **pandas**  # высокоуровневая библиотека для анализа данных
-	from **openpyxl** import **load_workbook**   # библиотека для работы с файлами Excel
-	import **requests**   # библиотека запросов
-	from **bs4** import **BeautifulSoup**   # библиотека запросов для извлечения данных из файлов HTML и XML
-	from **copy** import **copy**   # метод поверхностное и глубокое копирование объектов
-	from **time** import **sleep**   # для симуляции задержки в выполнении программы

Для более подробной информации по версиям библиотек можно увидеть в файле ***requirements.txt*** загруженном в этом же проекте.
## Описание работы бота

При старте бота Вы получаете информацию, которая помогает понять, какую информацию нужно ввести.
Необходимо ввести номер детали для Вашей модели оборудования. Пример как пользоваться
Бот находит деталь в файле **склад.xlsx** и выдает Вам информацию о названии детали, что она существует в базе данных. Дальше, предоставляется возможность клиенту проверить наличие по складу, цену, заказать нужное ему количество сообщив свой емайл для обратной связи. Для удобства созданы кнопки с наименованием запроса. Сам код записывает заказ в файл **Заказы.xlsx** и менеджеры фирмы продавца могут извлекать запросы в любое время.
Этот бот можно дорабатывать на свои усмотрения, делать дополнительные кнопки по локации, где можно забрать товар и так далее.
