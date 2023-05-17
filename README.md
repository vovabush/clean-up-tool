# Что это
Утилита для очистки следов работы с Офисными файлами. Более подробное описание функционала дано ниже.

# Как запускать
`python clean.py` или `clean.exe`

# Функционал:
1. Для windows:
	1. Удаляет в Офисе историю открытых документов;
	2. Удаляет в истории Последних файлах все документы;
	3. Удаляет в меню пуск в приложениях Офиса историю последних открытых документов;
	4. Чистит реестр для Офисных приложений;
	5. Удаляет несохраненные файлы Офиса по пути `C:\Users\mVirt\AppData\Local\Microsoft\Office\UnsavedFiles`
	6. Удаляет список последних открытых файлов из того что кеширует Windows: `%APPDATA%\Microsoft\Windows\Recent\AutomaticDestinations`
	7. Осуществляет полную Очистку диска.
2. Для Linux: пока ничего нет.

# Можно добавить:
1. Для windows:
	1. Проверять статус гибернации, отключать и если надо включать гибернацию (чтобы дампы не были на жестких дисках).
	2. Искать и выводить пользователю на жестких дисках файлы: 
		- \*.doc*
		- \*.ppt*
		- \*.xls*
2. Для Linux:
	1. Искать и выводить пользователю на жестких дисках файлы: 
		- \*.doc*
		- \*.ppt*
		- \*.xls*
	2. Очищать разные метаданные.
