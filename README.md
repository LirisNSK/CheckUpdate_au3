# Скрипт CheckUpdate

## Описание
Здесь представлен скрипт, упрощающий работу с РБД 1С:Предприятие 8.
Код скрипта написан на языке [AutoIT 3](https://www.autoitscript.com/site/) 

## Использование
Перед началом использования установить параметры подключения к ИБ и пути к исполняемым файлам.

### Описание параметров
```bsl
LogfileName - Строка (без кавычек). Желаемое имя для файла-журнала
IBConn      - Строка (без кавычек). Строка подключения к ИБ вида Srvr=ИмяСервера;Ref=ИмяИБ;Usr=Пользователь;Pwd=ПарольПользователя; 
V8exePath   - Строка (в кавычках). Указывает путь к исполняемому файлу 1С:Предприятие 8
ComConnectorObj - Строка (без кавычек). Имя COM-объекта, например: V83.COMConnector
ArcLogfileName  - Строка (без кавычек). Имя файла, в который будет записан текущий журнал, при достижении лимита по размеру. Например: 512000
LogMaxSize  - Число. Максимальный размер файла журнала в байтах. При достижении указанного размера, файл журнала будет отправлен в архив с именем ArcLogfileName
PIDFileExt  - Строка (без кавычек). Указывает расширение для служебных файлов (pid)
EmailTo     - Строка (без кавычек). Содержит список получателей письма-уведомления
EmailCc     - Строка (без кавычек). Содержит список получателей "Копия" письма-уведомления
```
