# Получение внешних переменных 3D модели | api "Компас 3D"
##### _Возвращает таблицу переменных в файл parametrsDetails.xlsx_ 

### Настройка

| Файл | Назначение |
| ------ | ------ |
| main.py -> class Config() ->  dir | указать путь до папки с проектом |

### Алгоритм работы:
    1) Для работы программы необходимо в классе Config в переменную dir установить путь рабочей папки, в рабочую папку скопировать файл parametrsDetails.xlsx.
    2) Отредактировать dataFrame = pd.DataFrame(documents, columns=['D','d','h','g','lg','path']). Указать количество и наименование колонок соответственно Вашему проекту.
    3) При запуске программа перебет файлы рабочей папки, выберет ".m3d"(".m3d.bak" будут проигнорированы). В конечном файле parametrsDetails.xlsx сформирует таблицу параметров детали.
### Файлы для демонтсрации:
[Тестовый набор файлов](https://disk.yandex.ru/d/sbaW01LM5huC2w)



