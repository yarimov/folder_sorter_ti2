## Сортировщик папок - Титан2

Программа для персонального компьютера, позволяющая производить подготовку проектной документации к выгрузке в архив CRM
системы проектной организации. 

[Скачать](https://github.com/yarimov/folder_sorter_ti2/releases).  

### Описание алгоритма работы программы
В первое текстовое поле необходимо скопировать путь к проекту (или выбрать его через контекстное меню "Обзор"), который
необходимо подготовить к выгрузке в архив CRM системы проектной организации.

Нажать кнопку "Проверить". Выполнится проверка наличия PDF файлов, к рабочим файлам проекта (все редактируемые форматы).

В случае удачной проверки в информационном поле отобразится результат проверки. Напротив имен файлов отобразится
статус [Ок], означающий, наличие PDF файла к данному документу или в противном случае - [None].

Внешний вид программы. Результат проверки
![Внешний вид программы. Результат проверки](https://github.com/yarimov/folder_sorter_ti2/blob/master/img/folder_sorter_ti2.png)
Программа проверяет правильность введенного пути к проекту, при недействительном пути к папке отображаются
предупреждения о некорректности ссылки.
Если отчет проверки содержит статусы [None], необходимо проверить файлы проекта при необходимости провести повторную
проверку. 

Второй этап работы программы - "Сортировка".

Для ее выполнения необходимо указать путь сохранения отсортированных файлов.
Если путь оставить не заполненным, то в исходной папке с проектом создастся подкаталог с названием исходной папки и в
него будут скопированы отсортированные файлы. Если будет указан путь, то в нем создастся папка с именем исходной папки
проекта с отсортированными файлами.

Существует дополнительная возможность указать путь для сохранения по умолчанию. Для этого предварительно в папке
программы "Сортировщик папок - Титан2" необходимо заполнить путь по умолчанию в файле "config.yml". После чего при
нажатии чекбокса путь по умолчанию будет браться из "config.yml".

Результат сортировки.
![Результат сортировки](https://github.com/yarimov/folder_sorter_ti2/blob/master/img/folder_sorter_ti2_2.png)
Дополнительные настройки по сортировке описаны в разделе настройка [config-файла](#Настройки-config-файла).

### Алгоритма сортировки
'folder_from' - это папка текущего проекта. Обычно называется по шифру проекта, например "AKU.0179.40UQR.CYP.SS.MB0001_С01"

'folder_to' - это папка куда будет сохранен отсортированный проект.

В папке 'folder_to' создается папка с именем 'folder_from'. В ней создаются папки 'WORD' 'EXCEL', 'DWG', 'PDF',
соответствующие форматам файлов ('.doc', '.docx', '.dot', '.dotx'), ('.xlsx', '.xlsm'), ('.dwg') и ('pdf') соответственно.
Все файлы сортируются в папки соответствующим по формату, кроме PDF файлов содержащих в названии '-UL', '_UL', '
-RVI', '_RVI'. Они сохраняются в коревой папке 'folder_from'.

По умолчанию файлы начинающиеся, на "!" или "~" считаются служебными и исключены из сортируемых файлов. Это может быть
нужно для хранения общего редактируемого DWG файла (со всеми листами) до разбивки по листам. Для унификации предлагается
называть его в формате: "!AKU.0179.40UQR.CYP.SS.TB0001_С01-общий.dwg".

### Настройки config-файла
В конфиг файле - "config.yml" предусмотрены следующие настройки:

Можно указать путь к папке для сохранения по умолчанию. Не используется при не нажатом чекбоксе "Путь для сохранения по умолчанию"
```
# Укажите папку сохранения по умолчанию
"DEFAULT_FOLDER": 'C:\documents'
```

Предусмотрена возможность задать символы для служебных файлов, которые будет исключены для сортировки. 
Если предполагается использовать все файлы, можно написать "IGNORED_IN_FILENAME": ('*', )
```
# Первые символы в названии файлов, которые будут исключены из проверок и сортировок (Служебные файлы)
"IGNORED_IN_FILENAME": ('~', '!')
```

При добавлении новых форматов в проектной документации можно продолжить запись сопоставления форматов файлов соответствующим каталогам.
Например, 'txt': 'TXT'
Важно сохранять в файле конфигурации исходную табуляцию, т.е. количество пробелов в записи ниже должно быть как у строки выше.
```
# Указать форматы файлов и соответствующие им наименования папок для их последующей сортировки
FORMAT_DIRECTORY:
  'doc': 'WORD'
  'docx': 'WORD'
  'dot': 'WORD'
  'dotx': 'WORD'
  'xlsx': 'EXCEL'
  'xlsm': 'EXCEL'
  'dwg': 'DWG'
  'pdf': 'PDF'
```

### Особенности
- Программа не требует "прав администратора" для запуска;
- Работает по принципу "portable" - все необходимые компоненты упакованы в EXE-исполняемом файле;
- Предусмотрено автоматическое переключение на темную и светлую тему оформления, в зависимости от системной темы компьютера.
