# Payment
## Задача
Обновить xlsx-файл так, чтобы сохранялось исходное форматирование: выделение цветом шрифта, заливка ячейки  и заметки. 

## Исходные данные
1. Все платежи на.xlsx - перечень всех платежей за предыдущий день (текущие платежи). Файл представляет собой перечень всех неоплаченных платежей с различными пометками. 
2. Выгрузка.xlsx - перечень всех неоплаченных заявок (файл с обновленными данными по текущим платежам, а также содержащий новые созданные платежи).
3. Самозанятые.xlsx - перечень контрагентов, имеющих статус самозанятых.

Файлы Все платежи на.xlsx и Выгрузка.xlsx представляет собой таблицу со следующими столбцами:
1. **№** - уникальный 6-значный индентификатор заявки
2. **Сделка** - описание закупки (заявки)
3. **Статус** - стадия согласования
4. **Сумма к оплате** - сумма, подлежащая оплате, в рублях
5. **Контрагент** - наименование контрагента
6. **Название бюджета** - наименование центра финансовой ответственности (ЦФО), в рамках которого необходимо осущетсвить закупку
7. **Номер счета** - номер документа, выставленного контрагентом для оплаты
8. **Дата создания** - дата формирования заявки
9. **Автор заявки** - ФИО инициатора
10. **Договор** - реквизиты договора, в рамках которого осуществляется закупка
11. **Проект** - номер и наименование проекта у ЦФО, в рамках которого будет осуществлен платеж
12. **Статья бюджета** - номер и наименование статьи, отражающей суть расхода
## Результат
Xlsx-файл с форматированием, аналогичным файлу Все платежи на.xlsx, содержащий обновленные данные по текущим платежам, а также новые созданные платежи из файла Выгрузка.xlsx.
В итоговом файле не должно быть платежей, предназначенных физическим лицам (кроме ИП и самозанятых), в т.ч. нерезидентам РФ, а также некоторые определенные контрагенты из списка.
## Техника решения
### 1. Отбор контрагентов
Для того, чтобы в перечень попали только юридические лица - резиденты РФ, а также ИП и самозанятые, был применен отбор.
```
//python code
pmnts = pmnts[(pmnts['Контрагент'].str.contains('ООО')) | 
              (pmnts['Контрагент'].str.contains('АО')) |
              (pmnts['Контрагент'].str.contains('"'))|
              (pmnts['Контрагент'].str.contains('Самозанят'))|
              (pmnts['Контрагент'].str.contains('самозанят'))|
              (pmnts['Контрагент'].str.contains('ИП'))|
              (pmnts['Контрагент'].isin(individ_grant))|
              (pmnts['Контрагент'].isin(self_empld))|
              (pmnts['Контрагент'].str[0].isin(list))|
              (pmnts['Контрагент'] == "")]
```
В этом отборе поиск контрагетов осуществляется по ключевым словам, характерным для юридических ЛИЦ:
1. Организационно-правовая форма:
   1. **"ООО"**
   2. **"АО"**
2. Наличие кавычек в названии
3. Наличие слов, характеризующих статус ИП и Самозанятого (в названии контрагентов ИП и самозанятых обязательно фигурирует наименование его статуса):
   1. **ИП**
   2. **Самозанят** (без окончаний)
   3. **самозанят** (без окончаний)
Однако некоторые юридические лица, чаще всего гос. учреждения, не имеют в своем наименовании вышеуказанных признаков. Ввиду того, что таких исключений немного добавлен список *individ_grant*. Для пополнения этого списка необходима первая ячейка. Запустив код, получаем список контрагентов, которые будут исключены. При необходимости добавляем контрагентов, которых не нужно исключать в список *individ_grant*.
Также был создан список из латинских прописных букв для фильтрации иностранных контрагентов
```
x = string.ascii_uppercase
list = [ch for ch in x]
```
При этом встречаются иностранные контрагенты - физические лица, наименования которых введены латинскими строными буквами. Таких контрагентов немного, поэтому создан отдельный список *list_foreign*
### 2. Удаление тестовых заявок.
В процессе эксплуатации файла возникла необходимость дополнительного отбора заявок, созданных тех. поддержкой при тестировании функционала. Были выявлены следующие схожие моменты для тестовых заявок
1. Наличие незаполненного поля контрагент, поэтому был добавлен отбор по контрагента
```
pmnts = pmnts[pmnts['Контрагент'] == "")]
```
2. Наличие в описании заявки слова **Удалить**
```
pmnts = pmnts[pmnts['Описание'] != 'Удалить']
```
3. Числа в поле ***Сумма к оплате***: **0.0, 1.0, 2.0, 3.0, 5.0, 10.0, 20.0**, а также наличие слова **Тестов** в поле ***Сделка***
```
pmnts = pmnts[(round(pmnts['Сумма к оплате'], 2) != 0.0)&
              (round(pmnts['Сумма к оплате'], 2) != 1.0)&
              (round(pmnts['Сумма к оплате'], 2) != 10.0)&
              (round(pmnts['Сумма к оплате'], 2) != 20.0)&
              (round(pmnts['Сумма к оплате'], 2) != 3.0)&
              (round(pmnts['Сумма к оплате'], 2) != 2.0)&
              (round(pmnts['Сумма к оплате'], 2) != 5.0)&
              (~pmnts['Сделка'].str.contains('Тестов'))]
```
### 3. Удаление оплаченных заявок.
1.Заявки, которые были оплачены будут присутствовать в файле Все платежи на.xlsx, но будут отсутствовать в файле Выгрузка.xlsx. Поэтому для удаления делаем *merge* двух файлов, отбирая платежи, которые есть в файле Все платежи на.xlsx, но которых нет в файле Выгрузка.xlsx (столбец "№" -  уникальный идентификатор заявки).
```
del_pmnts = curr_pmnts.merge(pmnts, how='outer', left_on='№',right_on='№')
del_pmnts = del_pmnts[(del_pmnts['Статус_y'].isna() == True) &
                      (del_pmnts['Статья бюджета_x'].isna() == False) & 
                      (~del_pmnts['Контрагент_x'].isin(cntr_except))]
```
2. Для работы с xlsx-файл его необходимо прочитать с помощью *load_workbook*, который необходимо импортировать  из *openpyxl*
```
wb = load_workbook(<path>)
ws = wb[<name_of_sheet>]
```
3. Для удаления отобранных заявок из xlsx-файла, необходимо создать список строк для удаления *cell.row*. Для прохождения по всем строкам xlsx-файла используется *iter_rows*, по всем строкам DataFrame - *iterrows()*
```
del_row = []
if del_pmnts.shape[0] != 0:
  for row in ws.iter_rows(min_row=2, max_col=1):
    for cell in row:
      for idx, r in del_pmnts.iterrows():
        if cell.value == del_pmnts.loc[idx]['№']:
          del_row.append(cell.row)
```
### 4. Обновление полей файла Все платежи на.xlsx.
Для обновления данных по заявкам как в предыдущем шаге делаем *merge*, отбирая поля, которые могли измениться из файла Выгрузка.xlsx.
```
curr_pmnts = curr_pmnts.merge(pmnts, left_on='№',right_on='№')
curr_pmnts = curr_pmnts[['№', 'Сделка_x', 'Статус_y','Сумма к оплате_y','Контрагент_y',
           'Название бюджета_y','Номер счета_y', 'Дата создания_y',
           'Автор заявки_y', 'Договор_y', 'Проект_y', 'Статья бюджета_y']]
```
### 5. Добавление новых платежей.
1. С помощью *merge*, аналогичного в предыдущем шаге, но, отбирая только строки из файла Выгрузка.xlsx, формируется DataFrame новых платежей *new_pmnts*
```
new_pmnts = curr_pmnts.merge(pmnts, how='outer', left_on='№',right_on='№')
new_pmnts = new_pmnts[new_pmnts['Сделка_x'].isna() == True]
new_pmnts = new_pmnts [['№', 'Сделка', 'Статус','Сумма к оплате','Контрагент',
           'Название бюджета','Номер счета', 'Дата создания',
           'Автор заявки', 'Договор', 'Проект', 'Статья бюджета']]
```
2. Для добавления новых платежей в xlsx-файл необходимо найти строку, с которой нужно добавлять новые строчки. Для этого находим номер строки *position* в файле Все платежи на.xlsx, который соответствует "№" из первой строки *curr_pmnts*.
```
for row in ws.iter_rows(min_row=2, max_col=1, max_row=curr_pmnts.shape[0]):
   for cell in row:
     if cell.value == curr_pmnts.iloc[0,0]:
       position = cell.row
       break
```
3. Обновление данных, содержащихся в файле Все платежи на.xlsx в соответствие с обновленным датафреймом *curr_pmnts*, осуществляется присвоением ячейки из xlsx-файла значения, соответствующего значению на пересечении *row*, *col* из *curr_pmnts*. При этом индекс строки ячейки сдвигается на *position* для определения корректной начальной строки, в которую необходимо внести изменения. В столбце **Дата создания** необходимо отдельно задать формат ячейки как даты, в связи с чем выделено отдельно условием *if col == 7:*
```
for col in range(curr_pmnts.shape[1]):
  for row in range(curr_pmnts.shape[0]):
    if col == 7:
      value = curr_pmnts.iloc[row, col]
      cell = ws.cell(position+row, col+1)
      cell.value = value
      cell.number_format = 'DD.MM.YYYY'
    else:
      ws.cell(position+row, col+1).value = curr_pmnts.iloc[row, col]  
```
5. Вставка новых платежей реализована через функцию *def ins_row (a_pmnts, numb)*, где *a_pmnts* - датафрейм с новыми платежами,  *numb* - номер строки, с которой необходимо вставлять новые строки. Реализация функции аналогична предыдущему шагу с отличием
i. в добавлении строки при помощи *insert_rows*
```
ws.insert_rows(numb, a_pmnts.shape[0])
```
ii. в смещении строки ячейки для вставления на *numb*
```
ws.cell(numb+row, col+1).value = a_pmnts.iloc[row, col]
```
iii. в привидении новых строк столбца **Сумма к оплате** к числовому виду с разделителями и 2 знаками после запятой
```
if col == 3:
          ws.cell(numb+row, col+1).number_format = '# ### 000.00'
```
### 6. Добавление гиперссылок в столбец **№**
Все заявки имеют гиперссылки https://portal.university.innopolis.ru/processes/list/111/element/0/{}/, {} - номер из столбца **№**. В каждой строке с помощью *iter_rows* проставляются гиперссылки.
```
url = "https://portal.university.innopolis.ru/processes/list/111/element/0/{}/"
for row in ws.iter_rows(min_row=2, max_col=1):
  for cell in row:
    value = cell.value
    cell.value = '=HYPERLINK("%s", "%s")' % (url.format(value), value)
```
### 7. Сохранение файла
Для сохранения файла используется *wb.save*.
