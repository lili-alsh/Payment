# Payment
## Description
To update xlsx-file with current formatting: cell filling, highlighting and notes. 
Output  - xlsx-file that have formating exactly like as in the file Все платежи на.xlsx. This file also consists of current payments' update items (data in columns) and new payments (newly created) exactly like as in the file Выгрузка.xlsx. 
In update file there should be no payments for individuals (except of self-employed and indidual enterpreuners) including non-residents and for some contarctors from exception list.
В результате должен быть xlsx-файл с форматированием, аналогичным файлу Все платежи на.xlsx, содержащий обновленные данные по текущим платежам, а также новые созданные платежи из файла Выгрузка.xlsx.
## Input data
1. Все платежи на.xlsx - list of all previoys payments (yesterday payments) with cell filling, highlighting and notes. This formatting should be saved in the output file.
2. Выгрузка.xlsx - list of all unpaid invoices (file consists of current payments with update data and new payments (today payments)). These update details of payments should be saved in the output file. Newly created payments should be displayed in the output file.
3. Самозанятые.xlsx - list of self-employed contrastors.

Files Все платежи на.xlsx and Выгрузка.xlsx are table that cocnsists of the columns: 
1. **№** - unique 6-number id of payment request
2. **Сделка** - description of payment request
3. **Статус** - stage of coordination
4. **Сумма к оплате** - summ of payment
5. **Контрагент** - contractor's name
6. **Название бюджета** - name of responsibility centre that is required this payment
7. **Номер счета** - number of invoice
8. **Дата создания** - date of payment request's creation
9. **Автор заявки** - full name of payment request's author
10. **Договор** - number, date and name of contract that is signed for doing this payment
11. **Проект** - number and name of project that have need for this payment
12. **Статья бюджета** - number and name of objects of expenditure

### 1. Contractors' selection
The filter to select legal entities, individual entrepreneurs and selg-employed person
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
There are keywords for selection:
1.  The legal form:
   1. **"ООО"**
   2. **"АО"**
2. The presence of quotation mark in contractor's name
3. The presence of words that are characteristic for individual entrepreneurs and selg-employed person in contractor's name (these words are storng required in name by organistions' rules):
   1. **ИП**
   2. **Самозанят** (without ending of the word)
   3. **самозанят** (without ending of the word)
Some companies, often government companies, don't have the features above. Those companies are not so much. That's why there is exception list that is named legal entities. 
This list can be extended. It is doing by starting first block of the code. The result of it is the list of contractors that wouldn't be selected. If there will be legal entities in list they should be added in *individ_grant*.
To select non-resinet's contractots the list of latin capital letters is created.
```
x = string.ascii_uppercase
list = [ch for ch in x]
```
To except non resident's individuals (their names are also consist of capital letters) the list of their names (*list_foreign*) is created.
### 2. Deleting of tests' requests of payment.
To except test's requests of payment that's formed by technical support it is needed to set additional filter. There are the same features for these tests' requests of payment:
1. Empty field **Контрагент**, that's why this filter is used
```
pmnts = pmnts[pmnts['Контрагент'] == "")]
```
2. The presence of the word **Удалить** in the description of tests' requests of payment
```
pmnts = pmnts[pmnts['Описание'] != 'Удалить']
```
3. Thera are numbers: **0.0, 1.0, 2.0, 3.0, 5.0, 10.0, 20.0** in the field **Сумма к оплате**, also the presence of the word **Тестов** in the field **Сделка**
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
### 3. Deleting paid requests of payment.
1. Paid requests of payment will be in the file Все платежи на.xlsx, but it won't be  in the file Выгрузка.xlsx. To delete paid requests of payment both files (Все платежи на.xlsx and Выгрузка.xlsx) are merged. Merge it selecting payments that are presented in the file Все платежи на.xlsx, but are ubsented in the file Выгрузка.xlsx.
```
del_pmnts = curr_pmnts.merge(pmnts, how='outer', left_on='№',right_on='№')
del_pmnts = del_pmnts[(del_pmnts['Статус_y'].isna() == True) &
                      (del_pmnts['Статья бюджета_x'].isna() == False) & 
                      (~del_pmnts['Контрагент_x'].isin(cntr_except))]
```
2. To read xlsx-file for saving current formatting it's used *load_workbook* that is imported from *openpyxl*.
```
wb = load_workbook(<path>)
ws = wb[<name_of_sheet>]
```
3. To delete selected payments list *cell.row* of deleting rows is created. *iter_rows* is needed for step-by-step iteration into xlsx-file, *iterrows()* is used into DataFrame.
```
del_row = []
if del_pmnts.shape[0] != 0:
  for row in ws.iter_rows(min_row=2, max_col=1):
    for cell in row:
      for idx, r in del_pmnts.iterrows():
        if cell.value == del_pmnts.loc[idx]['№']:
          del_row.append(cell.row)
```
### 4. Updating the fields of Все платежи на.xlsx.
To update the requests' of payment data merge 2 files (Все платежи на.xlsx and Выгрузка.xlsx) selecting columns from Выгрузка.xlsx (with *_y* suffix). All updating data are into the Выгрузка.xlsx, that's why columns from this file are selected only.
```
curr_pmnts = curr_pmnts.merge(pmnts, left_on='№',right_on='№')
curr_pmnts = curr_pmnts[['№', 'Сделка_x', 'Статус_y','Сумма к оплате_y','Контрагент_y',
           'Название бюджета_y','Номер счета_y', 'Дата создания_y',
           'Автор заявки_y', 'Договор_y', 'Проект_y', 'Статья бюджета_y']]
```
### 5. Adding new requests of payment.
1. If rows from file Выгрузка.xlsx (mask *new_pmnts['Сделка_x'].isna()== True* is used for it) will be selected only the dataframe of newly created payments is formed. 
```
new_pmnts = curr_pmnts.merge(pmnts, how='outer', left_on='№',right_on='№')
new_pmnts = new_pmnts[new_pmnts['Сделка_x'].isna() == True]
new_pmnts = new_pmnts [['№', 'Сделка', 'Статус','Сумма к оплате','Контрагент',
           'Название бюджета','Номер счета', 'Дата создания',
           'Автор заявки', 'Договор', 'Проект', 'Статья бюджета']]
```
2. To add new payments it's needed to find position. It's a number of row into the file Все платежи на.xlsx for adding new paymnets. This position is equal to value from a intersection of the first row and column **№** in the dataframe *curr_pmnts*.
```
for row in ws.iter_rows(min_row=2, max_col=1, max_row=curr_pmnts.shape[0]):
   for cell in row:
     if cell.value == curr_pmnts.iloc[0,0]:
       position = cell.row
       break
```
3. To update the data into the xlsx-file every cell is filled by the value from a intersection  of the row and the column from dataframe *curr_pmnts*. To insert right value into the xlsx-file index of the starting row is offseted on *position*. The cells in column **Дата создания** must have date format that's why there is condition *if col == 7:*.
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
5. The inserting of new requests of paymnet is doing by function *def ins_row (a_pmnts, numb)*, where *a_pmnts* - the updating dataframe,  *numb* - the number of row for adding new payments. The function is similar to the code in the previous step. But there are some differncies: 
i. *insert_rows* is used to add new empty rows into the xlsx-file
```
ws.insert_rows(numb, a_pmnts.shape[0])
```
ii. *numb* is the same as the *position*
```
ws.cell(numb+row, col+1).value = a_pmnts.iloc[row, col]
```
iii. the certain format of the cells of column **Сумма к оплате** (numeric form with separators and 2 decimal places)
```
if col == 3:
          ws.cell(numb+row, col+1).number_format = '# ### 000.00'
```
### 6. Adding hyperlinks into column **№**
All requests of paymnets have the hyperlinks with the body https://portal.university.innopolis.ru/processes/list/111/element/0/{}/, where {} - the value from column **№**. The hyperlinks are filled in the xlsx-file with *iter_rows*.
```
url = "https://portal.university.innopolis.ru/processes/list/111/element/0/{}/"
for row in ws.iter_rows(min_row=2, max_col=1):
  for cell in row:
    value = cell.value
    cell.value = '=HYPERLINK("%s", "%s")' % (url.format(value), value)
```
### 7. Saving the xlsx-file
*wb.save* is used to save the changes into the xlsx-file.
