# Payment
## Description
Update xlsx-file with current formatting: cell filling, highlighting and notes. 
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
