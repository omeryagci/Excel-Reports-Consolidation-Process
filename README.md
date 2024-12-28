# EXCEL REPORTS CONSOLIDATION PROCESS

## İçindekiler - Contents

1. [TR](#tr)
   - [Genel Açıklama](#genel-açıklama)
   - [Adım Adım Süreç](#adım-adım-süreç)
   - [Robot Yolu](#robot-yolu)
   - ["Teams" Excel Örneği](#teams-excel-örneği)
   - [Çıktı Excel Örneği](#çıktı-excel-örneği)
2. [ENG](#eng)
   - [General Description](#general-description)
   - [Step-by-Step Process](#step-by-step-process)
   - [Robot Path](#robot-path)
   - ["Teams" Excel Example](#teams-excel-example)
   - [Output Excel Example](#output-excel-example)

## [TR]

### Genel Açıklama

Merhabalar,

Excel Reports Consolidation süreci için dinamik bir şekilde tarihe göre, verilerin yazdırılacağı "Sales Results" exceli oluşturulur.
Sonrasında "Item", "Unit Price ($)", "Items Sold" ve "Total Income" başlıkları yazdırılır. Başlık oldukları için "Font Style", "Font Size"
ve "Fill Color" ayarları yapılır. İşlem yapacağımız exceller için bir For Each oluşturulur. İşlem yapılan excellerin sheet isimleri 
kontrol edilir. Eğer uygun isimde değilse, işlem yapacağımız türde sheet ismine çevrilir. İşlem yapılacak excellerden alınan veriler
"Sales Results" exceline yazdırılır. Yazdırılan veriler için Autofit işlemi gerçekleştirilir. "Sales Results" exceli için yeni bir sheet
oluşturulur. Yeni sheet içerisinde pivot table oluşturulur. "Item" ve "Total Income" kısımları eklenir ve toplam işlemi yapılır. Yeni sheet 
içerisinde ekstra olarak pasta grafiği eklenir.

Teşekkürler.

### Adım Adım Süreç

1. Çıktı olarak alınacak olan excel varsa silinir.
2. İşlem yapılan tarihe göre "Sales Reports" exceli oluşturulur.
3. "Item", "Unit Price ($)", "Items Sold" ve "Total Income" başlıkları oluşturulur.
4. "Font Style", "Font Size" ve "Fill Color" ayarları yapılır.
5. İşlem yapılacak exceller için For Each oluşturulur.
6. İşlem yapılacak olan excel sheet'inin isminin uygun olmaması durumu göze alınarak if-else bloğu oluşturulur ve eğer isim uygun değilse, uygun hale getirilir.
7. Alınan veriler "Sales Results" exceline yazdırılır. Bu işlemler işlem yapılacak excel sayısı kadar tekrarlanır.
9. Tüm exceller için işlem gerçekleştikten sonra Autofit işlemi uygulanır.
10. "Sales Results" excelinde yeni bir sheet açılarak "Item" ve "Total Income" verileri ile pivot table oluşturulur.
11. Son olarak tüm veriler ile bir pasta grafiği excele eklenir.

### Robot Yolu

![image](https://github.com/user-attachments/assets/87f4b227-ee6c-448b-9794-3923f6b56740)

### "Coffee Shop" Excel Örneği

![image](https://github.com/user-attachments/assets/d80621b5-8b93-4fc9-8ca8-8d63da915bb8)

### "Sales Results" Excel Örneği

![image](https://github.com/user-attachments/assets/d1c7cf98-72fd-46c2-bb52-7c0707f567a7)

![image](https://github.com/user-attachments/assets/c65c09af-2a1b-4ce3-821f-bdb3910e4720)

## [ENG]

### General Description

Hello,

For the Excel Reports Consolidation process, a "Sales Results" Excel file is dynamically created based on the date for printing data. 
Subsequently, the headers "Item," "Unit Price ($)," "Items Sold," and "Total Income" are printed. Since these are headers, their "Font Style," "Font Size," and "Fill Color" settings are configured. 
A For Each loop is created for the Excel files to be processed. The sheet names of the processed Excel files are checked, and if they are not appropriately named, they are converted to a suitable sheet name. 
The data taken from the Excel files to be processed is written to the "Sales Results" Excel file. Autofit is applied to the written data. A new sheet is created for the "Sales Results" Excel file. 
In the new sheet, a pivot table is created. The "Item" and "Total Income" sections are added, and the total is calculated. Additionally, a pie chart is inserted into the new sheet.

Thanks.

### Step-by-Step Process

1. If there is an existing Excel file to be generated as output, it is deleted.
2. A "Sales Reports" Excel file is created based on the date of the operation.
3. Headers "Item," "Unit Price ($)," "Items Sold," and "Total Income" are created.
4. "Font Style," "Font Size," and "Fill Color" settings are configured.
5. A For Each loop is created for the Excel files to be processed.
6. An if-else block is constructed to handle cases where the sheet name of the Excel file being processed is not appropriate, and if the name is unsuitable, it is adjusted to be appropriate.
7. The extracted data is written to the "Sales Results" Excel file. These operations are repeated for the number of Excel files being processed.
8. After processing all Excel files, Autofit is applied.
9. A new sheet is opened in the "Sales Results" Excel file, and a pivot table is created using the "Item" and "Total Income" data.
10. Finally, a pie chart is added to the Excel file with all the data.

### Robot Path

![image](https://github.com/user-attachments/assets/d7658dea-dbdf-4024-825c-2948bfadb41b)

### "Coffee Shop" Excel Example

![image](https://github.com/user-attachments/assets/d80621b5-8b93-4fc9-8ca8-8d63da915bb8)

### "Sales Results" Excel Example

![image](https://github.com/user-attachments/assets/d1c7cf98-72fd-46c2-bb52-7c0707f567a7)

![image](https://github.com/user-attachments/assets/c65c09af-2a1b-4ce3-821f-bdb3910e4720)
