Dim startDate
Dim endDate

startDate = CDate("20/08/2011")
endDate = CDate("31/12/2011")

Const ForWriting = 2
Const fileName = ".\test_Invoice.csv"

Dim fso, f1, ts
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile (fileName)
Set f1 = fso.GetFile(fileName)
Set ts = f1.OpenAsTextStream(ForWriting, TristateFalse)

ts.WriteLine "Дата;Номер;Проведен;Вход. номер;Вход. дата;Поставщик (ИНН);Поставщик (код);" & _
        "Поставщик (наименование);Сотрудник (код);Сотрудник (наименование);Склад(код);" & _
        "Склад (наименование);Счет-фактура;Номенклатура (код);Номенклатура (наименование);" & _
        "Группа учета (название);Группа учета (код);Количество;Цена с НДС;Сумма с НДС;Сумма НДС;Ставка НДС;" & _
        "Торговое предприятие(код);Торговое предприятие (наименование);ЮЛ (ИНН);ЮЛ (наименование);" & _
        "Тип номенклатуры (код);Тип номенклатуры (название);Единица измерения (код);Единица измерения (название);" & _
        "Концепция (название);Концепция (код);" & _
        "Номенклатура поставщика (код);Номенклатура поставщика (название);Тара (код);Тара (название);Количество в таре;" & _
        "Склад поставщик (наименование);Склад поставщик (код);Комментарий"

Public Sub RunTest()

Dim loader
Dim oInvoce
Dim oInvoces

Set loader = CreateObject("iiko1CInterface.DocumentLoader")

Dim filter
Set filter = CreateObject("iiko1CInterface.DocumentFilter")
filter.DateFrom = startDate
filter.DateTo = endDate
'Set filter.IncludeUnprocessed = true
'Set filter.UseBusinessSettings = true
'filter.AddDepartmentId("1")
Set oInvoces = loader.LoadInvoiceData2("http://localhost:8080/resto/", "admin", "resto#test", filter)

Set oInvoce = oInvoces.GetFirstInvoiceData()

Do While Not oInvoce Is Nothing
    PrintInvoice(oInvoce)
    Set oInvoce = oInvoces.GetNextInvoiceData()   
Loop

MsgBox("Выгрузка завершена!")

End Sub

Public Sub PrintInvoice(invoice)

Set oInvoiceItem = invoice.GetFirstInvoiceItem()
Do While Not oInvoiceItem Is Nothing
    PrintInvoiceItem invoice, oInvoiceItem
    Set oInvoiceItem = invoice.GetNextInvoiceItem()   
Loop
	
End Sub


Public Sub PrintInvoiceItem(document, documentItem)

ts.Write document.Date
ts.Write ";"
ts.Write document.Number
ts.Write ";"
ts.Write document.Processed
ts.Write ";"
ts.Write document.IncomingNumber
ts.Write ";"
ts.Write document.IncomingDate_DateTimeAsString
ts.Write ";"
ts.Write document.SupplierINN
ts.Write ";"
ts.Write document.SupplierCode
ts.Write ";"
ts.Write document.SupplierName
ts.Write ";"
ts.Write document.StaffCode
ts.Write ";"
ts.Write document.StaffName
ts.Write ";"
ts.Write documentItem.StoreCode
ts.Write ";"
ts.Write documentItem.StoreName
ts.Write ";"
ts.Write document.Invoice
ts.Write ";"
ts.Write documentItem.Article
ts.Write ";"
ts.Write documentItem.Nomenclature
ts.Write ";"
ts.Write documentItem.AccountingCategory
ts.Write ";"
ts.Write documentItem.AccountingCategoryCode
ts.Write ";"
ts.Write documentItem.Amount_DecimalAsString
ts.Write ";"
ts.Write documentItem.PriceWithNds_DecimalAsString
ts.Write ";"
ts.Write documentItem.SumWithNds_DecimalAsString
ts.Write ";"
ts.Write documentItem.Nds_DecimalAsString
ts.Write ";"
ts.Write documentItem.NdsPercent_DecimalAsString
ts.Write ";"
ts.Write document.DepartmentCode
ts.Write ";"
ts.Write document.DepartmentName
ts.Write ";"
ts.Write document.JuristicPersonINN
ts.Write ";"
ts.Write document.JuristicPersonName
ts.Write ";"
ts.Write documentItem.NomenclatureType
ts.Write ";"
ts.Write documentItem.NomenclatureTypeName
ts.Write ";"
ts.Write documentItem.MeasureUnitCode
ts.Write ";"
ts.Write documentItem.MeasureUnitName
ts.Write ";"
ts.Write document.ConceptionName
ts.Write ";"
ts.Write document.ConceptionCode
ts.Write ";"
ts.Write documentItem.OuterProductCode
ts.Write ";"
ts.Write documentItem.OuterProductName
ts.Write ";"
ts.Write documentItem.ContainerCode
ts.Write ";"
ts.Write documentItem.ContainerName
ts.Write ";"
ts.Write documentItem.AmountInContainer
ts.Write ";"
ts.Write document.SupplierStoreName
ts.Write ";"
ts.Write document.SupplierStoreCode
ts.Write ";"
ts.Write document.Comment

ts.WriteLine

End Sub

RunTest