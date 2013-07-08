Dim startDate
Dim endDate

startDate = CDate("01/01/2011")
endDate = CDate("31/12/2011")

Const ForWriting = 2
Const fileName = ".\test_ReturnedInvoices.csv"

Dim fso, f1, ts
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile (fileName)
Set f1 = fso.GetFile(fileName)
Set ts = f1.OpenAsTextStream(ForWriting, TristateFalse)

ts.WriteLine "Дата;Номер;Проведен;Склад (код);Склад (наименование);Номенклатура (код);Номенклатура (наименование);" & _
        "Группа учета (название);Группа учета (код);Количество;Цена продажи с НДС;Сумма продажи с НДС;Сумма НДС с продаж;Ставка НДС с продаж;" & _
        "Себестоимость за ед. без НДС;Себестоимость без НДС;Ставка НДС;Торговое предприятие(код);" & _
        "Торговое предприятие (наименование);ЮЛ (ИНН);ЮЛ (наименование);" & _
        "Движение денежных средств (код);Движение денежных средств; " & _
        "Статья расходов(код);Статья расходов;Покупатель (ИНН);Покупатель (код);Покупатель (наименование);" & _
        "Приходная накладная (дата);Приходная накладная (номер);Счет-фактура;" & _
        "Тип номенклатуры (код);Тип номенклатуры (название);Единица измерения (код);Единица измерения (название);" & _
        "Концепция (название);Концепция (код);Комментарий"

Public Sub RunTest()

Dim loader
Dim oDocument
Dim oDocumentsc

Set loader = CreateObject("iiko1CInterface.DocumentLoader")

Dim filter
Set filter = CreateObject("iiko1CInterface.DocumentFilter")
filter.DateFrom = startDate
filter.DateTo = endDate
'Set filter.IncludeUnprocessed = true
'Set filter.UseBusinessSettings = true
'filter.AddDepartmentId("1")

Set oDocuments = loader.LoadReturnedInvoicesData2("http://localhost:8080/resto/", "admin", "resto#test", filter)

Set oDocument = oDocuments.GetFirstReturnedInvoice()

Do While Not oDocument Is Nothing
    PrintDocument(oDocument)
    Set oDocument = oDocuments.GetNextReturnedInvoice()   
Loop

MsgBox("Выгрузка завершена!")

End Sub

Public Sub PrintDocument(document)

Set oDocumentItem = document.GetFirstReturnedInvoiceItem()
Do While Not oDocumentItem Is Nothing
    PrintDocumentItem document, oDocumentItem
    Set oDocumentItem = document.GetNextReturnedInvoiceItem()   
Loop
	
End Sub


Public Sub PrintDocumentItem(document, documentItem)

ts.Write document.Date
ts.Write ";"
ts.Write document.Number
ts.Write ";"
ts.Write document.Processed
ts.Write ";"
ts.Write documentItem.StoreCode
ts.Write ";"
ts.Write documentItem.StoreName
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
ts.Write documentItem.CostPriceByUnit_DecimalAsString
ts.Write ";"
ts.Write documentItem.CostPrice_DecimalAsString
ts.Write ";"
ts.Write documentItem.NdsProductPercent_DecimalAsString
ts.Write ";"
ts.Write document.DepartmentCode
ts.Write ";"
ts.Write document.DepartmentName
ts.Write ";"
ts.Write document.JuristicPersonINN
ts.Write ";"
ts.Write document.JuristicPersonName
ts.Write ";"
ts.Write document.RevenueAccountCode
ts.Write ";"
ts.Write document.RevenueAccountName
ts.Write ";"
ts.Write document.WriteoffAccountCode
ts.Write ";"
ts.Write document.WriteoffAccountName
ts.Write ";"
ts.Write document.CustomerINN
ts.Write ";"
ts.Write document.CustomerCode
ts.Write ";"
ts.Write document.CustomerName
ts.Write ";"
ts.Write document.IncomingInvoiceNumber
ts.Write ";"
ts.Write document.IncomingInvoiceDate_DateTimeAsString
ts.Write ";"
ts.Write document.Invoice
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
ts.Write document.Comment

ts.WriteLine

End Sub

RunTest