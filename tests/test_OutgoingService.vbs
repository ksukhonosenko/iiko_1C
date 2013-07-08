Dim startDate
Dim endDate

startDate = CDate("01/01/2013")
endDate = CDate("31/12/2013")

Const ForWriting = 2
Const fileName = ".\test_OutgoingSerice.csv"

Dim fso, f1, ts
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile (fileName)
Set f1 = fso.GetFile(fileName)
Set ts = f1.OpenAsTextStream(ForWriting, TristateFalse)

ts.WriteLine "Дата;Номер;Проведен;Движение денежных средств (код);" & _
        "Движение денежных средств;Номенклатура (код);Номенклатура (наименование);" & _
        "Группа учета (название);Группа учета (код);Количество;Цена с НДС;Сумма с НДС;Сумма НДС;Ставка НДС с продаж;" & _
        "Ставка НДС;Торговое предприятие(код);Торговое предприятие (наименование);ЮЛ (ИНН);ЮЛ (наименование);" & _
	"Покупатель (ИНН);Покупатель (код);Покупатель (наименование);" & _
        "Тип номенклатуры (код);Тип номенклатуры (название);Единица измерения (код);Единица измерения (название);" & _
        "Номер кассы;Номер смены;Концепция (название);Концепция (код);" & _
        "Комментарий"

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

Set oInvoces = loader.LoadOutgoingServiceData2("http://localhost:8080/resto/", "admin", "resto#test", filter)

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
ts.Write document.AccountName
ts.Write ";"
ts.Write document.AccountCode
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
ts.Write document.SupplierINN
ts.Write ";"
ts.Write document.SupplierCode
ts.Write ";"
ts.Write document.SupplierName
ts.Write ";"
ts.Write documentItem.NomenclatureType
ts.Write ";"
ts.Write documentItem.NomenclatureTypeName
ts.Write ";"
ts.Write documentItem.MeasureUnitCode
ts.Write ";"
ts.Write documentItem.MeasureUnitName
ts.Write ";"
ts.Write document.RegisterNumber
ts.Write ";"
ts.Write document.SessionNumber
ts.Write ";"
ts.Write document.ConceptionName
ts.Write ";"
ts.Write document.ConceptionCode
ts.Write ";"
ts.Write document.Comment

ts.WriteLine

End Sub

RunTest