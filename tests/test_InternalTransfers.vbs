Dim startDate
Dim endDate

startDate = CDate("01/01/2011")
endDate = CDate("31/12/2011")

Const ForWriting = 2
Const fileName = ".\test_InternalTransfers.csv"

Dim fso, f1, ts
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile (fileName)
Set f1 = fso.GetFile(fileName)
' TristateUseDefault -2	Uses the system default file format
' TristateTrue 	     -1	Opens the file using the Unicode format
' TristateFalse       0	Opens the file in AscII format
Set ts = f1.OpenAsTextStream(ForWriting, TristateFalse)

ts.WriteLine "Дата;Номер;Проведен;Торговое предприятие отправитель (код);Торговое предприятие отправитель (наименование);" & _
        "ЮЛ отправитель (ИНН);ЮЛ отправитель (наименование);Склад отправитель (код);Склад отправитель (наименование);" & _
        "Торговое предприятие получатель(код);Торговое предприятие получатель (наименование);ЮЛ получатель (ИНН);" & _
        "ЮЛ получатель (наименование);Склад получатель (код);Склад получатель (наименование);Номенклатура (код);" & _
        "Номенклатура (наименование);Группа учета (название);Группа учета (код);Количество;Себестоимость за ед. без НДС;Себестоимость без НДС;Ставка НДС;" & _
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

Set oDocuments = loader.LoadInternalTransfersData2("http://localhost:8080/resto/", "admin", "resto#test", filter)

Set oDocument = oDocuments.GetFirstInternalTransfer()

Do While Not oDocument Is Nothing
    PrintDocument(oDocument)
    Set oDocument = oDocuments.GetNextInternalTransfer()   
Loop

MsgBox("Выгрузка завершена!")

End Sub

Public Sub PrintDocument(document)

Set oDocumentItem = document.GetFirstInternalTransferItem()
Do While Not oDocumentItem Is Nothing
    PrintDocumentItem document, oDocumentItem
    Set oDocumentItem = document.GetNextInternalTransferItem()   
Loop
	
End Sub


Public Sub PrintDocumentItem(document, documentItem)

ts.Write document.Date
ts.Write ";"
ts.Write document.Number
ts.Write ";"
ts.Write document.Processed
ts.Write ";"
ts.Write document.DepartmentOutcomeCode
ts.Write ";"
ts.Write document.DepartmentOutcomeName
ts.Write ";"
ts.Write document.JuristicPersonOutcomeINN
ts.Write ";"
ts.Write document.JuristicPersonOutcomeName
ts.Write ";"
ts.Write document.StoreOutcomeCode
ts.Write ";"
ts.Write document.StoreOutcomeName
ts.Write ";"
ts.Write document.DepartmentIncomeCode
ts.Write ";"
ts.Write document.DepartmentIncomeName
ts.Write ";"
ts.Write document.JuristicPersonIncomeINN
ts.Write ";"
ts.Write document.JuristicPersonIncomeName
ts.Write ";"
ts.Write document.StoreIncomeCode
ts.Write ";"
ts.Write document.StoreIncomeName
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
ts.Write documentItem.CostPriceByUnit_DecimalAsString
ts.Write ";"
ts.Write documentItem.CostPrice_DecimalAsString
ts.Write ";"
ts.Write documentItem.NdsProductPercent_DecimalAsString
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