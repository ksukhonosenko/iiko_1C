Dim startDate
Dim endDate

startDate = CDate("01/01/2011")
endDate = CDate("31/12/2011")

Const ForWriting = 2
Const fileName = ".\test_Productions.csv"

Dim fso, f1, ts
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile (fileName)
Set f1 = fso.GetFile(fileName)
Set ts = f1.OpenAsTextStream(ForWriting, TristateFalse)

ts.WriteLine "Дата;Номер;Проведен;Склад списания(код);Склад списания (наименование);Склад прихода(код);Склад прихода (наименование);" & _
        "Знак операции;Знак операции (название);Номенклатура (код);Номенклатура (наименование);Группа учета (код);Группа учета (название);" & _
        "Количество;Себестоимость за ед. без НДС;Себестоимость без НДС;Ставка НДС;Торговое предприятие(код);" & _
        "Торговое предприятие (наименование);ЮЛ (ИНН);ЮЛ (наименование);Тип номенклатуры (код);Тип номенклатуры (название);" & _
        "Целевое блюдо (код);Целевое блюдо (название);Единица измерения (код);Единица измерения (название);" & _
        "Концепция (название);Концепция (код);Группа учета целевого блюда (название);Группа учета целевого блюда (код);" & _
        "Комментарий"

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

Set oDocuments = loader.LoadProductionsData2("http://localhost:8080/resto/", "admin", "resto#test", filter)

Set oDocument = oDocuments.GetFirstProduction()

Do While Not oDocument Is Nothing
    PrintDocument(oDocument)
    Set oDocument = oDocuments.GetNextProduction()   
Loop

MsgBox("Выгрузка завершена!")

End Sub

Public Sub PrintDocument(document)

Dim isOutcome
Set oDocumentItem = document.GetFirstIncomeItems()
Do While Not oDocumentItem Is Nothing
    isOutcome = -1    
    PrintDocumentItem document, oDocumentItem, isOutcome
    Set oDocumentItem = document.GetNextIncomeItems()   
Loop

Set oDocumentItem = document.GetFirstOutcomeItems()
Do While Not oDocumentItem Is Nothing
    isOutcome = 1
    PrintDocumentItem document, oDocumentItem, isOutcome
    Set oDocumentItem = document.GetNextOutcomeItems()   
Loop

End Sub


Public Sub PrintDocumentItem(document, documentItem, isOutcome)

ts.Write document.Date
ts.Write ";"
ts.Write document.Number
ts.Write ";"
ts.Write document.Processed
ts.Write ";"

If isOutcome > 0 Then
ts.Write documentItem.StoreCode
ts.Write ";"
ts.Write documentItem.StoreName
ts.Write ";"
ts.Write ""
ts.Write ";"
ts.Write ""
Else
ts.Write ""
ts.Write ";"
ts.Write ""
ts.Write ";"
ts.Write documentItem.StoreCode
ts.Write ";"
ts.Write documentItem.StoreName

End If
ts.Write ";"

ts.Write documentItem.Operation
ts.Write ";"
ts.Write documentItem.OperationName
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
ts.Write documentItem.TargetDishCode
ts.Write ";"
ts.Write documentItem.TargetDishName
ts.Write ";"
ts.Write documentItem.MeasureUnitCode
ts.Write ";"
ts.Write documentItem.MeasureUnitName
ts.Write ";"
ts.Write document.ConceptionName
ts.Write ";"
ts.Write document.ConceptionCode
ts.Write ";"
ts.Write documentItem.TargetDishAccountingCategoryName
ts.Write ";"
ts.Write documentItem.TargetDishAccountingCategoryCode
ts.Write ";"
ts.Write document.Comment

ts.WriteLine

End Sub

RunTest