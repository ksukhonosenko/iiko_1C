Dim startDate
Dim endDate

startDate = CDate("01/01/2012")
endDate = CDate("31/12/2012")

Const ForWriting = 2
Const fileName = ".\test_Sales.csv"

Dim fso, f1, ts
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile (fileName)
Set f1 = fso.GetFile(fileName)
Set ts = f1.OpenAsTextStream(ForWriting, TristateFalse)

ts.WriteLine "Дата;Номер;Проведен;Склад(код);Склад (наименование);Операция;Операция (название);" & _
    "Номенклатура (код);Номенклатура (наименование);Группа учета (название);Группа учета (код);Количество;" & _
    "Цена продажи с НДС;Сумма продажи с НДС;Сумма НДС с продаж;Ставка НДС с продаж;" & _
    "Себестоимость за ед. без НДС;Себестоимость без НДС;Ставка НДС;Торговое предприятие(код);" & _
    "Торговое предприятие (наименование);ЮЛ (ИНН);ЮЛ (наименование);" &_
    "Тип списания;Тип списания (название);Движение денежных средств (код);Движение денежных средств;" & _
    "Статья расходов(код);Статья расходов;Тип номенклатуры (код);Тип номенклатуры (название);" & _
    "Целевое блюдо (код);Целевое блюдо (название);Единица измерения (код);Единица измерения (название);" & _
    "Номер смены;Номер кассы;Концепция (название);Концепция (код);Группа учета целевого блюда (название);Группа учета целевого блюда (код);" & _
    "Комментарий"

Public Sub RunTest()

Dim loader
Dim oSale
Dim oSales

Set loader = CreateObject("iiko1CInterface.DocumentLoader")

Dim filter
Set filter = CreateObject("iiko1CInterface.DocumentFilter")
filter.DateFrom = startDate
filter.DateTo = endDate
'Set filter.IncludeUnprocessed = true
'Set filter.UseBusinessSettings = true
'filter.AddDepartmentId("1")

Set oSales = loader.LoadSalesData2("http://localhost:8080/resto/", "admin", "resto#test", filter)

Set oSale = oSales.GetFirstSaleData()

Do While Not oSale Is Nothing
    PrintSale(oSale)
    Set oSale = oSales.GetNextSaleData()   
Loop

MsgBox("Выгрузка завершена!")

End Sub

Public Sub PrintSale(sale)

Set item = sale.GetFirstSaleDataItem()
Do While Not item Is Nothing
    PrintSaleItem sale, item
    Set item = sale.GetNextSaleDataItem()   
Loop

End Sub


Public Sub PrintSaleItem(sale, saleItem)

ts.Write sale.Date
ts.Write ";"
ts.Write sale.Number
ts.Write ";"
ts.Write sale.Processed
ts.Write ";"
ts.Write sale.StoreCode
ts.Write ";"
ts.Write sale.StoreName
ts.Write ";"
ts.Write saleItem.Operation
ts.Write ";"
ts.Write saleItem.OperationName
ts.Write ";"
ts.Write saleItem.Article
ts.Write ";"
ts.Write saleItem.Nomenclature
ts.Write ";"
ts.Write saleItem.AccountingCategory
ts.Write ";"
ts.Write saleItem.AccountingCategoryCode
ts.Write ";"
ts.Write saleItem.Amount_DecimalAsString
ts.Write ";"
ts.Write saleItem.PriceWithNds_DecimalAsString
ts.Write ";"	
ts.Write saleItem.SumWithNds_DecimalAsString
ts.Write ";"
ts.Write saleItem.Nds_DecimalAsString
ts.Write ";"
ts.Write saleItem.NdsPercent_DecimalAsString
ts.Write ";"
ts.Write saleItem.CostPriceByUnit_DecimalAsString
ts.Write ";"
ts.Write saleItem.CostPrice_DecimalAsString
ts.Write ";"
ts.Write saleItem.NdsNomenclaturePercent_DecimalAsString
ts.Write ";"
ts.Write sale.DepartmentCode
ts.Write ";"
ts.Write sale.DepartmentName
ts.Write ";"
ts.Write sale.JuristicPersonINN
ts.Write ";"
ts.Write sale.JuristicPersonName
ts.Write ";"
ts.Write saleItem.WriteoffType
ts.Write ";"
ts.Write saleItem.WriteoffTypeName
ts.Write ";"
ts.Write sale.RevenueAccountCode
ts.Write ";"
ts.Write sale.RevenueAccountName
ts.Write ";"
ts.Write sale.WriteoffAccountCode
ts.Write ";"
ts.Write sale.WriteoffAccountName
ts.Write ";"
ts.Write saleItem.NomenclatureType
ts.Write ";"
ts.Write saleItem.NomenclatureTypeName
ts.Write ";"
ts.Write saleItem.TargetDishCode
ts.Write ";"
ts.Write saleItem.TargetDishName
ts.Write ";"
ts.Write saleItem.MeasureUnitCode
ts.Write ";"
ts.Write saleItem.MeasureUnitName
ts.Write ";"
ts.Write sale.SessionNumber
ts.Write ";"
ts.Write sale.CashRegNumber
ts.Write ";"
ts.Write sale.ConceptionName
ts.Write ";"
ts.Write sale.ConceptionCode
ts.Write ";"
ts.Write saleItem.TargetDishAccountingCategoryName
ts.Write ";"
ts.Write saleItem.TargetDishAccountingCategoryCode
ts.Write ";"
ts.Write sale.Comment

ts.WriteLine

End Sub


RunTest