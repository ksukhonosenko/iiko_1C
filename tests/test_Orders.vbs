Dim startDate
Dim endDate

startDate = CDate("01/01/2011")
endDate = CDate("31/12/2011")

Const ForWriting = 2
Const fileName = ".\test_Orders.csv"

Dim fso, f1, ts
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile (fileName)
Set f1 = fso.GetFile(fileName)
Set ts = f1.OpenAsTextStream(ForWriting, TristateFalse)

ts.WriteLine "Дата учета;Номер;Проведен ди документ;Тип строки (название);Тип строки (код);Торговое предприятие(код);Торговое предприятие (наименование);" & _
        "ЮЛ (ИНН);ЮЛ (наименование);Концепция (название);Концепция (код);Номер смены;Фискальный номер смены;Номер кассы;Серийный номер кассы;" & _
        "Номер заказа;Guid заказа;Номер чека;Дата и время регистрации;Вид деятельности;Комментарий;" & _
        "Номенклатура (код);Номенклатура (наименование);Группа учета (код);Группа учета (название);Тип номенклатуры (код);Тип номенклатуры (название);"  & _
        "Единица измерения (код);Единица измерения (название);Целевое блюдо (код);Целевое блюдо (название);Цена номенклатуры;Количество;" & _
        "Сумма продажи;Сумма скидки по позиции;Ставка НДС по позиции;Сумма НДС по позиции;" & _
        "Тип оплаты (код);Тип оплаты (название);Сумма оплаты;Фискальная оплата;Контрагент (код);Контрагент (наименование);"  & _
        "Тип скидки (код);Тип скидки (название);Сумма скидки;" & _
        "Причина удаления (код);Причина удаления (название)"
        

Public Sub RunTest()

Dim loader
Dim oDocument
Dim oDocuments
Set loader = CreateObject("iiko1CInterface.DocumentLoader")

Dim filter
Set filter = CreateObject("iiko1CInterface.DocumentFilter")
filter.DateFrom = startDate
filter.DateTo = endDate

Set oDocuments = loader.LoadOrders2("http://localhost:8080/resto/", "admin", "resto#test", filter)

Set oDocument = oDocuments.GetFirstOrdersDocument()

Do While Not oDocument Is Nothing
    PrintDocument(oDocument)
    Set oDocument = oDocuments.GetNextOrdersDocument()   
Loop

Dim version
Set version = CreateObject("iiko1CInterface.Version")

MsgBox ("Выгрузка завершена! Версия протокола: " + version.ProtocolVersion)

End Sub

Public Sub PrintDocument(document)
Set oDocumentItem = document.GetFirstOrderDocumentItem()
Do While Not oDocumentItem Is Nothing
	PrintCommonOrderInfo oDocumentItem, document
	if oDocumentItem.ItemTypeCode = 1 then
	    PrintItemSale oDocumentItem
        End If
	if oDocumentItem.ItemTypeCode = 2 then
	    PrintItemPayment oDocumentItem
        End If
	if oDocumentItem.ItemTypeCode = 3 then
	    PrintItemDiscount oDocumentItem
        End If
	if oDocumentItem.ItemTypeCode = 4 then
	    PrintItemWriteoff oDocumentItem
	End If
    ts.WriteLine
    Set oDocumentItem = document.GetNextOrderDocumentItem()   
Loop
	
End Sub


Public Sub PrintCommonOrderInfo(docItem, document)
ts.Write document.Date
ts.Write ";"
ts.Write document.Number
ts.Write ";"
ts.Write document.Processed
ts.Write ";"
ts.Write docItem.ItemTypeName
ts.Write ";"
ts.Write docItem.ItemTypeCode
ts.Write ";"
ts.Write document.DepartmentCode
ts.Write ";"
ts.Write document.DepartmentName
ts.Write ";"
ts.Write document.JuristicPersonINN
ts.Write ";"
ts.Write document.JuristicPersonName
ts.Write ";"
ts.Write document.ConceptionName
ts.Write ";"
ts.Write document.ConceptionCode
ts.Write ";"
ts.Write document.SessionNumber
ts.Write ";"
ts.Write document.SessionFiscalNumber
ts.Write ";"
ts.Write document.CashRegisterNumber
ts.Write ";"
ts.Write document.CashRegisterSerialNumber
ts.Write ";"
ts.Write docItem.OrderNumber
ts.Write ";"
ts.Write docItem.OrderId
ts.Write ";"
ts.Write docItem.CheckNumber
ts.Write ";"
ts.Write docItem.OrderClosed
ts.Write ";"
ts.Write docItem.TypeOfActivity
ts.Write ";"
ts.Write document.Comment
ts.Write ";"
End Sub

Public Sub PrintItemSale(docItem)
ts.Write docItem.Article
ts.Write ";"
ts.Write docItem.Nomenclature
ts.Write ";"
ts.Write docItem.AccountingCategoryCode
ts.Write ";"
ts.Write docItem.AccountingCategory
ts.Write ";"
ts.Write docItem.NomenclatureType
ts.Write ";"
ts.Write docItem.NomenclatureTypeName
ts.Write ";"
ts.Write docItem.MeasureUnitCode
ts.Write ";"
ts.Write docItem.MeasureUnitName
ts.Write ";"
ts.Write docItem.SoldWithDishCode
ts.Write ";"
ts.Write docItem.SoldWithDishName
ts.Write ";"
ts.Write docItem.NomenclaturePrice
ts.Write ";"
ts.Write docItem.NomenclatureAmount
ts.Write ";"
ts.Write docItem.SumWithNds
ts.Write ";"
ts.Write docItem.NomenclatureDiscountSum
ts.Write ";"
ts.Write docItem.NdsPercent
ts.Write ";;;;;;;;;"
End Sub

Public Sub PrintItemPayment(docItem)
ts.Write ";;;;;;;;;;;;;;;;"
ts.Write docItem.PaymentTypeId
ts.Write ";"
ts.Write docItem.PaymentTypeName
ts.Write ";"
ts.Write docItem.PaymentSum
ts.Write ";"
ts.Write docItem.IsPaymentFiscal
ts.Write ";"
ts.Write docItem.CounteragentCode
ts.Write ";"
ts.Write docItem.CounteragentName
ts.Write ";;;;"
End Sub

Public Sub PrintItemDiscount(docItem)
ts.Write ";;;;;;;;;;;;;;;;;;;;;;"
ts.Write docItem.DiscountId
ts.Write ";"
ts.Write docItem.DiscountName
ts.Write ";"
ts.Write docItem.DiscountSum
ts.Write ";;"
End Sub

Public Sub PrintItemWriteoff(docItem)
ts.Write ";;;;;;;;;;;;;;;;;;;;;;;;;"
ts.Write docItem.RemovalTypeId
ts.Write ";"
ts.Write docItem.RemovalTypeName
End Sub

RunTest