Dim startDate
Dim endDate

startDate = CDate("01/01/2011")
endDate = CDate("31/12/2011")

Const ForWriting = 2
Const fileName = ".\test_ProfitTaking.csv"

Dim fso, f1, ts
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile (fileName)
Set f1 = fso.GetFile(fileName)
Set ts = f1.OpenAsTextStream(ForWriting, TristateFalse)

ts.WriteLine "Дата;Номер;Проведен;Торговое предприятие(код);Торговое предприятие (наименование);" & _
        "ЮЛ (ИНН);ЮЛ (наименование);Сумма;Вид оплаты(код);Вид оплаты(наименование);" & _
        "Номер карты;Номер кассы;Рег. номер ККМ;Номер смены;Покупатель(ИНН);Покупатель (наименование);Покупатель (код);" &_
        "Покупатель (фамилия);Покупатель (имя);Покупатель (отчество);Покупатель (дата рождения);" &_
        "Авансовый платеж (название);Авансовый платеж (код);" & _
        "Концепция (название);Концепция (код);Комментарий;Вид деятельности (название);Вид деятельности (код)"

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

Set oDocuments = loader.LoadProfitTakingData2("http://localhost:8080/resto/", "admin", "resto#test", filter)

Set oDocument = oDocuments.GetFirstProfitTaking()

Do While Not oDocument Is Nothing
    PrintDocument(oDocument)
    Set oDocument = oDocuments.GetNextProfitTaking()   
Loop

MsgBox("Выгрузка завершена!")

End Sub

Public Sub PrintDocument(document)

Set oDocumentItem = document.GetFirstProfitTakingItem()
Do While Not oDocumentItem Is Nothing
    PrintDocumentItem document, oDocumentItem
    Set oDocumentItem = document.GetNextProfitTakingItem()   
Loop
	
End Sub


Public Sub PrintDocumentItem(document, documentItem)

ts.Write document.Date
ts.Write ";"
ts.Write document.Number
ts.Write ";"
ts.Write document.Processed
ts.Write ";"
ts.Write document.DepartmentCode
ts.Write ";"
ts.Write document.DepartmentName
ts.Write ";"
ts.Write document.JuristicPersonINN
ts.Write ";"
ts.Write document.JuristicPersonName
ts.Write ";"
ts.Write documentItem.Sum_DecimalAsString
ts.Write ";"
ts.Write documentItem.PaymentTypeId
ts.Write ";"
ts.Write documentItem.PaymentTypeName
ts.Write ";"
ts.Write documentItem.Card
ts.Write ";"
ts.Write document.CashRegisterNumber
ts.Write ";"
ts.Write document.CashRegisterSerial
ts.Write ";"
ts.Write document.SessionNumber
ts.Write ";"
ts.Write documentItem.CardHolderCompanyINN
ts.Write ";"
ts.Write documentItem.CardHolderCompanyName
ts.Write ";"
ts.Write documentItem.CardHolderCode
ts.Write ";"
ts.Write documentItem.CardHolderLastName
ts.Write ";"
ts.Write documentItem.CardHolderFirstName
ts.Write ";"
ts.Write documentItem.CardHolderMiddleName
ts.Write ";"
ts.Write documentItem.CardHolderBirthday_DateTimeAsString
ts.Write ";"
ts.Write documentItem.AdvanceProductName
ts.Write ";"
ts.Write documentItem.AdvanceProductCode
ts.Write ";"
ts.Write document.ConceptionName
ts.Write ";"
ts.Write document.ConceptionCode
ts.Write ";"
ts.Write document.Comment
ts.Write ";"
ts.Write documentItem.TypeOfActivityName
ts.Write ";"
ts.Write documentItem.TypeOfActivityCode

ts.WriteLine

End Sub

RunTest