Dim startDate
Dim endDate

startDate = CDate("01/01/2011")
endDate = CDate("31/12/2011")


Const ForWriting = 2
Const fileName = ".\test_CashFlow.csv"

Dim fso, f1, ts
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile (fileName)
Set f1 = fso.GetFile(fileName)
Set ts = f1.OpenAsTextStream(ForWriting, TristateFalse)

ts.WriteLine "Дата;Номер;Проведен;Сумма;Дебет (наименование);Дебет (код);Дебет (тип);Дебет, подсчет (наименование);" & _
        "Дебет, подсчет (код);Организация, дебет (ИНН);Организация, дебет (код);Организация, дебет (наименование);" & _
        "Кредит (наименование);Кредит (код);Кредит (тип);Кредит, подсчет (наименование);Кредит, подсчет (код);" & _
        "Организация, кредит (ИНН);Организация, кредит (код);Организация, кредит (наименование);"  &_
        "Торговое предприятие(код);Торговое предприятие (наименование);ЮЛ (ИНН);ЮЛ (наименование);" & _
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


Set oDocuments = loader.LoadCashFlowData2("http://localhost:8080/resto/", "admin", "resto#test", filter)

Set oDocument = oDocuments.GetFirstCashFlow()

Do While Not oDocument Is Nothing
    PrintDocument(oDocument)
    Set oDocument = oDocuments.GetNextCashFlow()   
Loop

MsgBox("Выгрузка завершена!")

End Sub

Public Sub PrintDocument(document)

ts.Write document.Date
ts.Write ";"
ts.Write document.Number
ts.Write ";"
ts.Write document.Processed
ts.Write ";"
ts.Write document.Sum
ts.Write ";"
ts.Write document.DebitAccountName
ts.Write ";"
ts.Write document.DebitAccountCode
ts.Write ";"
ts.Write document.DebitAccountType
ts.Write ";"
ts.Write document.DebitSubAccountName
ts.Write ";"
ts.Write document.DebitSubAccountCode
ts.Write ";"
ts.Write document.CounteragentToINN
ts.Write ";"
ts.Write document.CounteragentToCode
ts.Write ";"
ts.Write document.CounteragentToName
ts.Write ";"
ts.Write document.CreditAccountName
ts.Write ";"
ts.Write document.CreditAccountCode
ts.Write ";"
ts.Write document.CreditAccountType
ts.Write ";"
ts.Write document.CreditSubAccountName
ts.Write ";"
ts.Write document.CreditSubAccountCode
ts.Write ";"
ts.Write document.CounteragentFromINN
ts.Write ";"
ts.Write document.CounteragentFromCode
ts.Write ";"
ts.Write document.CounteragentFromName
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
ts.Write document.Comment

ts.WriteLine
	
End Sub

RunTest