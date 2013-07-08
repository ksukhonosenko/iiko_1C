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

ts.WriteLine "����;�����;��������;����� ��������(���);����� �������� (������������);����� �������(���);����� ������� (������������);" & _
        "���� ��������;���� �������� (��������);������������ (���);������������ (������������);������ ����� (���);������ ����� (��������);" & _
        "����������;������������� �� ��. ��� ���;������������� ��� ���;������ ���;�������� �����������(���);" & _
        "�������� ����������� (������������);�� (���);�� (������������);��� ������������ (���);��� ������������ (��������);" & _
        "������� ����� (���);������� ����� (��������);������� ��������� (���);������� ��������� (��������);" & _
        "��������� (��������);��������� (���);������ ����� �������� ����� (��������);������ ����� �������� ����� (���);" & _
        "�����������"

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

MsgBox("�������� ���������!")

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