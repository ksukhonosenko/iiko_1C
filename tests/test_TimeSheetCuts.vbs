Dim startDate
Dim endDate

startDate = CDate("01/04/2012")
endDate = CDate("30/06/2012")

Const ForWriting = 2
Const fileName = ".\test_TimeSheet_Cuts.csv"

Dim fso, f1, ts
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile (fileName)
Set f1 = fso.GetFile(fileName)
Set ts = f1.OpenAsTextStream(ForWriting, TristateFalse)

ts.WriteLine "Дата;Торговое Предприятие;Сотрудник;Табельный номер;Тип явки;Дневной (DAY)/Ночной (NIGHT) период;Длительность за вычетом обеда;Длительность;Длительность обеда"
        

Public Sub RunTest()

Dim loader
Dim oDocument
Dim oDocuments

Set loader = CreateObject("iiko1CInterface.DocumentLoader")

Dim filter
Set filter = CreateObject("iiko1CInterface.DocumentFilter")
filter.DateFrom = startDate
filter.DateTo = endDate
'filter.AddDepartmentId("1")

Set oDocuments = loader.LoadEmloyeeTimeSheetWithDaysNightsAndSchedulers("http://localhost:8080/resto/", "admin", "resto#test", filter)

Set oDocument = oDocuments.GetFirstEmloyeeTimeSheet()

Do While Not oDocument Is Nothing
    PrintDocument(oDocument)
    Set oDocument = oDocuments.GetNextEmloyeeTimeSheet()   
Loop

Dim version
Set version = CreateObject("iiko1CInterface.Version")

MsgBox ("Выгрузка завершена! Версия протокола: " + version.ProtocolVersion)

End Sub

Public Sub PrintDocument(document)
    PrintDocument document	
End Sub

Public Sub PrintDocument(document)
ts.Write document.AttendanceDate
ts.Write ";"
ts.Write document.Department
ts.Write ";"
ts.Write document.Employee
ts.Write ";"
ts.Write document.EmployeeCode
ts.Write ";"
ts.Write document.AttendanceType
ts.Write ";"
ts.Write document.DayNightType
ts.Write ";"
ts.Write document.DinnerTimeString
ts.Write ";"
ts.Write document.DurationString
ts.Write ";"
ts.Write document.DurationOfDinnerString
ts.WriteLine
End Sub

RunTest