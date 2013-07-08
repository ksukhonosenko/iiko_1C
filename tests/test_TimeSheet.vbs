Dim startDate
Dim endDate

startDate = CDate("01/06/2012")
endDate = CDate("30/06/2012")

Const ForWriting = 2
Const fileName = ".\test_TimeSheet.csv"

Dim fso, f1, ts
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile (fileName)
Set f1 = fso.GetFile(fileName)
Set ts = f1.OpenAsTextStream(ForWriting, TristateFalse)

ts.WriteLine "Вид журнала (название);Вид журнала (код);Дата регистрации;Дата учета;Сотрудник (код);Сотрудник (наименование);Тип смены (код);" & _
        "Тип смены (наименование);Тип смены (длительность);Тип смены (неоплачиваемое время);Тип смены (время начала);Тип смены (время окончания);" & _
        "Время прихода;Время ухода;Длительность;Неоплачиваемое время;Тип явки (наименование);Тип явки (код);Тип явки (коэффициент расчета, %);"  &_
        "Тип явки (явка/неявка);Комментарий;Концепция (название);Концепция (код);" &_
        "Торговое предприятие (код);Торговое предприятие (наименование);ЮЛ (ИНН);ЮЛ (наименование);" & _
        "Комментарий"
        

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

Set oDocuments = loader.LoadEmloyeeTimeSheet2("http://localhost:8080/resto/", "admin", "resto#test", filter)

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
Set oDocumentItem = document.GetFirstEmloyeeTimeSheetItem()
Do While Not oDocumentItem Is Nothing
    PrintDocumentItem oDocumentItem, document
    ts.WriteLine
    Set oDocumentItem = document.GetNextEmloyeeTimeSheetItem()   
Loop
	
End Sub

Public Sub PrintDocumentItem(docItem, document)
ts.Write docItem.TypeCode
ts.Write ";"
ts.Write docItem.TypeName
ts.Write ";"
ts.Write docItem.Date
ts.Write ";"
ts.Write docItem.AccountingDate
ts.Write ";"
ts.Write docItem.EmployeeName
ts.Write ";"
ts.Write docItem.EmployeeCode
ts.Write ";"
ts.Write docItem.ScheduleCode
ts.Write ";"
ts.Write docItem.ScheduleName
ts.Write ";"
ts.Write docItem.ScheduleDuration
ts.Write ";"
ts.Write docItem.ScheduleNonPaid
ts.Write ";"
ts.Write docItem.ScheduleStart
ts.Write ";"
ts.Write docItem.ScheduleEnd
ts.Write ";"
ts.Write docItem.StartDate_DateTimeAsString
ts.Write ";"
ts.Write docItem.EndDate_DateTimeAsString 
ts.Write ";"
ts.Write docItem.Duration
ts.Write ";"
ts.Write docItem.NonPaid
ts.Write ";"
ts.Write docItem.AttendanceTypeName
ts.Write ";"
ts.Write docItem.AttendanceTypeCode
ts.Write ";"
ts.Write docItem.AttendancePayRatePercent_Int32AsString
ts.Write ";"
ts.Write docItem.IsAttendance_BooleanAsString
ts.Write ";"
ts.Write docItem.Comment
ts.Write ";"
ts.Write docItem.ConceptionName
ts.Write ";"
ts.Write docItem.ConceptionCode
ts.Write ";"
ts.Write document.DepartmentCode
ts.Write ";"
ts.Write document.DepartmentName
ts.Write ";"
ts.Write document.JuristicPersonINN
ts.Write ";"
ts.Write document.JuristicPersonName
ts.Write ";"
ts.Write document.Comment

End Sub

RunTest