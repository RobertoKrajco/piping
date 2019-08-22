Imports System.Data.OleDb
Imports System.IO

Module MyModule
    Public cn1 As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\krajcovic\Documents\piping\PipingPlan_v02.MDB")
    'Public cn1 As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\enpemaskafs001\lhgroups\01_LEAN_ACTIVITY\Thermal Management Lean Workshops\= Piping\PipingPlan\PipingPlan_v01.MDB")
    'Public cn1 As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\frantisek.horecny\Documents\Lean\Thermal Management\= Piping\PipingPlan Aplikacia\PipingPlan_v01.MDB")

    'Public MyConnection As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & HlavnyPlanLocation & ";Extended Properties=Excel 8.0;")
    'Public MyConnection As New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & HlavnyPlanLocation & ";Extended Properties=Excel 8.0;")
    Public MyConnection As New System.Data.OleDb.OleDbConnection
    Public NazovAplikacie As String = "Piping Plán"
    Public HlavnyPlanLocation As String
    Public pickPN As String
    Public filepath As String
    Public webAddress As String
    Public str1 As String
    Public DRWLink As String
    Public indexHlavnyPlan As Integer
    Public indexRezanie As Integer
    Public indexOhybanie As Integer
    Public indexPriprava As Integer
    Public indexSpajkovanie As Integer
    Public indexVychystanie As Integer
    Public indexInziniering As Integer

    Public Sub ProfileVisibility()

    End Sub

    Public Sub SetLinesColorOhybanie()

        'vyfarbenie riadkov, kde je operacia ukoncena
        For i As Integer = 0 To PipingPlanForm.gridOhybanie.Rows.Count - 1
            If IsDBNull(PipingPlanForm.gridOhybanie.Rows(i).Cells(11).Value) Then
                Continue For
            ElseIf PipingPlanForm.gridOhybanie.Rows(i).Cells(11).Value = True Then
                PipingPlanForm.gridOhybanie.Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
            End If

        Next

    End Sub

    Public Sub SetLinesColorPriprava()

        'vyfarbenie riadkov, kde je operacia ukoncena

        For i As Integer = 0 To PipingPlanForm.gridPriprava.Rows.Count - 1
            If IsDBNull(PipingPlanForm.gridPriprava.Rows(i).Cells(6).Value) Then
                Continue For
            ElseIf PipingPlanForm.gridPriprava.Rows(i).Cells(6).Value = True Then
                PipingPlanForm.gridPriprava.Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
            End If

        Next

    End Sub

    Public Sub SetLinesColorSpajkovanie()

        'vyfarbenie riadkov, kde je operacia ukoncena
        For i As Integer = 0 To PipingPlanForm.gridSpajkovanie.Rows.Count - 1
            If IsDBNull(PipingPlanForm.gridSpajkovanie.Rows(i).Cells(6).Value) Then
                Continue For
            ElseIf PipingPlanForm.gridSpajkovanie.Rows(i).Cells(6).Value = True Then
                PipingPlanForm.gridSpajkovanie.Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
            End If

        Next
    End Sub

    Public Sub SetLinesColorVychystanie()

        'vyfarbenie riadkov, kde je vyplanovaný WJ
        For i As Integer = 0 To PipingPlanForm.gridVychystanie.Rows.Count - 1
            If IsDBNull(PipingPlanForm.gridVychystanie.Rows(i).Cells(10).Value) Then
                Continue For
            ElseIf PipingPlanForm.gridVychystanie.Rows(i).Cells(10).Value = True Then
                PipingPlanForm.gridVychystanie.Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
            End If

        Next

    End Sub

    Public Sub SetLinesColorVyplanovany()

        'vyfarbenie riadkov, kde je vyplanovaný WJ
        For i As Integer = 0 To PipingPlanForm.gridVychystanie.Rows.Count - 1
            If IsDBNull(PipingPlanForm.gridVychystanie.Rows(i).Cells(6).Value) Then
                Continue For
            ElseIf PipingPlanForm.gridVychystanie.Rows(i).Cells(6).Value = True Then
                PipingPlanForm.gridVychystanie.Rows(i).DefaultCellStyle.BackColor = Color.Orange
            End If

        Next

    End Sub

    Public Sub SetLinesColorRezanie()

        ''vyfarbenie riadkov, kde je operacia ukoncena
        'For i As Integer = 0 To PipingPlanForm.gridRezanie.Rows.Count - 1
        '    If (Not PipingPlanForm.gridRezanie.Rows(i).IsNewRow) Then
        '        If PipingPlanForm.gridRezanie.Rows(i).Cells(8).Value = True Then
        '            PipingPlanForm.gridRezanie.Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
        '        End If
        '    End If
        'Next

        'vyfarbenie riadkov, kde je operacia ukoncena
        'PipingPlanForm.Button1_Click()
        Try
            For i As Integer = 0 To PipingPlanForm.gridRezanie.Rows.Count - 1
                If IsDBNull(PipingPlanForm.gridRezanie.Rows(i).Cells(9).Value) Then
                    Continue For
                ElseIf PipingPlanForm.gridRezanie.Rows(i).Cells(9).Value = True Then
                    PipingPlanForm.gridRezanie.Rows(i).DefaultCellStyle.BackColor = Color.LightGreen
                End If

            Next
        Catch ex As Exception

        End Try
    End Sub

    Public Function HyperlinkSplitFunction(ByRef Str1 As String) As String
        'funkcia volana z kliknutia na riadok v gridOhybanie
        'po kliknuti sa do premennej Str1 ulozi hyperlink a zavola sa TATO funkcia
        'funkcia rozdeli nespravny hyperlink z premennej Str1 na retazce a vyseparuje spravny format hyperlinku a vrati spravnu hodnotu tam, odkial sa funkcia zavolala

        'split funkcia, ktora rozlozi retazec medzi znaky #
        Dim parts As String() = DRWLink.Split(New Char() {"#"c})

        Dim i As Integer = 0
        HyperlinkSplitFunction = parts(1)

    End Function

    Public Function DbNullOrStringValue(ByVal value As String) As Object
        If String.IsNullOrEmpty(value) Then
            Return DBNull.Value
        Else
            Return value
        End If
    End Function

    Public Function CalculateBusinessDaysFromInputDate(ByVal StartDate As Date, ByVal NumberOfBusinessDays As Integer) As Date
        'Knock the start date down one day if it is on a weekend.
        If StartDate.DayOfWeek = DayOfWeek.Saturday Or StartDate.DayOfWeek =
        DayOfWeek.Sunday Then
            NumberOfBusinessDays -= 1
        End If

        For index = 1 To NumberOfBusinessDays
            Select Case StartDate.DayOfWeek
                Case DayOfWeek.Sunday
                    StartDate = StartDate.AddDays(2)
                Case DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday,
                    DayOfWeek.Thursday, DayOfWeek.Friday
                    StartDate = StartDate.AddDays(1)
                Case DayOfWeek.Saturday
                    StartDate = StartDate.AddDays(3)

            End Select

        Next

        'check to see if the end date is on a weekend.
        'If so move it ahead to Monday.
        'You could also bump it back to the Friday before if you desired to. 
        'Just change the code to -2 and -1.
        If StartDate.DayOfWeek = DayOfWeek.Saturday Then
            StartDate = StartDate.AddDays(2)
        ElseIf StartDate.DayOfWeek = DayOfWeek.Sunday Then
            StartDate = StartDate.AddDays(1)
        End If

        Return StartDate

    End Function
    Public Sub ErrorHandling(ByVal msg As String,
       ByVal stkTrace As String, ByVal title As String)

        'check and make the directory if necessary; this is set to look in 
        'the Application folder, you may wish to place the error log in 
        'another Location depending upon the user's role and write access to 
        'different areas of the file system
        If Not System.IO.Directory.Exists(Application.StartupPath &
    "\Errors\") Then
            System.IO.Directory.CreateDirectory(Application.StartupPath &
            "\Errors\")
        End If

        'check the file
        Dim fs As FileStream = New FileStream(Application.StartupPath &
        "\Errors\errlog.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim s As StreamWriter = New StreamWriter(fs)
        s.Close()
        fs.Close()

        'log it
        Dim fs1 As FileStream = New FileStream(Application.StartupPath &
        "\Errors\errlog.txt", FileMode.Append, FileAccess.Write)

        Dim s1 As StreamWriter = New StreamWriter(fs1)

        s1.Write("MachineName: " & Environment.MachineName & vbCrLf)
        s1.Write("Title: " & title & vbCrLf)
        s1.Write("Message: " & msg & vbCrLf)
        s1.Write("StackTrace: " & stkTrace & vbCrLf)
        s1.Write("Date/Time: " & DateTime.Now.ToString() & vbCrLf)
        s1.Write("================================================" & vbCrLf)
        s1.Close()
        fs1.Close()

    End Sub
End Module

