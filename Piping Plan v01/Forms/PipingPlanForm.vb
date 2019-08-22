Imports System.Data.OleDb
Imports ExcelDataReader
Imports System.IO
Imports Microsoft.Reporting.WinForms

'tlac stitkov
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Printing
Imports System.Windows.Forms.DataVisualization.Charting

Public Class PipingPlanForm

    'Private dt As DataTable
    'Private filteredPN As Integer
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        populateGridRezanie()
        populateGridOhybanie()
        populategridPriprava()
        populateGridSpajkovanie()
        populateGridInziniering()
        populateGridKanban()
        populateGridPNTyp()
        FillCboSkupina()
    End Sub
    Public Sub Form1_Load()

        populateGridRezanie()
        populateGridOhybanie()
        populategridPriprava()
        populateGridSpajkovanie()
        populateGridInziniering()
        populateGridKanban()
        populateGridPNTyp()
        'populategridVychystanie()
        FillCboSkupina()
        LoginForm1.Visible = False

    End Sub
    Private Sub tabOhybanie_click(sender As Object, e As EventArgs) Handles tabPiping.Click
        SetLinesColorOhybanie()
    End Sub

    Private Sub tabRezanie_click(sender As Object, e As EventArgs) Handles tabPiping.Click
        SetLinesColorRezanie()
    End Sub

    Private Sub tabPriprava_click(sender As Object, e As EventArgs) Handles tabPiping.Click
        SetLinesColorPriprava()
    End Sub

    Private Sub tabSpajkovanie_click(sender As Object, e As EventArgs) Handles tabPiping.Click
        SetLinesColorSpajkovanie()
    End Sub

    Private Function populateGridRezanie()

        'nastavenie atributu pre oznacovanie celeho riadku
        gridRezanie.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        'pri spusteni okna chceme, aby datagridview bol odfiltrovany len pre jedinu otvorenu polozku
        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        If cn1.State = ConnectionState.Closed Then cn1.Open()

        'Dim queryRezanie As String = "SELECT DISTINCT tblHlavnyPlan.DatumStavby, tblKompletnyzoznam.PN, tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka, tblHlavnyPlan.QTY, " _
        '    & "tblHlavnyPlan.RezanieStatus, tblKompletnyzoznam.OhybRovna, tblKanban.Kanban, tblRezanieStatus.RezanieStatus, IIf([tblRezanieStatus.RezanieStatus]=True,DateValue([ulozene])) AS HotovoCas, " _
        '    & "tblKompletnyzoznam.ID, tblHlavnyPlan.JednotkaWJ, tblHlavnyPlan.PipingWJ " _
        '    & "FROM ((tblHlavnyPlan LEFT JOIN tblKompletnyzoznam ON tblHlavnyPlan.PN = tblKompletnyzoznam.PN) LEFT JOIN tblKanban ON (tblKompletnyzoznam.Dlzka = tblKanban.Dlzka) " _
        '    & "AND (tblKompletnyzoznam.Hrubka = tblKanban.Hrubka) AND (tblKompletnyzoznam.Priemer = tblKanban.Priemer)) LEFT JOIN tblRezanieStatus ON (tblHlavnyPlan.PipingWJ = tblRezanieStatus.PipingWJ) " _
        '    & "AND (tblHlavnyPlan.DatumStavby = tblRezanieStatus.DatumStavby) AND (tblHlavnyPlan.PN = tblRezanieStatus.PN) AND (tblHlavnyPlan.JednotkaWJ = tblRezanieStatus.JednotkaWJ) " _
        '    & "WHERE (((tblKompletnyzoznam.Priemer)=22) AND ((tblKompletnyzoznam.Hrubka)=1.5) AND ((tblKompletnyzoznam.OhybRovna)='Ohyb') AND ((tblKanban.Kanban) Is Null Or (tblKanban.Kanban)<>Yes) " _
        '    & "AND ((IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'dd\.mm hh:nn')))>DateValue((Now()-1)) Or (IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'dd\.mm hh:nn'))) Is Null)) " _
        '    & "OR (((tblKompletnyzoznam.Priemer)>=42) AND ((tblKompletnyzoznam.Dlzka)>=300) AND ((tblKompletnyzoznam.OhybRovna)='Rovna') AND ((tblKanban.Kanban) Is Null Or (tblKanban.Kanban)<>Yes) " _
        '    & "AND ((IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'dd\.mm hh:nn')))>DateValue((Now()-1)) Or (IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'dd\.mm hh:nn'))) Is Null)) " _
        '    & "OR (((tblKompletnyzoznam.Priemer)>22) AND ((tblKompletnyzoznam.OhybRovna)='Ohyb') AND ((tblKanban.Kanban) Is Null Or (tblKanban.Kanban)<>Yes) " _
        '    & "AND ((IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'dd\.mm hh:nn')))>DateValue((Now()-1)) Or (IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'dd\.mm hh:nn'))) Is Null)) " _
        '    & "ORDER BY tblHlavnyPlan.DatumStavby DESC , tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka;"


        Dim queryRezanie As String = "SELECT [qrySubGridRezanie 02].DatumStavby, [qrySubGridRezanie 02].PN, [qrySubGridRezanie 02].Priemer, [qrySubGridRezanie 02].Hrubka, [qrySubGridRezanie 02].Dlzka, " _
            & "[qrySubGridRezanie 02].QTY, [qrySubGridRezanie 02].RezanieStatus, [qrySubGridRezanie 02].OhybRovna, [qrySubGridRezanie 02].Kanban, tblRezanieStatus.RezanieStatus, " _
            & "IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date')) AS HotovoCas, [qrySubGridRezanie 02].ID, [qrySubGridRezanie 02].JednotkaWJ " _
            & "FROM [qrySubGridRezanie 02] LEFT JOIN tblRezanieStatus ON ([qrySubGridRezanie 02].DatumStavby = tblRezanieStatus.DatumStavby) AND ([qrySubGridRezanie 02].PN = tblRezanieStatus.PN) " _
            & "AND ([qrySubGridRezanie 02].JednotkaWJ = tblRezanieStatus.JednotkaWJ) And ([qrySubGridRezanie 02].ID = tblRezanieStatus.IDKomplZ) " _
            & "WHERE (((IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date')))>DateValue((Now()-1)) Or (IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date'))) Is Null)) " _
            & "OR (((IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date')))>DateValue((Now()-1)) Or (IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date'))) Is Null)) " _
            & "OR (((IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date')))>DateValue((Now()-1)) Or (IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date'))) Is Null)) " _
            & "ORDER BY [qrySubGridRezanie 02].DatumStavby DESC , [qrySubGridRezanie 02].Priemer, [qrySubGridRezanie 02].Hrubka, [qrySubGridRezanie 02].Dlzka;"

        Dim gridcmdRezanie As OleDbCommand = New OleDbCommand(queryRezanie, cn1)
        Dim sdaRezanie As OleDbDataAdapter = New OleDbDataAdapter(gridcmdRezanie)
        gridcmdRezanie.Parameters.AddWithValue("@OhybRovna", "Ohyb")
        Dim dtRezanie As DataTable = New DataTable()
        sdaRezanie.Fill(dtRezanie)

        gridRezanie.DataSource = dtRezanie
        If cn1.State = ConnectionState.Open Then cn1.Close()

        'gridRezanie.Columns(0).DefaultCellStyle.Format = "dd.MM"

        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridRezanie.RowTemplate.Height = 35

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridRezanie
            .Columns(0).Width = 110
            .Columns(1).Width = 110
            .Columns(2).Width = 110
            .Columns(3).Width = 110
            .Columns(4).Width = 110
            .Columns(6).Visible = False
            '.Columns(9).Visible = False
            .Columns(10).Visible = False
            .Columns(11).Visible = False
            .Columns(12).Visible = False
            '.Columns(13).Visible = False


            .Columns(0).DefaultCellStyle.Format = "dd.MM"
            .Columns(10).DefaultCellStyle.Format = "dd.MM hh:mm"

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(0).HeaderCell.Value = "Dátum Stavby"
            .Columns(1).HeaderCell.Value = "Piping PN"
            .Columns(2).HeaderCell.Value = "Priemer [mm]"
            .Columns(3).HeaderCell.Value = "Hrúbka [mm]"
            .Columns(4).HeaderCell.Value = "Dĺžka [mm]"
            .Columns(5).HeaderCell.Value = "Množstvo"
            .Columns(6).HeaderCell.Value = "Hotovo" 'PN hotovo
            .Columns(7).HeaderCell.Value = "Ohyb/Rovná"
            .Columns(8).HeaderCell.Value = "Kanban" 'Kanban
            .Columns(9).HeaderCell.Value = "Hotovo" 'tblRezanieStatus
            .Columns(10).HeaderCell.Value = "Cas ukoncenia"
            '.Columns(11).HeaderCell.Value = "ID KomplPl"
            '.Columns(12).HeaderCell.Value = "WJ Jednotky"
            '.Columns(13).HeaderCell.Value = "WJ Pipingu"



            .ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)

        End With

        SetLinesColorRezanie()

        txtRezanieCount.Text = gridRezanie.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()
        Return dtRezanie
    End Function


    Public Sub populateGridInziniering()


        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        If cn1.State = ConnectionState.Closed Then cn1.Open()

        Dim queryInziniering As String = "Select DISTINCT tblHlavnyPlan.PN, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka, " _
            & "tblKompletnyzoznam.OhybRovna, tblKanban.Kanban, tblKompletnyzoznam.Skontrolovane, tblPNTyp.Typ " _
            & "FROM ((tblHlavnyPlan LEFT JOIN tblKompletnyzoznam On tblHlavnyPlan.PN = tblKompletnyzoznam.PN) LEFT JOIN tblKanban On (tblKompletnyzoznam.Dlzka = tblKanban.Dlzka) " _
            & "And (tblKompletnyzoznam.Hrubka = tblKanban.Hrubka) And (tblKompletnyzoznam.Priemer = tblKanban.Priemer)) LEFT JOIN tblPNTyp On tblHlavnyPlan.PN = tblPNTyp.PN " _
            & "WHERE (((tblKompletnyzoznam.Priemer) Is Null) And ((tblPNTyp.Typ) Is Null)) Or (((tblKompletnyzoznam.Skontrolovane)=No) And ((tblPNTyp.Typ) Is Null)) " _
            & "Or (((tblKompletnyzoznam.Dlzka) Is Null) And ((tblPNTyp.Typ) Is Null)) " _
            & "ORDER BY tblHlavnyPlan.PN, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer;"

        Dim gridcmdInziniering As OleDbCommand = New OleDbCommand(queryInziniering, cn1)
        Dim sdaInziniering As OleDbDataAdapter = New OleDbDataAdapter(gridcmdInziniering)
        Dim dtInziniering As DataTable = New DataTable()
        sdaInziniering.Fill(dtInziniering)
        gridInziniering.DataSource = dtInziniering
        If cn1.State = ConnectionState.Open Then cn1.Close()

        gridInziniering.Columns(0).DefaultCellStyle.Format = "dd.MM"

        'nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridInziniering

            .RowTemplate.Height = 25

            .Columns(0).Width = 90
            .Columns(1).Width = 90
            .Columns(2).Width = 90
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 90
            .Columns(6).Width = 90
            .Columns(7).Width = 120
            .Columns(8).Visible = False

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter



            .RowHeadersVisible = False
            'ID,PN,QTY,subPN,Pocetnost, Priemer, Hrubka, Dlzka, OhybRovna, Kanban, AMOB
            .Columns(0).HeaderCell.Value = "Piping"
            .Columns(1).HeaderCell.Value = "subPiping"
            .Columns(2).HeaderCell.Value = "Priemer [mm]"
            .Columns(3).HeaderCell.Value = "Hrúbka [mm]"
            .Columns(4).HeaderCell.Value = "Dĺžka [mm]"
            .Columns(5).HeaderCell.Value = "Ohyb / Rovná"
            .Columns(6).HeaderCell.Value = "Kanban"
            .Columns(7).HeaderCell.Value = "Skontrolované"

            .ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)

        End With



        txtInzinieringCount.Text = gridInziniering.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()
    End Sub

    Private Sub populateGridKanban()


        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        If cn1.State = ConnectionState.Closed Then cn1.Open()
        Dim queryKanban As String = "Select * from tblKanban;"

        Dim gridcmdKanban As OleDbCommand = New OleDbCommand(queryKanban, cn1)
        Dim sdaKanban As OleDbDataAdapter = New OleDbDataAdapter(gridcmdKanban)
        Dim dtKanban As DataTable = New DataTable()
        sdaKanban.Fill(dtKanban)
        gridKanban.DataSource = dtKanban
        If cn1.State = ConnectionState.Open Then cn1.Close()


        'nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridKanban

            .RowTemplate.Height = 25

            .Columns(0).Visible = False
            .Columns(1).Width = 70
            .Columns(2).Width = 70
            .Columns(3).Width = 70
            .Columns(4).Visible = False

            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


            .RowHeadersVisible = False
            'ID,PN,QTY,subPN,Pocetnost, Priemer, Hrubka, Dlzka, OhybRovna, Kanban, AMOB
            .Columns(1).HeaderCell.Value = "Priemer [mm]"
            .Columns(2).HeaderCell.Value = "Hrúbka [mm]"
            .Columns(3).HeaderCell.Value = "Dĺžka [mm]"


            .ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)

        End With



        txtKanbanCount.Text = gridKanban.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()
    End Sub

    Private Sub populateGridPNTyp()

        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        If cn1.State = ConnectionState.Closed Then cn1.Open()
        Dim queryPNTyp As String = "Select ID,PN,Description,Typ from tblPNTyp ORDER BY PN;"

        Dim gridcmdPNTyp As OleDbCommand = New OleDbCommand(queryPNTyp, cn1)
        Dim sdaPNTyp As OleDbDataAdapter = New OleDbDataAdapter(gridcmdPNTyp)
        Dim dtPNTyp As DataTable = New DataTable()
        sdaPNTyp.Fill(dtPNTyp)
        gridPNTyp.DataSource = dtPNTyp
        If cn1.State = ConnectionState.Open Then cn1.Close()


        'nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridPNTyp

            .RowTemplate.Height = 25

            .Columns(0).Visible = False
            .Columns(1).Width = 90
            .Columns(2).Width = 250
            .Columns(3).Width = 120



            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter




            .RowHeadersVisible = False
            'ID,PN,QTY,subPN,Pocetnost, Priemer, Hrubka, Dlzka, OhybRovna, PNTyp, AMOB
            .Columns(1).HeaderCell.Value = "Piping PN"
            .Columns(2).HeaderCell.Value = "Popis"
            .Columns(3).HeaderCell.Value = "Typ trubky"



            .ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)

        End With



        txtPNTypCount.Text = gridPNTyp.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()

    End Sub

    Public Function populateGridOhybanie()

        'nastavenie atributu pre oznacovanie celeho riadku
        gridOhybanie.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        'vyplnenie gridu pre Ohybanie
        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        cn1.Open()


        'Dim query As String = "SELECT qrySubGridOhybanie.IDHlavnyPlan AS ID, qrySubGridOhybanie.DatumStavby, qrySubGridOhybanie.Skupina, qrySubGridOhybanie.PN, qrySubGridOhybanie.JednotkaWJ, " _
        '    & "qrySubGridOhybanie.QTY, qrySubGridOhybanie.subPN, qrySubGridOhybanie.Priemer, qrySubGridOhybanie.Hrubka, qrySubGridOhybanie.Dlzka, qrySubGridOhybanie.OhybRovna, " _
        '    & "tblOhybanieStatus.OhybanieStatus, tblOhybanieStatus.Ulozene, qrySubGridOhybanie.Narezane, qrySubGridOhybanie.IDKomplZ, " _
        '    & "IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([ulozene]),Null) AS HotovoCas " _
        '    & "FROM tblOhybanieStatus RIGHT JOIN subQRYGridOhybanie ON (tblOhybanieStatus.IDHlavnyPlan = qrySubGridOhybanie.IDHlavnyPlan) And (tblOhybanieStatus.IDKomplZ = qrySubGridOhybanie.IDKomplZ) " _
        '    & "WHERE (((qrySubGridOhybanie.Priemer)>=12) And ((qrySubGridOhybanie.OhybRovna)='Ohyb') " _
        '    & "AND ((IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([ulozene]),Null))>DateValue((Now()-1)) Or (IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([ulozene]),Null)) Is Null)) OR (((qrySubGridOhybanie.Priemer)>=42) " _
        '    & "AND ((qrySubGridOhybanie.Dlzka)>=300) And ((qrySubGridOhybanie.OhybRovna)='Rovna') AND ((IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([ulozene]),Null))>DateValue((Now()-1)) Or (IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([ulozene]),Null)) Is Null)) " _
        '    & "ORDER BY qrySubGridOhybanie.DatumStavby DESC , qrySubGridOhybanie.Priemer, qrySubGridOhybanie.PN, qrySubGridOhybanie.subPN;"

        Dim query As String = "SELECT qrySubGridOhybanie.IDHlavnyPlan AS ID, qrySubGridOhybanie.DatumStavby, qrySubGridOhybanie.Skupina, qrySubGridOhybanie.PN, " _
            & "qrySubGridOhybanie.JednotkaWJ, qrySubGridOhybanie.QTY, qrySubGridOhybanie.subPN, qrySubGridOhybanie.Priemer, qrySubGridOhybanie.Hrubka, qrySubGridOhybanie.Dlzka, " _
            & "qrySubGridOhybanie.OhybRovna, tblOhybanieStatus.OhybanieStatus, tblOhybanieStatus.Ulozene, IIf(IsNull([tblKanban.Kanban])=True " _
            & "Or [tblKanban.Kanban]=No,IIf([tblRezanieStatus.RezanieStatus]=Yes,'ANO','NIE'),'ANO-Kanban') As Narezane, qrySubGridOhybanie.IDKomplZ, " _
            & "IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null) As HotovoCas " _
            & "FROM (qrySubGridOhybanie LEFT JOIN tblOhybanieStatus On (qrySubGridOhybanie.IDKomplZ = tblOhybanieStatus.IDKomplZ) " _
            & "And (qrySubGridOhybanie.IDHlavnyPlan = tblOhybanieStatus.IDHlavnyPlan)) LEFT JOIN tblRezanieStatus On (qrySubGridOhybanie.JednotkaWJ = tblRezanieStatus.JednotkaWJ) " _
            & "And (qrySubGridOhybanie.PN = tblRezanieStatus.PN) And (qrySubGridOhybanie.IDKomplZ = tblRezanieStatus.IDKomplZ) " _
            & "WHERE (((qrySubGridOhybanie.Priemer)>=12) And ((qrySubGridOhybanie.OhybRovna)='Ohyb') AND ((IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null))>DateValue((Now()-1)) " _
            & "Or (IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null)) Is Null)) Or (((qrySubGridOhybanie.Priemer)>=42) And ((qrySubGridOhybanie.Dlzka)>=300) " _
            & "And ((qrySubGridOhybanie.OhybRovna)='Rovna') AND ((IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null))>DateValue((Now()-1)) " _
            & "Or (IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null)) Is Null)) " _
            & "ORDER BY qrySubGridOhybanie.DatumStavby DESC , qrySubGridOhybanie.Priemer, qrySubGridOhybanie.PN, qrySubGridOhybanie.subPN;"


        Dim gridcmd As OleDbCommand = New OleDbCommand(query, cn1)
        Dim sda As OleDbDataAdapter = New OleDbDataAdapter(gridcmd)
        Dim dt As DataTable = New DataTable()
        cn1.Close()
        sda.Fill(dt)
        gridOhybanie.DataSource = dt
        If cn1.State = ConnectionState.Open Then cn1.Close()

        'Dim lnk As New DataGridViewLinkColumn()
        'gridOhybanie.Columns.Add(lnk)

        With gridOhybanie
            .RowTemplate.Height = 35
            .ForeColor = Color.Black
            .Columns(0).Visible = False

            .Columns(1).Width = 80
            .Columns(2).Width = 120
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 75
            .Columns(6).Width = 80
            .Columns(7).Width = 80
            .Columns(8).Width = 80
            .Columns(9).Width = 80
            .Columns(10).Width = 80
            .Columns(13).Width = 130


            .Columns(12).Visible = False
            .Columns(14).Visible = False
            .Columns(15).Visible = False
            '.Columns(13).Visible = False

            .Columns(1).DefaultCellStyle.Format = "dd.MM"
            '.Columns(11).DefaultCellStyle.Format = "dd.MM"

            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(1).HeaderCell.Value = "Dátum Stavby"
            .Columns(2).HeaderCell.Value = "Skupina"
            .Columns(3).HeaderCell.Value = "Piping PN"
            .Columns(4).HeaderCell.Value = "WJ Jednotky"
            .Columns(5).HeaderCell.Value = "Množ."
            .Columns(6).HeaderCell.Value = "sub Piping"
            .Columns(7).HeaderCell.Value = "Priemer [mm]"
            .Columns(8).HeaderCell.Value = "Hrúbka [mm]"
            .Columns(9).HeaderCell.Value = "Dĺžka [mm]"
            .Columns(10).HeaderCell.Value = "Ohýbaná / Rovná"
            .Columns(11).HeaderCell.Value = "Hotovo"

            .Columns(12).HeaderCell.Value = "ulozene v tblOhybanieStatus"
            .Columns(13).HeaderCell.Value = "Narezane"

            .ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        End With

        txtOhybanieCount.Text = gridOhybanie.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()
        Return dt
    End Function

    Private Function populategridPriprava()

        'pri spusteni okna chceme, aby datagridview bol odfiltrovany len pre jedinu otvorenu polozku

        cn1.Open()

        'Dim query As String = "SELECT DISTINCT qryPripravaFilterHiLevel.DatumStavby, tblHlavnyPlan.Skupina, qryPripravaFilterHiLevel.PN, tblHlavnyPlan.QTY, " _
        '                        & "tblHlavnyPlan.JednotkaWJ, qryPripravaFilterHiLevel.CountOfOhybanieStatus, qryPripravaFilterHiLevel.UncompletedPNs, tblHlavnyPlan.PripravaStatus " _
        '                        & "FROM tblHlavnyPlan RIGHT JOIN qryPripravaFilterHiLevel ON tblHlavnyPlan.PN = qryPripravaFilterHiLevel.PN;"

        Dim query As String = "SELECT tblHlavnyPlan.ID, tblHlavnyPlan.DatumStavby, tblHlavnyPlan.Skupina, tblHlavnyPlan.PN, tblHlavnyPlan.QTY, tblHlavnyPlan.JednotkaWJ, " _
                            & "tblHlavnyPlan.PripravaStatus " _
                            & "FROM tblHlavnyPlan " _
                            & "WHERE (((tblHlavnyPlan.OhybanieStatus)=Yes)) " _
                            & "ORDER BY tblHlavnyPlan.DatumStavby DESC , tblHlavnyPlan.Skupina, tblHlavnyPlan.PN;"

        Dim gridcmd As OleDbCommand = New OleDbCommand(query, cn1)
        Dim sda As OleDbDataAdapter = New OleDbDataAdapter(gridcmd)
        Dim dt As DataTable = New DataTable()
        cn1.Close()
        sda.Fill(dt)
        gridPriprava.DataSource = dt
        If cn1.State = ConnectionState.Open Then cn1.Close()

        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridPriprava.RowTemplate.Height = 35

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridPriprava

            .Columns(1).Width = 90
            .Columns(2).Width = 90
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 90
            .Columns(6).Width = 90

            .Columns(0).Visible = False
            '.Columns(7).Visible = False

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(1).HeaderCell.Value = "Dátum Stavby"
            .Columns(2).HeaderCell.Value = "Skupina"
            .Columns(3).HeaderCell.Value = "Piping PN"
            .Columns(4).HeaderCell.Value = "Množstvo"
            .Columns(5).HeaderCell.Value = "WJ Jednotky"
            .Columns(6).HeaderCell.Value = "Hotovo"

        End With
        gridPriprava.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        txtPripravaCount.Text = gridPriprava.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()
        Return dt
    End Function

    Private Function populateGridSpajkovanie()

        'pri spusteni okna chceme, aby datagridview bol odfiltrovany len pre jedinu otvorenu polozku

        cn1.Open()
        'Dim query As String = "Select DatumStavby,Skupina,PN,QTY,JednotkaWJ FROM tblHlavnyPlan " _
        '                    & " WHERE OhybanieStatus=YES And PripravaStatus=NO ORDER BY DatumStavby DESC, Skupina, PN;"

        'Dim query As String = "SELECT DatumStavby, Skupina, PN, QTY, JednotkaWJ, " _
        '        & "Count(PripravaStatus) AS TotalPNs, Count([PripravaStatus])+Sum([PripravaStatus]) AS UncompletedPNs " _
        '        & "FROM tblHlavnyPlan " _
        '        & "WHERE ((([SpajkovanieStatus])=No)) " _
        '        & "GROUP BY DatumStavby, Skupina, QTY, PN, JednotkaWJ " _
        '        & "HAVING (((Count([PripravaStatus])+Sum([PripravaStatus]))=0));"

        Dim query As String = "SELECT tblHlavnyPlan.ID, tblHlavnyPlan.DatumStavby, tblHlavnyPlan.Skupina, tblHlavnyPlan.PN, tblHlavnyPlan.QTY, " _
                            & "tblHlavnyPlan.JednotkaWJ, tblHlavnyPlan.SpajkovanieStatus " _
                            & "FROM tblHlavnyPlan " _
                            & "WHERE (((tblHlavnyPlan.PripravaStatus)=Yes)) " _
                            & "ORDER BY tblHlavnyPlan.DatumStavby DESC , tblHlavnyPlan.Skupina, tblHlavnyPlan.PN;"

        Dim gridcmd As OleDbCommand = New OleDbCommand(query, cn1)
        Dim sda As OleDbDataAdapter = New OleDbDataAdapter(gridcmd)
        Dim dt As DataTable = New DataTable()
        cn1.Close()
        sda.Fill(dt)
        gridSpajkovanie.DataSource = dt
        If cn1.State = ConnectionState.Open Then cn1.Close()

        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridSpajkovanie.RowTemplate.Height = 35

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridSpajkovanie

            .Columns(1).Width = 90
            .Columns(2).Width = 90
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 90
            .Columns(6).Width = 90

            .Columns(0).Visible = False
            '.Columns(7).Visible = False

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(1).HeaderCell.Value = "Dátum Stavby"
            .Columns(2).HeaderCell.Value = "Skupina"
            .Columns(3).HeaderCell.Value = "Piping PN"
            .Columns(4).HeaderCell.Value = "Množstvo"
            .Columns(5).HeaderCell.Value = "WJ Jednotky"
            .Columns(6).HeaderCell.Value = "Hotovo"

        End With
        gridSpajkovanie.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        txtSpajkovanieCount.Text = gridSpajkovanie.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()
        Return dt
    End Function

    Private Function populategridVychystanie()

        If cboSkupinaVychystanie.SelectedValue = Nothing Or dtpDatumStavby.Value.Date = Nothing Then Return Nothing

        'pri spusteni okna chceme, aby datagridview bol odfiltrovany len pre jedinu otvorenu polozku

        If cn1.State = ConnectionState.Closed Then cn1.Open()

        'Dim query As String = "SELECT tblHlavnyPlan.ID, tblHlavnyPlan.DatumStavby, tblHlavnyPlan.Skupina, tblHlavnyPlan.JednotkaWJ, tblHlavnyPlan.PN, tblHlavnyPlan.QTY, " _
        '        & "tblHlavnyPlan.VyplanovanyWJ, tblHlavnyPlan.VyplanovaneMnozstvo, " _
        '        & "tblHlavnyPlan.VychystanieStatus, tblSkupina.Linka, tblDocuments.Name AS Vykres,IIf(tblHlavnyPlan.ZaplanovanyWJ=TRUE,'Zobrat zo skladu',FALSE) AS ZaplanovanyPN " _
        '        & "FROM tblDocuments RIGHT JOIN (tblSkupina INNER JOIN tblHlavnyPlan ON tblSkupina.Skupina = tblHlavnyPlan.Skupina) ON tblDocuments.Title = tblHlavnyPlan.PN " _
        '        & "WHERE tblHlavnyPlan.DatumStavby=@DatumStavby AND tblSkupina.Linka=@Linka " _
        '        & "ORDER BY tblHlavnyPlan.Skupina, tblHlavnyPlan.PN, tblHlavnyPlan.QTY;"

        Dim query As String = "SELECT tblHlavnyPlan.ID, tblHlavnyPlan.DatumStavby, tblHlavnyPlan.Skupina, tblHlavnyPlan.JednotkaWJ, tblHlavnyPlan.PN, " _
            & "tblHlavnyPlan.QTY, tblHlavnyPlan.VyplanovanyWJ, tblHlavnyPlan.VyplanovaneMnozstvo,IIf(tblHlavnyPlan.ZaplanovanyWJ=True,'Áno - zobrat zo skladu','Nie') AS ZaplanovanyPN, tblSkupina.Linka, " _
            & "tblHlavnyPlan.VychystanieStatus " _
            & "From tblSkupina INNER Join tblHlavnyPlan On tblSkupina.Skupina = tblHlavnyPlan.Skupina " _
            & "Where (((tblHlavnyPlan.DatumStavby) =[@DatumStavby]) And ((tblSkupina.Linka)=[@Linka])) " _
            & "Order By tblHlavnyPlan.Skupina, tblHlavnyPlan.PN, tblHlavnyPlan.QTY;"

        Dim gridcmdVychystanie As OleDbCommand = New OleDbCommand(query, cn1)
        Dim sda As OleDbDataAdapter = New OleDbDataAdapter(gridcmdVychystanie)

        'gridcmdVychystanie.Parameters.AddWithValue("@Nevychystavat", cboSkupinaVychystanie.SelectedValue.ToString)
        gridcmdVychystanie.Parameters.AddWithValue("@DatumStavby", Format(dtpDatumStavby.Value.Date, "dd/MM/yyyy"))
        'MsgBox(cboSkupinaVychystanie.SelectedValue.ToString)
        gridcmdVychystanie.Parameters.AddWithValue("@Linka", cboSkupinaVychystanie.SelectedValue.ToString)
        Dim dtVychystanie As DataTable = New DataTable()
        cn1.Close()
        sda.Fill(dtVychystanie)
        gridVychystanie.DataSource = dtVychystanie
        If cn1.State = ConnectionState.Open Then cn1.Close()

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridVychystanie

            .RowTemplate.Height = 35

            .Columns(1).Width = 120
            .Columns(2).Width = 120
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 120
            '.Columns(6).Width = 100
            .Columns(7).Width = 70
            .Columns(8).Width = 150
            .Columns(10).Width = 100

            .Columns(0).Visible = False
            .Columns(1).Visible = False
            .Columns(9).Visible = False
            '.Columns(10).Visible = False
            '.Columns(2).Visible = False
            '.Columns(6).Visible = False

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(1).HeaderCell.Value = "Dátum Stavby"
            .Columns(2).HeaderCell.Value = "Skupina"
            .Columns(3).HeaderCell.Value = "WJ Jednotky"
            .Columns(4).HeaderCell.Value = "Piping PN"
            .Columns(5).HeaderCell.Value = "Pôvodné množstvo"
            .Columns(6).HeaderCell.Value = "Vyplánovaný WJ"
            .Columns(7).HeaderCell.Value = "Odložiť množstvo"
            .Columns(8).HeaderCell.Value = "Zaplánovaný WJ"
            .Columns(10).HeaderCell.Value = "Vychystanie Hotovo"

            .ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)

        End With

        'txtPripravaCount.Text = gridVychystanie.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()
        SetLinesColorVyplanovany()
        SetLinesColorVychystanie()
        Return dtVychystanie
    End Function

    Private Sub populategridReport()


        'pri spusteni okna chceme, aby datagridview bol odfiltrovany len pre jedinu otvorenu polozku
        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        cn1.Open()
        Dim queryStatus As String = "Select DatumStavby,Skupina,PipingWJ,PN,QTY,subPN,Pocetnost, Priemer, Hrubka, Dlzka, OhybRovna, Kanban, AMOB,RezanieStatus,RezanieCompleted,OhybanieStatus,OhybanieCompleted,PripravaStatus,PripravaCompleted,SpajkovanieStatus,SpajkovanieCompleted FROM tblHlavnyPlan;"

        Dim gridCmdStatus As OleDbCommand = New OleDbCommand(queryStatus, cn1)
        'gridcmd.Parameters.AddWithValue("@today", MyDateIsToday)
        'gridcmd.Parameters.AddWithValue("@MCName", Environment.MachineName)
        Dim sdaStatus As OleDbDataAdapter = New OleDbDataAdapter(gridCmdStatus)
        Dim dtStatus As DataTable = New DataTable()
        sdaStatus.Fill(dtStatus)
        gridReport.DataSource = dtStatus
        If cn1.State = ConnectionState.Open Then cn1.Close()

        gridRezanie.Columns(0).DefaultCellStyle.Format = "dd.MM"

    End Sub


    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click

        Try
            If cn1.State = ConnectionState.Closed Then cn1.Open()
            '----------------------------------------------------------------------------------------------------
            'import planu z excelu vyberom suboru cez openfiledialog
            Dim OpenDLG As New OpenFileDialog
            OpenDLG.ShowDialog()
            Dim xPath As String = OpenDLG.FileName
            Dim wPathLong = System.IO.Path.GetDirectoryName(xPath)
            Dim stream As FileStream = File.Open(xPath, FileMode.Open, FileAccess.Read)
            Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
            Dim result As DataSet = excelReader.AsDataSet()
            excelReader.Close()
            gridHlavnyPlan.DataSource = result.Tables(0)
        Catch ex As Exception
            If MessageBox.Show(ex.Message, NazovAplikacie, MessageBoxButtons.OK, MessageBoxIcon.Question) = vbOK Then
                ErrorHandling(ex.Message, ex.StackTrace, ex.Source)
                Exit Sub
            End If

            If cn1.State = ConnectionState.Open Then cn1.Close()
        End Try



        'premenovat hlavicky stlpcov z hodnot z prveho riadku kazdeho stlpca 
        Dim a As Integer = gridHlavnyPlan.ColumnCount
        For i = 0 To a - 1
            gridHlavnyPlan.Columns(i).HeaderCell.Value = gridHlavnyPlan.Rows(0).Cells(i).Value
        Next
        gridHlavnyPlan.Rows.Remove(gridHlavnyPlan.Rows(0)) 'vymazanie prveho riadku, kedze su tam len nazvy stlpcov
        gridHlavnyPlan.Columns(0).DefaultCellStyle.Format = "dd.MM"
        txtPlanCount.Text = gridHlavnyPlan.RowCount

        gridHlavnyPlan.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        If (MessageBox.Show("Data z excelu nahrate do zoznamu, pokracovat?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbNo Then
            cn1.Close()

            Exit Sub
        Else
        End If

        '----------------------------------------------------------------------------------------------------

        'nasype vsetky importovane riadky z excelu do tblHlavnyPlan tabulky
        Dim rowcount As Double = 0
        For i As Integer = 0 To gridHlavnyPlan.Rows.Count - 1 Step +1
            ' If IsDBNull(gridHlavnyPlan.Rows(i).Cells(0).Value()) Then
            If IsDBNull(gridHlavnyPlan.Rows(i).Cells(0).ToString()) Then
                Continue For
            Else


                Dim cmdImport As New OleDbCommand("INSERT INTO tblHlavnyPlan (DatumStavby,Skupina,JednotkaWJ,PipingWJ,PN,QTY,Drazkovanie,RoutingCas) " _
                                                      & "VALUES (@DatumStavby,@Skupina,@JednotkaWJ,@PipingWJ,@PN,@QTY,@Drazkovanie,@RoutingCas)", cn1)
                cmdImport.Parameters.AddWithValue("@DatumStavby", gridHlavnyPlan.Rows(i).Cells(0).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Skupina", gridHlavnyPlan.Rows(i).Cells(1).Value().ToString)
                cmdImport.Parameters.AddWithValue("@JednotkaWJ", gridHlavnyPlan.Rows(i).Cells(2).Value().ToString)
                cmdImport.Parameters.AddWithValue("@PipingWJ", gridHlavnyPlan.Rows(i).Cells(3).Value().ToString)
                cmdImport.Parameters.AddWithValue("@PN", gridHlavnyPlan.Rows(i).Cells(4).Value().ToString)
                cmdImport.Parameters.AddWithValue("@QTY", gridHlavnyPlan.Rows(i).Cells(5).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Drazkovanie", CBool(gridHlavnyPlan.Rows(i).Cells(6).Value))
                cmdImport.Parameters.AddWithValue("@RoutingCas", gridHlavnyPlan.Rows(i).Cells(7).Value)
                cmdImport.ExecuteNonQuery()

            End If
            rowcount = rowcount + 1
        Next
        txtPlanCount.Text = gridHlavnyPlan.RowCount
        gridHlavnyPlan.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        If (MessageBox.Show(rowcount & " riadkov z excelu bolo ulozenych do Hlavneho Planu, pokracovat?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbNo Then
            cn1.Close()

            Exit Sub
        Else
        End If

        '----------------------------------------------------------------------------------------------------

        'Zluci obsah importovaneho planu s tblKompletnyZoznam
        Dim dt As New DataTable()
        Dim query As String = "SELECT tblHlavnyPlan.DatumStavby, tblHlavnyPlan.Skupina, tblHlavnyPlan.JednotkaWJ, tblHlavnyPlan.PipingWJ, tblHlavnyPlan.PN, tblHlavnyPlan.QTY, " _
            & "tblHlavnyPlan.Drazkovanie, tblHlavnyPlan.RoutingCas, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka, " _
            & "tblKompletnyzoznam.OhybRovna, tblKanban.Kanban, tblKompletnyzoznam.AMOB " _
            & "FROM (tblKompletnyzoznam RIGHT JOIN tblHlavnyPlan ON tblKompletnyzoznam.PN = tblHlavnyPlan.PN) " _
            & "LEFT JOIN tblKanban ON (tblKompletnyzoznam.Dlzka = tblKanban.Dlzka) AND (tblKompletnyzoznam.Hrubka = tblKanban.Hrubka) AND (tblKompletnyzoznam.Priemer = tblKanban.Priemer);"

        Dim cmd As New OleDbCommand(query, cn1)
        Dim adapter As New OleDbDataAdapter(cmd)
        adapter.Fill(dt)
        gridHlavnyPlan.DataSource = dt
        txtPlanCount.Text = gridHlavnyPlan.RowCount

        With gridHlavnyPlan
            .Columns(7).Width = 80

            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Format = "0.00"

        End With

        gridHlavnyPlan.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)

        txtPlanCount.Text = gridHlavnyPlan.RowCount
        If (MessageBox.Show("Importovane data zlucene s Kompletnym Zoznamom, pokračovať v uložení plánu do databázy?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbNo Then
            cn1.Close()
            Exit Sub
        Else
        End If

        populateGridRezanie()
        populateGridOhybanie()
        populateGridInziniering()


    End Sub

    Private Sub btnSavePlan_Click(sender As Object, e As EventArgs)

        'vymaze zaznamy v temp tabulke
        cn1.Open()
        Dim cmdDel As New OleDbCommand("delete from tblTempPlan", cn1)
        cmdDel.ExecuteNonQuery()

        'nasype vsetky zaznamy z importovaneho excelu (denny plan)


        For i As Integer = 0 To gridHlavnyPlan.Rows.Count - 2 Step +1
            If IsDBNull(gridHlavnyPlan.Rows(i).Cells(0).Value()) Then
                Continue For
            Else

                Dim cmd1 As New OleDbCommand("SELECT * from tblHlavnyPlan where PipingWJ=@PipingWJ", cn1)
                'Dim cmd1 As New OleDbCommand("SELECT * from tblHlavnyPlan where PN=@a)", cn1)
                cmd1.Parameters.AddWithValue("@PipingWJ", gridHlavnyPlan.Rows(i).Cells(3).Value().ToString)
                Dim check As Boolean = cmd1.ExecuteScalar()
                If check = True Then
                    Continue For
                Else



                    Dim cmd As New OleDbCommand("INSERT INTO tblTempPlan (DatumStavby,Skupina,JednotkaWJ,PipingWJ,PN,QTY) VALUES (@DatumStavby,@Skupina,@JednotkaWJ,@PipingWJ,@PN,@QTY)", cn1)
                    cmd.Parameters.AddWithValue("@DatumStavby", gridHlavnyPlan.Rows(i).Cells(0).Value().ToString)
                    cmd.Parameters.AddWithValue("@Skupina", CStr(gridHlavnyPlan.Rows(i).Cells(1).Value().ToString))
                    cmd.Parameters.AddWithValue("@JednotkaWJ", gridHlavnyPlan.Rows(i).Cells(2).Value().ToString)
                    cmd.Parameters.AddWithValue("@PipingWJ", gridHlavnyPlan.Rows(i).Cells(3).Value().ToString)
                    cmd.Parameters.AddWithValue("@PN", gridHlavnyPlan.Rows(i).Cells(4).Value().ToString)
                    cmd.Parameters.AddWithValue("@QTY", CInt(gridHlavnyPlan.Rows(i).Cells(5).Value().ToString))
                    cmd.ExecuteNonQuery()
                End If
            End If

        Next
        cn1.Close()
        MessageBox.Show("Plán bol uložený do databázy", NazovAplikacie, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub



    Private Sub gridRezanie_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridRezanie.CellClick

        indexRezanie = e.RowIndex
        Dim i As Integer

        With gridRezanie
            If e.RowIndex >= 0 Then
                i = .CurrentRow.Index
                txtDatumStavbyRezanie.Text = .Rows(i).Cells("DatumStavby").Value.ToString
                txtDatumStavbyRezanie.Text = Format(CDate(txtDatumStavbyRezanie.Text), "dd.MM")
                txtQTYRezanie.Text = .Rows(i).Cells("QTY").Value.ToString
                txtPriemerRezanie.Text = .Rows(i).Cells("Priemer").Value.ToString
                txtHrubkaRezanie.Text = .Rows(i).Cells("Hrubka").Value.ToString
                txtDlzkaRezanie.Text = .Rows(i).Cells("Dlzka").Value.ToString
                txtPNRezanie.Text = .Rows(i).Cells("PN").Value.ToString
                pickPN = .Rows(i).Cells("PN").Value.ToString
            End If
        End With

    End Sub

    Private Sub btnUkonciRezanie_Click(sender As Object, e As EventArgs) Handles btnUkonciRezanie.Click


        Try


            If (MessageBox.Show("Naozaj chcete uzavrieť vyznačenú prácu?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbNo Then
                Exit Sub
            Else
            End If

            Dim MyDateTimeVariable = Date.Now.ToString("dd MMM yyyy HH:mm:ss") 'format musi byt spravne zadefinovany, inak sa nezobrazi v MDB databaze
            cn1.Open()

            For index As Integer = 0 To gridRezanie.SelectedRows.Count - 1
                'Dim rownum As Integer = gridRezanie.CurrentRow.Index
                Dim rownum As Integer = gridRezanie.SelectedRows.Item(index).Index


                Dim cmdImport As New OleDbCommand("INSERT INTO tblRezanieStatus (DatumStavby, Skupina, PN, QTY,subPN, Priemer, Hrubka, Dlzka, JednotkaWJ, IDHlavnyPLan, IDKomplZ, RezanieStatus, Ulozene) " _
                                            & "SELECT tblHlavnyPlan.DatumStavby, tblHlavnyPlan.Skupina, tblHlavnyPlan.PN, tblHlavnyPlan.QTY, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer, " _
                                            & "tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka, tblHlavnyPlan.JednotkaWJ, tblHlavnyPlan.ID AS IDHlavnyPLan, " _
                                            & "tblKompletnyzoznam.ID AS IDKomplZ, True, Now() " _
                                            & "FROM tblKompletnyzoznam RIGHT JOIN tblHlavnyPlan ON tblKompletnyzoznam.PN = tblHlavnyPlan.PN " _
                                            & "WHERE tblHlavnyPlan.DatumStavby=@DatumStavby AND tblKompletnyzoznam.Priemer=@Priemer AND tblKompletnyzoznam.Hrubka=@Hrubka " _
                                            & "AND tblKompletnyzoznam.Dlzka=@Dlzka AND tblKompletnyzoznam.ID = @IDKomplZ AND tblHlavnyPlan.JednotkaWJ=@JednotkaWJ", cn1)
                '                           & "WHERE tblKompletnyzoznam.ID = @IDKomplZ ", cn1)

                cmdImport.Parameters.AddWithValue("@DatumStavby", gridRezanie.Rows(rownum).Cells(0).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Priemer", gridRezanie.Rows(rownum).Cells(2).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Hrubka", gridRezanie.Rows(rownum).Cells(3).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Dlzka", gridRezanie.Rows(rownum).Cells(4).Value().ToString)
                cmdImport.Parameters.AddWithValue("@IDKomplZ", gridRezanie.Rows(rownum).Cells(11).Value().ToString)
                cmdImport.Parameters.AddWithValue("@JednotkaWJ", gridRezanie.Rows(rownum).Cells(12).Value().ToString)

                cmdImport.ExecuteNonQuery()

                'kod zistuje, ci dany PN je kompletne vyrobeny v dany den. Ak ano, tak ho v tblHlavnyPlan ukonci, aby ho videl Ohybac v zozname
                Dim rownumPN As Integer = gridRezanie.CurrentRow.Index

                'Dim queryCheck As String = "SELECT tblHlavnyPlan.DatumStavby, tblKompletnyzoznam.PN, tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka, " _
                '    & "tblHlavnyPlan.QTY, tblHlavnyPlan.RezanieStatus, tblKompletnyzoznam.OhybRovna, tblKanban.Kanban, tblKompletnyzoznam.AMOB, tblRezanieStatus.RezanieStatus " _
                '    & "FROM ((tblHlavnyPlan LEFT JOIN tblKompletnyzoznam ON tblHlavnyPlan.PN = tblKompletnyzoznam.PN) " _
                '    & "LEFT JOIN tblRezanieStatus ON (tblKompletnyzoznam.Dlzka = tblRezanieStatus.Dlzka) AND (tblKompletnyzoznam.Hrubka = tblRezanieStatus.Hrubka) " _
                '    & "AND (tblKompletnyzoznam.Priemer = tblRezanieStatus.Priemer)) LEFT JOIN tblKanban ON (tblKompletnyzoznam.Dlzka = tblKanban.Dlzka) " _
                '    & "AND (tblKompletnyzoznam.Hrubka = tblKanban.Hrubka) AND (tblKompletnyzoznam.Priemer = tblKanban.Priemer) " _
                '    & "WHERE (((tblHlavnyPlan.DatumStavby)=[@DatumStavby]) AND ((tblKompletnyzoznam.PN)=[@PN]) AND ((tblKompletnyzoznam.Priemer)>=12) " _
                '    & "AND ((tblKompletnyzoznam.OhybRovna)='Ohyb') AND ((tblKanban.Kanban) Is Null Or (tblKanban.Kanban)<>Yes) AND ((tblRezanieStatus.RezanieStatus) Is Null)) " _
                '    & "OR (((tblHlavnyPlan.DatumStavby)=[@DatumStavby]) AND ((tblKompletnyzoznam.PN)=[@PN]) AND ((tblKompletnyzoznam.Priemer)>=42) AND ((tblKompletnyzoznam.Dlzka)>=300) " _
                '    & "AND ((tblKompletnyzoznam.OhybRovna)='Rovna') AND ((tblKanban.Kanban) Is Null Or (tblKanban.Kanban)<>Yes) AND ((tblRezanieStatus.RezanieStatus) Is Null)) " _
                '    & "ORDER BY tblHlavnyPlan.DatumStavby DESC , tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka;"

                Dim queryCheck As String = "SELECT tblHlavnyPlan.DatumStavby, tblKompletnyzoznam.PN, tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka, " _
                    & "tblHlavnyPlan.QTY, tblHlavnyPlan.RezanieStatus, tblKompletnyzoznam.OhybRovna, tblKanban.Kanban, tblKompletnyzoznam.AMOB, tblRezanieStatus.RezanieStatus, " _
                    & "tblHlavnyPlan.JednotkaWJ " _
                    & "FROM ((tblHlavnyPlan LEFT JOIN tblKompletnyzoznam ON tblHlavnyPlan.PN = tblKompletnyzoznam.PN) " _
                    & "LEFT JOIN tblKanban ON (tblKompletnyzoznam.Priemer = tblKanban.Priemer) And (tblKompletnyzoznam.Hrubka = tblKanban.Hrubka) And (tblKompletnyzoznam.Dlzka = tblKanban.Dlzka)) " _
                    & "LEFT JOIN tblRezanieStatus ON (tblHlavnyPlan.PN = tblRezanieStatus.PN) And (tblHlavnyPlan.JednotkaWJ = tblRezanieStatus.JednotkaWJ) " _
                    & "WHERE (((tblHlavnyPlan.DatumStavby)=[@DatumStavby]) And ((tblKompletnyzoznam.PN)=[@PN]) And ((tblKompletnyzoznam.OhybRovna)='Ohyb') " _
                    & "AND ((tblKanban.Kanban) Is Null Or (tblKanban.Kanban)<>Yes) AND ((tblRezanieStatus.RezanieStatus) Is Null)) " _
                    & "ORDER BY tblHlavnyPlan.DatumStavby DESC , tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka;"

                Dim gridcmd As OleDbCommand = New OleDbCommand(queryCheck, cn1)
                Dim sda As OleDbDataAdapter = New OleDbDataAdapter(gridcmd)
                Dim dt As DataTable = New DataTable()
                gridcmd.Parameters.AddWithValue("@DatumStavby", gridRezanie.Rows(rownum).Cells(0).Value().ToString)
                'cmdImport.Parameters.AddWithValue("@Priemer", gridRezanie.Rows(rownum).Cells(2).Value().ToString)
                'cmdImport.Parameters.AddWithValue("@Hrubka", gridRezanie.Rows(rownum).Cells(3).Value().ToString)
                'cmdImport.Parameters.AddWithValue("@Dlzka", gridRezanie.Rows(rownum).Cells(4).Value().ToString)
                gridcmd.Parameters.AddWithValue("@PN", gridRezanie.Rows(rownum).Cells(1).Value().ToString)
                sda.Fill(dt)


                If dt.Rows.Count = 0 Then

                    Dim cmd As New OleDbCommand("UPDATE tblHlavnyPlan SET RezanieCompleted = @DateTimeStop, RezanieStatus = YES " _
                                            & "WHERE PN=@PN And DatumStavby=@DatumStavby", cn1)
                    cmd.Parameters.AddWithValue("@DateTimeStop", MyDateTimeVariable)
                    cmd.Parameters.AddWithValue("@PN", gridRezanie.Rows(rownum).Cells(1).Value().ToString)
                    cmd.Parameters.AddWithValue("@DatumStavby", gridRezanie.Rows(rownum).Cells(0).Value().ToString)
                    'cmd.Parameters.AddWithValue("@JednotkaWJ", gridRezanie.Rows(rownum).Cells(12).Value().ToString)


                    cmd.ExecuteNonQuery()

                End If

                '    If count = True Then
                'Else
                'End If
            Next
            If cn1.State = ConnectionState.Open Then cn1.Close()
            'cn1.Close()

            populateGridRezanie()
            populateGridOhybanie()



            txtQTYRezanie.Text = ""
            txtDatumStavbyRezanie.Text = ""
            txtPriemerRezanie.Text = ""
            txtHrubkaRezanie.Text = ""
            txtDlzkaRezanie.Text = ""

            gridRezanie.FirstDisplayedScrollingRowIndex = indexRezanie

        Catch ex As Exception
            If MessageBox.Show("Prepáčte, nastala chyba. Skúste opakovať akciu neskôr" & ex.Message, NazovAplikacie, MessageBoxButtons.OK, MessageBoxIcon.Question) = vbOK Then
                ErrorHandling(ex.Message, ex.StackTrace, ex.Source)
                Exit Sub
            End If

            If cn1.State = ConnectionState.Open Then cn1.Close()
        End Try

    End Sub


    Private Sub gridOhybanie_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridOhybanie.CellClick

        indexOhybanie = e.RowIndex

        Dim j As Integer



        With gridOhybanie
            If e.RowIndex >= 0 Then
                j = .CurrentRow.Index

                txtPNOhybanie.Text = .Rows(j).Cells("PN").Value.ToString
                txtQTYOhybanie.Text = .Rows(j).Cells("QTY").Value.ToString
                txtIDOhybanie.Text = .Rows(j).Cells("ID").Value.ToString
                txtJednotkaWJOhybanie.Text = .Rows(j).Cells("JednotkaWJ").Value.ToString
                txtSkupinaOhybanie.Text = .Rows(j).Cells("Skupina").Value.ToString
                txtDatumStavbyOhybanie.Text = .Rows(j).Cells("DatumStavby").Value.ToString
                txtDatumStavbyOhybanie.Text = Format(CDate(txtDatumStavbyOhybanie.Text), "dd.MM")
                txtPriemerOhybanie.Text = .Rows(j).Cells("Priemer").Value.ToString
                txtHrubkaOhybanie.Text = .Rows(j).Cells("Hrubka").Value.ToString
                txtDlzkaOhybanie.Text = .Rows(j).Cells("Dlzka").Value.ToString
                pickPN = gridOhybanie.Rows(j).Cells("PN").Value.ToString

            End If
        End With



    End Sub

    Private Sub gridPriprava_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridPriprava.CellClick
        indexPriprava = e.RowIndex
        Dim j As Integer

        With gridPriprava
            If e.RowIndex >= 0 Then
                j = .CurrentRow.Index
                'DatumStavby,Skupina,PN,QTY,JednotkaWJ
                txtIDPriprava.Text = .Rows(j).Cells("ID").Value.ToString
                txtPNPriprava.Text = .Rows(j).Cells("PN").Value.ToString
                txtQTYPriprava.Text = .Rows(j).Cells("QTY").Value.ToString
                txtJednotkaWJPriprava.Text = .Rows(j).Cells("JednotkaWJ").Value.ToString
                txtSkupinaPriprava.Text = .Rows(j).Cells("Skupina").Value.ToString
                txtDatumStavbyPriprava.Text = .Rows(j).Cells("DatumStavby").Value.ToString
                txtDatumStavbyPriprava.Text = Format(CDate(txtDatumStavbyPriprava.Text), "dd.MM")
                pickPN = .Rows(j).Cells("PN").Value.ToString
            End If
        End With

        SetLinesColorPriprava()


    End Sub
    Private Sub btnProcessPlan_Click(sender As Object, e As EventArgs)

        'Zluci obsah tblTempPlan s tblKompletnyZoznam
        Dim dt As New DataTable()
        cn1.Open()
        Dim query As String = "SELECT tblTempPlan.DatumStavby, tblTempPlan.Skupina, tblTempPlan.JednotkaWJ, tblTempPlan.PipingWJ, tblTempPlan.PN, " _
                              & "tblTempPlan.QTY, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Pocetnost, tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka, " _
                              & "tblKompletnyzoznam.Dlzka, tblKompletnyzoznam.OhybRovna, tblKompletnyzoznam.Kanban, tblKompletnyzoznam.AMOB " _
                              & "FROM tblKompletnyzoznam INNER JOIN tblTempPlan ON tblKompletnyzoznam.PN = tblTempPlan.PN"
        Dim cmd As New OleDbCommand(query, cn1)
        Dim adapter As New OleDbDataAdapter(cmd)
        adapter.Fill(dt)
        gridHlavnyPlan.DataSource = dt
        If cn1.State = ConnectionState.Open Then cn1.Close()

    End Sub

    Private Sub lvPriemeryRezanie_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvPriemeryRezanie.SelectedIndexChanged



        'pri spusteni okna chceme, aby datagridview bol odfiltrovany len pre jedinu otvorenu polozku
        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        cn1.Open()
        'Dim queryRezanie As String = "Select ID,DatumStavby,Skupina,PN,QTY,subPN,Pocetnost, Priemer, Hrubka, Dlzka, OhybRovna,JednotkaWJ " _
        '                            & " FROM tblHlavnyPlan where RezanieStatus=NO And Kanban=NO And AMOB=YES ORDER BY DatumStavby DESC , Skupina, PN, subPN;"

        Dim queryRezanie As String = "SELECT [qrySubGridRezanie 02].DatumStavby, [qrySubGridRezanie 02].PN, [qrySubGridRezanie 02].Priemer, [qrySubGridRezanie 02].Hrubka, [qrySubGridRezanie 02].Dlzka, [qrySubGridRezanie 02].QTY, [qrySubGridRezanie 02].RezanieStatus, [qrySubGridRezanie 02].OhybRovna, [qrySubGridRezanie 02].Kanban, tblRezanieStatus.RezanieStatus, IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date')) AS HotovoCas, [qrySubGridRezanie 02].ID, [qrySubGridRezanie 02].JednotkaWJ
FROM [qrySubGridRezanie 02] LEFT JOIN tblRezanieStatus ON ([qrySubGridRezanie 02].ID = tblRezanieStatus.IDKomplZ) AND ([qrySubGridRezanie 02].JednotkaWJ = tblRezanieStatus.JednotkaWJ) AND ([qrySubGridRezanie 02].PN = tblRezanieStatus.PN) AND ([qrySubGridRezanie 02].DatumStavby = tblRezanieStatus.DatumStavby)
WHERE ((([qrySubGridRezanie 02].Priemer)=[@Priemer]) AND ((IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date')))>DateValue((Now()-1)) Or (IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date'))) Is Null Or ((IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date')))>DateValue((Now()-1)) Or (IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date'))) Is Null) Or ((IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date')))>DateValue((Now()-1)) Or (IIf([tblRezanieStatus.RezanieStatus]=True,Format([ulozene],'Short Date'))) Is Null)))
ORDER BY [qrySubGridRezanie 02].DatumStavby DESC , [qrySubGridRezanie 02].Priemer, [qrySubGridRezanie 02].Hrubka, [qrySubGridRezanie 02].Dlzka;"

        Dim gridcmdRezanie As OleDbCommand = New OleDbCommand(queryRezanie, cn1)
        gridcmdRezanie.Parameters.AddWithValue("@Priemer", lvPriemeryRezanie.Items(lvPriemeryRezanie.FocusedItem.Index).SubItems(0).Text)
        'gridcmdRezanie.Parameters.AddWithValue("@OhybRovna", "Ohyb")
        Dim sdaRezanie As OleDbDataAdapter = New OleDbDataAdapter(gridcmdRezanie)
        Dim dtRezanie As DataTable = New DataTable()
        sdaRezanie.Fill(dtRezanie)
        gridRezanie.DataSource = dtRezanie
        If cn1.State = ConnectionState.Open Then cn1.Close()

        gridRezanie.Columns(0).DefaultCellStyle.Format = "dd.MM"

        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridRezanie.RowTemplate.Height = 35

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridRezanie
            .Columns(0).Width = 110
            .Columns(1).Width = 110
            .Columns(2).Width = 110
            .Columns(3).Width = 110
            .Columns(4).Width = 110
            .Columns(6).Visible = False

            .Columns(0).DefaultCellStyle.Format = "dd.MM"

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(0).HeaderCell.Value = "Dátum Stavby"
            .Columns(1).HeaderCell.Value = "Piping PN"
            .Columns(2).HeaderCell.Value = "Priemer [mm]"
            .Columns(3).HeaderCell.Value = "Hrúbka [mm]"
            .Columns(4).HeaderCell.Value = "Dĺžka [mm]"
            .Columns(5).HeaderCell.Value = "Množstvo"
            .Columns(6).HeaderCell.Value = "Hotovo" 'PN hotovo
            .Columns(7).HeaderCell.Value = "Ohyb/Rovná"
            .Columns(8).HeaderCell.Value = "Kanban" 'Kanban
            '.Columns(9).HeaderCell.Value = "AMOB"
            .Columns(9).HeaderCell.Value = "Hotovo" 'tblRezanieStatus

            .ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)

        End With

        gridRezanie.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        SetLinesColorRezanie()
        'nastavenie sortovania
        'gridRezanie.Sort(gridRezanie.Columns(1), System.ComponentModel.ListSortDirection.Descending)
        'gridRezanie.Sort(gridRezanie.Columns(2), System.ComponentModel.ListSortDirection.Ascending)
        'gridRezanie.Sort(gridRezanie.Columns(3), System.ComponentModel.ListSortDirection.Ascending)
        'gridRezanie.Sort(gridRezanie.Columns(5), System.ComponentModel.ListSortDirection.Ascending)
        txtRezanieCount.Text = gridRezanie.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()
        txtFilterPNRezanie.Text = ""
    End Sub



    Private Sub BtnLinkImport_Click(sender As Object, e As EventArgs)

        cn1.Open()

        'Zluci obsah tblTempPlan s tblKompletnyZoznam
        Dim dt As New DataTable()

        Dim query As String = "SELECT tblDennyPlan.*, tblKompletnyZoznam.subPN, tblKompletnyZoznam.Pocetnost, tblKompletnyZoznam.Priemer, tblKompletnyZoznam.Hrubka, " _
                            & "tblKompletnyZoznam.Dlzka, tblKompletnyZoznam.OhybRovna, tblKompletnyZoznam.Kanban, tblKompletnyZoznam.AMOB " _
                            & "FROM tblDennyPlan LEFT JOIN tblKompletnyZoznam ON tblDennyPlan.PN = tblKompletnyZoznam.PN;"


        Dim cmd As New OleDbCommand(query, cn1)
        Dim adapter As New OleDbDataAdapter(cmd)
        adapter.Fill(dt)
        gridHlavnyPlan.DataSource = dt
        txtPlanCount.Text = gridHlavnyPlan.RowCount


        If (MessageBox.Show("Importovane data zlucene s Kompletnym Zoznamom", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbNo Then
            cn1.Close()
            Exit Sub
        Else
        End If

        'nahra obsah z gridu (po riadku) do tabulky
        For i As Integer = 0 To gridHlavnyPlan.Rows.Count - 2 Step +1
            ' If IsDBNull(gridHlavnyPlan.Rows(i).Cells(0).Value()) Then
            If IsDBNull(gridHlavnyPlan.Rows(i).Cells(0).ToString()) Then
                Continue For
            Else

                Dim cmdImport As New OleDbCommand("INSERT INTO tblHlavnyPlan (DatumStavby,Skupina,JednotkaWJ,PipingWJ,PN,QTY,subPN,Priemer,Hrubka,Dlzka,OhybRovna,Kanban,AMOB) " _
                                                  & "VALUES (@DatumStavby,@Skupina,@JednotkaWJ,@PipingWJ,@PN,@QTY,@subPN,@Priemer,@Hrubka,@Dlzka,@OhybRovna,@Kanban,@AMOB)", cn1)
                cmdImport.Parameters.AddWithValue("@DatumStavby", gridHlavnyPlan.Rows(i).Cells(0).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Skupina", gridHlavnyPlan.Rows(i).Cells(1).Value().ToString)
                cmdImport.Parameters.AddWithValue("@JednotkaWJ", gridHlavnyPlan.Rows(i).Cells(2).Value().ToString)
                cmdImport.Parameters.AddWithValue("@PipingWJ", gridHlavnyPlan.Rows(i).Cells(3).Value().ToString)
                cmdImport.Parameters.AddWithValue("@PN", gridHlavnyPlan.Rows(i).Cells(4).Value().ToString)
                cmdImport.Parameters.AddWithValue("@QTY", CInt(gridHlavnyPlan.Rows(i).Cells(5).Value().ToString))
                cmdImport.Parameters.AddWithValue("@subPN", gridHlavnyPlan.Rows(i).Cells(6).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Pocetnost", gridHlavnyPlan.Rows(i).Cells(7).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Priemer", gridHlavnyPlan.Rows(i).Cells(8).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Hrubka", gridHlavnyPlan.Rows(i).Cells(9).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Dlzka", gridHlavnyPlan.Rows(i).Cells(10).Value().ToString)
                cmdImport.Parameters.AddWithValue("@OhybRovna", gridHlavnyPlan.Rows(i).Cells(11).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Kanban", gridHlavnyPlan.Rows(i).Cells(12).Value().ToString)
                cmdImport.Parameters.AddWithValue("@AMOB", gridHlavnyPlan.Rows(i).Cells(13).Value().ToString)

                cmdImport.ExecuteNonQuery()

            End If



        Next

        'pri spusteni okna chceme, aby datagridview bol odfiltrovany len pre jedinu otvorenu polozku
        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        'cn1.Open()
        Dim queryStatus As String = "Select DatumStavby,Skupina,PipingWJ,PN,QTY,subPN,Pocetnost, Priemer, Hrubka, Dlzka, " _
                                    & " OhybRovna, Kanban, AMOB,RezanieStatus,RezanieCompleted,OhybanieStatus,OhybanieCompleted, " _
                                    & " PripravaStatus,PripravaCompleted,SpajkovanieStatus,SpajkovanieCompleted FROM tblHlavnyPlan " _
                                    & " ORDER BY DatumStavby DESC , Skupina, PN, subPN;"

        Dim gridCmdStatus As OleDbCommand = New OleDbCommand(queryStatus, cn1)
        'gridcmd.Parameters.AddWithValue("@today", MyDateIsToday)
        'gridcmd.Parameters.AddWithValue("@MCName", Environment.MachineName)
        Dim sdaStatus As OleDbDataAdapter = New OleDbDataAdapter(gridCmdStatus)
        Dim dtStatus As DataTable = New DataTable()
        sdaStatus.Fill(dtStatus)
        gridHlavnyPlan.DataSource = dtStatus
        If cn1.State = ConnectionState.Open Then cn1.Close()

        gridHlavnyPlan.Columns(0).DefaultCellStyle.Format = "dd.MM"
    End Sub



    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        lvPriemeryRezanie.SelectedItems.Clear() 'odstranenie filtra
        txtFilterPNRezanie.Text = ""
        populateGridRezanie()


    End Sub

    Private Sub ListPriemeryOhybanie_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvPriemeryOhybanie.SelectedIndexChanged

        'vyplnenie gridu pre Ohybanie
        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        cn1.Open()
        'Dim query As String = "Select ID,DatumStavby,Skupina,PN,QTY,subPN,Priemer, Hrubka,Dlzka,JednotkaWJ,OhybanieStatus FROM tblHlavnyPlan " _
        '                    & " where (Kanban=YES or (Kanban=NO and RezanieStatus=YES)) ORDER BY DatumStavby DESC , PN, subPN, Priemer;"

        Dim query As String = "SELECT qrySubGridOhybanie.IDHlavnyPlan AS ID, qrySubGridOhybanie.DatumStavby, qrySubGridOhybanie.Skupina, qrySubGridOhybanie.PN, qrySubGridOhybanie.JednotkaWJ, qrySubGridOhybanie.QTY, qrySubGridOhybanie.subPN, qrySubGridOhybanie.Priemer, qrySubGridOhybanie.Hrubka, qrySubGridOhybanie.Dlzka, qrySubGridOhybanie.OhybRovna, tblOhybanieStatus.OhybanieStatus, tblOhybanieStatus.Ulozene, IIf(IsNull([tblKanban.Kanban])=True Or [tblKanban.Kanban]=No,IIf([tblRezanieStatus.RezanieStatus]=Yes,'ANO','NIE'),'ANO-Kanban') AS Narezane, qrySubGridOhybanie.IDKomplZ, IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null) AS HotovoCas " _
            & "FROM (qrySubGridOhybanie LEFT JOIN tblOhybanieStatus ON (qrySubGridOhybanie.IDKomplZ = tblOhybanieStatus.IDKomplZ) And (qrySubGridOhybanie.IDHlavnyPlan = tblOhybanieStatus.IDHlavnyPlan)) LEFT JOIN tblRezanieStatus ON (qrySubGridOhybanie.JednotkaWJ = tblRezanieStatus.JednotkaWJ) And (qrySubGridOhybanie.PN = tblRezanieStatus.PN) And (qrySubGridOhybanie.IDKomplZ = tblRezanieStatus.IDKomplZ) " _
        & "WHERE (((qrySubGridOhybanie.Priemer)>=12 And (qrySubGridOhybanie.Priemer)=[@Priemer]) And ((qrySubGridOhybanie.OhybRovna)='Ohyb') AND ((IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null))>DateValue((Now()-1)) Or (IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null)) Is Null)) OR (((qrySubGridOhybanie.Priemer)>=42 And (qrySubGridOhybanie.Priemer)=[@Priemer]) AND ((qrySubGridOhybanie.Dlzka)>=300) AND ((qrySubGridOhybanie.OhybRovna)='Rovna') AND ((IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null))>DateValue((Now()-1)) Or (IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null)) Is Null)) " _
        & "ORDER BY qrySubGridOhybanie.DatumStavby DESC , qrySubGridOhybanie.Priemer, qrySubGridOhybanie.PN, qrySubGridOhybanie.subPN;"

        'WHERE (((qrySubGridOhybanie.Priemer)>=12)) AND qrySubGridOhybanie.Priemer=@Priemer 

        Dim gridcmd As OleDbCommand = New OleDbCommand(query, cn1)
        Dim sda As OleDbDataAdapter = New OleDbDataAdapter(gridcmd)
        gridcmd.Parameters.AddWithValue("@Priemer", lvPriemeryOhybanie.Items(lvPriemeryOhybanie.FocusedItem.Index).SubItems(0).Text)
        Dim dt As DataTable = New DataTable()
        cn1.Close()
        sda.Fill(dt)
        gridOhybanie.DataSource = dt
        If cn1.State = ConnectionState.Open Then cn1.Close()

        'Dim lnk As New DataGridViewLinkColumn()
        'gridOhybanie.Columns.Add(lnk)

        With gridOhybanie
            .RowTemplate.Height = 35
            .ForeColor = Color.Black
            .Columns(0).Visible = False

            .Columns(1).Width = 80
            .Columns(2).Width = 120
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 75
            .Columns(6).Width = 80
            .Columns(7).Width = 80
            .Columns(8).Width = 80
            .Columns(9).Width = 80
            .Columns(10).Width = 80
            .Columns(13).Width = 130


            .Columns(12).Visible = False
            .Columns(14).Visible = False
            '.Columns(13).Visible = False

            .Columns(1).DefaultCellStyle.Format = "dd.MM"
            '.Columns(11).DefaultCellStyle.Format = "dd.MM"

            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(1).HeaderCell.Value = "Dátum Stavby"
            .Columns(2).HeaderCell.Value = "Skupina"
            .Columns(3).HeaderCell.Value = "Piping PN"
            .Columns(4).HeaderCell.Value = "WJ Jednotky"
            .Columns(5).HeaderCell.Value = "Množ."
            .Columns(6).HeaderCell.Value = "sub Piping"
            .Columns(7).HeaderCell.Value = "Priemer [mm]"
            .Columns(8).HeaderCell.Value = "Hrúbka [mm]"
            .Columns(9).HeaderCell.Value = "Dĺžka [mm]"
            .Columns(10).HeaderCell.Value = "Ohýbaná / Rovná"
            .Columns(11).HeaderCell.Value = "Hotovo"

            .Columns(12).HeaderCell.Value = "ulozene v tblOhybanieStatus"
            .Columns(13).HeaderCell.Value = "Narezane"

            .ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        End With

        txtOhybanieCount.Text = gridOhybanie.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()
        txtFilterPNOhybanie.Text = ""
        SetLinesColorOhybanie()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        lvPriemeryOhybanie.SelectedItems.Clear()
        txtFilterPNOhybanie.Text = ""
        populateGridOhybanie()
        SetLinesColorOhybanie()
    End Sub

    Private Sub BtnFindPNRezanie_Click_1(sender As Object, e As EventArgs) Handles btnFindPNRezanie.Click

        lvPriemeryRezanie.SelectedItems.Clear()


        'pri spusteni okna chceme, aby datagridview bol odfiltrovany len pre jedinu otvorenu polozku
        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        cn1.Open()
        Dim queryRezanie As String = "SELECT tblHlavnyPlan.ID, tblHlavnyPlan.DatumStavby, tblHlavnyPlan.Skupina, tblHlavnyPlan.PN, tblHlavnyPlan.QTY, " _
                    & "tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka, tblKompletnyzoznam.OhybRovna " _
                    & "FROM tblKompletnyzoznam RIGHT JOIN tblHlavnyPlan ON tblKompletnyzoznam.PN = tblHlavnyPlan.PN " _
                    & "WHERE (((tblHlavnyPlan.[RezanieStatus])=No) And ((tblKompletnyzoznam.[Kanban])=No) And ((tblHlavnyPlan.[PN])=[@PN])) " _
                    & "ORDER BY tblHlavnyPlan.DatumStavby DESC , tblHlavnyPlan.Skupina, tblHlavnyPlan.PN, tblKompletnyzoznam.subPN;"

        Dim gridcmdRezanie As OleDbCommand = New OleDbCommand(queryRezanie, cn1)
        gridcmdRezanie.Parameters.AddWithValue("@PN", txtFilterPNRezanie.Text)
        Dim sdaRezanie As OleDbDataAdapter = New OleDbDataAdapter(gridcmdRezanie)
        Dim dtRezanie As DataTable = New DataTable()
        sdaRezanie.Fill(dtRezanie)
        gridRezanie.DataSource = dtRezanie
        If cn1.State = ConnectionState.Open Then cn1.Close()

        gridRezanie.Columns(0).DefaultCellStyle.Format = "dd.MM"

        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridRezanie.RowTemplate.Height = 35

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridRezanie
            .Columns(0).Visible = False
            .Columns(1).Width = 90
            .Columns(2).Width = 90
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 110
            .Columns(6).Width = 110
            .Columns(7).Width = 80
            .Columns(8).Width = 80
            .Columns(9).Width = 90
            .Columns(10).Width = 90
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Format = "dd.MM"
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With

        With gridRezanie
            .RowHeadersVisible = False

            .Columns(1).HeaderCell.Value = "Dátum Stavby"
            .Columns(2).HeaderCell.Value = "Skupina"
            .Columns(3).HeaderCell.Value = "Piping"
            .Columns(4).HeaderCell.Value = "Množstvo"
            .Columns(5).HeaderCell.Value = "subPiping"
            .Columns(6).HeaderCell.Value = "Početnosť"
            .Columns(7).HeaderCell.Value = "Priemer [mm]"
            .Columns(8).HeaderCell.Value = "Hrúbka [mm]"
            .Columns(9).HeaderCell.Value = "Dĺžka [mm]"
            .Columns(10).HeaderCell.Value = "Ohyb / Rovná"
        End With

        'nastavenie sortovania
        'gridRezanie.Sort(gridRezanie.Columns(1), System.ComponentModel.ListSortDirection.Descending)
        'gridRezanie.Sort(gridRezanie.Columns(2), System.ComponentModel.ListSortDirection.Descending)

    End Sub

    Private Sub BtnFindPNOhybanie_Click(sender As Object, e As EventArgs) Handles btnFindPNOhybanie.Click

        lvPriemeryOhybanie.SelectedItems.Clear()


        'pri spusteni okna chceme, aby datagridview bol odfiltrovany len pre jedinu otvorenu polozku
        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        cn1.Open()

        'Dim queryOhybanie As String = "Select qrySubGridOhybanie.IDHlavnyPlan As ID, qrySubGridOhybanie.DatumStavby, qrySubGridOhybanie.Skupina, qrySubGridOhybanie.PN, " _
        '    & "qrySubGridOhybanie.JednotkaWJ, qrySubGridOhybanie.QTY, qrySubGridOhybanie.subPN, qrySubGridOhybanie.Priemer, qrySubGridOhybanie.Hrubka, qrySubGridOhybanie.Dlzka, " _
        '    & "qrySubGridOhybanie.OhybRovna, tblOhybanieStatus.OhybanieStatus, tblOhybanieStatus.Ulozene, qrySubGridOhybanie.Narezane, qrySubGridOhybanie.IDKomplZ " _
        '    & "FROM tblOhybanieStatus RIGHT JOIN subQRYGridOhybanie On (tblOhybanieStatus.IDKomplZ = qrySubGridOhybanie.IDKomplZ) " _
        '    & "And (tblOhybanieStatus.IDHlavnyPlan = qrySubGridOhybanie.IDHlavnyPlan) " _
        '    & "WHERE (((qrySubGridOhybanie.Priemer)>=12)) And qrySubGridOhybanie.PN=@PN " _
        '    & "ORDER BY qrySubGridOhybanie.DatumStavby DESC,qrySubGridOhybanie.Priemer,qrySubGridOhybanie.PN,qrySubGridOhybanie.subPN"

        Dim queryOhybanie As String = "SELECT qrySubGridOhybanie.IDHlavnyPlan AS ID, qrySubGridOhybanie.DatumStavby, qrySubGridOhybanie.Skupina, qrySubGridOhybanie.PN, qrySubGridOhybanie.JednotkaWJ, qrySubGridOhybanie.QTY, qrySubGridOhybanie.subPN, qrySubGridOhybanie.Priemer, qrySubGridOhybanie.Hrubka, qrySubGridOhybanie.Dlzka, qrySubGridOhybanie.OhybRovna, tblOhybanieStatus.OhybanieStatus, tblOhybanieStatus.Ulozene, IIf(IsNull([tblKanban.Kanban])=True Or [tblKanban.Kanban]=No,IIf([tblRezanieStatus.RezanieStatus]=Yes,'ANO','NIE'),'ANO-Kanban') AS Narezane, qrySubGridOhybanie.IDKomplZ, IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null) AS HotovoCas " _
            & "FROM (qrySubGridOhybanie LEFT JOIN tblOhybanieStatus ON (qrySubGridOhybanie.IDKomplZ = tblOhybanieStatus.IDKomplZ) And (qrySubGridOhybanie.IDHlavnyPlan = tblOhybanieStatus.IDHlavnyPlan)) LEFT JOIN tblRezanieStatus ON (qrySubGridOhybanie.JednotkaWJ = tblRezanieStatus.JednotkaWJ) And (qrySubGridOhybanie.PN = tblRezanieStatus.PN) And (qrySubGridOhybanie.IDKomplZ = tblRezanieStatus.IDKomplZ) " _
        & "WHERE (((qrySubGridOhybanie.PN)=[@PN]) And ((qrySubGridOhybanie.Priemer)>=12) And ((qrySubGridOhybanie.OhybRovna)='Ohyb') AND ((IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null))>DateValue((Now()-1)) Or (IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null)) Is Null)) OR (((qrySubGridOhybanie.PN)=[@PN]) AND ((qrySubGridOhybanie.Priemer)>=42) AND ((qrySubGridOhybanie.Dlzka)>=300) AND ((qrySubGridOhybanie.OhybRovna)='Rovna') AND ((IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null))>DateValue((Now()-1)) Or (IIf([tblOhybanieStatus.OhybanieStatus]=True,DateValue([tblOhybanieStatus.ulozene]),Null)) Is Null)) " _
        & "ORDER BY qrySubGridOhybanie.DatumStavby DESC , qrySubGridOhybanie.Priemer, qrySubGridOhybanie.PN, qrySubGridOhybanie.subPN;"

        Dim gridcmdOhybanie As OleDbCommand = New OleDbCommand(queryOhybanie, cn1)
        'ID,DatumStavby,Skupina,PN,QTY,subPN,Priemer, Hrubka, Dlzka
        gridcmdOhybanie.Parameters.AddWithValue("@PN", txtFilterPNOhybanie.Text)
        'filteredPN = txtFilterPNOhybanie.Text
        Dim sdaOhybanie As OleDbDataAdapter = New OleDbDataAdapter(gridcmdOhybanie)
        Dim dtOhybanie As DataTable = New DataTable()
        sdaOhybanie.Fill(dtOhybanie)
        gridOhybanie.DataSource = dtOhybanie
        If cn1.State = ConnectionState.Open Then cn1.Close()

        gridOhybanie.Columns(0).DefaultCellStyle.Format = "dd.MM"
        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridOhybanie.RowTemplate.Height = 35

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridOhybanie
            .RowTemplate.Height = 35
            .ForeColor = Color.Black
            .Columns(0).Visible = False

            .Columns(1).Width = 80
            .Columns(2).Width = 120
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 75
            .Columns(6).Width = 80
            .Columns(7).Width = 80
            .Columns(8).Width = 80
            .Columns(9).Width = 80
            .Columns(10).Width = 80
            .Columns(13).Width = 130


            .Columns(12).Visible = False
            .Columns(14).Visible = False
            '.Columns(13).Visible = False

            .Columns(1).DefaultCellStyle.Format = "dd.MM"
            '.Columns(11).DefaultCellStyle.Format = "dd.MM"

            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(1).HeaderCell.Value = "Dátum Stavby"
            .Columns(2).HeaderCell.Value = "Skupina"
            .Columns(3).HeaderCell.Value = "Piping PN"
            .Columns(4).HeaderCell.Value = "WJ Jednotky"
            .Columns(5).HeaderCell.Value = "Množ."
            .Columns(6).HeaderCell.Value = "sub Piping"
            .Columns(7).HeaderCell.Value = "Priemer [mm]"
            .Columns(8).HeaderCell.Value = "Hrúbka [mm]"
            .Columns(9).HeaderCell.Value = "Dĺžka [mm]"
            .Columns(10).HeaderCell.Value = "Ohýbaná / Rovná"
            .Columns(11).HeaderCell.Value = "Hotovo"

            .Columns(12).HeaderCell.Value = "ulozene v tblOhybanieStatus"
            .Columns(13).HeaderCell.Value = "Narezane"

            .ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        End With

        'nastavenie sortovania
        'gridOhybanie.Sort(gridOhybanie.Columns(1), System.ComponentModel.ListSortDirection.Descending)
        'gridOhybanie.Sort(gridOhybanie.Columns(2), System.ComponentModel.ListSortDirection.Descending)

        SetLinesColorOhybanie()
    End Sub



    Private Sub btnHlavnyPlan_Click(sender As Object, e As EventArgs) Handles btnHlavnyPlan.Click

        populateFullPlan()


    End Sub

    Private Sub populateFullPlan()

        If cn1.State = ConnectionState.Closed Then cn1.Open()
        'Dim queryStatus As String = "SELECT tblHlavnyPlan.ID, tblHlavnyPlan.DatumStavby, tblHlavnyPlan.Skupina, tblHlavnyPlan.JednotkaWJ, tblHlavnyPlan.PipingWJ, " _
        '            & "tblHlavnyPlan.PN, tblHlavnyPlan.QTY, tblHlavnyPlan.Drazkovanie, tblHlavnyPlan.RoutingCas, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer, " _
        '            & "tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka, tblKompletnyzoznam.OhybRovna, tblKompletnyzoznam.Kanban, tblKompletnyzoznam.AMOB, " _
        '            & "tblHlavnyPlan.RezanieStatus, tblHlavnyPlan.RezanieCompleted, tblHlavnyPlan.OhybanieStatus, tblHlavnyPlan.OhybanieCompleted, tblHlavnyPlan.PripravaStatus, " _
        '            & "tblHlavnyPlan.PripravaCompleted, tblHlavnyPlan.SpajkovanieStatus, tblHlavnyPlan.SpajkovanieCompleted, tblHlavnyPlan.VychystanieStatus, " _
        '            & "tblHlavnyPlan.VychystanieCompleted, tblHlavnyPlan.VyplanovanyWJ, tblHlavnyPlan.VyplanovaneMnozstvo, tblHlavnyPlan.VyplanovaneDovod, tblHlavnyPlan.VyplanovaneDatum, " _
        '            & "tblHlavnyPlan.Vykres " _
        '            & "From tblHlavnyPlan LEFT Join tblKompletnyzoznam On tblHlavnyPlan.PN = tblKompletnyzoznam.PN " _
        '            & "ORDER BY tblHlavnyPlan.DatumStavby DESC , tblHlavnyPlan.Skupina, tblHlavnyPlan.PN, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer;"
        Dim queryStatus As String = "SELECT tblHlavnyPlan.ID, tblHlavnyPlan.DatumStavby, tblHlavnyPlan.Skupina, tblHlavnyPlan.JednotkaWJ, tblHlavnyPlan.PipingWJ, " _
                & "tblHlavnyPlan.PN, tblHlavnyPlan.QTY, tblHlavnyPlan.Drazkovanie, tblHlavnyPlan.RoutingCas, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer, " _
                & "tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka, tblKompletnyzoznam.OhybRovna, tblKompletnyzoznam.AMOB, " _
                & "tblHlavnyPlan.RezanieStatus, tblHlavnyPlan.RezanieCompleted, tblHlavnyPlan.OhybanieStatus, tblHlavnyPlan.OhybanieCompleted, tblHlavnyPlan.PripravaStatus, " _
                & "tblHlavnyPlan.PripravaCompleted, tblHlavnyPlan.SpajkovanieStatus, tblHlavnyPlan.SpajkovanieCompleted, tblHlavnyPlan.VychystanieStatus, " _
                & "tblHlavnyPlan.VychystanieCompleted, tblHlavnyPlan.VyplanovanyWJ, tblHlavnyPlan.VyplanovaneMnozstvo, tblHlavnyPlan.VyplanovaneDovod, " _
                & "tblHlavnyPlan.VyplanovaneDatum, tblHlavnyPlan.Vykres, tblKanban.Kanban,tblHlavnyPlan.Urgent " _
                & "FROM (tblHlavnyPlan LEFT JOIN tblKompletnyzoznam ON tblHlavnyPlan.PN = tblKompletnyzoznam.PN) " _
                & "LEFT JOIN tblKanban ON (tblKompletnyzoznam.Dlzka = tblKanban.Dlzka) AND (tblKompletnyzoznam.Hrubka = tblKanban.Hrubka) " _
                & "AND (tblKompletnyzoznam.Priemer = tblKanban.Priemer) " _
                & "ORDER BY tblHlavnyPlan.DatumStavby DESC , tblHlavnyPlan.Skupina, tblHlavnyPlan.PN, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer;"

        Dim gridCmdStatus As OleDbCommand = New OleDbCommand(queryStatus, cn1)
        Dim sdaStatus As OleDbDataAdapter = New OleDbDataAdapter(gridCmdStatus)
        Dim dtStatus As DataTable = New DataTable()
        sdaStatus.Fill(dtStatus)

        gridHlavnyPlan.DataSource = dtStatus
        If cn1.State = ConnectionState.Open Then cn1.Close()
        gridHlavnyPlan.Columns(0).DefaultCellStyle.Format = "dd.MM"
        txtPlanCount.Text = gridHlavnyPlan.RowCount

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridHlavnyPlan
            .Columns(0).Visible = False
            .Columns(1).Width = 80
            .Columns(2).Width = 80
            .Columns(3).Width = 80
            .Columns(4).Width = 80
            .Columns(5).Width = 80
            .Columns(6).Width = 50
            .Columns(7).Width = 100
            .Columns(8).Width = 80
            .Columns(9).Width = 60
            .Columns(10).Width = 70
            .Columns(11).Width = 70
            .Columns(12).Width = 70
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Format = "dd.MM"
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With

        With gridHlavnyPlan
            .RowHeadersVisible = False

            .Columns(1).HeaderCell.Value = "Dátum Stavby"
            .Columns(2).HeaderCell.Value = "Skupina"
            .Columns(3).HeaderCell.Value = "WJ Jednotky"
            .Columns(4).HeaderCell.Value = "WJ Pipingu"
            .Columns(5).HeaderCell.Value = "Piping"
            .Columns(6).HeaderCell.Value = "Mn."
            .Columns(7).HeaderCell.Value = "Drážkovanie"
            .Columns(8).HeaderCell.Value = "Routing"
            .Columns(9).HeaderCell.Value = "subPN"
            .Columns(10).HeaderCell.Value = "Priemer [mm]"
            .Columns(11).HeaderCell.Value = "Hrúbka [mm]"
            .Columns(12).HeaderCell.Value = "Dĺžka [mm]"
            .Columns(13).HeaderCell.Value = "Ohyb/Rovná"
            .Columns(14).HeaderCell.Value = "Kanban"
            .Columns(15).HeaderCell.Value = "AMOB"
            .Columns(16).HeaderCell.Value = "Rezanie Status"
            .Columns(17).HeaderCell.Value = "Rezanie ukončené"
            .Columns(18).HeaderCell.Value = "Ohýbanie Status"
            .Columns(19).HeaderCell.Value = "Ohýbanie ukončené"
            .Columns(20).HeaderCell.Value = "Príprava Status"
            .Columns(21).HeaderCell.Value = "Príprava ukončená"
            .Columns(22).HeaderCell.Value = "Spájkovanie Status"
            .Columns(23).HeaderCell.Value = "Spájkovanie ukončené"
        End With
        gridHlavnyPlan.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        'nastavenie sortovania
        'gridHlavnyPlan.Sort(gridHlavnyPlan.Columns(1), System.ComponentModel.ListSortDirection.Descending)
        'gridHlavnyPlan.Sort(gridHlavnyPlan.Columns(2), System.ComponentModel.ListSortDirection.Descending)

    End Sub

    Private Sub populateFullPlanGridReport()

        If cn1.State = ConnectionState.Closed Then cn1.Open()

        Dim queryStatus As String = "SELECT tblHlavnyPlan.ID, tblHlavnyPlan.DatumStavby, tblHlavnyPlan.Skupina, tblHlavnyPlan.JednotkaWJ, tblHlavnyPlan.PipingWJ, " _
                & "tblHlavnyPlan.PN, tblHlavnyPlan.QTY, tblHlavnyPlan.Drazkovanie, tblHlavnyPlan.RoutingCas, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer, " _
                & "tblKompletnyzoznam.Hrubka, tblKompletnyzoznam.Dlzka, tblKompletnyzoznam.OhybRovna, tblKompletnyzoznam.AMOB, " _
                & "tblHlavnyPlan.RezanieStatus, tblHlavnyPlan.RezanieCompleted, tblHlavnyPlan.OhybanieStatus, tblHlavnyPlan.OhybanieCompleted, tblHlavnyPlan.PripravaStatus, " _
                & "tblHlavnyPlan.PripravaCompleted, tblHlavnyPlan.SpajkovanieStatus, tblHlavnyPlan.SpajkovanieCompleted, tblHlavnyPlan.VychystanieStatus, " _
                & "tblHlavnyPlan.VychystanieCompleted, tblHlavnyPlan.VyplanovanyWJ, tblHlavnyPlan.VyplanovaneMnozstvo, tblHlavnyPlan.VyplanovaneDovod, " _
                & "tblHlavnyPlan.VyplanovaneDatum, tblHlavnyPlan.Vykres, tblKanban.Kanban,tblHlavnyPlan.Urgent " _
                & "FROM (tblHlavnyPlan LEFT JOIN tblKompletnyzoznam ON tblHlavnyPlan.PN = tblKompletnyzoznam.PN) " _
                & "LEFT JOIN tblKanban ON (tblKompletnyzoznam.Dlzka = tblKanban.Dlzka) AND (tblKompletnyzoznam.Hrubka = tblKanban.Hrubka) " _
                & "AND (tblKompletnyzoznam.Priemer = tblKanban.Priemer) " _
                & "WHERE tblHlavnyPlan.DatumStavby=@DatumStavby " _
                & "ORDER BY tblHlavnyPlan.DatumStavby DESC , tblHlavnyPlan.Skupina, tblHlavnyPlan.PN, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer;"

        Dim gridCmdStatus As OleDbCommand = New OleDbCommand(queryStatus, cn1)
        Dim sdaStatus As OleDbDataAdapter = New OleDbDataAdapter(gridCmdStatus)
        gridCmdStatus.Parameters.AddWithValue("@DatumStavby", Format(dtpDatumStavbyPrehlady.Value.Date, "dd/MM/yyyy"))
        Dim dtStatus As DataTable = New DataTable()
        sdaStatus.Fill(dtStatus)

        gridReport.DataSource = dtStatus
        If cn1.State = ConnectionState.Open Then cn1.Close()
        gridReport.Columns(0).DefaultCellStyle.Format = "dd.MM"
        txtPlanCount.Text = gridReport.RowCount

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridReport
            .Columns(0).Visible = False
            .Columns(1).Width = 80
            .Columns(2).Width = 120
            .Columns(3).Width = 80
            .Columns(4).Width = 80
            .Columns(5).Width = 80
            .Columns(6).Width = 50
            .Columns(7).Width = 100
            .Columns(8).Width = 80
            .Columns(9).Width = 60
            .Columns(10).Width = 70
            .Columns(11).Width = 70
            .Columns(12).Width = 70
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Format = "dd.MM"
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(1).HeaderCell.Value = "Dátum Stavby"
            .Columns(2).HeaderCell.Value = "Skupina"
            .Columns(3).HeaderCell.Value = "WJ Jednotky"
            .Columns(4).HeaderCell.Value = "WJ Pipingu"
            .Columns(5).HeaderCell.Value = "Piping"
            .Columns(6).HeaderCell.Value = "Mn."
            .Columns(7).HeaderCell.Value = "Drážkovanie"
            .Columns(8).HeaderCell.Value = "Routing"
            .Columns(9).HeaderCell.Value = "subPN"
            .Columns(10).HeaderCell.Value = "Priemer [mm]"
            .Columns(11).HeaderCell.Value = "Hrúbka [mm]"
            .Columns(12).HeaderCell.Value = "Dĺžka [mm]"
            .Columns(13).HeaderCell.Value = "Ohyb/Rovná"
            .Columns(14).HeaderCell.Value = "Kanban"
            .Columns(15).HeaderCell.Value = "AMOB"
            .Columns(16).HeaderCell.Value = "Rezanie Status"
            .Columns(17).HeaderCell.Value = "Rezanie ukončené"
            .Columns(18).HeaderCell.Value = "Ohýbanie Status"
            .Columns(19).HeaderCell.Value = "Ohýbanie ukončené"
            .Columns(20).HeaderCell.Value = "Príprava Status"
            .Columns(21).HeaderCell.Value = "Príprava ukončená"
            .Columns(22).HeaderCell.Value = "Spájkovanie Status"
            .Columns(23).HeaderCell.Value = "Spájkovanie ukončené"
        End With
        gridReport.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)


    End Sub

    Private Sub BtnFindPNInziniering_Click(sender As Object, e As EventArgs) Handles btnFindPNInziniering.Click

        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        cn1.Open()
        Dim queryInziniering As String = "SELECT DISTINCT tblHlavnyPlan.PN, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka, " _
                & "tblKompletnyzoznam.Dlzka, tblKompletnyzoznam.OhybRovna, tblKanban.Kanban, tblKompletnyZoznam.Skontrolovane " _
                & "FROM ((tblHlavnyPlan LEFT JOIN tblKompletnyzoznam ON tblHlavnyPlan.PN = tblKompletnyzoznam.PN) LEFT JOIN tblKanban ON (tblKompletnyzoznam.Priemer = tblKanban.Priemer) " _
                & "AND (tblKompletnyzoznam.Hrubka = tblKanban.Hrubka) AND (tblKompletnyzoznam.Dlzka = tblKanban.Dlzka)) LEFT JOIN tblDocuments ON tblHlavnyPlan.PN = tblDocuments.Title " _
                & "WHERE tblKompletnyZoznam.PN=@PN " _
                & "Order by tblHlavnyPlan.PN, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka;"

        Dim gridcmdInziniering As OleDbCommand = New OleDbCommand(queryInziniering, cn1)
        Dim sdaInziniering As OleDbDataAdapter = New OleDbDataAdapter(gridcmdInziniering)
        gridcmdInziniering.Parameters.AddWithValue("@PN", txtFilterPNInziniering.Text)
        Dim dtInziniering As DataTable = New DataTable()
        sdaInziniering.Fill(dtInziniering)
        gridInziniering.DataSource = dtInziniering
        If cn1.State = ConnectionState.Open Then cn1.Close()

        'gridInziniering.Columns(0).DefaultCellStyle.Format = "dd.MM"

        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridInziniering.RowTemplate.Height = 25

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridInziniering

            .RowTemplate.Height = 25

            .Columns(0).Width = 90
            .Columns(1).Width = 90
            .Columns(2).Width = 90
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 90
            .Columns(6).Width = 90

            .Columns(7).Width = 120
            '.Columns(9).Visible = False
            '.Columns(9).Width = 110

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter



            .RowHeadersVisible = False
            'ID,PN,QTY,subPN,Pocetnost, Priemer, Hrubka, Dlzka, OhybRovna, Kanban, AMOB
            .Columns(0).HeaderCell.Value = "Piping"
            .Columns(1).HeaderCell.Value = "subPiping"
            .Columns(2).HeaderCell.Value = "Priemer [mm]"
            .Columns(3).HeaderCell.Value = "Hrúbka [mm]"
            .Columns(4).HeaderCell.Value = "Dĺžka [mm]"
            .Columns(5).HeaderCell.Value = "Ohyb / Rovná"
            .Columns(6).HeaderCell.Value = "Kanban"
            .Columns(7).HeaderCell.Value = "Skontrolované"

            .ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)

        End With

        'nastavenie sortovania
        'gridInziniering.Sort(gridInziniering.Columns(0), System.ComponentModel.ListSortDirection.Descending)
        'gridInziniering.Sort(gridInziniering.Columns(1), System.ComponentModel.ListSortDirection.Descending)
        txtInzinieringCount.Text = gridInziniering.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()

    End Sub

    Private Sub btnInziniering_Click(sender As Object, e As EventArgs) Handles btnInziniering.Click
        populateGridInziniering()
    End Sub

    Private Sub BtnKompletZoznam_Click(sender As Object, e As EventArgs) Handles btnKompletZoznam.Click
        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        If cn1.State = ConnectionState.Closed Then cn1.Open()
        Dim queryInziniering As String = "SELECT DISTINCT tblHlavnyPlan.PN, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka, " _
                & "tblKompletnyzoznam.Dlzka, tblKompletnyzoznam.OhybRovna, tblKanban.Kanban, tblKompletnyZoznam.Skontrolovane " _
                & "FROM ((tblHlavnyPlan LEFT JOIN tblKompletnyzoznam ON tblHlavnyPlan.PN = tblKompletnyzoznam.PN) LEFT JOIN tblKanban ON (tblKompletnyzoznam.Priemer = tblKanban.Priemer) " _
                & "AND (tblKompletnyzoznam.Hrubka = tblKanban.Hrubka) AND (tblKompletnyzoznam.Dlzka = tblKanban.Dlzka)) LEFT JOIN tblDocuments ON tblHlavnyPlan.PN = tblDocuments.Title " _
                & "Order by tblHlavnyPlan.PN, tblKompletnyzoznam.subPN, tblKompletnyzoznam.Priemer, tblKompletnyzoznam.Hrubka;"

        Dim gridcmdInziniering As OleDbCommand = New OleDbCommand(queryInziniering, cn1)
        Dim sdaInziniering As OleDbDataAdapter = New OleDbDataAdapter(gridcmdInziniering)
        Dim dtInziniering As DataTable = New DataTable()
        sdaInziniering.Fill(dtInziniering)
        gridInziniering.DataSource = dtInziniering
        If cn1.State = ConnectionState.Open Then cn1.Close()

        'gridInziniering.Columns(0).DefaultCellStyle.Format = "dd.MM"

        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridInziniering.RowTemplate.Height = 25

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridInziniering

            .RowTemplate.Height = 25

            .Columns(0).Width = 90
            .Columns(1).Width = 90
            .Columns(2).Width = 90
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 90
            .Columns(6).Width = 90

            .Columns(7).Width = 120
            '.Columns(9).Visible = False
            '.Columns(9).Width = 110

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            '.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False
            'ID,PN,QTY,subPN,Pocetnost, Priemer, Hrubka, Dlzka, OhybRovna, Kanban, AMOB
            .Columns(0).HeaderCell.Value = "Piping"
            .Columns(1).HeaderCell.Value = "subPiping"
            .Columns(2).HeaderCell.Value = "Priemer [mm]"
            .Columns(3).HeaderCell.Value = "Hrúbka [mm]"
            .Columns(4).HeaderCell.Value = "Dĺžka [mm]"
            .Columns(5).HeaderCell.Value = "Ohyb / Rovná"
            .Columns(6).HeaderCell.Value = "Kanban"
            .Columns(7).HeaderCell.Value = "Skontrolované"

            .ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)

        End With

        'nastavenie sortovania
        'gridInziniering.Sort(gridInziniering.Columns(0), System.ComponentModel.ListSortDirection.Descending)
        'gridInziniering.Sort(gridInziniering.Columns(1), System.ComponentModel.ListSortDirection.Descending)
        txtInzinieringCount.Text = gridInziniering.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()
    End Sub

    Private Function PreparePrintDocumentOhybanie() As PrintDocument
        ' Make the PrintDocument object.
        Dim print_document As New PrintDocument

        ' Install the PrintPage event handler.
        AddHandler print_document.PrintPage, AddressOf _
        PrintOhybanie_PrintPage
        print_document.DefaultPageSettings.Landscape = True
        ' Return the object.
        Return print_document
    End Function



    ' Print the next page.
    Private Sub PrintOhybanie_PrintPage(ByVal sender As Object, ByVal e _
    As System.Drawing.Printing.PrintPageEventArgs)

        txtDatumStavbyOhybanie.Text = Format(CDate(txtDatumStavbyOhybanie.Text), "dd.MM")

        Dim fnt As New Font("Arial", 160, FontStyle.Regular, GraphicsUnit.Point)
        e.Graphics.DrawString(txtPNOhybanie.Text.ToString, fnt, Brushes.Black, 100, 100)

        Dim fntSmall As New Font("Arial", 40, FontStyle.Regular, GraphicsUnit.Point)
        Dim fntSmallValue As New Font("Arial", 50, FontStyle.Bold, GraphicsUnit.Point)

        e.Graphics.DrawString("Množstvo: ", fntSmall, Brushes.DarkGray, 200, 400)
        e.Graphics.DrawString(txtQTYOhybanie.Text.ToString, fntSmallValue, Brushes.Black, 600, 400)

        e.Graphics.DrawString("Dátum stavby: ", fntSmall, Brushes.DarkGray, 200, 480)
        e.Graphics.DrawString(txtDatumStavbyOhybanie.Text.ToString, fntSmallValue, Brushes.Black, 600, 480)

        e.Graphics.DrawString("Skupina: ", fntSmall, Brushes.DarkGray, 200, 560)
        e.Graphics.DrawString(txtSkupinaOhybanie.Text, fntSmallValue, Brushes.Black, 600, 560)

        e.Graphics.DrawString("WJ Jednotky: ", fntSmall, Brushes.DarkGray, 200, 650)
        e.Graphics.DrawString(txtJednotkaWJOhybanie.Text.ToString, fntSmallValue, Brushes.Black, 600, 650)

        ' There are no more pages.
        e.HasMorePages = False
    End Sub
    Private Sub btnPrintPreview_Click(ByVal sender As _
    System.Object, ByVal e As System.EventArgs)
        ' Make a PrintDocument and attach it to 
        ' the PrintPreview dialog.
        dlgPrintPreview.Document = PreparePrintDocumentOhybanie()

        ' Preview.
        dlgPrintPreview.WindowState = FormWindowState.Maximized
        dlgPrintPreview.ShowDialog()
    End Sub
    Private Sub btnPrintWithDialog_Click(ByVal sender As _
    System.Object, ByVal e As System.EventArgs)
        ' Make a PrintDocument and attach it to 
        ' the Print dialog.
        dlgPrint.Document = PreparePrintDocumentOhybanie()

        ' Display the print dialog.
        dlgPrint.ShowDialog()
    End Sub
    Private Sub btnPrintNow_Click(ByVal sender As _
    System.Object, ByVal e As System.EventArgs)
        ' Make a PrintDocument object.
        Dim print_document As PrintDocument =
            PreparePrintDocumentOhybanie()
        print_document.DefaultPageSettings.Landscape = True

        ' Print immediately.
        print_document.Print()
    End Sub

    Private Sub BtnSprievodnyListokOhybanie_Click(sender As Object, e As EventArgs) Handles btnSprievodnyListokOhybanie.Click

        ' Make a PrintDocument object.
        For index As Integer = 0 To gridOhybanie.SelectedRows.Count - 1
            Dim row As DataGridViewRow = gridOhybanie.SelectedRows.Item(index)
            txtPNOhybanie.Text = row.Cells(3).Value
            txtDatumStavbyOhybanie.Text = row.Cells(1).Value
            txtQTYOhybanie.Text = row.Cells(5).Value
            txtSkupinaOhybanie.Text = row.Cells(2).Value
            txtJednotkaWJOhybanie.Text = row.Cells(4).Value
            Dim print_document As PrintDocument =
            PreparePrintDocumentOhybanie()
            print_document.DefaultPageSettings.Landscape = True

            ' Print immediately.
            print_document.Print()
        Next


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

        'Vyfarbenie duplikatnych riadkov

        For i = 0 To gridHlavnyPlan.RowCount - 2
            For j = i + 1 To gridHlavnyPlan.RowCount - 1
                Dim row2 = gridHlavnyPlan.Rows(j)
                If Not row2.IsNewRow Then
                    Dim row1 = gridHlavnyPlan.Rows(i)
                    If row1.Cells("PN").Value.ToString() = row2.Cells("PN").Value.ToString() Then

                        row1.DefaultCellStyle.BackColor = Color.LightGray
                        row2.DefaultCellStyle.BackColor = Color.LightGray
                    End If
                End If
            Next
        Next



    End Sub

    Private Sub BtnFindPNVyplanovanie_Click(sender As Object, e As EventArgs) Handles btnFindPNVyplanovanie.Click

        populateGridVyplanovanie()

    End Sub

    Private Sub populateGridVyplanovanie()

        'pri spusteni okna chceme, aby datagridview bol odfiltrovany len pre jedinu otvorenu polozku
        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        If cn1.State = ConnectionState.Closed Then cn1.Open()

        Dim queryVyplanovanie As String = "SELECT DISTINCT JednotkaWJ, PipingWJ, DatumStavby, Skupina,PN, QTY, Drazkovanie,Urgent,RoutingCas,RezanieStatus, OhybanieStatus, PripravaStatus, " _
                                        & "SpajkovanieStatus, VychystanieStatus, VyplanovanyWJ,VyplanovaneMnozstvo,VyplanovaneDovod FROM tblHlavnyPlan WHERE JednotkaWJ=@JednotkaWJ " _
                                        & "GROUP BY JednotkaWJ, PipingWJ, DatumStavby, Skupina, PN, QTY, Drazkovanie,Urgent,RoutingCas,RezanieStatus, OhybanieStatus, PripravaStatus, SpajkovanieStatus, " _
                                        & "VychystanieStatus,VyplanovanyWJ,VyplanovaneMnozstvo,VyplanovaneDovod;"

        Dim gridcmdVyplanovanie As OleDbCommand = New OleDbCommand(queryVyplanovanie, cn1)
        gridcmdVyplanovanie.Parameters.AddWithValue("@JednotkaWJ", txtFilterWJJednotkyVyplanovanie.Text)
        Dim sdaVyplanovanie As OleDbDataAdapter = New OleDbDataAdapter(gridcmdVyplanovanie)
        Dim dtVyplanovanie As DataTable = New DataTable()
        sdaVyplanovanie.Fill(dtVyplanovanie)
        gridVyplanovanie.DataSource = dtVyplanovanie
        If cn1.State = ConnectionState.Open Then cn1.Close()

        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridVyplanovanie.RowTemplate.Height = 25

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridVyplanovanie

            .Columns(0).Width = 110
            .Columns(1).Width = 110
            .Columns(2).Width = 110
            .Columns(3).Width = 110
            .Columns(4).Width = 110
            .Columns(5).Width = 80
            .Columns(6).Width = 110
            .Columns(7).Width = 110
            '.Columns(8).Width = 110
            .Columns(9).Width = 110
            .Columns(10).Width = 110
            .Columns(11).Width = 110
            .Columns(12).Width = 110
            .Columns(13).Width = 110
            .Columns(8).Visible = False
            .Columns(2).DefaultCellStyle.Format = "dd.MM.yy"

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False
            'JednotkaWJ,PipingWJ,DatumStavby,Skupina,PN,QTY,RezanieStatus,OhybanieStatus,PripravaStatus,SpajkovanieStatus
            .Columns(0).HeaderCell.Value = "WJ Jednotky"
            .Columns(1).HeaderCell.Value = "WJ Pipingu"
            .Columns(2).HeaderCell.Value = "Dátum Stavby"
            .Columns(3).HeaderCell.Value = "Skupina"
            .Columns(4).HeaderCell.Value = "Piping PN"
            .Columns(5).HeaderCell.Value = "Mn."
            .Columns(6).HeaderCell.Value = "Drazkovanie"
            .Columns(7).HeaderCell.Value = "Urgent"
            .Columns(8).HeaderCell.Value = "Routing"
            .Columns(9).HeaderCell.Value = "Rezanie Status"
            .Columns(10).HeaderCell.Value = "Ohýbanie Status"
            .Columns(11).HeaderCell.Value = "Príprava Status"
            .Columns(12).HeaderCell.Value = "Spájkovanie Status"
            .Columns(13).HeaderCell.Value = "Vychystanie Status"
            .Columns(14).HeaderCell.Value = "Vyplanovaný WJ jedn."
            .Columns(15).HeaderCell.Value = "Vyplánované Mn."
            .Columns(16).HeaderCell.Value = "Dôvod vyplánovania"

        End With

        txtVyplanovanieCount.Text = gridVyplanovanie.RowCount
        Try
            txtJednotkaWJVyplanovanie.Text = gridVyplanovanie.Rows(0).Cells(0).Value().ToString
            txtQTYVyplanovanie.Text = gridVyplanovanie.Rows(0).Cells(5).Value().ToString
            txtDatumStavbyVyplanovanie.Text = Format(CDate(gridVyplanovanie.Rows(0).Cells(2).Value()), "dd.MM.yy")
            txtSkupinaVyplanovanie.Text = gridVyplanovanie.Rows(0).Cells(3).Value().ToString
        Catch ex As Exception
            MessageBox.Show("Zadaný WJ jednotky sa nenachádza v pláne.", NazovAplikacie, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
            If cn1.State = ConnectionState.Open Then cn1.Close()
        End Try
        'vyplnenie policok pre vyhladany WJ jednotky



        'txtDatumStavbyVyplanovanie.Text = Format(CDate(txtDatumStavbyVyplanovanie.Text), "dd.MM.YYYY")

        If cn1.State = ConnectionState.Open Then cn1.Close()

    End Sub

    Private Sub BtnVyplanovat_Click(sender As Object, e As EventArgs) Handles btnVyplanovat.Click

        If (MessageBox.Show("Ste si istý, že chcete vyplánovať jednotky podľa zadaných informácií?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbNo Then
            Exit Sub
        End If

        If txtNewQTYVyplanovanie.Text > txtQTYVyplanovanie.Text Then
            MessageBox.Show("Množstvo vyplánovaných WJ jednotiek je väčšie ako množstvo WJ jednotiek v pláne. Zmeňte množstvo.", NazovAplikacie, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        cn1.Open()
        Dim MyDateTimeVariable = Date.Now.ToString("dd MMM yyyy HH:mm:ss") 'format musi byt spravne zadefinovany, inak sa nezobrazi v MDB databaze
        Dim rowcount As Double = 0
        For i As Integer = 0 To gridVyplanovanie.Rows.Count - 1 Step +1
            If IsDBNull(gridVyplanovanie.Rows(i).Cells(0).ToString()) Then
                Continue For
            Else



                Dim cmdImport As New OleDbCommand("INSERT INTO tblVyplanovaneJednotky (OriginalDatumStavby,Skupina,JednotkaWJ,PipingWJ,PN,OdlozeneQTY,Drazkovanie,Urgent,RoutingCas,OriginalQTY,DatumVyplanovania,Dovod) " _
                                                  & "VALUES (@OriginalDatumStavby,@Skupina,@JednotkaWJ,@PipingWJ,@PN,@OdlozeneQTY,@Drazkovanie,@Urgent,@RoutingCas,@OriginalQTY,@DatumVyplanovania,@Dovod)", cn1)
                cmdImport.Parameters.AddWithValue("@OriginalDatumStavby", gridVyplanovanie.Rows(i).Cells(2).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Skupina", gridVyplanovanie.Rows(i).Cells(3).Value().ToString)
                cmdImport.Parameters.AddWithValue("@JednotkaWJ", gridVyplanovanie.Rows(i).Cells(0).Value().ToString)
                cmdImport.Parameters.AddWithValue("@PipingWJ", gridVyplanovanie.Rows(i).Cells(1).Value().ToString)
                cmdImport.Parameters.AddWithValue("@PN", gridVyplanovanie.Rows(i).Cells(4).Value().ToString)
                cmdImport.Parameters.AddWithValue("@OdlozeneQTY", txtNewQTYVyplanovanie.Text)
                cmdImport.Parameters.AddWithValue("@Drazkovanie", CBool(gridVyplanovanie.Rows(i).Cells(6).Value))
                cmdImport.Parameters.AddWithValue("@Urgent", CBool(gridVyplanovanie.Rows(i).Cells(7).Value))
                cmdImport.Parameters.AddWithValue("@RoutingCas", gridVyplanovanie.Rows(i).Cells(8).Value)
                cmdImport.Parameters.AddWithValue("@OriginalQTY", gridVyplanovanie.Rows(i).Cells(5).Value)
                cmdImport.Parameters.AddWithValue("@DatumVyplanovania", MyDateTimeVariable)
                cmdImport.Parameters.AddWithValue("@Dovod", txtDovodVyplanovanie.Text.ToString)
                cmdImport.ExecuteNonQuery()

            End If
            rowcount = rowcount + 1
        Next

        'kod pre ukladanie hodnot z formularu do mdb tabulky - 
        Dim cmdVyplanovanie As New OleDbCommand("UPDATE tblHlavnyPlan SET VyplanovanyWJ=@VyplanovanyWJ, VyplanovaneMnozstvo = @NewQTY, VyplanovaneDovod = @Dovod, VyplanovaneDatum = @VyplanovaneDatum WHERE JednotkaWJ = @JednotkaWJ", cn1)
        cmdVyplanovanie.Parameters.AddWithValue("@VyplanovanyWJ", CBool("TRUE"))
        cmdVyplanovanie.Parameters.AddWithValue("@NewQTY", CDbl(txtNewQTYVyplanovanie.Text))
        cmdVyplanovanie.Parameters.AddWithValue("@Dovod", txtDovodVyplanovanie.Text.ToString)
        cmdVyplanovanie.Parameters.AddWithValue("@VyplanovaneDatum", MyDateTimeVariable)
        cmdVyplanovanie.Parameters.AddWithValue("@JednotkaWJ", txtJednotkaWJVyplanovanie.Text.ToString)
        cmdVyplanovanie.ExecuteNonQuery()

        cn1.Close()

        MessageBox.Show("Informácia o vyplánovaní bola uložená", NazovAplikacie, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        populateGridVyplanovanie()

    End Sub

    Private Sub BtnUkonciOhybanie_Click(sender As Object, e As EventArgs) Handles btnUkonciOhybanie.Click
        Try


            If (MessageBox.Show("Naozaj chcete uzavrieť vyznačený Piping PN?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbNo Then
                Exit Sub
            Else
            End If

            Dim MyDateTimeVariable = Date.Now.ToString("dd MMM yyyy HH:mm:ss") 'format musi byt spravne zadefinovany, inak sa nezobrazi v MDB databaze
            If cn1.State = ConnectionState.Closed Then cn1.Open()

            For index As Integer = 0 To gridOhybanie.SelectedRows.Count - 1
                'Dim rownum As Integer = gridOhybanie.CurrentRow.Index
                Dim rownum As Integer = gridOhybanie.SelectedRows.Item(index).Index

                Dim cmdImport As New OleDbCommand("INSERT INTO tblOhybanieStatus (DatumStavby,Skupina,PN,JednotkaWJ,QTY,subPN,Priemer,Hrubka,Dlzka,OhybanieStatus,Ulozene,IDHlavnyPlan,IDKomplZ) " _
                                              & "VALUES (@DatumStavby,@Skupina,@PN,@JednotkaWJ,@QTY,@subPN,@Priemer,@Hrubka,@Dlzka,@OhybanieStatus,@Ulozene,@IDHlavnyPlan,@IDKomplZ)", cn1)

                cmdImport.Parameters.AddWithValue("@DatumStavby", gridOhybanie.Rows(rownum).Cells(1).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Skupina", gridOhybanie.Rows(rownum).Cells(2).Value().ToString)
                cmdImport.Parameters.AddWithValue("@PN", gridOhybanie.Rows(rownum).Cells(3).Value().ToString)
                cmdImport.Parameters.AddWithValue("@JednotkaWJ", gridOhybanie.Rows(rownum).Cells(4).Value().ToString)
                cmdImport.Parameters.AddWithValue("@QTY", gridOhybanie.Rows(rownum).Cells(5).Value().ToString)
                cmdImport.Parameters.AddWithValue("@subPN", gridOhybanie.Rows(rownum).Cells(6).Value().ToString)
                cmdImport.Parameters.AddWithValue("@Priemer", gridOhybanie.Rows(rownum).Cells(7).Value)
                cmdImport.Parameters.AddWithValue("@Hrubka", CDbl(gridOhybanie.Rows(rownum).Cells(8).Value))
                cmdImport.Parameters.AddWithValue("@Dlzka", gridOhybanie.Rows(rownum).Cells(9).Value)
                cmdImport.Parameters.AddWithValue("@OhybanieStatus", CBool("TRUE"))
                cmdImport.Parameters.AddWithValue("@Ulozene", MyDateTimeVariable)
                cmdImport.Parameters.AddWithValue("@IDHlavnyPlan", gridOhybanie.Rows(rownum).Cells(0).Value)
                cmdImport.Parameters.AddWithValue("@IDKomplZ", gridOhybanie.Rows(rownum).Cells(14).Value)
                cmdImport.ExecuteNonQuery()


                'kod zistuje, ci dany PN je kompletne vyrobeny v dany den. Ak no, tak ho v tblHlavnyPlan ukonci, aby ho videl Pripravar v zozname
                Dim rownumPN As Integer = gridOhybanie.CurrentRow.Index

                'Dim queryCheck As String = "SELECT qrySubGridOhybanie.DatumStavby, qrySubGridOhybanie.PN, qrySubGridOhybanie.Priemer, qrySubGridOhybanie.dlzka, " _
                '        & "qrySubGridOhybanie.OhybRovna, tblOhybanieStatus.OhybanieStatus " _
                '        & "FROM tblOhybanieStatus RIGHT JOIN subQRYGridOhybanie ON (tblOhybanieStatus.IDHlavnyPlan = qrySubGridOhybanie.IDHlavnyPlan) " _
                '        & "AND (tblOhybanieStatus.IDKomplZ = qrySubGridOhybanie.IDKomplZ) " _
                '        & "WHERE (((qrySubGridOhybanie.DatumStavby)=[@DatumStavby]) AND ((qrySubGridOhybanie.PN)=[@PN]) AND ((qrySubGridOhybanie.Priemer)>=12) " _
                '        & "AND ((qrySubGridOhybanie.OhybRovna)='Ohyb') AND ((tblOhybanieStatus.OhybanieStatus) Is Null)) OR (((qrySubGridOhybanie.DatumStavby)=[@DatumStavby]) " _
                '        & "AND ((qrySubGridOhybanie.PN)=[@PN]) AND ((qrySubGridOhybanie.Priemer)>=42) AND ((qrySubGridOhybanie.dlzka)>=300) AND ((qrySubGridOhybanie.OhybRovna)='Rovna') " _
                '        & "AND ((tblOhybanieStatus.OhybanieStatus) Is Null));"

                Dim queryCheck As String = "SELECT qrySubGridOhybanie.DatumStavby, qrySubGridOhybanie.PN, qrySubGridOhybanie.Priemer, qrySubGridOhybanie.dlzka, qrySubGridOhybanie.OhybRovna, tblOhybanieStatus.OhybanieStatus " _
                    & "FROM tblOhybanieStatus RIGHT JOIN qrySubGridOhybanie ON (tblOhybanieStatus.IDKomplZ = qrySubGridOhybanie.IDKomplZ) AND (tblOhybanieStatus.IDHlavnyPlan = qrySubGridOhybanie.IDHlavnyPlan) " _
                    & "WHERE (((qrySubGridOhybanie.DatumStavby)=[@DatumStavby]) AND ((qrySubGridOhybanie.PN)=[@PN]) AND ((qrySubGridOhybanie.Priemer)>=12) AND ((qrySubGridOhybanie.OhybRovna)='Ohyb') AND ((tblOhybanieStatus.OhybanieStatus) Is Null)) OR (((qrySubGridOhybanie.DatumStavby)=[@DatumStavby]) AND ((qrySubGridOhybanie.PN)=[@PN]) AND ((qrySubGridOhybanie.Priemer)>=42) AND ((qrySubGridOhybanie.dlzka)>=300) AND ((qrySubGridOhybanie.OhybRovna)='Rovna') AND ((tblOhybanieStatus.OhybanieStatus) Is Null));"

                Dim gridcmd As OleDbCommand = New OleDbCommand(queryCheck, cn1)
                Dim sda As OleDbDataAdapter = New OleDbDataAdapter(gridcmd)
                Dim dt As DataTable = New DataTable()
                gridcmd.Parameters.AddWithValue("@DatumStavby", gridOhybanie.Rows(rownum).Cells(1).Value().ToString)
                gridcmd.Parameters.AddWithValue("@PN", gridOhybanie.Rows(rownum).Cells(3).Value().ToString)
                sda.Fill(dt)


                If dt.Rows.Count = 0 Then
                    Dim cmd As New OleDbCommand("UPDATE tblHlavnyPlan SET OhybanieCompleted = @DateTimeStop, OhybanieStatus = YES " _
                                            & "WHERE PN=@PN AND DatumStavby=@DatumStavby", cn1)
                    cmd.Parameters.AddWithValue("@DateTimeStop", MyDateTimeVariable)
                    cmd.Parameters.AddWithValue("@PN", gridOhybanie.Rows(rownum).Cells(3).Value().ToString)
                    cmd.Parameters.AddWithValue("@DatumStavby", gridOhybanie.Rows(rownum).Cells(1).Value().ToString)
                    cmd.ExecuteNonQuery()
                End If

                '    If count = True Then
                'Else
                'End If
            Next
            If cn1.State = ConnectionState.Open Then cn1.Close()

            'populateGridRezanie()
            populateGridOhybanie()
            populategridPriprava()

            'ak bol pred ukoncenim vybraty filter v lvPriemery, tak sa znovu nahodi
            If lvPriemeryOhybanie.SelectedItems.Count > 0 Then
                ListPriemeryOhybanie_SelectedIndexChanged(sender, e)
            End If
            txtPNOhybanie.Text = ""
            txtQTYOhybanie.Text = ""
            txtIDOhybanie.Text = ""
            txtJednotkaWJOhybanie.Text = ""
            txtDatumStavbyOhybanie.Text = ""
            txtSkupinaOhybanie.Text = ""

            'ak bol pred ukoncenim vyfiltrovany konkretny PN, tak sa znovu nahodi
            If txtFilterPNOhybanie.Equals("") Then
                BtnFindPNOhybanie_Click(sender, e)
            End If
            SetLinesColorOhybanie()

            gridOhybanie.FirstDisplayedScrollingRowIndex = indexOhybanie

        Catch ex As Exception
            If MessageBox.Show("Prepáčte, nastala chyba. Prosím opakujte akciu neskôr." & ex.Message, NazovAplikacie, MessageBoxButtons.OK, MessageBoxIcon.Question) = vbOK Then
                ErrorHandling(ex.Message, ex.StackTrace, ex.Source)
                Exit Sub
            End If

            If cn1.State = ConnectionState.Open Then cn1.Close()
        End Try

    End Sub

    Private Sub BtnSprievodnyListokRezanie_Click(sender As Object, e As EventArgs) Handles btnSprievodnyListokRezanie.Click

        ' Make a PrintDocument object.
        Dim print_document As PrintDocument =
            PrepareSprievodnyListokRezanie()
        print_document.DefaultPageSettings.Landscape = True

        ' Print immediately.
        print_document.Print()

    End Sub

    Private Function PrepareSprievodnyListokRezanie() As PrintDocument
        ' Make the PrintDocument object.
        Dim print_document As New PrintDocument

        ' Install the PrintPage event handler.
        AddHandler print_document.PrintPage, AddressOf _
        PrintRezanie_PrintPage
        print_document.DefaultPageSettings.Landscape = True
        ' Return the object.
        Return print_document
    End Function

    Private Sub PrintRezanie_PrintPage(ByVal sender As Object, ByVal e _
    As System.Drawing.Printing.PrintPageEventArgs)

        txtDatumStavbyOhybanie.Text = Format(CDate(txtDatumStavbyRezanie.Text), "dd.MM")


        Dim fntSmall As New Font("Arial", 50, FontStyle.Regular, GraphicsUnit.Point)
        Dim fntSmallValue As New Font("Arial", 80, FontStyle.Bold, GraphicsUnit.Point)


        e.Graphics.DrawString("Dátum stavby: ", fntSmall, Brushes.DarkGray, 100, 100)
        e.Graphics.DrawString(txtDatumStavbyRezanie.Text.ToString, fntSmallValue, Brushes.Black, 600, 100)

        e.Graphics.DrawString("Priemer: ", fntSmall, Brushes.DarkGray, 100, 220)
        e.Graphics.DrawString(txtPriemerRezanie.Text.ToString, fntSmallValue, Brushes.Black, 600, 220)

        e.Graphics.DrawString("Hrúbka: ", fntSmall, Brushes.DarkGray, 100, 340)
        e.Graphics.DrawString(txtHrubkaRezanie.Text.ToString, fntSmallValue, Brushes.Black, 600, 340)

        e.Graphics.DrawString("Dĺžka: ", fntSmall, Brushes.DarkGray, 100, 460)
        e.Graphics.DrawString(txtDlzkaRezanie.Text.ToString, fntSmallValue, Brushes.Black, 600, 460)

        e.Graphics.DrawString("Množstvo: ", fntSmall, Brushes.DarkGray, 100, 580)
        e.Graphics.DrawString(txtQTYRezanie.Text.ToString, fntSmallValue, Brushes.Black, 600, 580)

        ' There are no more pages.
        e.HasMorePages = False
    End Sub

    Private Sub BtnUkonciPriprava_Click(sender As Object, e As EventArgs) Handles btnUkonciPriprava.Click
        If (MessageBox.Show("Naozaj chcete uzavrieť vyznačený Piping PN?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbNo Then
            Exit Sub
        Else
        End If

        Dim MyDateTimeVariable = Date.Now.ToString("dd MMM yyyy HH:mm:ss") 'format musi byt spravne zadefinovany, inak sa nezobrazi v MDB databaze
        If cn1.State = ConnectionState.Closed Then cn1.Open()

        Dim cmd As New OleDbCommand("UPDATE tblHlavnyPlan SET PripravaCompleted = @DateTimeStop, PripravaStatus = YES " _
                                    & "WHERE ID=@ID", cn1)
        cmd.Parameters.AddWithValue("@DateTimeStop", MyDateTimeVariable)
        cmd.Parameters.AddWithValue("@ID", txtIDPriprava.Text)
        'cmd.Parameters.AddWithValue("@Skupina", txtSkupinaPriprava.Text)
        'cmd.Parameters.AddWithValue("@DatumStavby", txtDatumStavbyPriprava.Text)
        'cmd.Parameters.AddWithValue("@JednotkaWJ", txtJednotkaWJPriprava.Text)
        cmd.ExecuteNonQuery()

        cn1.Close()


        populategridPriprava()
        populateGridSpajkovanie()
        SetLinesColorPriprava()

        txtPNPriprava.Text = ""
        txtQTYPriprava.Text = ""
        txtIDPriprava.Text = ""
        txtDatumStavbyPriprava.Text = ""
        txtSkupinaPriprava.Text = ""
        txtJednotkaWJPriprava.Text = ""

        gridPriprava.FirstDisplayedScrollingRowIndex = indexPriprava

    End Sub

    Private Sub BtnFindPNPriprava_Click(sender As Object, e As EventArgs) Handles btnFindPNPriprava.Click



        cn1.Open()

        'Dim query As String = "SELECT DISTINCT qryPripravaFilterHiLevel.DatumStavby, tblHlavnyPlan.Skupina, qryPripravaFilterHiLevel.PN, tblHlavnyPlan.QTY, " _
        '                        & "tblHlavnyPlan.JednotkaWJ, qryPripravaFilterHiLevel.CountOfOhybanieStatus, qryPripravaFilterHiLevel.UncompletedPNs, tblHlavnyPlan.PripravaStatus " _
        '                        & "FROM tblHlavnyPlan RIGHT JOIN qryPripravaFilterHiLevel ON tblHlavnyPlan.PN = qryPripravaFilterHiLevel.PN;"

        Dim query As String = "SELECT ID, DatumStavby, Skupina, PN, QTY, JednotkaWJ,PripravaStatus FROM tblHlavnyPlan " _
                                & "Where PN=@PN AND OhybanieStatus=YES;"

        Dim gridcmd As OleDbCommand = New OleDbCommand(query, cn1)
        Dim sda As OleDbDataAdapter = New OleDbDataAdapter(gridcmd)
        gridcmd.Parameters.AddWithValue("@PN", txtFindPNPriprava.Text)
        Dim dt As DataTable = New DataTable()
        cn1.Close()
        sda.Fill(dt)
        gridPriprava.DataSource = dt
        If cn1.State = ConnectionState.Open Then cn1.Close()

        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridPriprava.RowTemplate.Height = 35

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridPriprava

            .Columns(1).Width = 90
            .Columns(2).Width = 90
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 90
            .Columns(6).Width = 90

            .Columns(0).Visible = False

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(1).HeaderCell.Value = "Dátum Stavby"
            .Columns(2).HeaderCell.Value = "Skupina"
            .Columns(3).HeaderCell.Value = "Piping PN"
            .Columns(4).HeaderCell.Value = "Množstvo"
            .Columns(5).HeaderCell.Value = "WJ Jednotky"
            .Columns(6).HeaderCell.Value = "Hotovo"

        End With
        gridPriprava.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        txtPripravaCount.Text = gridPriprava.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()
        SetLinesColorPriprava()

    End Sub

    Private Sub btnImportExcel_Click(sender As Object, e As EventArgs)

        ''----------------------------------------------------------------------------------------------------
        ''otvori okno pre oznacenie suboru XLS (nie XLSX!), z ktoreho budu data importovane z harku 'Denny Plan'
        If OpenFileDialog1.ShowDialog() = DialogResult.Cancel Then
            MsgBox("Súbor nebol označený")
        Else
            Dim HlavnyPlanLocation As String = OpenFileDialog1.FileName
            If HlavnyPlanLocation <> "" Then
                'file is selected
                Dim parser As New FileIO.TextFieldParser(HlavnyPlanLocation)
            End If
        End If
        MsgBox(HlavnyPlanLocation)

        'Dim stream As FileStream = File.Open("C:\0\Denny plan.xls", FileMode.Open, FileAccess.Read)
        Dim stream As FileStream = File.Open(HlavnyPlanLocation, FileMode.Open, FileAccess.Read)
        Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)

        'excelReader.IsFirstRowAsColumnNames = True
        Dim result As DataSet = excelReader.AsDataSet()
        gridHlavnyPlan.DataSource = result.Tables(0)

        excelReader.Close()
        stream.Close()

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs)

        Dim OpenDLG As New OpenFileDialog
        OpenDLG.ShowDialog()
        Dim xPath As String = OpenDLG.FileName
        Dim wPathLong = System.IO.Path.GetDirectoryName(xPath)
        'MsgBox(xPath)
        Dim stream As FileStream = File.Open(xPath, FileMode.Open, FileAccess.Read)
        'Dim stream As FileStream = File.Open("c:\download\ZOZ.XLSX", FileMode.Open, FileAccess.Read)
        Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)



        Dim result As DataSet = excelReader.AsDataSet()

        excelReader.Close()

        gridHlavnyPlan.DataSource = result.Tables(0)
        Dim a As Integer = gridHlavnyPlan.ColumnCount

        For i = 0 To a - 1
            gridHlavnyPlan.Columns(i).HeaderCell.Value = gridHlavnyPlan.Rows(0).Cells(i).Value
        Next
        gridHlavnyPlan.Rows.Remove(gridHlavnyPlan.Rows(0))

        'gridHlavnyPlan.Columns(0).DefaultCellStyle.Format = "dd.MM"
        'Dim d As Date = Date.FromOADate(39051.4387847222)

    End Sub

    Private Sub GetDRWLink()

        Try
            If cn1.State = ConnectionState.Closed Then cn1.Open()
            Dim getDRW As New OleDbCommand("SELECT Name from tblDocuments where Title=@PN", cn1)
            getDRW.Parameters.AddWithValue("@PN", pickPN)
            DRWLink = getDRW.ExecuteScalar()

            webAddress = HyperlinkSplitFunction(DRWLink)
            Process.Start(webAddress)
            If cn1.State = ConnectionState.Open Then cn1.Close()

        Catch ex As Exception
            'MessageBox.Show(" ", NazovAplikacie, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            If MessageBox.Show("Výkres sa nepodarilo nájsť. Chcete, aby ste boli presmerovaný na server s výkresami?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
                'MessageBox.Show(ex.Message, NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                Dim webAddress As String = "https://vertivco.sharepoint.com/teams/engineering/emea/NMV/SitePages/Home.aspx"
                Process.Start(webAddress)

            End If

        End Try



    End Sub

    Private Sub BtnOpenDRWOhybanie_Click(sender As Object, e As EventArgs) Handles btnOpenDRWOhybanie.Click, btnOpenDRWPriprava.Click, btnOpenDRWSpajkovanie.Click, btnOpenDRWVychystanie.Click, btnOpenDRWInziniering.Click, btnOpenDRWRezanie.Click

        GetDRWLink()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        populategridPriprava()

    End Sub

    Private Sub BtnUkonciSpajkovanie_Click(sender As Object, e As EventArgs) Handles btnUkonciSpajkovanie.Click

        If (MessageBox.Show("Naozaj chcete uzavrieť vyznačený Piping PN?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbNo Then
            Exit Sub
        Else
        End If

        Dim MyDateTimeVariable = Date.Now.ToString("dd MMM yyyy HH:mm:ss") 'format musi byt spravne zadefinovany, inak sa nezobrazi v MDB databaze
        cn1.Open()

        Dim cmd As New OleDbCommand("UPDATE tblHlavnyPlan SET SpajkovanieCompleted = @DateTimeStop, SpajkovanieStatus = YES " _
                                    & "WHERE ID=@ID", cn1)
        cmd.Parameters.AddWithValue("@DateTimeStop", MyDateTimeVariable)
        cmd.Parameters.AddWithValue("@ID", txtIDSpajkovanie.Text)
        'cmd.Parameters.AddWithValue("@Skupina", txtSkupinaSpajkovanie.Text)
        'cmd.Parameters.AddWithValue("@DatumStavby", txtDatumStavbySpajkovanie.Text)
        'cmd.Parameters.AddWithValue("@JednotkaWJ", txtJednotkaWJSpajkovanie.Text)
        cmd.ExecuteNonQuery()

        cn1.Close()



        populateGridSpajkovanie()
        SetLinesColorSpajkovanie()

        txtPNSpajkovanie.Text = ""
        txtQTYSpajkovanie.Text = ""
        txtIDSpajkovanie.Text = ""
        txtDatumStavbySpajkovanie.Text = ""
        txtSkupinaSpajkovanie.Text = ""
        txtJednotkaWJSpajkovanie.Text = ""

        gridSpajkovanie.FirstDisplayedScrollingRowIndex = indexSpajkovanie

    End Sub

    Private Sub GridOhybanie_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles gridOhybanie.ColumnHeaderMouseClick

        SetLinesColorOhybanie()

    End Sub

    Private Sub GridRezanie_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles gridRezanie.ColumnHeaderMouseClick

        SetLinesColorRezanie()

    End Sub

    Private Sub GridPriprava_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles gridPriprava.ColumnHeaderMouseClick

        SetLinesColorPriprava()

    End Sub

    Private Sub GridSpajkovanie_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles gridSpajkovanie.ColumnHeaderMouseClick

        SetLinesColorSpajkovanie()

    End Sub

    Private Sub gridVychystanie_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs)

        SetLinesColorVyplanovany()
        SetLinesColorVychystanie()

    End Sub

    Private Sub GridSpajkovanie_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridSpajkovanie.CellClick
        indexSpajkovanie = e.RowIndex
        Dim j As Integer



        With gridSpajkovanie
            If e.RowIndex >= 0 Then
                j = .CurrentRow.Index

                txtPNSpajkovanie.Text = .Rows(j).Cells("PN").Value.ToString
                txtQTYSpajkovanie.Text = .Rows(j).Cells("QTY").Value.ToString
                txtIDSpajkovanie.Text = .Rows(j).Cells("ID").Value.ToString
                txtJednotkaWJSpajkovanie.Text = .Rows(j).Cells("JednotkaWJ").Value.ToString
                txtSkupinaSpajkovanie.Text = .Rows(j).Cells("Skupina").Value.ToString
                txtDatumStavbySpajkovanie.Text = .Rows(j).Cells("DatumStavby").Value.ToString
                txtDatumStavbySpajkovanie.Text = Format(CDate(txtDatumStavbySpajkovanie.Text), "dd.MM")
                pickPN = .Rows(j).Cells("PN").Value.ToString

            End If
        End With

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        txtFindPNPriprava.Text = ""
        populategridPriprava()
        SetLinesColorPriprava()
    End Sub

    Private Sub BtnFindPNSpajkovanie_Click(sender As Object, e As EventArgs) Handles btnFindPNSpajkovanie.Click


        cn1.Open()

        'Dim query As String = "SELECT DISTINCT qrySpajkovanieFilterHiLevel.DatumStavby, tblHlavnyPlan.Skupina, qrySpajkovanieFilterHiLevel.PN, tblHlavnyPlan.QTY, " _
        '                        & "tblHlavnyPlan.JednotkaWJ, qrySpajkovanieFilterHiLevel.CountOfOhybanieStatus, qrySpajkovanieFilterHiLevel.UncompletedPNs, tblHlavnyPlan.SpajkovanieStatus " _
        '                        & "FROM tblHlavnyPlan RIGHT JOIN qrySpajkovanieFilterHiLevel ON tblHlavnyPlan.PN = qrySpajkovanieFilterHiLevel.PN;"

        Dim query As String = "SELECT ID, DatumStavby, Skupina, PN, QTY, JednotkaWJ,SpajkovanieStatus FROM tblHlavnyPlan " _
                                & "Where PN=@PN AND OhybanieStatus=YES;"

        Dim gridcmd As OleDbCommand = New OleDbCommand(query, cn1)
        Dim sda As OleDbDataAdapter = New OleDbDataAdapter(gridcmd)
        gridcmd.Parameters.AddWithValue("@PN", txtFindPNSpajkovanie.Text)
        Dim dt As DataTable = New DataTable()
        cn1.Close()
        sda.Fill(dt)
        gridSpajkovanie.DataSource = dt
        If cn1.State = ConnectionState.Open Then cn1.Close()

        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridSpajkovanie.RowTemplate.Height = 35

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridSpajkovanie

            .Columns(1).Width = 90
            .Columns(2).Width = 90
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 90
            .Columns(6).Width = 90

            .Columns(0).Visible = False

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(1).HeaderCell.Value = "Dátum Stavby"
            .Columns(2).HeaderCell.Value = "Skupina"
            .Columns(3).HeaderCell.Value = "Piping PN"
            .Columns(4).HeaderCell.Value = "Množstvo"
            .Columns(5).HeaderCell.Value = "WJ Jednotky"
            .Columns(6).HeaderCell.Value = "Hotovo"

        End With
        gridSpajkovanie.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        txtSpajkovanieCount.Text = gridSpajkovanie.RowCount
        If cn1.State = ConnectionState.Open Then cn1.Close()
        SetLinesColorSpajkovanie()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        txtFindPNSpajkovanie.Text = ""
        populateGridSpajkovanie()
        SetLinesColorSpajkovanie()
    End Sub





    Private Sub btnSaveKanban_Click(sender As Object, e As EventArgs) Handles btnSaveKanban.Click

        'kontrola, ci priemer+hrubka+dlzka uz existuje
        cn1.Open()
        Dim queryKanbanCheck As String = "Select Priemer,Hrubka,Dlzka from tblKanban Where Priemer=@Priemer AND Hrubka=@Hrubka AND Dlzka=@Dlzka;"
        Dim gridcmdKanbanCheck As OleDbCommand = New OleDbCommand(queryKanbanCheck, cn1)
        gridcmdKanbanCheck.Parameters.AddWithValue("@Priemer", txtPriemerKanban.Text)
        gridcmdKanbanCheck.Parameters.AddWithValue("@Hrubka", txtHrubkaKanban.Text)
        gridcmdKanbanCheck.Parameters.AddWithValue("@Dlzka", txtDlzkaKanban.Text)
        Dim sdaKanbanCheck As OleDbDataAdapter = New OleDbDataAdapter(gridcmdKanbanCheck)
        Dim dtKanbanCheck As DataTable = New DataTable()
        cn1.Close()
        sdaKanbanCheck.Fill(dtKanbanCheck)

        If dtKanbanCheck.Rows.Count > 0 Then
            MessageBox.Show("Zadaný priemer/hrúbka/dĺžka už zozname existuje. Tentokráť ste nič neuložili.")
            Exit Sub
        End If

        'ulozenie polozky
        cn1.Open()
        Dim cmd As New OleDbCommand("INSERT INTO tblKanban (Priemer, Hrubka,Dlzka,Kanban) " _
                                    & " VALUES (@Priemer,@Hrubka,@Dlzka,@Kanban)", cn1)

        cmd.Parameters.AddWithValue("@Priemer", txtPriemerKanban.Text)
        cmd.Parameters.AddWithValue("@Hrubka", txtHrubkaKanban.Text)
        cmd.Parameters.AddWithValue("@Dlzka", txtDlzkaKanban.Text)
        cmd.Parameters.AddWithValue("@Kanban", CBool("TRUE"))


        cmd.ExecuteNonQuery()
        cn1.Close()

        txtPriemerKanban.Text = ""
        txtHrubkaKanban.Text = ""
        txtDlzkaKanban.Text = String.Empty
        txtIDKanban.Text = String.Empty

        populateGridKanban()

        MessageBox.Show("Položka bola uložená")

    End Sub





    Private Sub GridInziniering_Click(sender As Object, e As DataGridViewCellEventArgs) Handles gridInziniering.CellClick
        indexInziniering = e.RowIndex
        If e.RowIndex > -1 Then
            pickPN = gridInziniering.Rows(e.RowIndex).Cells("PN").Value.ToString
            ' str1 = gridInziniering.Rows(e.RowIndex).Cells("Vykres").Value.ToString
        End If


    End Sub

    Private Sub BtnAddPlanovanie_Click(sender As Object, e As EventArgs) Handles btnAddPlanovanie.Click

        If MessageBox.Show("Naozaj chcete uložiť novú položku?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbNo Then Exit Sub

        If cn1.State = ConnectionState.Closed Then cn1.Open()
        Dim cmd1 As New OleDbCommand("SELECT * from tblHlavnyPlan where PipingWJ=@PipingWJ", cn1)
        cmd1.Parameters.AddWithValue("@PipingWJ", txtPipingWJPlanovanie.Text.ToString)
        Dim check As Boolean = cmd1.ExecuteScalar()
        If check = True Then

        Else
            Dim cmdImport As New OleDbCommand("INSERT INTO tblHlavnyPlan (DatumStavby,Skupina,JednotkaWJ,PipingWJ,PN,QTY,Drazkovanie,RoutingCas,Urgent) " _
                                              & "VALUES (@DatumStavby,@Skupina,@JednotkaWJ,@PipingWJ,@PN,@QTY,@Drazkovanie,@RoutingCas,@Urgent)", cn1)
            cmdImport.Parameters.AddWithValue("@DatumStavby", Format(dtpDatumStavbyPlanovanie.Value.Date, "dd/MM/yyyy"))
            cmdImport.Parameters.AddWithValue("@Skupina", cboSkupinaPlanovanie.Text.ToString)
            cmdImport.Parameters.AddWithValue("@JednotkaWJ", txtJednotkaWJPlanovanie.Text.ToString)
            cmdImport.Parameters.AddWithValue("@PipingWJ", txtPipingWJPlanovanie.Text.ToString)
            cmdImport.Parameters.AddWithValue("@PN", txtPNPlanovanie.Text.ToString)
            cmdImport.Parameters.AddWithValue("@QTY", CDbl(txtQTYPlanovanie.Text))
            cmdImport.Parameters.AddWithValue("@Drazkovanie", CBool(ckbDrazkovaniePlanovanie.CheckState))
            cmdImport.Parameters.AddWithValue("@RoutingCas", CDec(txtRoutingPlanovanie.Text))
            cmdImport.Parameters.AddWithValue("@Urgent", CBool(ckbUrgentPlanovanie.CheckState))
            cmdImport.ExecuteNonQuery()
        End If

        txtPlanCount.Text = gridHlavnyPlan.RowCount
        gridHlavnyPlan.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        If cn1.State = ConnectionState.Open Then cn1.Close()
        populateFullPlan()

        MsgBox("Položka bola uložená")
    End Sub

    Private Sub BtnEditPlanovanie_Click(sender As Object, e As EventArgs) Handles btnEditPlanovanie.Click
        If MessageBox.Show("Naozaj chcete upraviť vybranú položku?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbNo Then Exit Sub

        If cn1.State = ConnectionState.Closed Then cn1.Open()
        Dim cmdImport As New OleDbCommand("UPDATE tblHlavnyPlan SET DatumStavby=@DatumStavby,Skupina=@Skupina,JednotkaWJ=@JednotkaWJ,PipingWJ=@PipingWJ,PN=@PN, " _
                                  & "QTY=@QTY,Drazkovanie=@Drazkovanie,RoutingCas=@RoutingCas,Urgent=@Urgent " _
                                  & "WHERE ID=@ID", cn1)
        cmdImport.Parameters.AddWithValue("@DatumStavby", Format(dtpDatumStavbyPlanovanie.Value.Date, "dd/MM/yyyy"))
        cmdImport.Parameters.AddWithValue("@Skupina", cboSkupinaPlanovanie.Text.ToString)
        cmdImport.Parameters.AddWithValue("@JednotkaWJ", txtJednotkaWJPlanovanie.Text.ToString)
        cmdImport.Parameters.AddWithValue("@PipingWJ", txtPipingWJPlanovanie.Text.ToString)
        cmdImport.Parameters.AddWithValue("@PN", txtPNPlanovanie.Text.ToString)
        cmdImport.Parameters.AddWithValue("@QTY", CDbl(txtQTYPlanovanie.Text))
        cmdImport.Parameters.AddWithValue("@Drazkovanie", CBool(ckbDrazkovaniePlanovanie.CheckState))
        cmdImport.Parameters.AddWithValue("@RoutingCas", CDec(txtRoutingPlanovanie.Text))
        cmdImport.Parameters.AddWithValue("@Urgent", CBool(ckbUrgentPlanovanie.CheckState))
        cmdImport.Parameters.AddWithValue("@ID", txtIDPlanovanie.Text)
        cmdImport.ExecuteNonQuery()


        txtPlanCount.Text = gridHlavnyPlan.RowCount
        gridHlavnyPlan.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        If cn1.State = ConnectionState.Open Then cn1.Close()
        populateFullPlan()

        gridHlavnyPlan.FirstDisplayedScrollingRowIndex = indexHlavnyPlan

        MsgBox("Položka bola upravená")
    End Sub

    Private Sub GridHlavnyPlan_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridHlavnyPlan.CellClick

        indexHlavnyPlan = e.RowIndex

        Dim j As Integer

        With gridHlavnyPlan
            If e.RowIndex >= 0 Then
                j = .CurrentRow.Index

                txtIDPlanovanie.Text = .Rows(j).Cells("ID").Value.ToString
                dtpDatumStavbyPlanovanie.Value = .Rows(j).Cells("DatumStavby").Value.ToString
                cboSkupinaPlanovanie.Text = .Rows(j).Cells("Skupina").Value()
                txtJednotkaWJPlanovanie.Text = .Rows(j).Cells("JednotkaWJ").Value.ToString
                txtPipingWJPlanovanie.Text = .Rows(j).Cells("PipingWJ").Value.ToString
                txtPNPlanovanie.Text = .Rows(j).Cells("PN").Value.ToString
                txtQTYPlanovanie.Text = .Rows(j).Cells("QTY").Value.ToString
                txtRoutingPlanovanie.Text = .Rows(j).Cells("RoutingCas").Value.ToString

                If .Rows(j).Cells("Drazkovanie").Value = True Then
                    ckbDrazkovaniePlanovanie.CheckState = CheckState.Checked
                Else
                    ckbDrazkovaniePlanovanie.CheckState = CheckState.Unchecked
                End If

                If .Rows(j).Cells("Urgent").Value = True Then
                    ckbUrgentPlanovanie.CheckState = CheckState.Checked
                Else
                    ckbUrgentPlanovanie.CheckState = CheckState.Unchecked
                End If



                'str1 = gridHlavnyPlan.Rows(e.RowIndex).Cells("Vykres").Value.ToString

            End If
        End With


    End Sub

    Private Sub btnDelPlanovanie_Click(sender As Object, e As EventArgs) Handles btnDelPlanovanie.Click
        If MessageBox.Show("Naozaj chcete vymazať novú položku?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbNo Then Exit Sub

        If cn1.State = ConnectionState.Closed Then cn1.Open()

        Dim cmd2 As New OleDbCommand("DELETE * From tblHlavnyPlan where ID=@ID", cn1)
        cmd2.Parameters.AddWithValue("@ID", txtIDPlanovanie.Text)
        cmd2.ExecuteNonQuery()

        txtPlanCount.Text = gridHlavnyPlan.RowCount
        gridHlavnyPlan.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        If cn1.State = ConnectionState.Open Then cn1.Close()

        populateFullPlan()

        txtIDPlanovanie.Text = ""
        dtpDatumStavbyPlanovanie.Value = ""
        cboSkupinaPlanovanie.Text = ""
        txtJednotkaWJPlanovanie.Text = ""
        txtPipingWJPlanovanie.Text = ""
        txtPNPlanovanie.Text = ""
        txtQTYPlanovanie.Text = ""
        txtRoutingPlanovanie.Text = ""
        ckbDrazkovaniePlanovanie.CheckState = CheckState.Unchecked
        ckbDrazkovaniePlanovanie.CheckState = CheckState.Unchecked

        MsgBox("Položka bola vymazaná")
    End Sub



    Public Sub FillCboSkupina()

        'naplnenie cbo pri vytvoreni noveho produktu (k dispozicii su vsetky operacie pre aktualny PC ID)

        cn1.Open()
        Dim dset2 As New DataTable
        Dim cboSkupinaContent As String = "SELECT DISTINCT Linka FROM tblSkupina;"
        Dim cboSkupinacmd As OleDbCommand = New OleDbCommand(cboSkupinaContent, cn1)
        cboSkupinacmd.ExecuteNonQuery()
        Dim sda2 As OleDbDataAdapter = New OleDbDataAdapter(cboSkupinacmd)
        sda2.Fill(dset2)
        'MsgBox(dset2.Rows.Count)
        cn1.Close()

        'naplnenie cboOperations
        cboSkupinaVychystanie.ValueMember = "Linka"     ' toto si pamata ako value napr cboSkupinaVychystanie.selectedvalue
        cboSkupinaVychystanie.DisplayMember = "Linka"   'toto tu berie ako cboSkupinaVychystanie.text
        cboSkupinaVychystanie.DataSource = dset2

        cboSkupinaVychystanie.SelectedIndex = -1
        cboSkupinaVychystanie.Text = String.Empty

        ''auto complete
        'cboSkupinaVychystanie.AutoCompleteMode = AutoCompleteMode.Append
        'cboSkupinaVychystanie.DropDownStyle = ComboBoxStyle.DropDown
        'cboSkupinaVychystanie.AutoCompleteSource = AutoCompleteSource.ListItems

    End Sub



    Private Sub GridVychystanie_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridVychystanie.CellClick
        indexVychystanie = e.RowIndex
        Dim j As Integer

        With gridVychystanie
            If e.RowIndex >= 0 Then
                j = .CurrentRow.Index

                txtPNVychystanie.Text = .Rows(j).Cells("PN").Value.ToString
                txtQTYVychystanie.Text = .Rows(j).Cells("QTY").Value.ToString
                txtIDVychystanie.Text = .Rows(j).Cells("ID").Value.ToString
                txtJednotkaWJVychystanie.Text = .Rows(j).Cells("JednotkaWJ").Value.ToString
                pickPN = .Rows(j).Cells("PN").Value.ToString

            End If
        End With

    End Sub

    Private Sub btnCompleteVychystanie_Click(sender As Object, e As EventArgs) Handles btnCompleteVychystanie.Click

        If (MessageBox.Show("Naozaj chcete vyznačený riadok označiť ako vychystaný?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbNo Then
            Exit Sub
        Else
        End If

        Dim MyDateTimeVariable = Date.Now.ToString("dd MMM yyyy HH:mm:ss") 'format musi byt spravne zadefinovany, inak sa nezobrazi v MDB databaze
        cn1.Open()


        If gridVychystanie.Rows(indexVychystanie).Cells(7).Value > 0 Then
            MessageBox.Show("Týmto potvrdzujete, že vyplánované množstvo " & gridVychystanie.Rows(indexVychystanie).Cells(7).Value & " ks pipingov odkladáte do skladu!", NazovAplikacie, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End If



        Dim cmd As New OleDbCommand("UPDATE tblHlavnyPlan SET VychystanieCompleted = @DateTimeStop, VychystanieStatus = YES " _
                                    & "WHERE ID=@ID", cn1)
        cmd.Parameters.AddWithValue("@DateTimeStop", MyDateTimeVariable)
        cmd.Parameters.AddWithValue("@ID", txtIDVychystanie.Text)
        cmd.ExecuteNonQuery()

        cn1.Close()



        populategridVychystanie()
        SetLinesColorVychystanie()

        txtPNVychystanie.Text = ""
        txtQTYVychystanie.Text = ""
        txtIDVychystanie.Text = ""
        txtJednotkaWJVychystanie.Text = ""

        gridVychystanie.FirstDisplayedScrollingRowIndex = indexVychystanie

    End Sub

    Private Sub Button2_Click_4(sender As Object, e As EventArgs) Handles btnKompletnyReport.Click
        populateFullPlanGridReport()
    End Sub

    Private Sub BtnVyplanovaneJednotky_Click(sender As Object, e As EventArgs) Handles btnVyplanovaneJednotky.Click

        If cn1.State = ConnectionState.Closed Then cn1.Open()

        Dim queryStatusVyplanovane As String = "SELECT * from tblVyplanovaneJednotky;"

        Dim gridCmdStatusVyplanovane As OleDbCommand = New OleDbCommand(queryStatusVyplanovane, cn1)
        Dim sdaStatusVyplanovane As OleDbDataAdapter = New OleDbDataAdapter(gridCmdStatusVyplanovane)

        Dim dtStatusVyplanovane As DataTable = New DataTable()
        sdaStatusVyplanovane.Fill(dtStatusVyplanovane)

        gridReport.DataSource = dtStatusVyplanovane
        If cn1.State = ConnectionState.Open Then cn1.Close()

        txtPlanCount.Text = gridReport.RowCount

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridReport
            .Columns(0).Visible = False
            .Columns(7).Visible = False
            .Columns(8).Visible = False
            .Columns(9).Visible = False
            .Columns(10).Visible = False
            .Columns(1).Width = 80
            .Columns(2).Width = 80
            .Columns(3).Width = 90
            .Columns(4).Width = 170
            .Columns(5).Width = 120
            .Columns(11).Width = 110
            .Columns(12).Width = 110
            .Columns(13).Width = 180

            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Format = "dd.MM"
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False


            .Columns(1).HeaderCell.Value = "WJ Jednotky"
            .Columns(2).HeaderCell.Value = "WJ Pipingu"
            .Columns(3).HeaderCell.Value = "Piping PN"
            .Columns(4).HeaderCell.Value = "Pôvodný dátum stavby"
            .Columns(5).HeaderCell.Value = "Skupina"
            .Columns(6).HeaderCell.Value = "Odložené množstvo"
            .Columns(7).HeaderCell.Value = "Drážkovanie"
            .Columns(8).HeaderCell.Value = "Urgent"
            .Columns(9).HeaderCell.Value = "RoutingCas"
            .Columns(11).HeaderCell.Value = "Pôvodné množstvo"
            .Columns(12).HeaderCell.Value = "Dátum vyplánovania"
            .Columns(13).HeaderCell.Value = "Dôvod vyplánovania"


        End With
        txtZaplanovanieCount.Text = gridReport.RowCount

        gridReport.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)


    End Sub

    Private Sub BtnJednotkaWJStatus_Click(sender As Object, e As EventArgs) Handles btnJednotkaWJStatus.Click

        If txtJednotkaWJPrehlady.Text = String.Empty Then Exit Sub

        'pri spusteni okna chceme, aby datagridview bol odfiltrovany len pre jedinu otvorenu polozku
        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        If cn1.State = ConnectionState.Closed Then cn1.Open()

        Dim queryVyplanovanie As String = "SELECT DISTINCT JednotkaWJ, PipingWJ, DatumStavby, Skupina,PN, QTY, RezanieStatus, OhybanieStatus, PripravaStatus, " _
                                        & "SpajkovanieStatus, VychystanieStatus, VyplanovanyWJ,VyplanovaneMnozstvo,VyplanovaneDovod FROM tblHlavnyPlan WHERE JednotkaWJ=@JednotkaWJ " _
                                        & "GROUP BY JednotkaWJ, PipingWJ, DatumStavby, Skupina, PN, QTY, RezanieStatus, OhybanieStatus, PripravaStatus, SpajkovanieStatus, " _
                                        & "VychystanieStatus,VyplanovanyWJ,VyplanovaneMnozstvo,VyplanovaneDovod;"

        Dim gridcmdVyplanovanie As OleDbCommand = New OleDbCommand(queryVyplanovanie, cn1)
        gridcmdVyplanovanie.Parameters.AddWithValue("@JednotkaWJ", txtJednotkaWJPrehlady.Text)
        Dim sdaVyplanovanie As OleDbDataAdapter = New OleDbDataAdapter(gridcmdVyplanovanie)
        Dim dtVyplanovanie As DataTable = New DataTable()
        sdaVyplanovanie.Fill(dtVyplanovanie)
        gridReport.DataSource = dtVyplanovanie
        If cn1.State = ConnectionState.Open Then cn1.Close()

        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridReport.RowTemplate.Height = 25

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridReport

            .Columns(0).Width = 110
            .Columns(1).Width = 110
            .Columns(2).Width = 110
            .Columns(3).Width = 110
            .Columns(4).Width = 110
            .Columns(5).Width = 80
            .Columns(6).Width = 110
            .Columns(7).Width = 110
            .Columns(8).Width = 110
            .Columns(9).Width = 110
            .Columns(10).Width = 110
            .Columns(11).Width = 110
            .Columns(12).Width = 110
            .Columns(13).Width = 110

            .Columns(2).DefaultCellStyle.Format = "dd.MM.yy"

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False
            'JednotkaWJ,PipingWJ,DatumStavby,Skupina,PN,QTY,RezanieStatus,OhybanieStatus,PripravaStatus,SpajkovanieStatus
            .Columns(0).HeaderCell.Value = "WJ Jednotky"
            .Columns(1).HeaderCell.Value = "WJ Pipingu"
            .Columns(2).HeaderCell.Value = "Dátum Stavby"
            .Columns(3).HeaderCell.Value = "Skupina"
            .Columns(4).HeaderCell.Value = "Piping PN"
            .Columns(5).HeaderCell.Value = "Mn."
            .Columns(6).HeaderCell.Value = "Rezanie Status"
            .Columns(7).HeaderCell.Value = "Ohýbanie Status"
            .Columns(8).HeaderCell.Value = "Príprava Status"
            .Columns(9).HeaderCell.Value = "Spájkovanie Status"
            .Columns(10).HeaderCell.Value = "Vychystanie Status"
            .Columns(11).HeaderCell.Value = "Vyplanovaný WJ jedn."
            .Columns(12).HeaderCell.Value = "Vyplánované Mn."
            .Columns(13).HeaderCell.Value = "Dôvod vyplánovania"

        End With

        txtPrehladyCount.Text = gridReport.RowCount

        'txtDatumStavbyPrehlady.Text = Format(CDate(txtDatumStavbyPrehlady.Text), "dd.MM")

        If cn1.State = ConnectionState.Open Then cn1.Close()

    End Sub

    Private Sub BtnVyplanovaneJednotkyZaplanovanie_Click(sender As Object, e As EventArgs) Handles btnVyplanovaneJednotkyZaplanovanie.Click

        If cn1.State = ConnectionState.Closed Then cn1.Open()

        Dim queryStatusVyplanovane As String = "SELECT * from tblVyplanovaneJednotky;"

        Dim gridCmdStatusVyplanovane As OleDbCommand = New OleDbCommand(queryStatusVyplanovane, cn1)
        Dim sdaStatusVyplanovane As OleDbDataAdapter = New OleDbDataAdapter(gridCmdStatusVyplanovane)

        Dim dtStatusVyplanovane As DataTable = New DataTable()
        sdaStatusVyplanovane.Fill(dtStatusVyplanovane)

        gridZaplanovanie.DataSource = dtStatusVyplanovane
        If cn1.State = ConnectionState.Open Then cn1.Close()

        txtPlanCount.Text = gridZaplanovanie.RowCount

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridZaplanovanie
            .Columns(0).Visible = False
            .Columns(7).Visible = False
            .Columns(8).Visible = False
            .Columns(9).Visible = False
            .Columns(10).Visible = False
            .Columns(1).Width = 80
            .Columns(2).Width = 80
            .Columns(3).Width = 90
            .Columns(4).Width = 170
            .Columns(5).Width = 120
            .Columns(11).Width = 110
            .Columns(12).Width = 110
            .Columns(13).Width = 180
            .Columns(14).Width = 180

            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Format = "dd.MM"
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False


            .Columns(1).HeaderCell.Value = "WJ Jednotky"
            .Columns(2).HeaderCell.Value = "WJ Pipingu"
            .Columns(3).HeaderCell.Value = "Piping PN"
            .Columns(4).HeaderCell.Value = "Pôvodný dátum stavby"
            .Columns(5).HeaderCell.Value = "Skupina"
            .Columns(6).HeaderCell.Value = "Odložené množstvo"
            .Columns(7).HeaderCell.Value = "Drážkovanie"
            .Columns(8).HeaderCell.Value = "Urgent"
            .Columns(9).HeaderCell.Value = "RoutingCas"
            .Columns(11).HeaderCell.Value = "Pôvodné množstvo"
            .Columns(12).HeaderCell.Value = "Dátum vyplánovania"
            .Columns(13).HeaderCell.Value = "Dôvod vyplánovania"
            .Columns(14).HeaderCell.Value = "Dátum zaplánovania"


        End With
        txtZaplanovanieCount.Text = gridZaplanovanie.RowCount

        gridZaplanovanie.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)

    End Sub

    Private Sub BtnFindPNZaplanovanie_Click(sender As Object, e As EventArgs) Handles btnFindPNZaplanovanie.Click
        populateGridZaplanovanie()

    End Sub

    Private Sub populateGridZaplanovanie()

        If cn1.State = ConnectionState.Closed Then cn1.Open()

        Dim queryStatusZaplanovane As String = "SELECT * from tblVyplanovaneJednotky WHERE JednotkaWJ=@JednotkaWJ;"

        Dim gridCmdStatusZaplanovane As OleDbCommand = New OleDbCommand(queryStatusZaplanovane, cn1)
        Dim sdaStatusZaplanovane As OleDbDataAdapter = New OleDbDataAdapter(gridCmdStatusZaplanovane)
        gridCmdStatusZaplanovane.Parameters.AddWithValue("@JednotkaWJ", txtFilterWJJednotkyZaplanovanie.Text.ToString)
        Dim dtStatusZaplanovane As DataTable = New DataTable()
        sdaStatusZaplanovane.Fill(dtStatusZaplanovane)
        gridZaplanovanie.DataSource = dtStatusZaplanovane
        If cn1.State = ConnectionState.Open Then cn1.Close()

        txtPlanCount.Text = gridZaplanovanie.RowCount

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridZaplanovanie
            .Columns(0).Visible = False
            .Columns(7).Visible = False
            .Columns(8).Visible = False
            .Columns(9).Visible = False
            .Columns(10).Visible = False
            .Columns(1).Width = 80
            .Columns(2).Width = 80
            .Columns(3).Width = 90
            .Columns(4).Width = 170
            .Columns(5).Width = 110
            .Columns(11).Width = 110
            .Columns(12).Width = 110
            .Columns(13).Width = 180

            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Format = "dd.MM"
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False


            .Columns(1).HeaderCell.Value = "WJ Jednotky"
            .Columns(2).HeaderCell.Value = "WJ Pipingu"
            .Columns(3).HeaderCell.Value = "Piping PN"
            .Columns(4).HeaderCell.Value = "Pôvodný dátum stavby"
            .Columns(5).HeaderCell.Value = "Skupina"
            .Columns(6).HeaderCell.Value = "Odložené množstvo"
            .Columns(7).HeaderCell.Value = "Drážkovanie"
            .Columns(8).HeaderCell.Value = "Urgent"
            .Columns(9).HeaderCell.Value = "RoutingCas"
            .Columns(11).HeaderCell.Value = "Pôvodné množstvo"
            .Columns(12).HeaderCell.Value = "Dátum vyplánovania"
            .Columns(13).HeaderCell.Value = "Dôvod vyplánovania"


        End With
        txtZaplanovanieCount.Text = gridZaplanovanie.RowCount

        If dtStatusZaplanovane.Rows.Count = 0 Then
            MsgBox("Zadaný WJ jednotky nebol nájdený v odložených WJoboch.")
            Exit Sub
        Else
            'MsgBox(Format(CDate(gridVyplanovanie.Rows(0).Cells(7).Value()), "dd.MM.yy"))
            txtIDZaplanovanie.Text = gridZaplanovanie.Rows(0).Cells(0).Value()
            txtDatumStavbyPovodnyZaplanovanie.Text = Format(CDate(gridZaplanovanie.Rows(0).Cells(4).Value()), "dd.MM.yy")
            gridZaplanovanie.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        End If

    End Sub

    Private Sub BtnZaplanovat_Click(sender As Object, e As EventArgs) Handles btnZaplanovat.Click



        If (MessageBox.Show("Ste si istý, že chcete zaplánovať jednotky podľa zadaných informácií?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question)) = vbNo Then
            Exit Sub
        Else
        End If
        If cn1.State = ConnectionState.Closed Then cn1.Open()
        Dim MyDateTimeVariable = Date.Now.ToString("dd MMM yyyy HH:mm:ss") 'format musi byt spravne zadefinovany, inak sa nezobrazi v MDB databaze

        If txtQTYZaplanovanie.Text = String.Empty Then
            MsgBox("Zadajte množstvo, ktoré chcete zaplánovať")
            Exit Sub
        End If

        If txtQTYZaplanovanie.Text > gridZaplanovanie.Rows(0).Cells(6).Value().ToString Then
            MsgBox("Pozor, množstvo, ktoré chcete zaplánovať je väčšie ako stav množstva na sklade. Upravte množstvo.")
            Exit Sub
        End If

        For i As Integer = 0 To gridZaplanovanie.Rows.Count - 1 Step +1
            If IsDBNull(gridZaplanovanie.Rows(i).Cells(0).ToString()) Then
                Continue For
            Else

                'Vlozenie zaplanovanych WJ Jednotiek do Hlavneho planu
                Dim cmdZaplanovanie As New OleDbCommand("INSERT INTO tblHlavnyPlan (DatumStavby,Skupina,JednotkaWJ,PipingWJ,PN,QTY,Drazkovanie,Urgent,RoutingCas, " _
                                                        & "RezanieStatus,OhybanieStatus,PripravaStatus,SpajkovanieStatus,ZaplanovanyWJ,ZaplanovaneDatum) " _
                                                  & "VALUES (@DatumStavby,@Skupina,@JednotkaWJ,@PipingWJ,@PN,@QTY,@Drazkovanie,@Urgent,@RoutingCas, " _
                                                  & "@RezanieStatus,@OhybanieStatus,@PripravaStatus,@SpajkovanieStatus,@ZaplanovanyWJ,@ZaplanovaneDatum)", cn1)

                cmdZaplanovanie.Parameters.AddWithValue("@DatumStavby", Format(dtpNovyDatumStavbyZaplanovanie.Value.Date, "dd/MM/yyyy"))
                cmdZaplanovanie.Parameters.AddWithValue("@Skupina", gridZaplanovanie.Rows(i).Cells(5).Value().ToString)
                If txtNovyWJJednotkyZaplanovanie.Text = String.Empty Then
                    cmdZaplanovanie.Parameters.AddWithValue("@JednotkaWJ", txtFilterWJJednotkyZaplanovanie.Text.ToString)
                Else
                    cmdZaplanovanie.Parameters.AddWithValue("@JednotkaWJ", txtNovyWJJednotkyZaplanovanie.Text.ToString)
                End If

                cmdZaplanovanie.Parameters.AddWithValue("@PipingWJ", gridZaplanovanie.Rows(i).Cells(2).Value().ToString)
                cmdZaplanovanie.Parameters.AddWithValue("@PN", gridZaplanovanie.Rows(i).Cells(3).Value().ToString)
                cmdZaplanovanie.Parameters.AddWithValue("@QTY", txtQTYZaplanovanie.Text)
                cmdZaplanovanie.Parameters.AddWithValue("@Drazkovanie", CBool(gridZaplanovanie.Rows(i).Cells(7).Value))
                cmdZaplanovanie.Parameters.AddWithValue("@Urgent", CBool("False"))
                cmdZaplanovanie.Parameters.AddWithValue("@RoutingCas", If(IsDBNull(gridZaplanovanie.Rows(i).Cells(9).Value), 0, IsDBNull(gridZaplanovanie.Rows(i).Cells(9).Value)))
                cmdZaplanovanie.Parameters.AddWithValue("@RezanieStatus", CBool("True"))
                cmdZaplanovanie.Parameters.AddWithValue("@OhybanieStatus", CBool("True"))
                cmdZaplanovanie.Parameters.AddWithValue("@PripravaStatus", CBool("True"))
                cmdZaplanovanie.Parameters.AddWithValue("@SpajkovanieStatus", CBool("True"))
                cmdZaplanovanie.Parameters.AddWithValue("@ZaplanovanyWJ", CBool("True"))
                cmdZaplanovanie.Parameters.AddWithValue("@ZaplanovaneDatum", MyDateTimeVariable)

                cmdZaplanovanie.ExecuteNonQuery()


                'update mnozstva jednotky WJ na odlozenom sklade
                Dim vysledok As Integer = gridZaplanovanie.Rows(0).Cells(6).Value - txtQTYZaplanovanie.Text
                Dim cmdUpdateStavu As New OleDbCommand("UPDATE tblVyplanovaneJednotky SET OdlozeneQTY = @NoveOdlozeneQTY,DatumZaplanovania=@DatumZaplanovania WHERE JednotkaWJ=@JednotkaWJ", cn1)
                cmdUpdateStavu.Parameters.AddWithValue("@NoveOdlozeneQTY", vysledok)
                cmdUpdateStavu.Parameters.AddWithValue("@DatumZaplanovania", MyDateTimeVariable)
                cmdUpdateStavu.Parameters.AddWithValue("@JednotkaWJ", txtFilterWJJednotkyZaplanovanie.Text.ToString)
                cmdUpdateStavu.ExecuteNonQuery()

            End If

        Next
        If cn1.State = ConnectionState.Open Then cn1.Close()

        populateGridZaplanovanie()

        MessageBox.Show("Informácia o zaplánovaní bola uložená", NazovAplikacie, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

    End Sub

    Private Sub GridZaplanovanie_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridZaplanovanie.CellClick

        Dim j As Integer = e.RowIndex
        'txtIDZaplanovanie.Text = gridZaplanovanie.Rows(j).Cells(0).Value()
        txtFilterWJJednotkyZaplanovanie.Text = gridZaplanovanie.Rows(j).Cells(1).Value()
        txtDatumStavbyPovodnyZaplanovanie.Text = Format(CDate(gridZaplanovanie.Rows(j).Cells(4).Value()), "dd.MM.yy")

        populateGridZaplanovanie()

    End Sub

    Private Sub GridKanban_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridKanban.CellClick


        Dim i As Integer

        With gridKanban
            If e.RowIndex >= 0 Then
                i = .CurrentRow.Index
                txtPriemerKanban.Text = .Rows(i).Cells("Priemer").Value.ToString
                txtHrubkaKanban.Text = .Rows(i).Cells("Hrubka").Value.ToString
                txtDlzkaKanban.Text = .Rows(i).Cells("Dlzka").Value.ToString
                txtIDKanban.Text = .Rows(i).Cells("ID").Value.ToString

            End If
        End With

    End Sub

    Private Sub BtnDelKanban_Click(sender As Object, e As EventArgs) Handles btnDelKanban.Click

        cn1.Open()

        Dim cmdKanban As New OleDbCommand("delete from tblKanban where ID=@ID", cn1)
        cmdKanban.Parameters.AddWithValue("ID", txtIDKanban.Text)
        cmdKanban.ExecuteNonQuery()

        cn1.Close()

        populateGridKanban()

        MessageBox.Show("Položka bola odstránená")

    End Sub

    Private Sub btnCompareWithDate_Click(sender As Object, e As EventArgs) Handles btnCompareWithDate.Click

        If cn1.State = ConnectionState.Closed Then cn1.Open()

        'zobrazenie obsahu, ktory sa nachadza v hlavnom plane, ale nenachadza sa v grupach
        Dim queryCompareHlavnyPlan As String = "SELECT DISTINCT tblHlavnyPlan.DatumStavby, tblHlavnyPlan.Skupina, tblHlavnyPlan.JednotkaWJ, tblHlavnyPlan.QTY " _
                                                & "FROM tblHlavnyPlan LEFT JOIN tblCompareGrupy ON tblHlavnyPlan.[JednotkaWJ] = tblCompareGrupy.[JednotkaWJ] " _
                                                & "WHERE (((tblHlavnyPlan.DatumStavby)=[@DatumStavby]) AND ((tblCompareGrupy.JednotkaWJ) Is Null));"

        Dim gridCmdCompareHlavnyPlan As OleDbCommand = New OleDbCommand(queryCompareHlavnyPlan, cn1)
        Dim sdaCompareHlavnyPlan As OleDbDataAdapter = New OleDbDataAdapter(gridCmdCompareHlavnyPlan)
        gridCmdCompareHlavnyPlan.Parameters.AddWithValue("@DatumStavby", Format(dtpPorovnanie.Value.Date, "dd/MM/yyyy"))
        Dim dtCompareHlavnyPlan As DataTable = New DataTable()
        sdaCompareHlavnyPlan.Fill(dtCompareHlavnyPlan)
        gridCompareHlavnyPlan.DataSource = dtCompareHlavnyPlan



        'zobrazenie obsahu, ktory sa nachadza v grupach, ale nenachadza sa v hlavnom plane
        Dim queryCompareGrupy As String = "SELECT DISTINCT tblCompareGrupy.DatumStavby, tblCompareGrupy.Skupina, tblCompareGrupy.JednotkaWJ, tblCompareGrupy.QTY " _
                                            & "FROM tblCompareGrupy LEFT JOIN tblHlavnyPlan ON tblCompareGrupy.[JednotkaWJ] = tblHlavnyPlan.[JednotkaWJ] " _
                                            & "WHERE (((tblCompareGrupy.DatumStavby)=[@DatumStavby]) AND ((tblHlavnyPlan.JednotkaWJ) Is Null));"

        Dim gridCmdCompareGrupy As OleDbCommand = New OleDbCommand(queryCompareGrupy, cn1)
        Dim sdaCompareGrupy As OleDbDataAdapter = New OleDbDataAdapter(gridCmdCompareGrupy)
        gridCmdCompareGrupy.Parameters.AddWithValue("@DatumStavby", Format(dtpPorovnanie.Value.Date, "dd/MM/yyyy"))
        Dim dtCompareGrupy As DataTable = New DataTable()
        sdaCompareGrupy.Fill(dtCompareGrupy)
        gridCompareGrupy.DataSource = dtCompareGrupy

        If cn1.State = ConnectionState.Open Then cn1.Close()

        With gridCompareHlavnyPlan
            .Columns(0).Width = 100
            .Columns(1).Width = 100
            .Columns(2).Width = 100
            .Columns(3).Width = 50


            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(0).HeaderCell.Value = "Dátum Stavby"
            .Columns(1).HeaderCell.Value = "Skupina"
            .Columns(2).HeaderCell.Value = "WJ Jednotky"
            .Columns(3).HeaderCell.Value = "Mn."

        End With

        With gridCompareGrupy
            .Columns(0).Width = 100
            .Columns(1).Width = 100
            .Columns(2).Width = 100
            .Columns(3).Width = 50


            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(0).HeaderCell.Value = "Dátum Stavby"
            .Columns(1).HeaderCell.Value = "Skupina"
            .Columns(2).HeaderCell.Value = "WJ Jednotky"
            .Columns(3).HeaderCell.Value = "Mn."

        End With

        gridCompareHlavnyPlan.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        gridCompareGrupy.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        txtPorovnanieHlavnyPlan.Text = gridCompareHlavnyPlan.RowCount
        txtPorovnanieGrupy.Text = gridCompareGrupy.RowCount



    End Sub



    Private Sub BtnNahrajGrupy_Click(sender As Object, e As EventArgs) Handles btnNahrajGrupy.Click

        If cn1.State = ConnectionState.Closed Then cn1.Open()
        '----------------------------------------------------------------------------------------------------
        'import planu z excelu vyberom suboru cez openfiledialog
        Dim OpenDLG As New OpenFileDialog
        OpenDLG.ShowDialog()
        Dim xPath As String = OpenDLG.FileName
        Dim wPathLong = System.IO.Path.GetDirectoryName(xPath)
        Dim stream As FileStream = File.Open(xPath, FileMode.Open, FileAccess.Read)
        Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
        Dim result As DataSet = excelReader.AsDataSet()
        excelReader.Close()
        'gridCompareGrupy.DataSource = result.Tables(0)
        Dim dtCompareGrupy As DataTable
        dtCompareGrupy = result.Tables(0)


        '----------------------------------------------------------------------------------------------------
        'vymaze zaznamy v temp tabulke tblCompareGrupy

        Dim cmdDel As New OleDbCommand("delete from tblCompareGrupy", cn1)
        cmdDel.ExecuteNonQuery()

        '----------------------------------------------------------------------------------------------------
        'Ulozi obsah datatable do temp tabulky tblCompareGrupy pre porovnavanie grup s hlavnym planom
        For i As Integer = 1 To dtCompareGrupy.Rows.Count - 1

            'MsgBox(gridCompareGrupy.Rows(i).Cells(11).toString())
            If IsDBNull(dtCompareGrupy.Rows(i)(0).ToString) Or IsDBNull(dtCompareGrupy.Rows(i)(11)) Then
                Continue For
            Else

                Dim cmdImport As New OleDbCommand("INSERT INTO tblCompareGrupy (DatumStavby,Skupina,JednotkaWJ,QTY) " _
                                                  & "VALUES (@DatumStavby,@Skupina,@JednotkaWJ,@QTY)", cn1)
                cmdImport.Parameters.AddWithValue("@DatumStavby", CalculateBusinessDaysFromInputDate(dtCompareGrupy.Rows(i)(3).ToString, 2))
                cmdImport.Parameters.AddWithValue("@Skupina", dtCompareGrupy.Rows(i)(1).ToString)
                cmdImport.Parameters.AddWithValue("@JednotkaWJ", dtCompareGrupy.Rows(i)(2).ToString)
                cmdImport.Parameters.AddWithValue("@QTY", dtCompareGrupy.Rows(i)(6).ToString)
                cmdImport.ExecuteNonQuery()


            End If
        Next
        MsgBox("Obsah naimportovany do tabulky")

    End Sub

    Private Sub BtnEditKanban_Click(sender As Object, e As EventArgs) Handles btnEditKanban.Click

        If MessageBox.Show("Naozaj chcete upraviť vybranú položku?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbNo Then Exit Sub


        'kontrola, ci priemer+hrubka+dlzka uz existuje
        cn1.Open()
        Dim queryKanbanCheck As String = "Select Priemer,Hrubka,Dlzka from tblKanban Where Priemer=@Priemer AND Hrubka=@Hrubka AND Dlzka=@Dlzka;"
        Dim cmdEditKanban As OleDbCommand = New OleDbCommand(queryKanbanCheck, cn1)
        cmdEditKanban.Parameters.AddWithValue("@Priemer", txtPriemerKanban.Text)
        cmdEditKanban.Parameters.AddWithValue("@Hrubka", txtHrubkaKanban.Text)
        cmdEditKanban.Parameters.AddWithValue("@Dlzka", txtDlzkaKanban.Text)
        Dim sdaKanbanCheck As OleDbDataAdapter = New OleDbDataAdapter(cmdEditKanban)
        Dim dtKanbanCheck As DataTable = New DataTable()
        cn1.Close()
        sdaKanbanCheck.Fill(dtKanbanCheck)

        If dtKanbanCheck.Rows.Count > 0 Then
            MessageBox.Show("Zadaný priemer/hrúbka/dĺžka už zozname existuje. Tentokráť ste nič neupravili.")
            Exit Sub
        End If

        'ulozenie polozky
        cn1.Open()
        Dim cmd As New OleDbCommand("UPDATE tblKanban SET Priemer=@Priemer, Hrubka=@Hrubka,Dlzka=@Dlzka,Kanban=@Kanban WHERE ID=@ID", cn1)

        cmd.Parameters.AddWithValue("@Priemer", txtPriemerKanban.Text)
        cmd.Parameters.AddWithValue("@Hrubka", txtHrubkaKanban.Text)
        cmd.Parameters.AddWithValue("@Dlzka", txtDlzkaKanban.Text)
        cmd.Parameters.AddWithValue("@Kanban", CBool("TRUE"))
        cmd.Parameters.AddWithValue("@ID", txtIDKanban.Text)


        cmd.ExecuteNonQuery()
        cn1.Close()

        txtPriemerKanban.Text = ""
        txtHrubkaKanban.Text = ""
        txtDlzkaKanban.Text = String.Empty
        txtIDKanban.Text = String.Empty

        populateGridKanban()

        MessageBox.Show("Položka bola upravená")

    End Sub

    Private Sub BtnSavePNTyp_Click(sender As Object, e As EventArgs) Handles btnSavePNTyp.Click

        'kontrola, ci priemer+hrubka+dlzka uz existuje
        cn1.Open()
        Dim queryPNTypCheck As String = "Select PN,Description,Typ from tblPNTyp Where PN=@PN AND Typ=@Typ and Description=@Description;"
        Dim gridcmdPNTypCheck As OleDbCommand = New OleDbCommand(queryPNTypCheck, cn1)
        gridcmdPNTypCheck.Parameters.AddWithValue("@PN", txtPNPNTyp.Text)
        gridcmdPNTypCheck.Parameters.AddWithValue("@Typ", cboTypPNTyp.Text)
        gridcmdPNTypCheck.Parameters.AddWithValue("@Description", txtDescriptionPNTyp.Text)
        Dim sdaPNTypCheck As OleDbDataAdapter = New OleDbDataAdapter(gridcmdPNTypCheck)
        Dim dtPNTypCheck As DataTable = New DataTable()
        cn1.Close()
        sdaPNTypCheck.Fill(dtPNTypCheck)

        If dtPNTypCheck.Rows.Count > 0 Then
            MessageBox.Show("Zadaný Piping PN / Typ už zozname existuje. Tentokráť ste nič neuložili.")
            Exit Sub
        End If

        'ulozenie polozky
        cn1.Open()
        Dim cmd As New OleDbCommand("INSERT INTO tblPNTyp (PN,Description,Typ) " _
                                    & " VALUES (@PN,@Description,@Typ)", cn1)

        cmd.Parameters.AddWithValue("@PN", txtPNPNTyp.Text)
        cmd.Parameters.AddWithValue("@Description", txtDescriptionPNTyp.Text)
        cmd.Parameters.AddWithValue("@Typ", cboTypPNTyp.Text)


        cmd.ExecuteNonQuery()
        cn1.Close()

        txtPNPNTyp.Text = ""
        txtDescriptionPNTyp.Text = ""
        cboTypPNTyp.Text = ""

        populateGridPNTyp()

        MessageBox.Show("Položka bola uložená")

    End Sub

    Private Sub BtnEditPNTyp_Click(sender As Object, e As EventArgs) Handles btnEditPNTyp.Click
        'kontrola, ci polozka uz existuje
        cn1.Open()
        Dim queryPNTypCheck As String = "Select PN,Typ from tblPNTyp Where PN=@PN AND Typ=@Typ AND Description=@Description;"
        Dim gridcmdPNTypCheck As OleDbCommand = New OleDbCommand(queryPNTypCheck, cn1)
        gridcmdPNTypCheck.Parameters.AddWithValue("@PN", txtPNPNTyp.Text)
        gridcmdPNTypCheck.Parameters.AddWithValue("@Typ", cboTypPNTyp.Text)
        gridcmdPNTypCheck.Parameters.AddWithValue("@Description", txtDescriptionPNTyp.Text)
        Dim sdaPNTypCheck As OleDbDataAdapter = New OleDbDataAdapter(gridcmdPNTypCheck)
        Dim dtPNTypCheck As DataTable = New DataTable()
        cn1.Close()
        sdaPNTypCheck.Fill(dtPNTypCheck)

        If dtPNTypCheck.Rows.Count > 0 Then
            MessageBox.Show("Zadaný Piping PN / Typ už zozname existuje. Tentokráť ste nič neupravili.")
            Exit Sub
        End If

        'ulozenie polozky
        cn1.Open()

        Dim cmd As New OleDbCommand("UPDATE tblPNTyp SET PN=@PN,Description=@Description,Typ=@Typ WHERE ID=@ID", cn1)
        Dim id As Integer = CType(txtIDPNTyp.Text, Int32)
        cmd.Parameters.AddWithValue("@PN", txtPNPNTyp.Text)
        cmd.Parameters.AddWithValue("@Description", txtDescriptionPNTyp.Text)
        cmd.Parameters.AddWithValue("@Typ", cboTypPNTyp.Text)
        cmd.Parameters.AddWithValue("@ID", id)

        Try
            cmd.ExecuteNonQuery()
            MessageBox.Show("Položka bola uložená")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        cn1.Close()

        txtPNPNTyp.Text = ""
        txtDescriptionPNTyp.Text = ""
        cboTypPNTyp.Text = ""

        populateGridPNTyp()


    End Sub

    Private Sub BtnDelPNTyp_Click(sender As Object, e As EventArgs) Handles btnDelPNTyp.Click

        'odstranenie polozky
        cn1.Open()
        Dim cmd As New OleDbCommand("DELETE * FROM tblPNTyp WHERE ID=@ID", cn1)

        cmd.Parameters.AddWithValue("@ID", txtIDPNTyp.Text)

        Dim affected As Integer = cmd.ExecuteNonQuery()
        cn1.Close()

        txtPNPNTyp.Text = ""
        txtDescriptionPNTyp.Text = ""
        cboTypPNTyp.Text = ""

        populateGridPNTyp()
        If affected > 0 Then
            MessageBox.Show("Položka bola vymazana")
        End If
    End Sub

    Private Sub GridPNTyp_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridPNTyp.CellClick

        Dim i As Integer

        With gridPNTyp
            If e.RowIndex >= 0 Then
                i = .CurrentRow.Index
                txtPNPNTyp.Text = .Rows(i).Cells("PN").Value.ToString
                cboTypPNTyp.Text = .Rows(i).Cells("Typ").Value.ToString
                txtDescriptionPNTyp.Text = .Rows(i).Cells("Description").Value.ToString
                txtIDPNTyp.Text = .Rows(i).Cells("ID").Value.ToString

            End If
        End With

    End Sub

    Private Sub btnShowReportSklad_Click(sender As Object, e As EventArgs) Handles btnShowReportSklad.Click



        'Dim frmReportSklad = New frmReportSklad()
        'Dim rds = New ReportDataSource("DataSet2", dtVychystanie)
        'frmReportSklad.rvSklad.LocalReport.DataSources.Clear()
        'frmReportSklad.rvSklad.LocalReport.DataSources.Add(rds)
        'frmReportSklad.rvSklad.LocalReport.ReportPath = "C:\Users\frantisek.horecny\Documents\Lean\Thermal Management\= Piping\PipingPlan Aplikacia\Piping Plan v01\Piping Plan v01\reportSklad.rdlc"
        'frmReportSklad.rvSklad.LocalReport.ReportEmbeddedResource = "Piping_Plan_v01.reportSklad.rdlc"
        'frmReportSklad.ShowDialog()


        frmReportSklad.rvSklad.LocalReport.DataSources.Clear()
        Dim rptDs As Microsoft.Reporting.WinForms.ReportDataSource
        frmReportSklad.rvSklad.LocalReport.DataSources.Clear()
        frmReportSklad.rvSklad.LocalReport.ReportEmbeddedResource = "Piping_Plan_v01.reportSklad.rdlc"
        'frmReportSklad.rvSklad.LocalReport.ReportPath = "C:\Users\frantisek.horecny\Documents\Lean\Thermal Management\= Piping\PipingPlan Aplikacia\Piping Plan v01\Piping Plan v01\reportSklad.rdlc"
        rptDs = New Microsoft.Reporting.WinForms.ReportDataSource("DataSet2", gridVychystanie.DataSource)

        frmReportSklad.rvSklad.LocalReport.DataSources.Add(rptDs)
        frmReportSklad.rvSklad.RefreshReport()

        frmReportSklad.Show()
    End Sub

    Private Sub TabPiping_TabIndexChanged(sender As Object, e As EventArgs) Handles tabPiping.TabIndexChanged

        pickPN = Nothing

    End Sub

    Private Sub BtnRefreshgridVychystanie_Click(sender As Object, e As EventArgs) Handles btnRefreshgridVychystanie.Click
        populategridVychystanie()
    End Sub

    Private Sub BtnCompareWithoutDate_Click(sender As Object, e As EventArgs) Handles btnCompareWithoutDate.Click

        If cn1.State = ConnectionState.Closed Then cn1.Open()

        'zobrazenie obsahu, ktory sa nachadza v hlavnom plane, ale nenachadza sa v grupach
        Dim queryCompareHlavnyPlan As String = "SELECT DISTINCT tblHlavnyPlan.DatumStavby, tblHlavnyPlan.Skupina, tblHlavnyPlan.JednotkaWJ, tblHlavnyPlan.QTY " _
                                                & "FROM tblHlavnyPlan LEFT JOIN tblCompareGrupy ON tblHlavnyPlan.[JednotkaWJ] = tblCompareGrupy.[JednotkaWJ] " _
                                                & "WHERE tblCompareGrupy.JednotkaWJ Is Null;"

        Dim gridCmdCompareHlavnyPlan As OleDbCommand = New OleDbCommand(queryCompareHlavnyPlan, cn1)
        Dim sdaCompareHlavnyPlan As OleDbDataAdapter = New OleDbDataAdapter(gridCmdCompareHlavnyPlan)
        'gridCmdCompareHlavnyPlan.Parameters.AddWithValue("@DatumStavby", Format(dtpPorovnanie.Value.Date, "dd/MM/yyyy"))
        Dim dtCompareHlavnyPlan As DataTable = New DataTable()
        sdaCompareHlavnyPlan.Fill(dtCompareHlavnyPlan)
        gridCompareHlavnyPlan.DataSource = dtCompareHlavnyPlan



        'zobrazenie obsahu, ktory sa nachadza v grupach, ale nenachadza sa v hlavnom plane
        Dim queryCompareGrupy As String = "SELECT DISTINCT tblCompareGrupy.DatumStavby, tblCompareGrupy.Skupina, tblCompareGrupy.JednotkaWJ, tblCompareGrupy.QTY " _
                                            & "FROM tblCompareGrupy LEFT JOIN tblHlavnyPlan ON tblCompareGrupy.[JednotkaWJ] = tblHlavnyPlan.[JednotkaWJ] " _
                                            & "WHERE tblHlavnyPlan.JednotkaWJ Is Null;"

        Dim gridCmdCompareGrupy As OleDbCommand = New OleDbCommand(queryCompareGrupy, cn1)
        Dim sdaCompareGrupy As OleDbDataAdapter = New OleDbDataAdapter(gridCmdCompareGrupy)
        'gridCmdCompareGrupy.Parameters.AddWithValue("@DatumStavby", Format(dtpPorovnanie.Value.Date, "dd/MM/yyyy"))
        Dim dtCompareGrupy As DataTable = New DataTable()
        sdaCompareGrupy.Fill(dtCompareGrupy)
        gridCompareGrupy.DataSource = dtCompareGrupy

        If cn1.State = ConnectionState.Open Then cn1.Close()

        With gridCompareHlavnyPlan
            .Columns(0).Width = 100
            .Columns(1).Width = 100
            .Columns(2).Width = 100
            .Columns(3).Width = 50


            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(0).HeaderCell.Value = "Dátum Stavby"
            .Columns(1).HeaderCell.Value = "Skupina"
            .Columns(2).HeaderCell.Value = "WJ Jednotky"
            .Columns(3).HeaderCell.Value = "Mn."

        End With

        With gridCompareGrupy
            .Columns(0).Width = 100
            .Columns(1).Width = 100
            .Columns(2).Width = 100
            .Columns(3).Width = 50


            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(0).HeaderCell.Value = "Dátum Stavby"
            .Columns(1).HeaderCell.Value = "Skupina"
            .Columns(2).HeaderCell.Value = "WJ Jednotky"
            .Columns(3).HeaderCell.Value = "Mn."

        End With

        gridCompareHlavnyPlan.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        gridCompareGrupy.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        txtPorovnanieHlavnyPlan.Text = gridCompareHlavnyPlan.RowCount
        txtPorovnanieGrupy.Text = gridCompareGrupy.RowCount

    End Sub

    Private Sub PipingPlanForm_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        LoginForm1.Visible = True
    End Sub

    Private Sub BtnCompareVyplanovane_Click(sender As Object, e As EventArgs) Handles btnCompareVyplanovane.Click

        If cn1.State = ConnectionState.Closed Then cn1.Open()

        'zobrazenie obsahu, ktory sa nachadza v hlavnom plane, ale nenachadza sa v grupach
        Dim queryCompareVyplanovane As String = "SELECT DISTINCT tblVyplanovaneJednotky.Skupina, tblVyplanovaneJednotky.JednotkaWJ, tblVyplanovaneJednotky.OdlozeneQTY " _
            & "FROM tblVyplanovaneJednotky RIGHT JOIN tblCompareGrupy ON tblVyplanovaneJednotky.[JednotkaWJ] = tblCompareGrupy.[JednotkaWJ] " _
            & "WHERE (((tblVyplanovaneJednotky.JednotkaWJ) Is Not Null) AND ((tblVyplanovaneJednotky.OdlozeneQTY)>0));"

        Dim CmdCompareVyplanovane As OleDbCommand = New OleDbCommand(queryCompareVyplanovane, cn1)
        Dim sdaCompareVyplanovane As OleDbDataAdapter = New OleDbDataAdapter(CmdCompareVyplanovane)
        'gridCmdCompareHlavnyPlan.Parameters.AddWithValue("@DatumStavby", Format(dtpPorovnanie.Value.Date, "dd/MM/yyyy"))
        Dim dtCompareVyplanovane As DataTable = New DataTable()
        sdaCompareVyplanovane.Fill(dtCompareVyplanovane)
        gridCompareGrupyANDVyplanovane.DataSource = dtCompareVyplanovane

        If cn1.State = ConnectionState.Open Then cn1.Close()

        With gridCompareGrupyANDVyplanovane
            .Columns(0).Width = 100
            .Columns(1).Width = 100
            .Columns(2).Width = 100


            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False

            .Columns(0).HeaderCell.Value = "Skupina"
            .Columns(1).HeaderCell.Value = "WJ Jednotky"
            .Columns(2).HeaderCell.Value = "Mn. na sklade"

        End With


        gridCompareGrupyANDVyplanovane.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft Sans Serif", 12, FontStyle.Regular)
        txtPorovnanieVyplanovane.Text = gridCompareGrupyANDVyplanovane.RowCount

    End Sub

    Private Sub GridOhybanie_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridOhybanie.CellDoubleClick

        'MsgBox(e.ColumnIndex)
        'MsgBox(e.RowIndex)
        pickPN = gridOhybanie.Rows(e.RowIndex).Cells("PN").Value.ToString
        'DTTimeStarted = Date.Now.ToString("dd MM yyyy HH:mm:ss")
        AddPipeRecordsForm.Show()

    End Sub

    Private Sub GridInziniering_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridInziniering.CellDoubleClick

        'MsgBox(e.ColumnIndex)
        'MsgBox(e.RowIndex)
        pickPN = gridInziniering.Rows(e.RowIndex).Cells("PN").Value.ToString
        'DTTimeStarted = Date.Now.ToString("dd MM yyyy HH:mm:ss")
        AddPipeRecordsForm.Show()

    End Sub

    Private Sub Label107_DoubleClick(sender As Object, e As EventArgs) Handles lblVersion.DoubleClick
        UpdatesForm.Show()
    End Sub

    Private Sub Chart1_Click(sender As Object, e As EventArgs) Handles Chart1.Click
        If Chart1.Series.Count = 0 Then
            Chart1.Series.Add("Dokoncene")
            Chart1.Series.Add("Nedokoncene")
            Chart1.Series.Add("Zameskane")
        End If

        Dim dtRezanie As DataTable = populateGridRezanie()
        Dim dtOhybanie As DataTable = populateGridOhybanie()
        Dim dtPriprava As DataTable = populategridPriprava()
        Dim dtZvaranie As DataTable = populateGridSpajkovanie()
        Dim dtVychystanie As DataTable = populategridVychystanie()
        Dim datum As Date = DateTimePicker1.Value.Date
        Dim dokoncene As Integer
        Dim count As Integer

        Dim completed As Double() = {0, 0, 0, 0, 0}
        Dim uncompleted As Double() = {0, 0, 0, 0, 0}

        For index As Integer = 0 To dtRezanie.Rows.Count - 1
            If dtRezanie.Rows(index).Item(0).Equals(datum) Then
                If dtRezanie.Rows(index).Item(6) = True Then
                    dokoncene = dtRezanie.Rows(index).Item(5)
                End If
                count += dtRezanie.Rows(index).Item(5)
            End If
        Next

        completed(0) = dokoncene
        uncompleted(0) = count - dokoncene


        count = 0
        dokoncene = 0
        For index As Integer = 0 To dtOhybanie.Rows.Count - 1
            If dtOhybanie.Rows(index).Item(1).Equals(datum) Then
                If dtOhybanie.Rows(index).Item(11) = True Then
                    dokoncene += dtOhybanie.Rows(index).Item(5)
                End If
                count += dtOhybanie.Rows(index).Item(5)
            End If
        Next



        completed(1) = dokoncene
        uncompleted(1) = count - dokoncene


        count = 0
        dokoncene = 0
        For index As Integer = 0 To dtPriprava.Rows.Count - 1
            If dtPriprava.Rows(index).Item(0).Equals(datum) Then
                If dtPriprava.Rows(index).Item(6) = True Then
                    dokoncene += dtPriprava.Rows(index).Item(4)
                End If
                count += dtPriprava.Rows(index).Item(4)
            End If
        Next



        completed(2) = dokoncene
        uncompleted(2) = count - dokoncene


        count = 0
        dokoncene = 0
        For index As Integer = 0 To dtZvaranie.Rows.Count - 1
            If dtPriprava.Rows(index).Item(0).Equals(datum) Then
                If dtZvaranie.Rows(index).Item(6) = True Then
                    dokoncene += dtZvaranie.Rows(index).Item(4)
                End If
                count += dtZvaranie.Rows(index).Item(4)
            End If
        Next



        completed(3) = dokoncene
        uncompleted(3) = count - dokoncene

        'For index As Integer = 0 To completed.Length - 1

        '    Dim coef As Double = 100 / uncompleted(index)
        '    Dim nedokoncene As Double = uncompleted(index) - completed(index)
        '    uncompleted(index) = nedokoncene * coef
        '    completed(index) = completed(index) * coef
        'Next

        Dim operations As String() = {"Rezanie", "Ohybanie", "Priprava", "Spajkovanie", "Vychystanie"}

        Chart1.Series("Dokoncene").IsValueShownAsLabel = True
        Chart1.Series("Dokoncene").ChartType = SeriesChartType.StackedColumn100
        Chart1.Series("Dokoncene").Points.DataBindXY(operations, completed)

        Chart1.Series("Nedokoncene").IsValueShownAsLabel = True
        Chart1.Series("Nedokoncene").ChartType = SeriesChartType.StackedColumn100
        Chart1.Series("Nedokoncene").Points.DataBindXY(operations, uncompleted)

        If Chart2.Series.Count = 0 Then
            Chart2.Series.Add("Dokoncene")
            Chart2.Series.Add("Nedokoncene")
            Chart2.Series.Add("Zameskane")
        End If

        Chart2.Series("Dokoncene").IsValueShownAsLabel = True
        Chart2.Series("Dokoncene").ChartType = SeriesChartType.RangeColumn
        Chart2.Series("Dokoncene").Points.DataBindXY(operations, completed)

        Chart2.Series("Nedokoncene").IsValueShownAsLabel = True
        Chart2.Series("Nedokoncene").ChartType = SeriesChartType.RangeColumn
        Chart2.Series("Nedokoncene").Points.DataBindXY(operations, uncompleted)
    End Sub

    Private Sub Chart2_Click(sender As Object, e As EventArgs) Handles Chart2.Click

    End Sub


End Class






