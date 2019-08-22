Imports System.Data.OleDb
Public Class AddPipeRecordsForm

    Public SelectedPipeID As Integer

    Private Sub AddPipeRecordsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        txtPipePN.Text = pickPN
        populateGridAddPipeRecord()

        'txtDTComponent.Select()

    End Sub

    Private Sub populateGridAddPipeRecord()


        Dim MyDateIsToday = Date.Now.ToString("dd MMM yyyy")
        cn1.Open()
        Dim queryAddPipeRecord As String = "SELECT tblKompletnyZoznam.ID, tblKompletnyZoznam.PN, tblKompletnyZoznam.subPN, tblKompletnyZoznam.Priemer, " _
                & "tblKompletnyZoznam.Hrubka, tblKompletnyZoznam.Dlzka, tblKompletnyZoznam.OhybRovna,tblKanban.Kanban,tblKompletnyZoznam.Skontrolovane  " _
                & "FROM tblKompletnyZoznam LEFT JOIN tblKanban ON (tblKompletnyZoznam.Dlzka = tblKanban.Dlzka) AND (tblKompletnyZoznam.Hrubka = tblKanban.Hrubka) " _
                & "AND (tblKompletnyZoznam.Priemer = tblKanban.Priemer) " _
                & "WHERE (((tblKompletnyZoznam.[PN])=[@PN])) " _
                & "ORDER BY tblKompletnyZoznam.subPN, tblKompletnyZoznam.Priemer, tblKompletnyZoznam.Hrubka, tblKompletnyZoznam.Dlzka;"

        ''pridanie linku pre vykres
        'Dim queryAddPipeRecord As String = "SELECT tblKompletnyZoznam.ID, tblKompletnyZoznam.PN, tblKompletnyZoznam.subPN, tblKompletnyZoznam.Priemer, " _
        '                & "tblKompletnyZoznam.Hrubka, tblKompletnyZoznam.Dlzka, tblKompletnyZoznam.OhybRovna, tblKanban.Kanban, tblKompletnyZoznam.Skontrolovane, tblDocuments.Name as Vykres " _
        '                & "FROM tblKompletnyZoznam LEFT JOIN tblKanban ON tblKompletnyZoznam.Priemer = tblKanban.Priemer AND tblKompletnyZoznam.Hrubka = tblKanban.Hrubka " _
        '                & "AND tblKompletnyZoznam.Dlzka = tblKanban.Dlzka LEFT JOIN tblDocuments ON tblKompletnyZoznam.PN = tblDocuments.Title " _
        '                & "WHERE tblKompletnyZoznam.PN=@PN " _
        '                & "ORDER BY tblKompletnyZoznam.subPN, tblKompletnyZoznam.Priemer, tblKompletnyZoznam.Hrubka, tblKompletnyZoznam.Dlzka;"

        Dim gridcmdInziniering As OleDbCommand = New OleDbCommand(queryAddPipeRecord, cn1)
        gridcmdInziniering.Parameters.AddWithValue("@PN", pickPN)
        Dim sdaInziniering As OleDbDataAdapter = New OleDbDataAdapter(gridcmdInziniering)
        Dim dtInziniering As DataTable = New DataTable()
        sdaInziniering.Fill(dtInziniering)
        gridAddPipeRecord.DataSource = dtInziniering
        If cn1.State = ConnectionState.Open Then cn1.Close()
        'MsgBox(dtInziniering.Rows.Count)
        'Dim query As String = "Select DatumStavby,Skupina,PipingWJ,PN,QTY,subPN,Pocetnost, Priemer, Hrubka, Dlzka, OhybRovna, Kanban, AMOB,InzinieringStatus,OhybanieStatus,PripravaStatus,SpajkovanieStatus FROM tblHlavnyPlan where InzinieringStatus=NO And Kanban Is Null;"

        gridAddPipeRecord.Columns(0).DefaultCellStyle.Format = "dd/MMM"
        'nastavenie vysky riadkov v DT gride v hlavnom okne
        gridAddPipeRecord.RowTemplate.Height = 25

        ''nastavenie sirky stlpca a zarovnanie pre datagridview historia
        With gridAddPipeRecord
            .Columns(0).Visible = False
            '.Columns(9).Visible = False
            .Columns(1).Width = 90
            .Columns(2).Width = 90
            .Columns(3).Width = 90
            .Columns(4).Width = 90
            .Columns(5).Width = 90
            .Columns(6).Width = 90
            .Columns(7).Width = 90
            .Columns(8).Width = 90
            '.Columns(9).Width = 90

            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            '.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .RowHeadersVisible = False
            'ID,PN,QTY,subPN,Pocetnost, Priemer, Hrubka, Dlzka, OhybRovna, Kanban, AMOB
            .Columns(1).HeaderCell.Value = "Piping"
            .Columns(2).HeaderCell.Value = "subPiping"
            .Columns(3).HeaderCell.Value = "Priemer"
            .Columns(4).HeaderCell.Value = "Hrúbka"
            .Columns(5).HeaderCell.Value = "Dĺžka"
            .Columns(6).HeaderCell.Value = "Ohyb / Rovná"
            .Columns(7).HeaderCell.Value = "Kanban"
            .Columns(8).HeaderCell.Value = "Skontrolované"

        End With

        'nastavenie sortovania
        gridAddPipeRecord.Sort(gridAddPipeRecord.Columns(1), System.ComponentModel.ListSortDirection.Descending)
        gridAddPipeRecord.Sort(gridAddPipeRecord.Columns(2), System.ComponentModel.ListSortDirection.Descending)

        If cn1.State = ConnectionState.Open Then cn1.Close()
    End Sub

    Private Sub btnSavePipe_Click(sender As Object, e As EventArgs) Handles btnSavePipe.Click
        If MessageBox.Show("Naozaj chcete uložiť novú položku?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbNo Then Exit Sub
        'Dim DTRecordedValue As Date = Date.Now.ToString("dd MMM yyyy HH:mm:ss")

        'kontrola, ci komponent bol vyplneny
        'If txtDTComponent.Text = "" Then
        '    ' ... tak vyhodime varovnu hlasku
        '    If MessageBox.Show("Komponent/komentar nie je vyplnený! Chcete pokračovať?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbNo Then Exit Sub

        'End If

        'kod pre ukladanie hodnot z formularu do mdb tabulky

        cn1.Open()
        Dim cmd As New OleDbCommand("INSERT INTO tblKompletnyZoznam (PN,subPN,Priemer,Hrubka,Dlzka,OhybRovna,Skontrolovane) " _
                                    & " VALUES (@txtPipePN,@txtPipeSubPN,@cboPipePriemer,@cboPipeHrubka,@txtPipeDlzka,@txtPipeOhybRovna,@ckbPipeSkontrolovane)", cn1)

        cmd.Parameters.AddWithValue("@txtPipePN", txtPipePN.Text)
        cmd.Parameters.AddWithValue("@txtPipeSubPN", txtPipeSubPN.Text)
        cmd.Parameters.AddWithValue("@cboPipePriemer", CDbl(cboPipePriemer.Text))
        cmd.Parameters.AddWithValue("@cboPipeHrubka", CDbl(cboPipeHrubka.Text))
        cmd.Parameters.AddWithValue("@txtPipeDlzka", If(txtPipeDlzka.Text = String.Empty, Nothing, CDbl(txtPipeDlzka.Text)))
        cmd.Parameters.AddWithValue("@cboPipeOhybRovna", cboPipeOhybRovna.Text)
        cmd.Parameters.AddWithValue("@ckbPipeSkontrolovane", CBool("TRUE"))

        cmd.ExecuteNonQuery()
        cn1.Close()
        'pridat insert do tblHlavnyPlan s 
        txtPipeSubPN.Text = ""
        cboPipePriemer.SelectedIndex = -1
        cboPipePriemer.Text = String.Empty
        cboPipeHrubka.SelectedIndex = -1
        cboPipeHrubka.Text = String.Empty
        txtPipeDlzka.Text = ""
        cboPipeOhybRovna.SelectedIndex = -1
        cboPipeOhybRovna.Text = String.Empty
        ckbPipeSkontrolovane.CheckState = 0

        populateGridAddPipeRecord()
        MsgBox("Položka bola uložená")
    End Sub

    Private Sub BtnDelPipe_Click(sender As Object, e As EventArgs) Handles btnDelPipe.Click

        If MessageBox.Show("Naozaj chcete vymazat vybraznú položku?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbNo Then Exit Sub


        cn1.Open()
        'vymazanie oznaceneho riadku

        Dim cmd2 As New OleDbCommand("DELETE * From tblKompletnyZoznam where ID=@ID", cn1)
        cmd2.Parameters.AddWithValue("@ID", SelectedPipeID.ToString)
        cmd2.ExecuteNonQuery()

        cn1.Close()
        populateGridAddPipeRecord()

        txtPipeSubPN.Text = ""
        cboPipePriemer.SelectedIndex = -1
        cboPipePriemer.Text = String.Empty
        cboPipeHrubka.SelectedIndex = -1
        cboPipeHrubka.Text = String.Empty
        txtPipeDlzka.Text = ""
        cboPipeOhybRovna.SelectedIndex = -1
        cboPipeOhybRovna.Text = String.Empty
        ckbPipeSkontrolovane.CheckState = 0

        MsgBox("Položka bola vymazaná")
    End Sub

    Private Sub GridAddPipeRecord_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridAddPipeRecord.CellClick

        SelectedPipeID = gridAddPipeRecord.Rows(e.RowIndex).Cells(0).Value.ToString
        txtPipeID.Text = SelectedPipeID
        txtPipeSubPN.Text = gridAddPipeRecord.Rows(e.RowIndex).Cells(2).Value.ToString
        'cboPipePriemer.SelectedIndex = gridAddPipeRecord.Rows(e.RowIndex).Cells(3).Value
        cboPipePriemer.Text = gridAddPipeRecord.Rows(e.RowIndex).Cells(3).Value.ToString
        'cboPipeHrubka.SelectedIndex = gridAddPipeRecord.Rows(e.RowIndex).Cells(5).Value
        cboPipeHrubka.Text = gridAddPipeRecord.Rows(e.RowIndex).Cells(4).Value.ToString
        txtPipeDlzka.Text = gridAddPipeRecord.Rows(e.RowIndex).Cells(5).Value.ToString
        'cboPipeOhybRovna.SelectedIndex = gridAddPipeRecord.Rows(e.RowIndex).Cells(7).Value
        cboPipeOhybRovna.Text = gridAddPipeRecord.Rows(e.RowIndex).Cells(6).Value.ToString

        'MsgBox(gridAddPipeRecord.Rows(e.RowIndex).Cells(8).Value)

        If gridAddPipeRecord.Rows(e.RowIndex).Cells(8).Value = True Then
            ckbPipeSkontrolovane.CheckState = CheckState.Checked
        Else
            ckbPipeSkontrolovane.CheckState = CheckState.Unchecked
        End If

        'If e.RowIndex > -1 Then
        '    str1 = gridAddPipeRecord.Rows(e.RowIndex).Cells("Vykres").Value.ToString
        'End If

    End Sub

    Private Sub BtnEditPipe_Click(sender As Object, e As EventArgs) Handles btnEditPipe.Click
        If MessageBox.Show("Naozaj chcete upraviť vybranú položku?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbNo Then Exit Sub
        cn1.Open()
        Dim cmd As New OleDbCommand("UPDATE tblKompletnyZoznam SET PN=@txtPipePN,subPN=txtPipeSubPN,Priemer=@cboPipePriemer,Hrubka=@cboPipeHrubka,Dlzka=@txtPipeDlzka, " _
                                    & "OhybRovna=@txtPipeOhybRovna,Skontrolovane=@ckbPipeSkontrolovane " _
                                    & "WHERE ID=@ID", cn1)

        cmd.Parameters.AddWithValue("@txtPipePN", txtPipePN.Text)
        cmd.Parameters.AddWithValue("@txtPipeSubPN", txtPipeSubPN.Text)
        cmd.Parameters.AddWithValue("@cboPipePriemer", CDbl(cboPipePriemer.Text))
        cmd.Parameters.AddWithValue("@cboPipeHrubka", CDbl(cboPipeHrubka.Text))
        cmd.Parameters.AddWithValue("@txtPipeDlzka", DbNullOrStringValue(txtPipeDlzka.Text))
        cmd.Parameters.AddWithValue("@cboPipeOhybRovna", cboPipeOhybRovna.Text)

        cmd.Parameters.AddWithValue("@ckbPipeSkontrolovane", CBool("TRUE"))
        cmd.Parameters.AddWithValue("@ID", txtPipeID.Text)

        cmd.ExecuteNonQuery()
        cn1.Close()

        txtPipeSubPN.Text = ""
        cboPipePriemer.SelectedIndex = -1
        cboPipePriemer.Text = String.Empty
        cboPipeHrubka.SelectedIndex = -1
        cboPipeHrubka.Text = String.Empty
        txtPipeDlzka.Text = ""
        cboPipeOhybRovna.SelectedIndex = -1
        cboPipeOhybRovna.Text = String.Empty
        ckbPipeSkontrolovane.CheckState = 0

        populateGridAddPipeRecord()
        MsgBox("Položka bola upravená")
    End Sub

    Private Sub AddPipeRecordsForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        PipingPlanForm.populateGridInziniering()
        PipingPlanForm.populateGridOhybanie()

    End Sub


End Class