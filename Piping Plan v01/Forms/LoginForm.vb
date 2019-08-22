Imports System.Data.OleDb
Public Class LoginForm1

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See https://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.
    Public Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click


        Dim profile As String = cmbProfil.Text
        Dim login As String = UsernameTextBox.Text
        Dim password As String = PasswordTextBox.Text
        Dim query As String = "SELECT * from tblProfile;"
        Dim gridcmd As OleDbCommand = New OleDbCommand(query, cn1)
        Dim sda As OleDbDataAdapter = New OleDbDataAdapter(gridcmd)
        Dim dt As DataTable = New DataTable()
        PipingPlanForm.ToolStripStatusLabel1.Text = "Profil: " & profile
        PipingPlanForm.ToolStripStatusLabel2.Text = "Login: " & login



        Try
            cn1.Open()
            sda.Fill(dt)
        Catch ex As Exception
            If MessageBox.Show("Pocitac sa nemoze pripojit k databaze chcete to skusit znova?", NazovAplikacie, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then OK_Click(sender, e)
        End Try
        cn1.Close()
        'cyklus preveri ci sa v databaze nachadza meno s danym heslom a priradi k nemu okienka ktore sa maju zobrazit
        'ak sa v databaze nenachadza meno ani heslo tak vyhodi hlasku
        For index As Integer = 0 To dt.Rows.Count - 1
            If dt.Rows(index).Item(1).Equals(profile) And dt.Rows(index).Item(2).Equals(password) Then

                PipingPlanForm.Form1_Load()
                Select Case profile
                    Case "operator"

                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabPlanovanie)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabVyplanovanie)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabZaplanovanie)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabInziniering)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabEdit)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabStatus)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabPrehlady)
                        PipingPlanForm.Show()
                        Exit Sub
                    Case "majster"

                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabPlanovanie)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabVyplanovanie)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabZaplanovanie)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabInziniering)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabEdit)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabStatus)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabPrehlady)
                        PipingPlanForm.Show()
                        Exit Sub
                    Case "admin"
                        PipingPlanForm.Show()
                        Exit Sub
                    Case "inzinier"
                        PipingPlanForm.Show()
                        Exit Sub
                    Case "planovac"
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.TabPriprava)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabRezanie)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabOhybanie)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabSpajkovanie)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabVychystanie)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabInziniering)
                        PipingPlanForm.tabPiping.TabPages.Remove(PipingPlanForm.tabEdit)
                        PipingPlanForm.Show()
                        Exit Sub
                End Select

                'PipingPlanForm.StatusStrip1.Text = "Login " & login


            End If
        Next

        MessageBox.Show("Meno alebo heslo sa nenachadza v databaze!", NazovAplikacie)
        PipingPlanForm.Close()

    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub PasswordTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles PasswordTextBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            OK_Click(sender, e)

        End If
    End Sub
End Class