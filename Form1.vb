Option Strict Off
Option Explicit On

Imports System.IO
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports iTextSharp.text.pdf
Imports iTextSharp.text

Public Class Form1

    Dim preference As String = "Jouer un son"
    Dim nom_audio As String = "Sons/wakeup.wav"

    Dim nomBD As String = "Database1.accdb"

    Dim ObjetConnection As OleDbConnection
    Dim ObjetCommand As OleDbCommand
    Dim ObjetDataAdapter As OleDbDataAdapter
    Dim ObjetDataSet As New DataSet
    Dim strSql As String
    Dim ObjetDataTable As DataTable
    Dim strConn As String
    Dim ObjetCommandBuilder As OleDbCommandBuilder
    Dim ObjetDataRow As DataRow

    Dim nom As String
    Dim horaire As Date
    Dim message As String
    Dim dateenreg As Date
    Dim niveau As Integer = 1

    Dim essaiSQL As String
    Dim commande As OleDbCommand
    Dim i, j As Integer

    Dim Connexion As String
    Dim ConnectionOLE As OleDbConnection = New OleDbConnection()
    Dim da As OleDbDataAdapter
    Dim ds As DataSet
    Dim dv As DataView
    Dim cb As OleDbCommandBuilder

    Public Sub AfficheTous()
        Connexion = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & nomBD & ";"
        ConnectionOLE.ConnectionString = Connexion 'Passage de la chaine à l'objet de type connection
        ConnectionOLE.Open() 'Demarrage de la connexion
        da = New OleDbDataAdapter("Select ID, NOM, HORAIRE, ACTIVATION, NIVEAU from TACHE", ConnectionOLE) 'Passage des elements de la table User dans un adaptateur
        ds = New DataSet() 'Instanciation du dataset
        da.Fill(ds, "TACHE") 'Remplissage du dataset par les elements retenus par l'adaptateur
        dv = ds.Tables("TACHE").DefaultView 'Copie du dataset dans un dataview pour qu'il puisse etre edité par l'application
        ConnectionOLE.Close() 'Arrêt de la connexion
        DataGridView1.DataSource = dv 'Le datagrid prend les elements du Dataview
        dv.AllowEdit = False 'Le dataview est rendu editable
    End Sub

    Public Sub SupprimeTous()
        Connexion = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & nomBD & ";"
        ConnectionOLE.ConnectionString = Connexion 'Passage de la chaine à l'objet de type connection
        ConnectionOLE.Open() 'Demarrage de la connexion
        da = New OleDbDataAdapter("DELETE * FROM TACHE", ConnectionOLE) 'Passage des elements de la table User dans un adaptateur
        ObjetDataSet = New DataSet() 'Instanciation du dataset
        da.Fill(ObjetDataSet, "TACHE") 'Remplissage du dataset par les elements retenus par l'adaptateur
        'dv = ds.Tables("TACHE").DefaultView 'Copie du dataset dans un dataview pour qu'il puisse etre edité par l'application
        ConnectionOLE.Close() 'Arrêt de la connexion
        'DataGridView1.DataSource = dv 'Le datagrid prend les elements du Dataview
        'dv.AllowEdit = False 'Le dataview est rendu editable
    End Sub

    Public Sub executer_routine(ByVal listB As ListBox)
        strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & nomBD & ";"
        strSql = "SELECT TACHE.* FROM TACHE"
        ObjetConnection = New OleDbConnection
        ObjetConnection.ConnectionString = strConn
        Try
            ObjetConnection.Open()
        Catch ex As OleDbException
            MsgBox(ex.Message)
            End
        End Try
        ObjetCommand = New OleDbCommand(strSql)
        ObjetDataAdapter = New OleDbDataAdapter(ObjetCommand)
        ObjetCommand.Connection() = ObjetConnection
        ObjetDataAdapter.Fill(ObjetDataSet, "TACHE")
        ObjetDataTable = ObjetDataSet.Tables("TACHE")
        listB.DataSource = ObjetDataSet.Tables("TACHE")
        listB.DisplayMember = "NOM"
        ObjetConnection.Close()
        i = 0
        j = 0
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        horaire = New Date(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month, DateTimePicker1.Value.Day, DateTimePicker2.Value.Hour, DateTimePicker2.Value.Minute, DateTimePicker2.Value.Second)
        If TextBox1.Text = Nothing Then
            MsgBox("Vous devez entrer une valeur dans le champ reservé au nom!", MsgBoxStyle.Information)
        ElseIf TextBox2.Text = Nothing Then
            MsgBox("Vous devez entrer une valeur dans le champ reservé au message!", MsgBoxStyle.Information)
        ElseIf horaire.CompareTo(Date.Now) < 0 Then
            MsgBox("Vous devez entrer une date qui est ultérieure à la date actuelle!", MsgBoxStyle.Information)
        Else

            nom = TextBox1.Text
            horaire = New Date(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month, DateTimePicker1.Value.Day, DateTimePicker2.Value.Hour, DateTimePicker2.Value.Minute, DateTimePicker2.Value.Second)
            message = TextBox2.Text
            dateenreg = Date.Now
            If RadioButton1.Checked Then
                niveau = 0
            ElseIf RadioButton2.Checked Then
                niveau = 1
            Else
                niveau = 2
            End If
            nom = nom.Replace("'", "`")
            message = message.Replace("'", "`")

            strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & nomBD & ";"
            strSql = "INSERT INTO TACHE(NOM, HORAIRE, MESSAGE, DATEENREG, NIVEAU, ACTIVATION) VALUES('" & nom & "', '" & CStr(horaire) & "', '" & message & "', '" & CStr(dateenreg) & "', '" & niveau & "', '" & "1" & "')"
            ObjetConnection = New OleDbConnection
            ObjetConnection.ConnectionString = strConn

            Try
                ObjetConnection.Open()
            Catch ex As OleDbException
                MsgBox(ex)
                End
            End Try

            ObjetCommand = New OleDbCommand(strSql)
            ObjetDataAdapter = New OleDbDataAdapter(ObjetCommand)
            ObjetCommand.Connection() = ObjetConnection
            ObjetCommandBuilder = New OleDbCommandBuilder(ObjetDataAdapter)
            ObjetDataAdapter.Update(ObjetDataSet, "TACHE")
            ObjetDataSet.Clear()

            ObjetDataAdapter.Fill(ObjetDataSet, "TACHE")
            ObjetDataTable = ObjetDataSet.Tables("TACHE")
            ListBox1.DataSource = ObjetDataSet.Tables("TACHE")
            ListBox1.DisplayMember = "NOM"
            ObjetConnection.Close()
            executer_routine(ListBox1)
            TextBox1.Text = ""
            TextBox2.Text = ""
            RadioButton2.Checked = True
            TextBox1.Focus()
            Label8.Text = "Nombre de tâches enregistrées : " & CStr(ObjetDataSet.Tables("TACHE").Rows.Count)
            hpp = HeureLaPlusProche()

        End If
    End Sub

    Dim hpp As Date 'Horaire le plus proche
    Dim idCouran As Integer
    
    Function HeureLaPlusProche() As Date

        Dim actual As Date = Date.Now
        Dim hour As Date = New Date(2090, 12, 1, actual.Hour, actual.Minute, actual.Second)
        'Dim hour As Date = New Date(actual.Year, actual.Month, actual.Day, actual.Hour, actual.Minute, actual.Second)
        Dim ii As Integer = 0
        'ObjetDataRow = ObjetDataSet.Tables("TACHE").Rows(ii)
        'hour = CDate(CStr(ObjetDataRow("HORAIRE")))

        For ii = 0 To ObjetDataSet.Tables("TACHE").Rows.Count - 1
            ObjetDataRow = ObjetDataSet.Tables("TACHE").Rows(ii)
            If hour.CompareTo(CDate(CStr(ObjetDataRow("HORAIRE")))) > 0 And CInt(ObjetDataRow("ACTIVATION")) = 1 Then
                hour = CDate(CStr(ObjetDataRow("HORAIRE")))
                idCouran = ii
            End If
        Next

        'Affichage d'une boite de dialogue pour vérifier l'heure suivante
        If hour <> actual Then

        Else
            MsgBox("Il n'ya aucune tâche à effectuer pour l'instant")
        End If
        'MsgBox(CStr(hour))
        If ObjetDataSet.Tables("TACHE").Rows.Count <> 0 And hour.Year <> 2090 Then
            Label10.Text = "Prochaine tâche le : " + hour.ToShortDateString() + " à " + hour.ToShortTimeString()
        Else
            Label10.Text = "Pas de prochaine tâche"
        End If

        Return hour

    End Function

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MsgBox("Voulez-vous réellement quitter cette application?", MsgBoxStyle.YesNo) = vbYes Then
            If ObjetConnection.State = ConnectionState.Open Then
                ObjetConnection.Close()
            End If
        Else
            e.Cancel = True
        End If
    End Sub

    Sub SyntheVocal(ByVal chaine As String)
        Dim voice As Object
        voice = CreateObject("sapi.spvoice")
        voice.Speak(chaine)
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'DataGridView1.ScrollBars.Horizontal

        AfficheTous()

        TextBox3.Text = My.Settings.son_minime
        TextBox4.Text = My.Settings.son_standard
        TextBox5.Text = My.Settings.son_urgent

        Me.ShowInTaskbar = True

        Select Case My.Settings.jouer_son
            Case "Son"
                RadioButton4.Checked = True
            Case "Vocal"
                RadioButton5.Checked = True
            Case Else
                RadioButton6.Checked = True
        End Select

        Me.ForeColor = My.Settings.couleurtexte
        Me.Font = My.Settings.policetexte

        'SyntheVocal("Bienvenue monsieur Penaye Cyrille!")

        NotifyIcon1 = New NotifyIcon()
        'NotifyIcon2.ShowBalloonTip(20000)
        TextBox1.Focus()
        'DateTimePicker1.MinDate = Date.Now
        executer_routine(ListBox1)
        hpp = HeureLaPlusProche()
        Timer2.Start()
        Label8.Text = "Nombre de tâches enregistrées : " & CStr(ObjetDataSet.Tables("TACHE").Rows.Count)

    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'DateTimePicker2.MinDate = New Date(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month, DateTimePicker1.Value.Day, 0, 0, 0)
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        j = ListBox1.SelectedIndex
        Label7.Text = CStr(j)
    End Sub

    Function EgaliteHeure(ByVal h1 As Date, ByVal h2 As Date) As Boolean
        If h1.Year = h2.Year And h1.Month = h2.Month And h1.Day = h2.Day And h1.Minute = h2.Minute And h1.Hour = h2.Hour And h1.Second = h2.Second Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick

        If EgaliteHeure(hpp, Date.Now) Then

            strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & nomBD & ";"
            ObjetConnection = New OleDbConnection
            ObjetConnection.ConnectionString = strConn
            ObjetConnection.Open()

            'hpp = HeureLaPlusProche()
            ObjetDataRow = ObjetDataSet.Tables("TACHE").Rows(idCouran)
            NotifyIcon2.BalloonTipText = CStr(ObjetDataRow("MESSAGE"))
            NotifyIcon2.BalloonTipTitle = CStr(ObjetDataRow("NOM"))
            'NotifyIcon1.ShowBalloonTip(200000)
            NotifyIcon2.ShowBalloonTip(20000)
            Timer2.Stop()

            If My.Settings.jouer_son = "Son" Then
                Select Case CStr(ObjetDataRow("NIVEAU"))
                    Case 1
                        nom_audio = My.Settings.son_standard
                    Case 0
                        nom_audio = My.Settings.son_minime
                    Case 2
                        nom_audio = My.Settings.son_urgent
                End Select
                My.Computer.Audio.Play(nom_audio & "", AudioPlayMode.BackgroundLoop)
            ElseIf My.Settings.jouer_son = "Vocal" Then
                SyntheVocal("C'est le moment d'effectuer la tâche " & CStr(ObjetDataRow("NOM")))
            End If

            If MsgBox("Nom de la tâche : " & CStr(ObjetDataRow("NOM")) & vbCrLf & "Horaire de la tâche : " & CStr(ObjetDataRow("HORAIRE")) & vbCrLf & "Niveau d'urgence : " & level(CInt(ObjetDataRow("NIVEAU"))) & vbCrLf & "Date de création : " & CStr(ObjetDataRow("DATEENREG"))) = vbOK Then
                Timer2.Start()
                ObjetDataRow("ACTIVATION") = 0
                'ObjetDataSet.Tables("TACHE").Rows(idCouran).Delete()
                hpp = HeureLaPlusProche()
                My.Computer.Audio.Stop()
            End If

            ObjetCommandBuilder = New OleDbCommandBuilder(ObjetDataAdapter)
            ObjetDataAdapter.Update(ObjetDataSet, "TACHE")
            ObjetDataSet.Clear()

            ObjetDataAdapter.Fill(ObjetDataSet, "TACHE")
            ObjetDataTable = ObjetDataSet.Tables("TACHE")
            ListBox1.DataSource = ObjetDataSet.Tables("TACHE")
            ListBox1.DisplayMember = "NOM"
            ObjetConnection.Close()
            i = 0
            j = 0
            'hpp = HeureLaPlusProche()


        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ObjetDataRow = ObjetDataSet.Tables("TACHE").Rows(j)
        If MsgBox("Voulez-vous réellement supprimer la tâche " & "<<" & CStr(ObjetDataRow("NOM")) & ">>" & "?", MsgBoxStyle.YesNo) = vbYes Then
            strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & nomBD & ";"
            ObjetConnection = New OleDbConnection
            ObjetConnection.ConnectionString = strConn
            ObjetConnection.Open()

            If Not j = -1 Then
                ObjetDataSet.Tables("TACHE").Rows(j).Delete()
            End If

            ObjetCommandBuilder = New OleDbCommandBuilder(ObjetDataAdapter)
            ObjetDataAdapter.Update(ObjetDataSet, "TACHE")
            ObjetDataSet.Clear()

            ObjetDataAdapter.Fill(ObjetDataSet, "TACHE")
            ObjetDataTable = ObjetDataSet.Tables("TACHE")
            ListBox1.DataSource = ObjetDataSet.Tables("TACHE")
            ListBox1.DisplayMember = "NOM"
            ObjetConnection.Close()
            i = 0
            j = 0
            Label8.Text = "Nombre de tâches enregistrées : " & CStr(ObjetDataSet.Tables("TACHE").Rows.Count)
            hpp = HeureLaPlusProche()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ObjetDataRow = ObjetDataSet.Tables("TACHE").Rows(j)
        If MsgBox("Voulez-vous réellement modifier la tâche " & "<<" & CStr(ObjetDataRow("NOM")) & ">>" & "?", MsgBoxStyle.YesNo) = vbYes Then
            horaire = New Date(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month, DateTimePicker1.Value.Day, DateTimePicker2.Value.Hour, DateTimePicker2.Value.Minute, DateTimePicker2.Value.Second)
            If TextBox1.Text = Nothing Then
                MsgBox("Vous devez entrer une valeur dans le champ reservé au nom!", MsgBoxStyle.Information)
            ElseIf TextBox2.Text = Nothing Then
                MsgBox("Vous devez entrer une valeur dans le champ reservé au message!", MsgBoxStyle.Information)
            ElseIf horaire.CompareTo(Date.Now) < 0 Then
                MsgBox("Vous devez entrer une date qui est ultérieure à la date actuelle!", MsgBoxStyle.Information)
            Else
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & nomBD & ";"
                ObjetConnection = New OleDbConnection
                ObjetConnection.ConnectionString = strConn
                ObjetConnection.Open()

                nom = TextBox1.Text
                horaire = New Date(DateTimePicker1.Value.Year, DateTimePicker1.Value.Month, DateTimePicker1.Value.Day, DateTimePicker2.Value.Hour, DateTimePicker2.Value.Minute, DateTimePicker2.Value.Second)
                message = TextBox2.Text
                dateenreg = Date.Now
                If RadioButton1.Checked Then
                    niveau = 0
                ElseIf RadioButton2.Checked Then
                    niveau = 1
                Else
                    niveau = 2
                End If

                ObjetDataRow("NOM") = nom
                ObjetDataRow("HORAIRE") = CStr(horaire)
                ObjetDataRow("MESSAGE") = message
                ObjetDataRow("DATEENREG") = CStr(dateenreg)
                ObjetDataRow("NIVEAU") = niveau

                ObjetCommandBuilder = New OleDbCommandBuilder(ObjetDataAdapter)
                ObjetDataAdapter.Update(ObjetDataSet, "TACHE")
                ObjetDataSet.Clear()

                ObjetDataAdapter.Fill(ObjetDataSet, "TACHE")
                ObjetDataTable = ObjetDataSet.Tables("TACHE")
                ListBox1.DataSource = ObjetDataSet.Tables("TACHE")
                ListBox1.DisplayMember = "NOM"
                ObjetConnection.Close()
                i = 0
                j = 0
                hpp = HeureLaPlusProche()
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Dim level() As String = {"Minime", "Standard", "Urgent"}
    Dim state() As String = {"Désactivée", "Activée"}

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        ObjetDataRow = ObjetDataSet.Tables("TACHE").Rows(j)
        MsgBox("Nom de la tâche : " & CStr(ObjetDataRow("NOM")) & vbCrLf & "Horaire de la tâche : " & CStr(ObjetDataRow("HORAIRE")) & vbCrLf & "Niveau d'urgence : " & level(CInt(ObjetDataRow("NIVEAU"))) & vbCrLf & "Date de création : " & CStr(ObjetDataRow("DATEENREG")) & vbCrLf & "Etat : " & state(CInt(ObjetDataRow("ACTIVATION"))))
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        ObjetDataRow = ObjetDataSet.Tables("TACHE").Rows(j)
        If MsgBox("Voulez-vous réellement copier la tâche " & "<<" & CStr(ObjetDataRow("NOM")) & ">>" & " dans le formulaire de saisie?", MsgBoxStyle.YesNo) = vbYes Then
            TextBox1.Text = CStr(ObjetDataRow("NOM"))
            TextBox2.Text = CStr(ObjetDataRow("MESSAGE"))
            horaire = CDate(CStr(ObjetDataRow("HORAIRE")))
            DateTimePicker1.Value = New Date(horaire.Year, horaire.Month, horaire.Day, 23, 59, 59)
            DateTimePicker2.Value = New Date(horaire.Year, horaire.Month, horaire.Day, horaire.Hour, horaire.Minute, horaire.Second)
            Select Case CInt(ObjetDataRow("NIVEAU"))
                Case 0
                    RadioButton1.Checked = True
                Case 1
                    RadioButton2.Checked = True
                Case 2
                    RadioButton3.Checked = True
            End Select
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & nomBD & ";"
        ObjetConnection = New OleDbConnection
        ObjetConnection.ConnectionString = strConn
        ObjetConnection.Open()
        ObjetDataRow = ObjetDataSet.Tables("TACHE").Rows(j)

        If CInt(ObjetDataRow("ACTIVATION")) = 1 Then
            If MsgBox("Voulez-vous réellement désactiver la tâche " & "<<" & CStr(ObjetDataRow("NOM")) & ">>", MsgBoxStyle.YesNo) = vbYes Then
                ObjetDataRow("ACTIVATION") = 0
            End If
        Else
            If MsgBox("Voulez-vous réellement activer la tâche " & "<<" & CStr(ObjetDataRow("NOM")) & ">>", MsgBoxStyle.YesNo) = vbYes Then

                If CDate(ObjetDataRow("HORAIRE")).CompareTo(Date.Now) > 0 Then
                    ObjetDataRow("ACTIVATION") = 1
                Else
                    MsgBox("Une tâche qui a déjà eu lieu ne peut être à nouveau activée", MsgBoxStyle.Information)
                End If

            End If
        End If

        ObjetCommandBuilder = New OleDbCommandBuilder(ObjetDataAdapter)
        ObjetDataAdapter.Update(ObjetDataSet, "TACHE")
        ObjetDataSet.Clear()

        ObjetDataAdapter.Fill(ObjetDataSet, "TACHE")
        ObjetDataTable = ObjetDataSet.Tables("TACHE")
        ListBox1.DataSource = ObjetDataSet.Tables("TACHE")
        ListBox1.DisplayMember = "NOM"
        ObjetConnection.Close()
        i = 0
        j = 0
        hpp = HeureLaPlusProche()
    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Button1_Click(sender, e)
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Button3_Click(sender, e)
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        Button4_Click(sender, e)
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        Button5_Click_1(sender, e)
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        TextBox2.Text = TextBox1.Text
    End Sub

    Private Sub AProposDeFICAgendaNumérique12ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AProposDeFICAgendaNumérique12ToolStripMenuItem.Click
        AboutBox1.ShowDialog()
    End Sub

    Private Sub FondDécranToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FondDécranToolStripMenuItem.Click

    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            TextBox4.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            TextBox5.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub Button6_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            TextBox3.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Me.Close()
    End Sub

    Private Sub WxcwxcwToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WxcwxcwToolStripMenuItem.Click
        'Code pour génerer la liste des tâches à effectuer dans un fichier texte ou PDF

        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then


            'Dim titre As String = "Liste des tâches du " + Now.ToString("d-M-yyyy h-m-s")
            Dim titre = SaveFileDialog1.FileName
            Dim nomdoc As String
            nomdoc = titre
            Dim Paragraph As New Paragraph
            Dim PdfFile As New Document(PageSize.A4, 40, 40, 40, 20)
            PdfFile.AddTitle(titre)
            Dim Write As PdfWriter = PdfWriter.GetInstance(PdfFile, New FileStream(nomdoc, FileMode.Create))
            PdfFile.Open()

            'Déclaration du type de la police
            Dim pTitle As New Font(iTextSharp.text.Font.FontFamily.HELVETICA, 14, iTextSharp.text.Font.BOLD, BaseColor.BLACK)
            Dim pTable As New Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK)

            'Insertion du titre dans le fichier PDF
            Paragraph = New Paragraph(New Chunk(titre, pTitle))
            Paragraph.Alignment = Element.ALIGN_CENTER
            Paragraph.SpacingAfter = 5.0F

            'Définition de la page munie des nouveaux paramètres
            PdfFile.Add(Paragraph)

            'Création d'une table de données
            Dim PdfTable As New PdfPTable(DataGridView1.Columns.Count)

            'Paramétrage de la taille du tableau
            PdfTable.TotalWidth = 500.0F
            PdfTable.LockedWidth = True

            Dim widths(0 To DataGridView1.Columns.Count - 1) As Single
            For i As Integer = 0 To DataGridView1.Columns.Count - 1
                widths(i) = 1.0F
            Next

            PdfTable.SetWidths(widths)
            PdfTable.HorizontalAlignment = 0
            PdfTable.SpacingBefore = 5.0F

            'Déclaration des cellules du PDF
            Dim PdfCell As PdfPCell = New PdfPCell

            'Création de l'entête du PDF
            For i As Integer = 0 To DataGridView1.Columns.Count - 1
                PdfCell = New PdfPCell(New Phrase(New Chunk(DataGridView1.Columns(i).HeaderText, pTable)))
                PdfCell.HorizontalAlignment = PdfPCell.ALIGN_LEFT
                PdfTable.AddCell(PdfCell)
            Next

            'Ajout des données dans la table du pdf
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                For j As Integer = 0 To DataGridView1.Columns.Count - 1
                    PdfCell = New PdfPCell(New Phrase(DataGridView1(j, i).Value.ToString(), pTable))
                    PdfTable.HorizontalAlignment = PdfPCell.ALIGN_LEFT
                    PdfTable.AddCell(PdfCell)
                Next
            Next

            'Ajout de la table dans le document PDF
            PdfFile.Add(PdfTable)
            PdfFile.Close()

            'Affichage du message de réussite
            MsgBox("La génération de la liste des tâches sous format PDF a réussi!", MsgBoxStyle.Information)

        End If
    End Sub

    Private Sub Button7_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        If Not TextBox3.Text = "" Then
            My.Settings.son_minime = TextBox3.Text
            MsgBox("Modification réussie!", MsgBoxStyle.Information)
        Else
            MsgBox("Vous devez choisir un fichier!", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        If Not TextBox4.Text = "" Then
            My.Settings.son_standard = TextBox4.Text
            MsgBox("Modification réussie!", MsgBoxStyle.Information)
        Else
            MsgBox("Vous devez choisir un fichier!", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        If Not TextBox5.Text = "" Then
            My.Settings.son_urgent = TextBox5.Text
            MsgBox("Modification réussie!", MsgBoxStyle.Information)
        Else
            MsgBox("Vous devez choisir un fichier!", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        If RadioButton4.Checked Then
            My.Settings.jouer_son = "Son"
        ElseIf RadioButton5.Checked Then
            My.Settings.jouer_son = "Vocal"
        Else
            My.Settings.jouer_son = "Rien"
        End If
        MsgBox("Modification enregistrée!", MsgBoxStyle.Information)
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        'Ici, je supprime tous les éléments de la base de données
        If MsgBox("Attention, cette opération supprimera toutes les informations enregistrées, voulez-vous tout de même continuer?", MsgBoxStyle.Exclamation) = MsgBoxResult.Ok Then
            SupprimeTous()
            AfficheTous()
            executer_routine(ListBox1)
            Label8.Text = "Nombre de tâches enregistrées : 0"
            MsgBox("La réinitialisation a fonctionné!", MsgBoxStyle.Information)
        End If
    End Sub

    Private Sub PoliceDesTextesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PoliceDesTextesToolStripMenuItem.Click
        If FontDialog1.ShowDialog() = DialogResult.OK Then
            Me.Font = FontDialog1.Font
            My.Settings.policetexte = FontDialog1.Font
        End If
    End Sub

    Private Sub CouleurDesBoutonsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CouleurDesBoutonsToolStripMenuItem.Click
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            Me.ForeColor = ColorDialog1.Color
            My.Settings.couleurtexte = ColorDialog1.Color
        End If
    End Sub

    Private Sub AfficherLaideToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AfficherLaideToolStripMenuItem.Click
        'Ici, on ouvre le fichier pdf dans notre interface graphique d'aide du logiciel
        'Process.Start("C:\Users\NPENAYEE\Documents\Visual Studio 2010\Projects\Agenda Electronique\Agenda Electronique\bin\Debug\Documents\PresentationAN.pdf")

        Process.Start(My.Application.Info.DirectoryPath + "\Documents\PresentationAN.pdf")
        'Dialog1.ShowDialog()
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        'Ici, on affiche tout encore
        AfficheTous()
    End Sub

    Private Sub TabControl1_TabIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.TabIndexChanged

    End Sub
End Class
