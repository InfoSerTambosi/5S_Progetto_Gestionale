Public Class Form1
    Private D1 As New Dialog1()     ' Dialog per inserimento prodotti
    

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim T As DataTable
        Dim query, DbFile As String
      
        With DataGridView1
            .AllowDrop = False
            .AllowUserToResizeRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeColumns = True
            .AllowUserToOrderColumns = False
            .RowHeadersVisible = False
            .MultiSelect = False
        End With

        With DataGridView2
            .AllowDrop = False
            .AllowUserToResizeRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToOrderColumns = False
            .RowHeadersVisible = False
        End With

        With DataGridView3
            .AllowDrop = False
            .AllowUserToResizeRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeColumns = True
            .AllowUserToOrderColumns = False
            .RowHeadersVisible = False
            .MultiSelect = False
        End With

        DataGridView1.DataSource = AggiornaGriglia()
          
        ' Aggiorna la griglia prodotti leggendo dal db:
        DbFile =  CurDir() + "\Data\Gestionale.db"
        query = "SELECT * FROM Fornitori;"
        T = Esegui_QuerySQL( DbFile , query )
        DataGridView3.DataSource = T             
    End Sub

    'Funzione "Visualizza"
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        
    End Sub

    ' Cerca Prodotto
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        
    End Sub

    ' Nuovo prodotto
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub ProdottiPerFatturatoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProdottiPerFatturatoToolStripMenuItem.Click

    End Sub

    ' Carico magazzino
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        
    End Sub

    ' Scarico magazzino
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        ' Visualizza tutti i prodotti
        DataGridView1.DataSource = AggiornaGriglia()
    End Sub
End Class
