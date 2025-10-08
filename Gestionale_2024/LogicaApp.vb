Imports System.Data.SQlite

Module LogicaApp
    Public Function Esegui_QuerySQL( DbFile As String , Query As String ) As DataTable     
        Dim connString As String
        Dim conn As SqliteConnection
        Dim cmd As SQLiteCommand
        Dim rdr As SQLiteDataReader
        Dim Tabella As DataTable
        
        ' utilizzo una stringa interpolata..
        connString = $"Data Source={DbFile};Version=3;Foreign Keys=True;FailIfMissing=True;"
        conn = New SQLiteConnection( connString ) 
        cmd = New SQLiteCommand( Query , conn )             
        cmd.Connection.Open()     
        rdr = cmd.ExecuteReader()
        Tabella = New DataTable()
        Tabella.Load( rdr )
        cmd.Connection.Close()
        Return Tabella     
    End Function

    Public Function Esegui_ComandoSQL( DbFile As String , SQLcmd As String ) As Integer
        Dim connString As String
        Dim conn As SqliteConnection
        Dim cmd As SQLiteCommand
        Dim N As Integer

        ' utilizzo una stringa interpolata..
        connString = $"Data Source={DbFile};Version=3;Foreign Keys=True;FailIfMissing=True;"
        conn = New SQLiteConnection( connString )
        cmd = New SQLiteCommand( SQLcmd , conn )
        cmd.Connection.Open()   
        N = cmd.ExecuteNonQuery()
        cmd.Connection.Close()
        Return N       
    End Function

    Function AggiornaGriglia() As DataTable
        Dim T As DataTable
        Dim query, DbFile As String
        
        ' Aggiorna la griglia prodotti leggendo dal db:
        DbFile =  CurDir() + "\Data\Gestionale.db"
        query = "SELECT id,descrizione,prezzo,fornitore,mag,riordino,soglia_riordino FROM prodotti;"
        T = Esegui_QuerySQL( DbFile , query )
        return T
    End Function

    Function ModificaProdotto( pk As String , campo As String , valore As String ) As Integer
        Dim query, DbFile As String
        Dim N As Integer
        
        try
            query = ""
            Select campo
                Case "descrizione": ' descrizione
                    query = "update prodotti set descrizione='" + valore + "' where id=" + pk +";"                           
                Case "prezzo": ' prezzo
                    ' sostituisco la virgola con il punto
                    valore = valore.Replace(",",".")
                    query = "update prodotti set prezzo=" + valore + " where id=" + pk + ";"
                Case "fornitore": ' fornitore
                    query = "update prodotti set fornitore=" + valore + " where id=" + pk + ";"
                Case "mag": ' mag
                    query = "update prodotti set mag=" + valore + " where id=" + pk + ";"
                Case "riordino": ' riordino
                    query = "update prodotti set riordino=" + valore + " where id=" + pk + ";"
                Case "soglia_riordino": ' soglia_riordino
                    query = "update prodotti set soglia_riordino=" + valore + " where id=" + pk + ";"
            End Select

            If query <> "" Then
                DbFile = "Gestionale.db"
                N = Esegui_ComandoSQL( DbFile , query )    
                return N
            Else 
                Return 0
            End If
        Catch exc As Exception
            Return 0
        End Try
    End Function

    Function AggiungiNuovoProdotto( Rec As Prodotto ) As Boolean
        Dim query As String
        Dim N As Integer

        query = "INSERT INTO Prodotti (descrizione,prezzo,fornitore,mag,categoria,stato,riordino,soglia_riordino) VALUES "
        query = query+"('"+Rec.descrizione+"',"+Rec.prezzo+","+Rec.fornitore+","+Rec.mag
        query = query+","+Rec.categoria+","+Rec.stato+","+Rec.riordino+","+Rec.soglia_riordino+");"
        N = Esegui_ComandoSQL( "Gestionale.db" , query )
        If N = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Function CaricoMagazzino( id As String , qt As Integer ) As Boolean
        Dim query, DbFile As String
        Dim N As Integer
  
        DbFile =  CurDir() + "\Data\Gestionale.db"
        query = $"UPDATE prodotti SET mag=mag+{CStr(qt)} WHERE id={id};"
        'query = "UPDATE prodotti SET mag=mag+" & CStr(qt) & " WHERE id=" & id & ";"           
        N = Esegui_ComandoSQL( DbFile , query )
        If N <> 1 Then
            Return False
        Else 
            Return True
        End If
    End Function

    Function ScaricoMagazzino( id As String , qt As Integer ) As Boolean
        Dim query, DbFile As String
        Dim N As Integer

        DbFile =  CurDir() + "\Data\Gestionale.db"   
        query = $"UPDATE prodotti SET mag=mag-{CStr(qt)} WHERE id={id} AND mag>{CStr(qt)};"
        'query = "UPDATE prodotti SET mag=mag-" & CStr(qt) & " WHERE id=" & id 
        'query = query & " AND mag>" & CStr(qt) & ";"        
        N = Esegui_ComandoSQL( DbFile , query )
        If N <> 1 Then
            Return False
        Else 
            Return True
        End If
    End Function


    Public Class Prodotto
        Public id As String
        Public descrizione As String
        Public prezzo As String
        Public fornitore As String
        Public mag As String
        Public categoria As String
        Public stato As String
        Public riordino As String
        Public soglia_riordino As String
    End Class
End Module
