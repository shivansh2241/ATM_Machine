Attribute VB_Name = "MOdcon"
Public con As Connection
Public rec As Recordset
Public sql, Keep As String
Public ins, trap As Integer
Public bool As Integer
Public pin As Long
Public accountnumber, atmnumber, atmpin As Long
Sub Network()
    Set con = New Connection
        Set rec = New Recordset
            con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ATM.mdb;jet oledb:database password=admin"
End Sub

