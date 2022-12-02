Attribute VB_Name = "modKoneksi"
Public con As New ADODB.Connection
Public tabel As New ADODB.Recordset
Public sql As String

Sub bukakoneksi()
Set con = New ADODB.Connection
If con.State = 1 Then con.Close
con.Open "DSN=ds_restaurant"
End Sub
