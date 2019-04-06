Attribute VB_Name = "Engine"
Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset
Public Objek As Control
Public X As Integer

Public Sub Nyambung()
If CN.State = adStateOpen Then CN.Close
    CN.CursorLocation = adUseClient
    CN.Open "Provider=MSDASQL.1;Persist Security Info=False;User ID=BambangAditya;Data Source=DBPenjualan"
End Sub

Public Sub PusatError()
    If Err.Number = -2147217900 Then
        MsgBox "Tidak boleh ada Kode Minuman yang sama!", vbCritical + vbOKOnly, "Error"
        With FormDataBarang
            .AdodcMain.Refresh
            .textKodeMinuman.Text = ""
            .textKodeMinuman.SetFocus
        End With
    End If
End Sub
