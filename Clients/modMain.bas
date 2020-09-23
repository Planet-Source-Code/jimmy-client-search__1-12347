Attribute VB_Name = "modMain"
Option Explicit
Global gvClient_ID As Integer
Global gvfrmAppointmentloaded As Boolean
Global gvSearched As Boolean
Global gvNewClient As Boolean
Global g_Cars As ADODB.Connection

Sub main()

    Set g_Cars = New Connection
    
    With g_Cars
                .Provider = "Microsoft.Jet.OLEDB.4.0"
                .ConnectionString = App.Path & "\Clients.Mdb"
                .CursorLocation = adUseClient
                .Open
End With

Load frmAppointment
frmAppointment.Show

End Sub

Public Function NullIt(Ctl As Control) As Variant

    If TypeOf Ctl Is TextBox Or _
        TypeOf Ctl Is ComboBox Or _
        TypeOf Ctl Is Label Then
        If Ctl = "" Then
            NullIt = Null
        Else
            NullIt = Ctl
        End If
    End If
    
End Function
