VERSION 5.00
Begin VB.Form frmClient1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Details"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      DataField       =   "client_ID"
      Height          =   285
      Index           =   8
      Left            =   600
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox chkFields 
      Caption         =   "Mailing List"
      DataField       =   "mailinglist"
      Height          =   315
      Left            =   2160
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4320
      Width           =   2535
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Email"
      Height          =   285
      Index           =   7
      Left            =   2160
      TabIndex        =   7
      Top             =   3315
      Width           =   2415
   End
   Begin VB.TextBox txtFields 
      DataField       =   "workphone"
      Height          =   285
      Index           =   6
      Left            =   2160
      TabIndex        =   6
      Top             =   2895
      Width           =   1215
   End
   Begin VB.TextBox txtFields 
      DataField       =   "homephone"
      Height          =   285
      Index           =   5
      Left            =   2160
      TabIndex        =   5
      Top             =   2505
      Width           =   1215
   End
   Begin VB.TextBox txtFields 
      DataField       =   "postcode"
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   4
      Top             =   2100
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Address"
      DataSource      =   "datClient"
      Height          =   285
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   1290
      Width           =   2895
   End
   Begin VB.TextBox txtFields 
      DataField       =   "firstname"
      Height          =   285
      Index           =   1
      Left            =   2175
      TabIndex        =   1
      Top             =   885
      Width           =   2295
   End
   Begin VB.TextBox txtFields 
      DataField       =   "lastname"
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   465
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox cboSuburb 
         DataField       =   "suburb"
         Height          =   315
         Left            =   1935
         Style           =   1  'Simple Combo
         TabIndex        =   3
         Top             =   1545
         Width           =   2520
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Date"
         Height          =   285
         Index           =   9
         Left            =   3945
         TabIndex        =   22
         Top             =   2010
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Client Details"
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Name:"
         Height          =   285
         Left            =   600
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name:"
         Height          =   285
         Left            =   600
         TabIndex        =   19
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Address:"
         Height          =   285
         Left            =   600
         TabIndex        =   18
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Suburb:"
         Height          =   285
         Left            =   600
         TabIndex        =   17
         Top             =   1575
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Post code:"
         Height          =   285
         Left            =   600
         TabIndex        =   16
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Home number:"
         Height          =   285
         Left            =   600
         TabIndex        =   15
         Top             =   2385
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Work / Mobile:"
         Height          =   285
         Left            =   600
         TabIndex        =   14
         Top             =   2790
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Email:"
         Height          =   285
         Left            =   600
         TabIndex        =   13
         Top             =   3195
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmClient1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim mbAddnew As Boolean 'are we adding a new record
Dim mbSearched As Boolean 'have we found a record
Dim rstemp As ADODB.Recordset
Dim ClientBookmark As Integer
Dim lngNewID As Long
Dim cn As ADODB.Connection

Private Sub cboSuburb_Click()
txtFields(4).Text = cboSuburb.ItemData(cboSuburb.ListIndex) 'fill the post code textbox with a postcode
End Sub

Private Sub cboSuburb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtFields(4).SetFocus
End Sub

Private Sub cboSuburb_LostFocus()
cboSuburb.Text = StrConv(cboSuburb, vbProperCase)
End Sub

Private Sub chkFields_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdSave.SetFocus
End Sub

Private Sub cmdExit_Click()
gvClient_ID = False
Set frmClient1 = Nothing
Unload Me
Unload frmClient1
End Sub
Private Sub cmdSave_Click()
If gvSearched Then
    subfinished
    Exit Sub
End If
If txtFields(0) = "" And txtFields(1) = "" Then             'do we or don't we search
            MsgBox "You must have a Lastname or a Firstname filled in", 64
        txtFields(0).SetFocus
        Exit Sub
    Else
If gvClient_ID = False Then
        frmClient.Show vbModal
        End If
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()

Me.Icon = LoadPicture(App.Path & "/Lookup.ico")
cmdExit.Picture = LoadPicture(App.Path & "/Delete.ico")
cmdSave.Picture = LoadPicture(App.Path & "/Save.ico")

Suburbs 'sub to fill cboSuburbs with records
chkFields.Value = 0
mbAddnew = False
mbSearched = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
gvClient_ID = False
gvSearched = False
Set rs = Nothing
Set rstemp = Nothing
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case Index
        Case 0
            txtFields(1).SetFocus
            KeyAscii = 0
        Case 1
            If gvSearched Then
            KeyAscii = 0
            txtFields(2).SetFocus
            Else
            KeyAscii = 0
            cmdSave_Click
            End If
        Case 2
            cboSuburb.SetFocus
            KeyAscii = 0
        Case 3
            txtFields(4).SetFocus
            KeyAscii = 0
        Case 4
            txtFields(5).SetFocus
            KeyAscii = 0
        Case 5
            txtFields(6).SetFocus
            KeyAscii = 0
        Case 6
            txtFields(7).SetFocus
            KeyAscii = 0
        Case 7
            chkFields.SetFocus
            KeyAscii = 0
    End Select
    End If
End Sub

Private Sub txtFields_LostFocus(Index As Integer) ' update all text to proper case
    txtFields(0).Text = StrConv(txtFields(0), vbProperCase)
    txtFields(1).Text = StrConv(txtFields(1), vbProperCase)
    txtFields(2).Text = StrConv(txtFields(2), vbProperCase)
End Sub

Public Sub GetClient()          'sub to retreive a client details
Dim cmd As ADODB.Command

    If gvClient_ID = False Then
        gvNewClient = True
     Exit Sub
    Else
         Set rs = Nothing
         Set rs = New Recordset
         Set cmd = New Command
         
         With cmd                                       'pass the parameter to the query
             Set .ActiveConnection = g_Cars
            .CommandText = "qry_sel_client"
            .CommandType = adCmdStoredProc
            .Parameters.Append .CreateParameter("client_ID", adInteger, adParamInput, , gvClient_ID)
            
         End With
         rs.CursorLocation = adUseClient        'open the recordset
         rs.Properties("Update Resync") = adResyncAutoIncrement
         rs.Open cmd, , adOpenKeyset, adLockBatchOptimistic
         
Dim oText As TextBox                            'bind the text boxes to the recordset
  For Each oText In Me.txtFields
    Set oText.DataSource = rs
Next
Set chkFields.DataSource = rs
Set cboSuburb.DataSource = rs

Set cmd = Nothing
Set rs = Nothing
End If
End Sub


Private Sub Suburbs() 'sub to fill the suburbs combo
                              ' the suburb field is a combo box that will lookup the postcode for each suburb
Set rstemp = New Recordset
With rstemp
    .Open "SELECT * FROM wasuburbs", g_Cars, adOpenForwardOnly, adLockReadOnly
    Do Until .EOF
        cboSuburb.AddItem ![suburb] & " , " & ![State]
        cboSuburb.ItemData(cboSuburb.NewIndex) = rstemp("Postcode")
        .MoveNext
    Loop
End With
rstemp.Close
Set rstemp = Nothing
End Sub
Private Sub subfinished()

If gvNewClient Then
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.Open "SELECT  * FROM client", g_Cars, adOpenKeyset, adLockOptimistic
    With rs
        .AddNew
        !Lastname = NullIt(txtFields(0))
        !firstname = NullIt(txtFields(1))
        !Address = NullIt(txtFields(2))
        !suburb = NullIt(cboSuburb)
        !Postcode = NullIt(txtFields(4))
        !homephone = NullIt(txtFields(5))
        !workphone = NullIt(txtFields(6))
        !Email = NullIt(txtFields(7))
        !mailinglist = chkFields.Value
        .Update
        .MoveLast
        txtFields(8).Text = rs!client_ID
    End With
Else
    Update
End If
Finish
End Sub

Private Sub Update()

g_Cars.BeginTrans
    Set rs = New Recordset
    rs.Source = "SELECT * FROM client WHERE client_ID = " & gvClient_ID
    rs.Open , g_Cars, adOpenKeyset, adLockOptimistic
    With rs
        rs("lastname") = NullIt(txtFields(0))
        rs("firstname") = NullIt(txtFields(1))
        rs("address") = NullIt(txtFields(2))
        rs("suburb") = NullIt(cboSuburb)
        rs("postcode") = NullIt(txtFields(4))
        rs("homephone") = NullIt(txtFields(5))
        rs("workphone") = NullIt(txtFields(6))
        rs("Email") = NullIt(txtFields(7))
        rs("date") = Now
        rs("mailinglist") = chkFields.Value
        .Update
        g_Cars.CommitTrans
        Exit Sub
    End With
  Set rs = Nothing
End Sub

Private Sub Finish()
If gvfrmAppointmentloaded Then frmClient1.Hide
End Sub


