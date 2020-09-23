VERSION 5.00
Begin VB.Form frmClient 
   Caption         =   "No records found"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   2655
   End
   Begin VB.ListBox lstResults 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2655
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'I have deleted all the address for obivous resons

Private Sub cmdSave_Click()
If lstResults.ListIndex = -1 Then 'this tells us if we are adding a new client or finding and updating a existing one
    gvNewClient = True
    gvClient_ID = False
    gvSearched = True
    Unload Me
    Exit Sub
Else
    gvClient_ID = True
    gvNewClient = False
    gvSearched = True
    gvClient_ID = lstResults.ItemData(lstResults.ListIndex)
    Call frmClient1.GetClient
    Unload Me
End If
End Sub

Public Sub Form_Load()
Dim rs As ADODB.Recordset 'main recordeset for the form
Dim strSQL As String         'sql string for the search

Set rs = New Recordset
gvClient_ID = False         'Tells us if we found a record or not

Dim strLastname As String 'Variables from the first form
Dim strFirstname As String
  
strLastname = frmClient1.txtFields(0).Text  'set the variables for what to search for
strFirstname = frmClient1.txtFields(1).Text
    
     strSQL = "SELECT client_ID,lastname,firstname,address,suburb FROM client " & _
        "WHERE [lastname] LIKE '" & strLastname & "%' " & _
         "AND [firstname] LIKE '" & strFirstname & "%' " & _
         "ORDER by [Lastname], [firstname]"
         
     rs.CursorLocation = adUseClient
     rs.Open [strSQL], g_Cars, adOpenForwardOnly, adLockReadOnly
     
    lstResults.AddItem "Add new record" 'this is used for adding a new client
        If rs.RecordCount > 0 Then          'fil the list box with records
            rs.MoveFirst
            Do Until rs.EOF
                lstResults.AddItem rs![Lastname] & ",   " & rs![firstname] & "  " & rs![Address] & "   " & rs![suburb]
                lstResults.ItemData(lstResults.NewIndex) = rs("client_ID")
                rs.MoveNext
            Loop
        End If
        
If lstResults.ListCount > 1 Then frmClient.Caption = lstResults.ListCount - 1 & " records found"

cmdExit.Picture = LoadPicture(App.Path & "/Delete.ico")
CmdSave.Picture = LoadPicture(App.Path & "/Save.ico")
Me.Icon = LoadPicture(App.Path & "/Lookup.ico")

rs.Close                'close the recordset and set it to nothing
Set rs = Nothing

Exit Sub
            
Loaderr:
MsgBox Err.Description
  
End Sub

Private Sub cmdExit_Click()
gvClient_ID = False
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmClient = Nothing
End Sub

Private Sub lstResults_DblClick()
cmdSave_Click
End Sub

Private Sub lstResults_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdSave_Click
End If
End Sub
