VERSION 5.00
Begin VB.Form frmAppointment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client Search"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Client Details"
      Height          =   2175
      Left            =   165
      TabIndex        =   2
      Top             =   90
      Width           =   3735
      Begin VB.Label cmdClient 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label txtWorkphone 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Top             =   1305
         Width           =   1575
      End
      Begin VB.Label txtHomephone 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1305
         Width           =   1695
      End
      Begin VB.Label txtPostcode 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2280
         TabIndex        =   7
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label txtSuburb 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   6
         Top             =   990
         Width           =   2055
      End
      Begin VB.Label txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   675
         Width           =   3255
      End
      Begin VB.Label txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   1620
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2445
      Width           =   1785
   End
   Begin VB.TextBox txtClient_ID 
      DataField       =   "client_ID"
      Height          =   285
      Left            =   1185
      TabIndex        =   1
      Top             =   345
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Click over the labels to search for a client"
      Height          =   465
      Left            =   240
      TabIndex        =   11
      Top             =   2430
      Width           =   1755
   End
End
Attribute VB_Name = "frmAppointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Private Sub cmdClient_Click()       'this will show the client form
If txtClient_ID = "" Then               'if there is a client already found it unload it and then reload it
        frmClient1.Show vbModal, Me
    Else: Unload frmClient1
        Load frmClient1
        frmClient1.Show vbModal
    End If
 Exit Sub
End Sub

Private Sub cmdExit_Click()
Set frmAppointment = Nothing
End
End Sub

Private Sub Form_Activate()
txtName = frmClient1.txtFields(1).Text & "  " & frmClient1.txtFields(0).Text
txtAddress = frmClient1.txtFields(2)
txtSuburb = frmClient1.cboSuburb
txtPostcode = frmClient1.txtFields(4)
txtHomephone = frmClient1.txtFields(5)
txtWorkphone = frmClient1.txtFields(6)
txtEmail = frmClient1.txtFields(7)
txtClient_ID = frmClient1.txtFields(8)
End Sub

Private Sub Form_Load()
gvfrmAppointmentloaded = True
cmdExit.Picture = LoadPicture(App.Path & "/Delete.ico")     'set the icons at runtime rather than a design time
Me.Icon = LoadPicture(App.Path & "/Lookup.ico")
End Sub

Private Sub Form_Unload(Cancel As Integer)
gvfrmAppointmentloaded = False
Unload frmClient1
End Sub


