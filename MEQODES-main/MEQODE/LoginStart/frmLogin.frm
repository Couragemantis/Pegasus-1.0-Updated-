VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H0016161D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   5865
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465.236
   ScaleMode       =   0  'User
   ScaleWidth      =   5070.308
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc loginado 
      Height          =   330
      Left            =   240
      Top             =   5400
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmLogin.frx":0000
      OLEDBString     =   $"frmLogin.frx":0087
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from tb_credentials"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtuser 
      Height          =   465
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   3285
   End
   Begin VB.CommandButton loginbtn 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   3720
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   3720
      Width           =   1500
   End
   Begin VB.TextBox txtpass 
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   3285
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub loginbtn_Click()
loginado.RecordSource = "select * from tb_credentials " & _
"where Username = '" & txtuser.Text & "' " & _
"and Password = '" & txtpass.Text & "'"

loginado.Refresh

If loginado.Recordset.EOF Then
MsgBox "Wrong login credentials.", vbCritical, "Login Failed"
Else
MsgBox "Successfully Logged in.", vbInformation, "Welcome"
frmMain.Show
Exit Sub
End If
End Sub
