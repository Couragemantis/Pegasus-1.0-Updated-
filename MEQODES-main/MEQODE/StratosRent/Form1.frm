VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H0016161D&
   BorderStyle     =   0  'None
   Caption         =   "MAIN"
   ClientHeight    =   12525
   ClientLeft      =   3075
   ClientTop       =   1140
   ClientWidth     =   22920
   ForeColor       =   &H80000013&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   12525
   ScaleWidth      =   22920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   11040
      TabIndex        =   10
      Top             =   9840
      Width           =   5775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "PEGASUS 1.0"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   1260
      TabIndex        =   9
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Overview"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1260
      TabIndex        =   8
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "STORAGE"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1260
      TabIndex        =   7
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "MANAGE"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1260
      TabIndex        =   6
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1260
      TabIndex        =   5
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lblinventoryMB 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1260
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label lblbilllingMB 
      BackStyle       =   0  'Transparent
      Caption         =   "Billing"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1260
      TabIndex        =   3
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label lblbookingMB 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Booking"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1260
      MouseIcon       =   "Form1.frx":0992
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label lblcustomerMB 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Customer"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1260
      MouseIcon       =   "Form1.frx":1A5C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label lblvehicleMB 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Vehicles"
      BeginProperty Font 
         Name            =   "@Microsoft YaHei UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1260
      MouseIcon       =   "Form1.frx":2B26
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      Height          =   12135
      Left            =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   8055
      Left            =   9000
      Picture         =   "Form1.frx":3BF0
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   9735
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H004040D8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000A&
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   -120
      Top             =   0
      Width           =   23175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Show
frmcar.Show vbModal
End Sub

Private Sub lblbookingMB_Click()
frmBooking.Show vbModal
End Sub

Private Sub lblcustomerMB_Click()
frmcustomer.Show VB.Modal
End Sub

