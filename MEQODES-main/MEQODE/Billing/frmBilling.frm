VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBilling 
   BackColor       =   &H0016161D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12525
   ScaleWidth      =   22920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7080
      TabIndex        =   31
      Top             =   6720
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      Height          =   495
      Left            =   5760
      TabIndex        =   30
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdMarkAsPaid 
      Caption         =   "Mark As &Paid"
      Height          =   615
      Left            =   5640
      TabIndex        =   29
      Top             =   5640
      Width           =   3015
   End
   Begin VB.TextBox txtReturnDate 
      Height          =   375
      Left            =   14880
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtCharges 
      Height          =   1575
      Left            =   14880
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtTotal 
      Height          =   375
      Left            =   18720
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox txtPickDate 
      Height          =   375
      Left            =   14880
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtCarPrice 
      Height          =   375
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtCarName 
      Height          =   375
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox txtCarBrand 
      Height          =   375
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtCusName 
      Height          =   375
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtbookingCode 
      Height          =   375
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   2535
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5175
      Left            =   5160
      TabIndex        =   0
      Top             =   7200
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   9128
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5160
      Picture         =   "frmBilling.frx":0000
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label Label13 
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
      TabIndex        =   28
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
      TabIndex        =   27
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label12 
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
      TabIndex        =   26
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label11 
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
      TabIndex        =   25
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label10 
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
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   5760
      Width           =   2415
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
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   1260
      TabIndex        =   22
      Top             =   4680
      Width           =   2415
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
      MouseIcon       =   "frmBilling.frx":78121
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   4200
      Width           =   2415
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
      MouseIcon       =   "frmBilling.frx":791EB
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   3720
      Width           =   2415
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
      MouseIcon       =   "frmBilling.frx":7A2B5
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      Height          =   12135
      Left            =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label9 
      Caption         =   "Total Cost"
      Height          =   375
      Left            =   18720
      TabIndex        =   18
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "Additional Charges"
      Height          =   375
      Left            =   12240
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label7 
      Caption         =   "Returned Day"
      Height          =   375
      Left            =   12240
      TabIndex        =   8
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Pick Up Day"
      Height          =   375
      Left            =   12240
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Car Price"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Car Name"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Car Brand"
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Customer Name"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Bookiing Code"
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   1
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H004040D8&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000A&
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   23175
   End
End
Attribute VB_Name = "frmBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =========================================
' Form-level declarations
' =========================================
Dim db As clsDB
Dim rsBilling As ADODB.Recordset

' =========================================
' Form_Load: Initialize DB and load DataGrid
' =========================================
Private Sub Form_Load()
    ' --- Create and open database connection ---
    Set db = New clsDB
    db.OpenDB

    ' --- Check connection ---
    If db.con Is Nothing Or db.con.State <> adStateOpen Then
        MsgBox "Database failed to open!", vbCritical
        Exit Sub
    End If

    ' --- Load bookings into DataGrid ---
    LoadBillingGrid

    ' --- Lock textboxes so user cannot edit ---
    txtbookingCode.Locked = True
    txtCarBrand.Locked = True
    txtCarName.Locked = True
    txtCarPrice.Locked = True
    txtCharges.Locked = True
    txtPickDate.Locked = True
    txtReturnDate.Locked = True
    txtTotal.Locked = True
End Sub

' =========================================
' LoadBillingGrid: Load bookings into DataGrid1
' =========================================
Private Sub LoadBillingGrid()
    Dim sql As String

    sql = "SELECT bookingCode, CusName, CarBrand, CarName, CarPrice, PickDate, ReturnDate, BillingStatus FROM bookings"

    ' Close previous recordset if open
    If Not rsBilling Is Nothing Then
        If rsBilling.State = adStateOpen Then rsBilling.Close
        Set rsBilling = Nothing
    End If

    ' Open new recordset
    Set rsBilling = New ADODB.Recordset
    rsBilling.CursorLocation = adUseClient
    rsBilling.Open sql, db.con, adOpenStatic, adLockReadOnly

    ' Bind to DataGrid
    Set DataGrid1.DataSource = Nothing
    Set DataGrid1.DataSource = rsBilling
    DataGrid1.Refresh
End Sub

' =========================================
' DataGrid1_Click: Show selected row in textboxes
' =========================================
Private Sub DataGrid1_Click()
    On Error Resume Next ' Prevent errors if no row selected

    ' Check if there is a current row
    If rsBilling.EOF Then Exit Sub

    ' Populate textboxes from current row
    txtbookingCode.Text = rsBilling!bookingCode
    txtCusName.Text = rsBilling!CusName
    txtCarBrand.Text = rsBilling!CarBrand
    txtCarName.Text = rsBilling!CarName
    txtCarPrice.Text = rsBilling!carPrice
    txtCharges.Text = "" ' If you have charges calculated separately
    txtPickDate.Text = Format(rsBilling!PickDate, "mm/dd/yyyy")
    txtReturnDate.Text = Format(rsBilling!ReturnDate, "mm/dd/yyyy")
    txtTotal.Text = "" ' If you want to calculate total later
End Sub

' =========================================
' Form_Unload: Clean up
' =========================================
Private Sub Form_Unload(Cancel As Integer)
    ' Close recordset
    If Not rsBilling Is Nothing Then
        If rsBilling.State = adStateOpen Then rsBilling.Close
        Set rsBilling = Nothing
    End If

    ' Close database connection
    If Not db Is Nothing Then
        db.CloseDB
        Set db = Nothing
    End If
End Sub

Private Sub lblbookingMB_Click()
Unload Me
frmBooking.Show vbModal
End Sub

Private Sub lblcustomerMB_Click()
Unload Me
frmcustomer.Show vbModal
End Sub

Private Sub lblvehicleMB_Click()
Unload Me
frmcar.Show vbModal

End Sub
