VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmBooking 
   BackColor       =   &H0016161D&
   BorderStyle     =   0  'None
   Caption         =   "Booking"
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      Height          =   495
      Left            =   5760
      TabIndex        =   30
      Top             =   7320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   7080
      TabIndex        =   29
      Top             =   7320
      Width           =   5175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4575
      Left            =   5160
      TabIndex        =   28
      Top             =   7800
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   8070
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
      Caption         =   "UNPAID BOOKINGS"
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
   Begin VB.CommandButton cmdArchive 
      Caption         =   "&Archive"
      Height          =   495
      Left            =   19800
      TabIndex        =   17
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   12360
      Top             =   7320
   End
   Begin MSComCtl2.DTPicker dtReturn 
      Height          =   375
      Left            =   10800
      TabIndex        =   16
      Top             =   4920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   129826817
      CurrentDate     =   46078
   End
   Begin MSComCtl2.DTPicker dtPick 
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   4920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   129826817
      CurrentDate     =   46078
   End
   Begin VB.CommandButton cmdBook 
      Caption         =   "&Book Now"
      Height          =   495
      Left            =   5400
      TabIndex        =   14
      Top             =   6360
      Width           =   2655
   End
   Begin VB.TextBox txtTotal 
      Height          =   375
      Left            =   20880
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtStatus 
      Height          =   375
      Left            =   20880
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtDays 
      Height          =   375
      Left            =   20880
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4200
      Width           =   1335
   End
   Begin VB.ComboBox cboPrice 
      Height          =   315
      Left            =   15960
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4560
      Width           =   1815
   End
   Begin VB.ComboBox cboSeater 
      Height          =   315
      Left            =   15960
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4080
      Width           =   1815
   End
   Begin VB.ComboBox cboCarName 
      Height          =   315
      Left            =   15960
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   3600
      Width           =   2295
   End
   Begin VB.ComboBox cboBrand 
      Height          =   315
      Left            =   15960
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.ComboBox cboPlate 
      Height          =   315
      Left            =   15960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtSearchCarName 
      Height          =   495
      Left            =   15960
      TabIndex        =   5
      Top             =   2040
      Width           =   3255
   End
   Begin VB.ComboBox cboCustomerName 
      Height          =   315
      Left            =   8760
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.ComboBox cboLicense 
      Height          =   315
      Left            =   8760
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtSearchCustomer 
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Booking"
      BeginProperty Font 
         Name            =   "Segoe UI Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5040
      TabIndex        =   43
      Top             =   720
      Width           =   9735
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
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Return Day"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   10800
      TabIndex        =   42
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pick Up Day"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   6960
      TabIndex        =   41
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Price"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   19200
      TabIndex        =   40
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   19200
      TabIndex        =   39
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   19200
      TabIndex        =   38
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   5520
      TabIndex        =   37
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "License Number:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   5520
      TabIndex        =   36
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   12720
      TabIndex        =   35
      Top             =   4560
      Width           =   3015
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Seater Type:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   12720
      TabIndex        =   34
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Model:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   12720
      TabIndex        =   33
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Brand:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   12720
      TabIndex        =   32
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Plate Number:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   12720
      TabIndex        =   31
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5160
      Picture         =   "frmbooking.frx":0000
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   495
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
      MouseIcon       =   "frmbooking.frx":78121
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   3240
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
      MouseIcon       =   "frmbooking.frx":791EB
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   3720
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
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   1260
      MouseIcon       =   "frmbooking.frx":7A2B5
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   4200
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1260
      TabIndex        =   24
      Top             =   4680
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
      TabIndex        =   22
      Top             =   2040
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
      TabIndex        =   21
      Top             =   3000
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
      TabIndex        =   20
      Top             =   5520
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
      TabIndex        =   19
      Top             =   2280
      Width           =   2415
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
      TabIndex        =   18
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label lblVehicle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Plate Number"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   12720
      TabIndex        =   4
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblCustomer 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer License Number"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      Height          =   12135
      Left            =   0
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "frmBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    
    Dim db As clsDB
    
    Dim rsBookings As ADODB.Recordset
    Private recBooking As ADODB.Recordset
    
    ' -----------------------------
    ' Public ADODB connections
    ' -----------------------------
    Public con As ADODB.Connection        ' Main database
    Public conArchive As ADODB.Connection ' Archive database
    Private isEditing As Boolean
    Private isAdding As Boolean
    
' --- NEW: track selected booking ---
Private selectedBookingPlate As String
Private Sub ClearGridSelection()
    On Error Resume Next
    ' Scroll to top, do NOT select any row
    DataGrid1.FirstRow = 0
    On Error GoTo 0
End Sub
    Public Sub OpenDB()
        ' --- Main DB ---
        If con Is Nothing Then Set con = New ADODB.Connection
        If con.State = adStateClosed Then
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MasterData.mdb;"
            con.Open
        End If
    
        ' --- Archive DB ---
        If conArchive Is Nothing Then Set conArchive = New ADODB.Connection
        If conArchive.State = adStateClosed Then
            conArchive.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\bookingarchive.mdb;"
            conArchive.Open
        End If
    
        ' Initialize table tracker
        If m_TableCounts Is Nothing Then
            Set m_TableCounts = CreateObject("Scripting.Dictionary")
        End If
    End Sub
    
    Public Sub CloseDB()
        If Not con Is Nothing Then
            If con.State = adStateOpen Then con.Close
            Set con = Nothing
        End If
    
        If Not conArchive Is Nothing Then
            If conArchive.State = adStateOpen Then conArchive.Close
            Set conArchive = Nothing
        End If
    
        If Not m_TableCounts Is Nothing Then
            m_TableCounts.RemoveAll
            Set m_TableCounts = Nothing
        End If
    End Sub
    '=========================
    ' Utility
    '=========================
    Private Function SafeText(ByVal s As String) As String
        SafeText = Replace(s, "'", "''")
    End Function
    
    Private Sub cmdDelete_Click()
    If recBooking Is Nothing Or recBooking.EOF Or recBooking.BOF Then
            MsgBox "Select a record first!", vbExclamation
            Exit Sub
        End If
    
        If MsgBox("Delete this record?", _
                  vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
        recBooking.Delete
        recBooking.Update
    
    
        MsgBox "Deleted successfully!", vbInformation
    End Sub
    
    
Private Sub cmdArchive_Click()
    Dim rsBooking As ADODB.Recordset
    Dim insertSQL As String

    On Error GoTo ErrHandler

    ' --- Ensure a booking is selected ---
    If selectedBookingPlate = "" Then
        MsgBox "Please select a booking first by clicking a row in the grid.", vbExclamation
        Exit Sub
    End If

    ' --- Load booking record from main DB ---
    Set rsBooking = New ADODB.Recordset
    rsBooking.Open "SELECT * FROM bookings WHERE CarPlate='" & db.SafeText(selectedBookingPlate) & "'", _
                   db.con, adOpenKeyset, adLockReadOnly

    If rsBooking.EOF Then
        MsgBox "Booking not found!", vbExclamation
        GoTo CleanUp
    End If

    ' --- Build INSERT SQL for archive DB ---
    insertSQL = "INSERT INTO bookingarchive (" & _
                "bookID, bookingCode, CusLicense, CusName, CusContact, CusType, CusExpiration, " & _
                "CarPlate, CarBrand, CarName, CarSeater, CarPrice, PickDate, ReturnDate, Days, Status, TotalPrice) " & _
                "VALUES (" & _
                "'" & db.SafeText(rsBooking!bookID) & "', " & _
                "'" & db.SafeText(rsBooking!bookingCode) & "', " & _
                "'" & db.SafeText(rsBooking!CusLicense) & "', " & _
                "'" & db.SafeText(rsBooking!CusName) & "', " & _
                "'" & db.SafeText(rsBooking!cusContact) & "', " & _
                "'" & db.SafeText(rsBooking!cusType) & "', " & _
                "#" & Format(rsBooking!CusExpiration, "mm/dd/yyyy") & "#, " & _
                "'" & db.SafeText(rsBooking!CarPlate) & "', " & _
                "'" & db.SafeText(rsBooking!CarBrand) & "', " & _
                "'" & db.SafeText(rsBooking!CarName) & "', " & _
                "'" & db.SafeText(rsBooking!CarSeater) & "', " & _
                rsBooking!carPrice & ", " & _
                "#" & Format(rsBooking!PickDate, "mm/dd/yyyy") & "#, " & _
                "#" & Format(rsBooking!ReturnDate, "mm/dd/yyyy") & "#, " & _
                rsBooking!days & ", " & _
                "'Available', " & _
                rsBooking!TotalPrice & ")"

    ' --- Insert into archive DB ---
    db.ExecuteSQL insertSQL, True   ' True = archive DB

    ' --- Update vehicle status in main DB to Available ---
    db.ExecuteSQL "UPDATE vehicles SET Status='Available' WHERE Plate='" & db.SafeText(selectedBookingPlate) & "'"

    ' --- Delete booking from main DB ---
    db.ExecuteSQL "DELETE FROM bookings WHERE CarPlate='" & db.SafeText(selectedBookingPlate) & "'"

    ' --- Refresh Booking DataGrid ---
    Set DataGrid1.DataSource = Nothing
    LoadBookingGrid

    ' --- Refresh Vehicle combo boxes ---
    LoadVehicleList ""  ' Reload available vehicles

    ' --- Optional: update computations ---
    UpdateComputation

    ' Clear selection after archiving
    selectedBookingPlate = ""
    ClearGridSelection

    MsgBox "Booking archived successfully! Vehicle is now available.", vbInformation

CleanUp:
    If Not rsBooking Is Nothing Then
        If rsBooking.State = adStateOpen Then rsBooking.Close
    End If
    Set rsBooking = Nothing
    Exit Sub

ErrHandler:
    MsgBox "Error archiving booking: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

Private Sub Command2_Click()

End Sub

Private Sub DataGrid1_Click()
    On Error GoTo ErrHandler
    
    ' Ensure recordset exists and has records
    If rsBookings Is Nothing Then Exit Sub
    If rsBookings.EOF Or rsBookings.BOF Then Exit Sub

    ' Move recordset to the row clicked in the grid
    rsBookings.Bookmark = DataGrid1.Bookmark

    ' Store selected CarPlate safely
    If Not IsNull(rsBookings!CarPlate) Then
        selectedBookingPlate = rsBookings!CarPlate
    Else
        selectedBookingPlate = ""
    End If

    isEditing = True
    isAdding = False

    Exit Sub

ErrHandler:
    MsgBox "Error selecting booking: " & Err.Description, vbExclamation
End Sub
    Private Sub Form_Load()
    
        Me.Left = (Screen.Width - Me.Width) / 2
        Me.Top = (Screen.Height - Me.Height) / 2
        Randomize
    
        ' --- Open databases ---
        Set db = New clsDB
        db.OpenDB
    
        dtPick.Value = Date
        dtReturn.Value = Date
    
        ' Timer
        Timer1.Interval = 60000
        Timer1.Enabled = True
    
        ' Update statuses
        UpdateBookingStatus
    
        ' Load lists
        LoadVehicleList ""
        LoadCustomerList
        LoadBookingGrid
    
        UpdateComputation
        ClearGridSelection
    
        ' DataGrid settings
        DataGrid1.AllowAddNew = False
        DataGrid1.AllowDelete = False
        DataGrid1.AllowUpdate = False
    
    End Sub
    Private Sub dtPick_Change()
        ' Prevent past date
        If dtPick.Value < Date Then
            MsgBox "Pick-up date cannot be in the past.", vbExclamation
            dtPick.Value = Date
        End If
    
        ' Ensure dtReturn is not before dtPick
        If dtReturn.Value < dtPick.Value Then
            dtReturn.Value = dtPick.Value
        End If
    
        UpdateBookingStatus
        UpdateComputation
    End Sub
    
Private Sub lblbilllingMB_Click()
 Unload Me
    frmBilling.Show vbModal
End Sub

    Private Sub lblcustomerMB_Click()
    Unload Me
    frmcustomer.Show vbModal
    End Sub
    
    Private Sub lblvehicleMB_Click()
    Unload Me
    
    frmcar.Show vbModal
    End Sub
    
    
    Private Sub Timer1_Timer()
        Dim today As Date
        today = Date
    
        ' Prevent past dates for pick-up and return
        If dtPick.Value < today Then dtPick.Value = today
        If dtReturn.Value < dtPick.Value Then dtReturn.Value = dtPick.Value
    
        ' Update booking status in DB
        UpdateBookingStatus
    
        ' Refresh DataGrid to show latest status
        LoadBookingGrid
        
        UpdateVehicleStatuses
    End Sub
    '=========================
    ' Date Validation
    '=========================
    Private Sub dtPick_Validate(Cancel As Boolean)
        If dtPick.Value < Date Then
            MsgBox "Pickup date cannot be in the past.", vbExclamation, "Invalid Date"
            dtPick.Value = Date
            Cancel = True
        End If
    End Sub
    
    Private Sub dtReturn_Validate(Cancel As Boolean)
        If dtReturn.Value < Date Then
            MsgBox "Return date cannot be in the past.", vbExclamation, "Invalid Date"
            dtReturn.Value = Date
            Cancel = True
        End If
    End Sub
    
    
    Private Sub dtReturn_Change()
        ' Prevent return date before pick-up date
        If dtReturn.Value < dtPick.Value Then
            MsgBox "Return date cannot be before pick-up date.", vbExclamation
            dtReturn.Value = dtPick.Value
        End If
    
        UpdateComputation
    End Sub
    '=========================
    ' CUSTOMER LIST FUNCTIONS
    '=========================
    Sub LoadCustomerList(Optional ByVal keyword As String = "")
        Dim rs As ADODB.Recordset
        Dim sql As String
        Dim sLicense As String, sName As String
        Dim lCusID As Long
    
        Set rs = New ADODB.Recordset
    
        sql = "SELECT License, Name, cusID FROM customer " & _
              "WHERE Name LIKE '%" & SafeText(keyword) & "%' ORDER BY Name"
    
        rs.CursorLocation = adUseClient
        rs.Open sql, db.con, adOpenStatic, adLockReadOnly
    
        cboLicense.Clear
        cboCustomerName.Clear
    
        Do While Not rs.EOF
            sLicense = IIf(IsNull(rs!License), "", rs!License)
            sName = IIf(IsNull(rs!Name), "", rs!Name)
            lCusID = IIf(IsNull(rs!cusID), 0, rs!cusID)
    
            cboLicense.AddItem sLicense
            cboLicense.ItemData(cboLicense.NewIndex) = lCusID
    
            cboCustomerName.AddItem sName
            cboCustomerName.ItemData(cboCustomerName.NewIndex) = lCusID
    
            rs.MoveNext
        Loop
    
        rs.Close
        Set rs = Nothing
    End Sub
    
    Private Sub txtSearchCarName_KeyPress(KeyAscii As Integer)
        Dim sText As String
        Dim iLen As Integer
        
        sText = txtSearchCarName.Text
        iLen = Len(sText)
        
        ' Allow Backspace (8) and Delete (46)
        If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
        
        ' Limit total length to 7 characters (3 letters + 4 digits)
        If iLen >= 7 Then
            KeyAscii = 0
            Beep
            Exit Sub
        End If
        
        ' For first 3 characters, only allow letters
        If iLen < 3 Then
            If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
                KeyAscii = Asc(UCase(Chr(KeyAscii))) ' convert to uppercase automatically
            Else
                KeyAscii = 0
                Beep
            End If
        Else
            ' Last 4 characters must be digits
            If KeyAscii >= 48 And KeyAscii <= 57 Then
                ' OK
            Else
                KeyAscii = 0
                Beep
            End If
        End If
    End Sub
    Private Sub txtSearchCustomer_Change()
        FilterCustomerByLicense txtSearchCustomer.Text
    End Sub
    
    Sub FilterCustomerByLicense(ByVal keyword As String)
        Dim rs As New ADODB.Recordset
        Dim sql As String
        Dim sLicense As String, sName As String
        Dim lCusID As Long
    
        sql = "SELECT License, Name, cusID FROM customer " & _
              "WHERE License LIKE '" & SafeText(keyword) & "%' ORDER BY License"
    
        rs.CursorLocation = adUseClient
        rs.Open sql, db.con, adOpenStatic, adLockReadOnly
    
        cboLicense.Clear
        cboCustomerName.Clear
    
        Do While Not rs.EOF
            sLicense = IIf(IsNull(rs!License), "", rs!License)
            sName = IIf(IsNull(rs!Name), "", rs!Name)
            lCusID = IIf(IsNull(rs!cusID), 0, rs!cusID)
    
            cboLicense.AddItem sLicense
            cboLicense.ItemData(cboLicense.NewIndex) = lCusID
    
            cboCustomerName.AddItem sName
            cboCustomerName.ItemData(cboCustomerName.NewIndex) = lCusID
    
            rs.MoveNext
        Loop
    
        rs.Close
        Set rs = Nothing
    
        ' Auto-select first
        If cboLicense.ListCount > 0 Then
            cboLicense.ListIndex = 0
            cboCustomerName.ListIndex = 0
        End If
    End Sub
    
    Sub SyncCustomer(ByVal id As Long)
        Dim i As Integer
        For i = 0 To cboLicense.ListCount - 1
            If cboLicense.ItemData(i) = id Then cboLicense.ListIndex = i
        Next
        For i = 0 To cboCustomerName.ListCount - 1
            If cboCustomerName.ItemData(i) = id Then cboCustomerName.ListIndex = i
        Next
    End Sub
    
    Private Sub cboLicense_Click()
        If cboLicense.ListIndex >= 0 Then SyncCustomer cboLicense.ItemData(cboLicense.ListIndex)
    End Sub
    
    Private Sub cboCustomerName_Click()
        If cboCustomerName.ListIndex >= 0 Then SyncCustomer cboCustomerName.ItemData(cboCustomerName.ListIndex)
    End Sub
    
    Sub LoadVehicleList(Optional ByVal keyword As String = "")
        Dim rs As ADODB.Recordset
        Dim sql As String
        Dim sStatus As String
    
        Set rs = New ADODB.Recordset
        
        ' Include vehicles that are Available or have empty/NULL status
        sql = "SELECT * FROM vehicles WHERE (Status IS NULL OR Status='' OR Status='Available') AND (" & _
              "Plate LIKE '%" & SafeText(keyword) & "%' OR " & _
              "Brand LIKE '%" & SafeText(keyword) & "%' OR " & _
              "Name LIKE '%" & SafeText(keyword) & "%' OR " & _
              "Seater LIKE '%" & SafeText(keyword) & "%') ORDER BY Plate"
    
        rs.CursorLocation = adUseClient
        rs.Open sql, db.con, adOpenStatic, adLockReadOnly
    
        ' Clear combo boxes before loading
        cboPlate.Clear
        cboBrand.Clear
        cboCarName.Clear
        cboSeater.Clear
        cboPrice.Clear
    
        ' Loop through vehicles and add to combo boxes
        Do While Not rs.EOF
            ' If status is NULL or empty, treat as Available
            If IsNull(rs!status) Or rs!status = "" Then
                sStatus = "Available"
                ' Update DB to reflect this
                db.con.Execute "UPDATE vehicles SET Status='Available' WHERE carID=" & rs!carID
            Else
                sStatus = rs!status
            End If
    
            ' Call AddVehicleRow to populate combo boxes
            AddVehicleRow rs
    
            rs.MoveNext
        Loop
    
        rs.Close
        Set rs = Nothing
    End Sub
    Sub AddVehicleRow(rs As ADODB.Recordset)
        Dim id As Long
        Dim sPlate As String, sBrand As String, sName As String, sSeater As String
        Dim dPrice As Double
        Dim status As String
    
        ' Null-safe conversions
        id = IIf(IsNull(rs!carID), 0, rs!carID)
        sPlate = IIf(IsNull(rs!Plate), "", rs!Plate)
        sBrand = IIf(IsNull(rs!brand), "", rs!brand)
        sName = IIf(IsNull(rs!Name), "", rs!Name)
        sSeater = IIf(IsNull(rs!Seater), "", rs!Seater)
        dPrice = IIf(IsNull(rs!price), 0, rs!price)
        status = IIf(IsNull(rs!status) Or rs!status = "", "Available", rs!status)
    
        ' Clear combo boxes only once before first load
        If cboPlate.ListCount = 0 Then
            cboPlate.Clear
            cboBrand.Clear
            cboCarName.Clear
            cboSeater.Clear
            cboPrice.Clear
        End If
    
        ' Add items to combo boxes
        cboPlate.AddItem sPlate
        cboPlate.ItemData(cboPlate.NewIndex) = id
    
        cboBrand.AddItem sBrand
        cboBrand.ItemData(cboBrand.NewIndex) = id
    
        cboCarName.AddItem sName
        cboCarName.ItemData(cboCarName.NewIndex) = id
    
        cboSeater.AddItem sSeater
        cboSeater.ItemData(cboSeater.NewIndex) = id
    
        cboPrice.AddItem Format(dPrice, "0.00")
        cboPrice.ItemData(cboPrice.NewIndex) = id
    
        ' Make sure vehicles with empty or NULL status show as Available
        If status = "Available" Then
            db.con.Execute "UPDATE vehicles SET Status='Available' WHERE carID=" & id
        End If
    End Sub
    Private Sub txtSearchCarName_Change()
        FilterVehiclesByPlate txtSearchCarName.Text
    End Sub
    
    Private Sub txtSearchBrand_Change()
        FilterVehicles txtSearchCarName.Text, txtSearchBrand.Text
    End Sub
    
    Sub UpdateVehicleStatuses()
        Dim sql As String
        sql = "UPDATE vehicles v " & _
              "INNER JOIN bookings b ON v.Plate = b.CarPlate " & _
              "SET v.Status = b.Status"
        db.con.Execute sql
    End Sub
    Sub FilterVehiclesByPlate(ByVal keyword As String)
        Dim rs As New ADODB.Recordset
        Dim sql As String
    
        sql = "SELECT * FROM vehicles WHERE Plate LIKE '" & SafeText(keyword) & "%' ORDER BY Plate"
    
        rs.CursorLocation = adUseClient
        rs.Open sql, db.con, adOpenStatic, adLockReadOnly
    
        cboPlate.Clear
        cboBrand.Clear
        cboCarName.Clear
        cboSeater.Clear
        cboPrice.Clear
    
        Do While Not rs.EOF
            AddVehicleRow rs
            rs.MoveNext
        Loop
    
        rs.Close
        Set rs = Nothing
    End Sub
    
    Sub FilterVehicles(Optional ByVal CarName As String = "", Optional ByVal brand As String = "")
        Dim rs As New ADODB.Recordset
        Dim sql As String
    
        sql = "SELECT * FROM vehicles WHERE Name LIKE '%" & SafeText(CarName) & _
              "%' AND Brand LIKE '%" & SafeText(brand) & "%'"
    
        rs.CursorLocation = adUseClient
        rs.Open sql, db.con, adOpenStatic, adLockReadOnly
    
        cboPlate.Clear
        cboBrand.Clear
        cboCarName.Clear
        cboSeater.Clear
        cboPrice.Clear
    
        Do While Not rs.EOF
            AddVehicleRow rs
            rs.MoveNext
        Loop
    
        rs.Close
        Set rs = Nothing
    End Sub
    
    Sub SyncVehicle(ByVal id As Long)
        Dim cboList(4) As ComboBox
        Dim i As Integer, j As Integer
    
        Set cboList(0) = cboPlate
        Set cboList(1) = cboBrand
        Set cboList(2) = cboCarName
        Set cboList(3) = cboSeater
        Set cboList(4) = cboPrice
    
        For i = 0 To 4
            For j = 0 To cboList(i).ListCount - 1
                If cboList(i).ItemData(j) = id Then
                    cboList(i).ListIndex = j
                    Exit For
                End If
            Next
        Next
    
        UpdateComputation
    End Sub
    
    Private Sub cboPlate_Click()
        If cboPlate.ListIndex >= 0 Then
            SyncVehicle cboPlate.ItemData(cboPlate.ListIndex)
            UpdateComputation
        End If
    End Sub
    
    Private Sub cboBrand_Click()
        If cboBrand.ListIndex >= 0 Then
            SyncVehicle cboBrand.ItemData(cboBrand.ListIndex)
            UpdateComputation
        End If
    End Sub
    
    Private Sub cboCarName_Click()
        If cboCarName.ListIndex >= 0 Then
            SyncVehicle cboCarName.ItemData(cboCarName.ListIndex)
            UpdateComputation
        End If
    End Sub
    
    Private Sub cboSeater_Click()
        If cboSeater.ListIndex >= 0 Then
            SyncVehicle cboSeater.ItemData(cboSeater.ListIndex)
            UpdateComputation
        End If
    End Sub
    
    Private Sub cboPrice_Click()
        If cboPrice.ListIndex >= 0 Then
            SyncVehicle cboPrice.ItemData(cboPrice.ListIndex)
        End If
    End Sub
    
    Sub UpdateComputation()
        On Error Resume Next
        
        Dim pickD As Date, returnD As Date
        Dim totalDays As Long
        Dim carID As Long, pricePerDay As Double
        Dim status As String
        
        ' Validate dates
        If Not IsDate(dtPick.Value) Or Not IsDate(dtReturn.Value) Then Exit Sub
        
        pickD = dtPick.Value
        returnD = dtReturn.Value
        
        ' Total planned days (minimum 1)
        totalDays = DateDiff("d", pickD, returnD)
        If totalDays <= 0 Then totalDays = 1
        
        ' Get selected carID
        carID = 0
        If cboPlate.ListIndex >= 0 Then carID = cboPlate.ItemData(cboPlate.ListIndex)
        
        ' Get car price per day
        pricePerDay = GetCarPrice(carID)
        
        ' Compute status dynamically
        status = ComputeStatus(pickD, returnD, Date)
        
        ' Update UI
        txtDays.Text = totalDays
        txtStatus.Text = status
        txtTotal.Text = Format(totalDays * pricePerDay, "0.00")
        
        ' Update vehicle table status immediately
        If carID > 0 Then
            db.con.Execute "UPDATE vehicles SET Status='" & SafeText(status) & "' WHERE carID=" & carID
        End If
    End Sub
    '=========================
    ' GET DATA FUNCTIONS
    '=========================
    Function GetCustomerContact(cusID As Long) As String
        Dim rs As New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "SELECT Contact FROM customer WHERE cusID=" & cusID, db.con, adOpenStatic, adLockReadOnly
        If Not rs.EOF Then GetCustomerContact = IIf(IsNull(rs!Contact), "", rs!Contact)
        rs.Close: Set rs = Nothing
    End Function
    
    '=========================
    ' Generate unique booking code
    '=========================
    Function GenerateBookingCode() As String
        Dim code As String
        Dim rs As New ADODB.Recordset
    
Retry:
        ' Generate random code like BOOK-1234A
        code = "BOOK-" & Int((9999 - 1000 + 1) * Rnd + 1000) & Chr(Int((90 - 65 + 1) * Rnd + 65))
        
        ' Check if code already exists
        rs.CursorLocation = adUseClient
        rs.Open "SELECT bookingCode FROM bookings WHERE bookingCode='" & code & "'", db.con, adOpenStatic, adLockReadOnly
        
        If Not rs.EOF Then
            rs.Close
            Set rs = Nothing
            GoTo Retry
        End If
        
        rs.Close
        Set rs = Nothing
        
        GenerateBookingCode = code
    End Function
    Function GetCustomerType(cusID As Long) As String
        Dim rs As New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "SELECT Type FROM customer WHERE cusID=" & cusID, db.con, adOpenStatic, adLockReadOnly
        If Not rs.EOF Then GetCustomerType = IIf(IsNull(rs!Type), "", rs!Type)
        rs.Close: Set rs = Nothing
    End Function
    
    Function GetCustomerExpiration(cusID As Long) As Date
        Dim rs As New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "SELECT Expiration FROM customer WHERE cusID=" & cusID, db.con, adOpenStatic, adLockReadOnly
        If Not rs.EOF Then
            If IsNull(rs!Expiration) Then GetCustomerExpiration = Date Else GetCustomerExpiration = rs!Expiration
        Else
            GetCustomerExpiration = Date
        End If
        rs.Close: Set rs = Nothing
    End Function
    
    Function GetCarPrice(carID As Long) As Double
        Dim rs As New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "SELECT Price FROM vehicles WHERE carID=" & carID, db.con, adOpenStatic, adLockReadOnly
        If Not rs.EOF Then
            If IsNull(rs!price) Then GetCarPrice = 0 Else GetCarPrice = rs!price
        Else
            GetCarPrice = 0
        End If
        rs.Close: Set rs = Nothing
    End Function
    
    '=========================
    ' BOOKING STATUS UPDATE
    '=========================
    Sub UpdateBookingStatus()
        Dim sql As String
    
        ' Update booking status based on date
        sql = "UPDATE bookings " & _
              "SET Status = IIf(Date() < PickDate,'Reserved'," & _
              "IIf(Date() <= ReturnDate,'On Going','Completed'))"
              
        db.con.Execute sql
    
        ' Set vehicles to BOOKED if booking is Reserved or On Going
        sql = "UPDATE vehicles INNER JOIN bookings ON vehicles.Plate = bookings.CarPlate " & _
              "SET vehicles.Status = 'Booked' " & _
              "WHERE bookings.Status='Reserved' OR bookings.Status='On Going'"
    
        db.con.Execute sql
    
        ' Set vehicles back to AVAILABLE if booking is completed
        sql = "UPDATE vehicles SET Status='' " & _
              "WHERE Plate NOT IN (" & _
              "SELECT CarPlate FROM bookings WHERE Status='Reserved' OR Status='On Going')"
    
        db.con.Execute sql
    End Sub
    '=========================
    ' MODULE-LEVEL VARIABLE
    '=========================
      ' Keep the recordset bound to the grid
    
Private Sub LoadBookingGrid()
    Set rsBookings = New ADODB.Recordset
    rsBookings.CursorLocation = adUseClient
    rsBookings.CursorType = adOpenStatic
    rsBookings.LockType = adLockReadOnly

    ' Fetch all bookings
    rsBookings.Open "SELECT * FROM bookings ORDER BY PickDate DESC", db.con

    ' Bind the recordset to DataGrid1
    Set DataGrid1.DataSource = rsBookings

    ' --- Hide the last column (TotalPrice) ---
    Dim lastCol As Integer
    lastCol = DataGrid1.Columns.Count - 1   ' zero-based index
    If lastCol >= 0 Then
        DataGrid1.Columns(lastCol).Width = 0
    End If

    ' --- Scroll to top WITHOUT selecting a row ---
    On Error Resume Next
    DataGrid1.FirstRow = 0
    ' Do NOT set DataGrid1.Row = 0
    On Error GoTo 0
End Sub
    '=========================
    ' REFRESH BOOKING GRID AFTER INSERT
    '=========================
    Private Sub RefreshBookingGrid()
        ' Requery the existing recordset to include new records
        rsBookings.Requery
        ClearGridSelection
    End Sub
    
    '=========================
    ' BOOKING BUTTON CLICK
    '=========================
    Private Sub cmdBook_Click()
    
        Dim code As String
        Dim days As Long, total As Double
        Dim cusID As Long, carID As Long
        Dim cusContact As String, cusType As String
        Dim cusExp As Date, carPrice As Double
        Dim status As String
        Dim sql As String
    
        On Error GoTo ErrHandler
    
        ' Validate selection
        If cboLicense.ListIndex < 0 Or cboPlate.ListIndex < 0 Then
            MsgBox "Select customer and vehicle first.", vbExclamation
            Exit Sub
        End If
    
        cusID = cboLicense.ItemData(cboLicense.ListIndex)
        carID = cboPlate.ItemData(cboPlate.ListIndex)
    
        If cusID = 0 Or carID = 0 Then
            MsgBox "Invalid customer or vehicle.", vbExclamation
            Exit Sub
        End If
    
        ' Generate booking code
        code = GenerateBookingCode()
    
        ' Get computed values
        days = CLng(txtDays.Text)
        total = CDbl(txtTotal.Text)
    
        ' Get customer info
        cusContact = GetCustomerContact(cusID)
        cusType = GetCustomerType(cusID)
        cusExp = GetCustomerExpiration(cusID)
    
        ' Get vehicle info
        carPrice = GetCarPrice(carID)
    
        ' Compute status dynamically
        status = ComputeStatus(dtPick.Value, dtReturn.Value, Date)
    
        ' Update vehicle status
        db.con.Execute "UPDATE vehicles SET Status='" & SafeText(status) & "' WHERE carID=" & carID
    
        ' Insert booking
        sql = "INSERT INTO bookings (" & _
            "bookingCode,CusLicense,CusName,CusContact,CusType,CusExpiration," & _
             "CarPlate,CarBrand,CarName,CarSeater,CarPrice," & _
            "PickDate,ReturnDate,Days,Status,TotalPrice,BillingStatus) VALUES (" & _
            "'" & SafeText(code) & "'," & _
            "'" & SafeText(cboLicense.Text) & "'," & _
            "'" & SafeText(cboCustomerName.Text) & "'," & _
            "'" & SafeText(cusContact) & "'," & _
            "'" & SafeText(cusType) & "'," & _
            "#" & Format(cusExp, "mm/dd/yyyy") & "#," & _
            "'" & SafeText(cboPlate.Text) & "'," & _
            "'" & SafeText(cboBrand.Text) & "'," & _
             "'" & SafeText(cboCarName.Text) & "'," & _
            "'" & SafeText(cboSeater.Text) & "'," & _
            carPrice & "," & _
            "#" & Format(dtPick.Value, "mm/dd/yyyy") & "#," & _
            "#" & Format(dtReturn.Value, "mm/dd/yyyy") & "#," & _
             days & "," & _
             "'" & SafeText(status) & "'," & _
             total & "," & _
                 "'Unpaid')"
    
        db.con.Execute sql
    
        ' Refresh DataGrid
        RefreshBookingGrid
    
        ' Update booking statuses
        UpdateBookingStatus
    
        ' Success message
        MsgBox "Booking saved! Code: " & code, vbInformation
    
    
        '=========================
        ' RELOAD COMBO LISTS
        '=========================
        LoadCustomerList
        LoadVehicleList ""
    
        ' Clear combo selections
        cboLicense.ListIndex = -1
        cboCustomerName.ListIndex = -1
        cboPlate.ListIndex = -1
        cboBrand.ListIndex = -1
        cboCarName.ListIndex = -1
        cboSeater.ListIndex = -1
        cboPrice.ListIndex = -1
    
        ' Reset fields
        txtDays.Text = ""
        txtTotal.Text = ""
    
        ' Reset dates
        dtPick.Value = Date
        dtReturn.Value = Date
    
        ' Recompute totals
        UpdateComputation
        
        ClearGridSelection
    
        Exit Sub
    
ErrHandler:
        MsgBox "Error saving booking: " & Err.Description & vbCrLf & vbCrLf & sql, vbCritical
    
    End Sub
'=========================
' REFRESH DATAGRID1
'=========================
Private Sub RefreshDataGrid1()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    ' Requery the bookings table
    rs.Open "SELECT * FROM bookings ORDER BY PickDate DESC", db.con, adOpenStatic, adLockReadOnly

    ' Bind to DataGrid1
    Set DataGrid1.DataSource = rs

    ' --- Scroll to top WITHOUT selecting any row ---
    On Error Resume Next
    DataGrid1.FirstRow = 0
    ' Do NOT set DataGrid1.Row = 0
    On Error GoTo 0
End Sub
    Function GetVehicleStatus(carID As Long) As String
        Dim rs As New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "SELECT Status FROM vehicles WHERE carID=" & carID, db.con, adOpenStatic, adLockReadOnly
        
        If Not rs.EOF Then
            If IsNull(rs!status) Then
                GetVehicleStatus = "Unknown"
            Else
                GetVehicleStatus = rs!status
            End If
        Else
            GetVehicleStatus = "Unknown"
        End If
        
        rs.Close
        Set rs = Nothing
    End Function
    
    Function ComputeStatus(pickD As Date, returnD As Date, Optional today As Date = 0) As String
        If today = 0 Then today = Date
        
        If today < pickD Then
            ComputeStatus = "Reserved"
        ElseIf today >= pickD And today <= returnD Then
            ComputeStatus = "On Going"
        Else
            ComputeStatus = "Overdue"
        End If
    End Function
    
    Private Sub txtSearchCustomer_KeyPress(KeyAscii As Integer)
     Dim pos As Integer
        pos = Len(txtSearchCustomer.Text) + 1   ' Next character position
    
        ' Allow Backspace
        If KeyAscii = 8 Then Exit Sub
    
        ' FIRST CHARACTER ? LETTER ONLY
        If pos = 1 Then
            If (KeyAscii >= 65 And KeyAscii <= 90) Or _
               (KeyAscii >= 97 And KeyAscii <= 122) Then
               
                ' Auto uppercase
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            Else
                KeyAscii = 0
            End If
            
            Exit Sub
        End If
    
        ' CHARACTERS 2 TO 11 ? DIGITS ONLY
        If pos >= 2 And pos <= 11 Then
            If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
                KeyAscii = 0
            End If
            Exit Sub
        End If
    
        ' LIMIT TO 11 TOTAL CHARACTERS
        If pos > 11 Then
            KeyAscii = 0
        End If
    End Sub
    
    
