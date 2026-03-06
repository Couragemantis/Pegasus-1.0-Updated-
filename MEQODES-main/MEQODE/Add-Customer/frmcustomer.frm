VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmcustomer 
   BackColor       =   &H0016161D&
   BorderStyle     =   0  'None
   Caption         =   "Student Registration Management System"
   ClientHeight    =   12525
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   22920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   12525
   ScaleWidth      =   22920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Command1"
      Height          =   735
      Left            =   14640
      TabIndex        =   32
      Top             =   3480
      Width           =   2415
   End
   Begin VB.TextBox txtaddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   31
      Top             =   1800
      Width           =   6495
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   7920
      TabIndex        =   28
      Top             =   4440
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   129826817
      CurrentDate     =   46082
   End
   Begin VB.Timer tmrAutoRefresh 
      Interval        =   2000
      Left            =   11640
      Top             =   6000
   End
   Begin VB.TextBox txtln 
      BackColor       =   &H80000002&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7920
      TabIndex        =   17
      Top             =   1800
      Width           =   4935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5895
      Left            =   5040
      TabIndex        =   16
      Top             =   6480
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   10398
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
      Caption         =   "CUSTOMERS"
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
   Begin VB.ComboBox combodt 
      Height          =   315
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0016161D&
      BorderStyle     =   0  'None
      Caption         =   "Search License Number"
      Height          =   855
      Left            =   5520
      TabIndex        =   11
      Top             =   5880
      Width           =   6375
      Begin VB.CommandButton Command9 
         Caption         =   "&Search"
         Height          =   495
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   525
         Left            =   1440
         TabIndex        =   12
         Top             =   120
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0016161D&
      BorderStyle     =   0  'None
      Caption         =   "Manipulator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   5040
      TabIndex        =   6
      Top             =   5280
      Width           =   17655
      Begin VB.CommandButton Command4 
         Caption         =   "&Archive"
         Height          =   495
         Left            =   15600
         TabIndex        =   10
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Update"
         Height          =   495
         Left            =   3720
         TabIndex        =   9
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   495
         Left            =   1560
         TabIndex        =   8
         Top             =   0
         Width           =   1935
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.TextBox txtcn 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7920
      TabIndex        =   1
      Top             =   3240
      Width           =   4935
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7920
      TabIndex        =   0
      Top             =   2520
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   12840
      TabIndex        =   30
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   5040
      Picture         =   "frmcustomer.frx":0000
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Customers"
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
      TabIndex        =   29
      Top             =   600
      Width           =   9735
   End
   Begin VB.Shape Shape1 
      Height          =   12135
      Left            =   0
      Top             =   360
      Width           =   4935
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
      MouseIcon       =   "frmcustomer.frx":78121
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
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   1260
      MouseIcon       =   "frmcustomer.frx":791EB
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1260
      MouseIcon       =   "frmcustomer.frx":7A2B5
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Expiration Date:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5715
      TabIndex        =   15
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Driver Type:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5715
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Number:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5715
      TabIndex        =   4
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5715
      TabIndex        =   3
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   5715
      TabIndex        =   2
      Top             =   1800
      Width           =   1935
   End
End
Attribute VB_Name = "frmcustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =========================================
' Form-level variables
' =========================================
Private db As clsDB
Private rec As ADODB.Recordset
Private isEditing As Boolean
Private isAdding As Boolean
Private bIsTyping As Boolean


Private Sub Command1_Click()

End Sub

Private Sub cmdDelete_Click()
    ' Check if the recordset is valid and has records
    If rec Is Nothing Or rec.EOF Or rec.BOF Then
        MsgBox "No record selected to delete!", vbExclamation
        Exit Sub
    End If

    ' Confirm deletion
    If MsgBox("Are you sure you want to delete this customer?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    ' Move to the current DataGrid row
    On Error Resume Next
    rec.Bookmark = DataGrid1.Bookmark
    On Error GoTo 0

    ' Delete the record
    rec.Delete
    rec.Update ' Only needed if you have a batch update (optional with client-side cursors)

    ' Refresh the grid to reflect changes
    RefreshGrid

    MsgBox "Customer deleted successfully!", vbInformation

    ' Clear input fields after deletion
    ClearInputFields
    isEditing = False
End Sub
Private Sub Command9_Click()

    Dim sql As String

    If db Is Nothing Then Set db = New clsDB
    db.OpenDB

    If Trim(Text5.Text) = "" Then
        sql = "SELECT * FROM customer"
    Else
        sql = "SELECT * FROM customer WHERE License LIKE '" & _
              db.SafeText(Text5.Text) & "%'"
    End If

    If rec Is Nothing Then
        Set rec = New ADODB.Recordset
    ElseIf rec.State = adStateOpen Then
        rec.Close
    End If

    rec.CursorLocation = adUseClient
    rec.Open sql, db.con
    
    If rec.EOF Then
        MsgBox "No records found.", vbInformation
        Exit Sub  ' Don't change the DataGrid if no records
    End If

    Set DataGrid1.DataSource = rec
    HidecusIDColumn
End Sub
Private Sub DTPicker1_Change()
    Dim selectedDate As Date
    selectedDate = DTPicker1.Value

    ' Minimum date is tomorrow
    If selectedDate <= Date Then
        MsgBox "You cannot select today or past dates!", vbExclamation
        DTPicker1.Value = Date + 1
        Exit Sub
    End If

    ' Maximum year is 3000
    If Year(selectedDate) > 3000 Then
        MsgBox "Year cannot exceed 3000!", vbExclamation
        DTPicker1.Value = Date + 1
        Exit Sub
    End If
End Sub
' =========================================
' =============== FORM LOAD ===============
' =========================================
Private Sub Form_Load()

    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

    Set db = New clsDB
    db.OpenDB

    Set rec = New ADODB.Recordset
    With rec
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open "SELECT * FROM customer", db.con
    End With

    Set DataGrid1.DataSource = rec
    HidecusIDColumn

    combodt.Clear
    combodt.AddItem "Non - Professional"
    combodt.AddItem "Professional"
    combodt.ListIndex = -1

    DTPicker1.Format = dtpCustom
    DTPicker1.CustomFormat = "MM/dd/yyyy"
    DTPicker1.Value = Date + 1

    isAdding = False
    isEditing = False

    SetFieldsEditable False
    ClearInputFields
    ClearGridSelection

End Sub

' =========================================
' =============== GRID ====================
' =========================================
Private Sub DataGrid1_Click()

    If rec Is Nothing Or rec.EOF Or rec.BOF Then Exit Sub

    On Error Resume Next
    rec.Bookmark = DataGrid1.Bookmark
    On Error GoTo 0

    PopulateTextBoxesFromGrid

    isEditing = True
    isAdding = False
    SetFieldsEditable True

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

' =========================================
' =============== BUTTONS =================
' =========================================
Private Sub cmdNew_Click()
    StartNew
End Sub

Private Sub cmdAdd_Click()

    Dim rsCheck As ADODB.Recordset

    ' Required fields including ADDRESS
    If Trim(txtln.Text) = "" Or _
       Trim(txtname.Text) = "" Or _
       Trim(txtcn.Text) = "" Or _
       Trim(txtaddress.Text) = "" Or _
       combodt.ListIndex = -1 Then

        MsgBox "Please fill in all required fields!", vbExclamation
        Exit Sub
    End If

    If Not txtln.Text Like "[A-Za-z]##########" Then
        MsgBox "License must be 1 letter followed by 10 digits!", vbExclamation
        Exit Sub
    End If

    If Not txtcn.Text Like "09#########" Then
        MsgBox "Invalid Contact Number!", vbExclamation
        Exit Sub
    End If

    If DTPicker1.Value <= Date Then
        MsgBox "Expiration must be future date!", vbExclamation
        DTPicker1.Value = Date + 1
        Exit Sub
    End If

    Set rsCheck = New ADODB.Recordset
    rsCheck.Open "SELECT License FROM customer WHERE License='" & _
                 db.SafeText(txtln.Text) & "'", _
                 db.con, adOpenForwardOnly, adLockReadOnly

    If Not rsCheck.EOF Then
        MsgBox "License already exists!", vbExclamation
        rsCheck.Close
        Set rsCheck = Nothing
        Exit Sub
    End If

    rsCheck.Close
    Set rsCheck = Nothing

    ' ADD RECORD including ADDRESS
    rec.AddNew
    rec!License = txtln.Text
    rec!Name = txtname.Text
    rec!Contact = txtcn.Text
    rec!Address = txtaddress.Text
    rec!Type = combodt.Text
    rec!Expiration = DTPicker1.Value
    rec.Update

    RefreshGrid
    ClearInputFields
    isAdding = False

    MsgBox "Customer added successfully!", vbInformation

End Sub

Private Sub cmdEdit_Click()

    If isAdding Then
        MsgBox "Cannot edit while adding.", vbExclamation
        Exit Sub
    End If

    If rec Is Nothing Or rec.EOF Or rec.BOF Then
        MsgBox "Please select a record to edit!", vbExclamation
        Exit Sub
    End If

    If txtln.Text = rec!License And _
       txtname.Text = rec!Name And _
       txtcn.Text = rec!Contact And _
       txtaddress.Text = rec!Address And _
       combodt.Text = rec!Type And _
       DTPicker1.Value = rec!Expiration Then

        MsgBox "No changes detected!", vbInformation
        Exit Sub
    End If

    If MsgBox("Save changes?", vbYesNo + vbQuestion) = vbYes Then

        rec!Name = txtname.Text
        rec!Contact = txtcn.Text
        rec!Address = txtaddress.Text
        rec!Type = combodt.Text
        rec!Expiration = DTPicker1.Value
        rec.Update

        RefreshGrid
        ClearInputFields
        isEditing = False

        MsgBox "Customer updated successfully!", vbInformation

    End If

End Sub

' =========================================
' =============== UTILITIES ===============
' =========================================
Private Sub RefreshGrid()
    Dim lastColIndex As Integer
    
    If Not rec Is Nothing Then
        rec.Requery
        Set DataGrid1.DataSource = rec
        
        ' Hide the specific column you already have
        HidecusIDColumn
        
        ' Hide the last column dynamically
        lastColIndex = DataGrid1.Columns.Count - 1
        If lastColIndex >= 0 Then
            DataGrid1.Columns(lastColIndex).Visible = False
        End If
        
        DataGrid1.Refresh
        ClearGridSelection
    End If
End Sub
Private Sub ClearGridSelection()

    If Not rec Is Nothing Then
        If rec.State = adStateOpen Then
            If rec.RecordCount > 0 Then
                rec.MoveLast
                rec.MoveNext
            End If
        End If
    End If

End Sub

Private Sub StartNew()

    ClearInputFields
    isAdding = True
    isEditing = False
    SetFieldsEditable False, True
    ClearGridSelection

End Sub

Private Sub SafeSetFocus(ctrl As Control)
    On Error Resume Next
    If ctrl.Enabled And ctrl.Visible And Not ctrl.Locked Then
        ctrl.SetFocus
    End If
    On Error GoTo 0
End Sub
Private Sub ClearInputFields()

    txtln.Text = ""
    txtname.Text = ""
    txtcn.Text = ""
    txtaddress.Text = ""
    combodt.ListIndex = -1
    DTPicker1.Value = Date + 1
    SafeSetFocus txtln

End Sub

Private Sub SetFieldsEditable(Optional ByVal editing As Boolean = False, _
                              Optional ByVal adding As Boolean = False)

    txtln.Locked = IIf(editing And Not adding, True, False)
    txtname.Locked = False
    txtcn.Locked = False
    txtaddress.Locked = False
    combodt.Enabled = True
    DTPicker1.Enabled = True

End Sub

Private Sub PopulateTextBoxesFromGrid()

    If rec Is Nothing Or rec.EOF Or rec.BOF Then Exit Sub

    txtln.Text = IIf(IsNull(rec!License), "", rec!License)
    txtname.Text = IIf(IsNull(rec!Name), "", rec!Name)
    txtcn.Text = IIf(IsNull(rec!Contact), "", rec!Contact)
    txtaddress.Text = IIf(IsNull(rec!Address), "", rec!Address)

    combodt.ListIndex = -1
    Dim i As Integer

    If Not IsNull(rec!Type) Then
        For i = 0 To combodt.ListCount - 1
            If combodt.List(i) = rec!Type Then
                combodt.ListIndex = i
                Exit For
            End If
        Next i
    End If

    If Not IsNull(rec!Expiration) Then
        DTPicker1.Value = rec!Expiration
    End If

End Sub

Private Sub HidecusIDColumn()
    On Error Resume Next
    DataGrid1.Columns(DataGrid1.Columns.Count - 1).Visible = False
End Sub

Private Sub lblbilllingMB_Click()
 Unload Me
    frmBilling.Show vbModal
End Sub

Private Sub lblbookingMB_Click()
Unload Me
frmBooking.Show vbModal
End Sub

Private Sub lblvehicleMB_Click()
Unload Me
frmcar.Show vbModal
End Sub


Private Sub txtcn_KeyPress(KeyAscii As Integer)

    Dim pos As Integer
    pos = Len(txtcn.Text) + 1   ' Next position

    ' Allow Backspace
    If KeyAscii = 8 Then Exit Sub

    ' Allow numbers only
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
        Exit Sub
    End If

    ' First digit must be 0
    If pos = 1 Then
        If KeyAscii <> 48 Then KeyAscii = 0
        Exit Sub
    End If

    ' Second digit must be 9
    If pos = 2 Then
        If KeyAscii <> 57 Then KeyAscii = 0
        Exit Sub
    End If

    ' Limit to 11 digits only
    If pos > 11 Then
        KeyAscii = 0
    End If

End Sub
Private Sub txtln_KeyPress(KeyAscii As Integer)

    Dim pos As Integer
    pos = Len(txtln.Text) + 1   ' Next character position

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

