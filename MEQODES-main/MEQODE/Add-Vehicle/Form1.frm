VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmcar 
   BackColor       =   &H0016161D&
   BorderStyle     =   0  'None
   Caption         =   "Manage Vehicles"
   ClientHeight    =   12525
   ClientLeft      =   2910
   ClientTop       =   1140
   ClientWidth     =   22920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   835
   ScaleMode       =   0  'User
   ScaleWidth      =   1528
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAutoRefresh 
      Interval        =   2000
      Left            =   11640
      Top             =   6000
   End
   Begin VB.TextBox txtplate 
      BackColor       =   &H80000002&
      Height          =   495
      Left            =   7920
      TabIndex        =   19
      Top             =   1800
      Width           =   4935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5895
      Left            =   5040
      TabIndex        =   17
      Top             =   6480
      Width           =   17775
      _ExtentX        =   31353
      _ExtentY        =   10398
      _Version        =   393216
      Appearance      =   0
      BackColor       =   -2147483642
      BorderStyle     =   0
      ForeColor       =   -2147483643
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   0
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
      Caption         =   "AVAILABLE VEHICLES"
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
   Begin VB.ComboBox comboseat 
      Height          =   315
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtprice 
      Height          =   375
      Left            =   7920
      TabIndex        =   14
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0016161D&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   5520
      TabIndex        =   11
      Top             =   6000
      Width           =   8415
      Begin VB.CommandButton Command9 
         Caption         =   "&Search"
         Height          =   495
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   525
         Left            =   1440
         TabIndex        =   12
         Top             =   0
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0016161D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   5040
      TabIndex        =   6
      Top             =   5280
      Width           =   17775
      Begin VB.CommandButton cmdArchive 
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
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   7920
      TabIndex        =   1
      Top             =   3120
      Width           =   4935
   End
   Begin VB.TextBox txtbrand 
      Height          =   495
      Left            =   7920
      TabIndex        =   0
      Top             =   2520
      Width           =   4935
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1260
      TabIndex        =   23
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
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   22
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
      MouseIcon       =   "Form1.frx":10CA
      MousePointer    =   99  'Custom
      TabIndex        =   21
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
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   1260
      MouseIcon       =   "Form1.frx":2194
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      Height          =   12135
      Left            =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   5040
      Picture         =   "Form1.frx":325E
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Manage Vehicles"
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
      TabIndex        =   16
      Top             =   600
      Width           =   9735
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6435
      TabIndex        =   15
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Seater:"
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
      Left            =   6435
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Left            =   6435
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6435
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   5955
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "frmcar"
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
    
    Private Sub Command9_Click()
    
        Dim sql As String
    
        If db Is Nothing Then Set db = New clsDB
        db.OpenDB
    
        ' Base SQL: only vehicles with empty or NULL Status
        sql = "SELECT * FROM vehicles WHERE (Status IS NULL OR Status='')"
    
        ' Apply search filter if user entered text
        If Trim(Text5.Text) <> "" Then
            sql = sql & " AND Plate LIKE '" & db.SafeText(Text5.Text) & "%'"
        End If
    
        ' Prepare recordset
        If rec Is Nothing Then
            Set rec = New ADODB.Recordset
        ElseIf rec.State = adStateOpen Then
            rec.Close
        End If
    
        rec.CursorLocation = adUseClient
        rec.CursorType = adOpenStatic
        rec.LockType = adLockOptimistic
        rec.Open sql, db.con
    
        ' Check if any records found
        If rec.EOF Then
            MsgBox "No records found.", vbInformation
            Exit Sub  ' Don't change the DataGrid if no records
        End If
    
        ' Bind to DataGrid and hide CarID + Status columns
        Set DataGrid1.DataSource = rec
        HidecarIDColumn
        ResizeDataGridColumns
        ClearGridSelection
    
    End Sub
    Private Sub Form_Load()
    
        Set db = New clsDB
        db.OpenDB
    
        ' Only load vehicles with Status = "Available"
        Set rec = New ADODB.Recordset
        With rec
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockReadOnly
            .Open "SELECT * FROM vehicles WHERE TRIM(Status)='Available'", db.con
        End With
    
        ' Bind to DataGrid
        Set DataGrid1.DataSource = rec
        HidecarIDColumn  ' Hide both CarID and Status columns
        ResizeDataGridColumns
    
        ' Make grid fully read-only
        DataGrid1.AllowAddNew = False
        DataGrid1.AllowDelete = False
        DataGrid1.AllowUpdate = False
    
        ' Initialize Seater ComboBox
        comboseat.Clear
        comboseat.AddItem "4 - Seater"
        comboseat.AddItem "6 - Seater"
        comboseat.AddItem "8 - Seater"
        comboseat.ListIndex = -1
    
        ' Clear input fields and set form state
        ClearInputFields
        isEditing = False
        isAdding = False
        SetFieldsEditable False
        ClearGridSelection
    
    End Sub
    Private Sub Form_Unload(Cancel As Integer)
    
        If Not rec Is Nothing Then
            If rec.State = adStateOpen Then rec.Close
            Set rec = Nothing
        End If
    
        If Not db Is Nothing Then
            db.CloseDB
            Set db = Nothing
        End If
    
    End Sub
    
    ' =========================================
    ' ============ Utility Methods =============
    ' =========================================
    Private Sub ClearInputFields()
        txtplate.Text = ""
        txtbrand.Text = ""
        txtname.Text = ""
        txtprice.Text = ""
        comboseat.ListIndex = -1
    End Sub
    
    Private Sub ClearGridSelection()
    
        If Not rec Is Nothing Then
            If rec.State = adStateOpen Then
                If rec.RecordCount > 0 Then
                    rec.MoveLast
                    rec.MoveNext   ' Move to EOF (no selected row)
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
    
    Private Sub SetFieldsEditable(Optional ByVal editing As Boolean = False, _
                                  Optional ByVal adding As Boolean = False)
    
        ' Plate editable ONLY when adding
        txtplate.Locked = IIf(editing And Not adding, True, False)
    
        txtbrand.Locked = False
        txtname.Locked = False
        txtprice.Locked = False
        comboseat.Enabled = True
    
    End Sub
    
    Public Sub RefreshGrid()
    
        If Not rec Is Nothing Then
            rec.Requery
            Set DataGrid1.DataSource = rec
            HidecarIDColumn
            ResizeDataGridColumns
            DataGrid1.Refresh
            ClearGridSelection
        End If
    
    End Sub
    
    Private Sub HidecarIDColumn()
        On Error Resume Next
        DataGrid1.Columns(DataGrid1.Columns.Count - 1).Visible = False
        DataGrid1.Columns(DataGrid1.Columns.Count - 2).Visible = False
    End Sub
    
    Private Sub ResizeDataGridColumns()
        Dim i As Long, colCount As Long, totalWidth As Long
    
        colCount = DataGrid1.Columns.Count
        If colCount = 0 Then Exit Sub
    
        totalWidth = DataGrid1.Width - 200
    
        For i = 0 To colCount - 1
            DataGrid1.Columns(i).Width = totalWidth / colCount
        Next i
    
    End Sub
    
    Private Sub PopulateTextBoxesFromGrid()
    
        If rec Is Nothing Or rec.EOF Or rec.BOF Then Exit Sub
    
        txtplate.Text = IIf(IsNull(rec!Plate), "", rec!Plate)
        txtbrand.Text = IIf(IsNull(rec!brand), "", rec!brand)
        txtname.Text = IIf(IsNull(rec!Name), "", rec!Name)
        txtprice.Text = IIf(IsNull(rec!price), "", CStr(rec!price))
    
        comboseat.ListIndex = -1
        Dim i As Integer
    
        If Not IsNull(rec!Seater) Then
            For i = 0 To comboseat.ListCount - 1
                If comboseat.List(i) = rec!Seater Then
                    comboseat.ListIndex = i
                    Exit For
                End If
            Next i
        End If
    
    End Sub
    
    ' =========================================
    ' ================= Buttons =================
    ' =========================================
    Private Sub cmdNew_Click()
        StartNew
    End Sub
    
    Private Sub cmdAdd_Click()
    
        Dim rsCheck As ADODB.Recordset
    
        ' --- Validate required fields ---
        If Trim(txtplate.Text) = "" Or Trim(txtbrand.Text) = "" Or _
           Trim(txtname.Text) = "" Or comboseat.ListIndex = -1 Or _
           Trim(txtprice.Text) = "" Then
    
            MsgBox "Please fill in all required fields!", vbExclamation
            Exit Sub
        End If
    
        ' --- Validate plate format ---
        If Not txtplate.Text Like "[A-Za-z][A-Za-z][A-Za-z]####" Then
            MsgBox "Plate must be 3 letters + 4 numbers!", vbExclamation
            Exit Sub
        End If
    
        ' --- Validate price ---
        If Not IsNumeric(txtprice.Text) Or Val(txtprice.Text) <= 0 Then
            MsgBox "Price must be positive!", vbExclamation
            Exit Sub
        End If
    
        ' --- Check if plate already exists ---
        Set rsCheck = New ADODB.Recordset
        rsCheck.Open "SELECT Plate FROM vehicles WHERE Plate='" & _
                     db.SafeText(txtplate.Text) & "'", _
                     db.con, adOpenForwardOnly, adLockReadOnly
    
        If Not rsCheck.EOF Then
            MsgBox "Plate already exists!", vbExclamation
            rsCheck.Close
            Set rsCheck = Nothing
            Exit Sub
        End If
    
        rsCheck.Close
        Set rsCheck = Nothing
    
        ' --- Add new vehicle record ---
        rec.AddNew
        rec!Plate = txtplate.Text
        rec!brand = txtbrand.Text
        rec!Name = txtname.Text
        rec!Seater = comboseat.Text
        rec!price = Val(txtprice.Text)
        rec!status = "Available"      ' <-- Set Status to Available
        rec.Update
    
        ' --- Refresh grid and reset form ---
        RefreshGrid
        StartNew
    
        MsgBox "Vehicle added successfully!", vbInformation
    
    End Sub
    Private Sub cmdEdit_Click()
    
        If isAdding Then
            MsgBox "Cannot edit while adding!", vbExclamation
            Exit Sub
        End If
    
        If rec Is Nothing Or rec.EOF Or rec.BOF Then
            MsgBox "Select a record first!", vbExclamation
            Exit Sub
        End If
    
        If MsgBox("Save changes?", vbYesNo + vbQuestion) = vbYes Then
    
            rec!brand = txtbrand.Text
            rec!Name = txtname.Text
            rec!Seater = comboseat.Text
            rec!price = Val(txtprice.Text)
            rec.Update
    
            RefreshGrid
            MsgBox "Record updated!", vbInformation
    
        End If
    
    End Sub
    
    Private Sub cmdArchive_Click()
    
        If rec Is Nothing Or rec.EOF Or rec.BOF Then Exit Sub
    
        Dim cnArchive As ADODB.Connection
        Dim rsArchive As ADODB.Recordset
    
        Set cnArchive = New ADODB.Connection
        cnArchive.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Vin\Documents\MEQODE\Archives\vehiclearchive\vehiclearchive.mdb;"
        cnArchive.Open
    
        Set rsArchive = New ADODB.Recordset
        rsArchive.Open "VehicleArchive", cnArchive, _
                       adOpenKeyset, adLockOptimistic, adCmdTable
    
        rsArchive.AddNew
        rsArchive!Plate = rec!Plate
        rsArchive!Name = rec!Name
        rsArchive!brand = rec!brand
        rsArchive!Seater = rec!Seater
        rsArchive!price = rec!price
        rsArchive.Update
    
        rsArchive.Close
        cnArchive.Close
    
        rec.Delete
        rec.Update
    
        RefreshGrid
        ClearInputFields
    
        MsgBox "Archived successfully!", vbInformation
    
    End Sub
    
    ' =========================================
    ' ================= Grid ====================
    ' =========================================
    Private Sub DataGrid1_Click()
    
        If rec Is Nothing Or rec.EOF Or rec.BOF Then Exit Sub
    
        On Error Resume Next
        rec.Bookmark = DataGrid1.Bookmark
        On Error GoTo 0
    
        PopulateTextBoxesFromGrid
    
        isEditing = True
        isAdding = False
        
    
    End Sub
    
    Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
    End Sub
    
    Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
        KeyAscii = 0
    End Sub
    
Private Sub lblbilllingMB_Click()
    Unload Me
    frmBilling.Show vbModal
End Sub

    Private Sub lblbookingMB_Click()
    Unload Me
    frmBooking.Show vbModal
    End Sub
    
    Private Sub lblcustomerMB_Click()
    
    Unload Me
    frmcustomer.Show vbModal
    End Sub
    
    ' =========================================
    ' ============ TextBox Validation ==========
    ' =========================================
    Private Sub txtplate_KeyPress(KeyAscii As Integer)
    
        Dim i As Integer
        i = Len(txtplate.Text) + 1
    
        If KeyAscii = 8 Then Exit Sub
    
        If i <= 3 Then
            If Not ((KeyAscii >= 65 And KeyAscii <= 90) _
            Or (KeyAscii >= 97 And KeyAscii <= 122)) Then
                KeyAscii = 0
                Exit Sub
            End If
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Else
            If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
    
        If i > 7 Then KeyAscii = 0
    
    End Sub
    
    Private Sub txtprice_KeyPress(KeyAscii As Integer)
    
        If KeyAscii = 8 Then Exit Sub
        If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
        If KeyAscii = 46 Then
            If InStr(txtprice.Text, ".") = 0 Then Exit Sub
        End If
    
        KeyAscii = 0
    
    End Sub
    
    Private Sub LoadVehiclesWithAvailableStatus(Optional ByVal keyword As String = "")
        Dim sql As String
    
        ' Only select vehicles with Status = "Available"
        sql = "SELECT * FROM vehicles WHERE TRIM(Status)='Available'"
    
        ' Apply search filter if keyword provided
        If Trim(keyword) <> "" Then
            sql = sql & " AND Plate LIKE '" & db.SafeText(keyword) & "%'"
        End If
    
        ' Open recordset
        If rec Is Nothing Then Set rec = New ADODB.Recordset
        If rec.State = adStateOpen Then rec.Close
    
        rec.CursorLocation = adUseClient
        rec.CursorType = adOpenStatic
        rec.LockType = adLockReadOnly
        rec.Open sql, db.con
    
        ' Bind to DataGrid
        Set DataGrid1.DataSource = rec
        HidecarIDColumn
        ResizeDataGridColumns
        DataGrid1.Refresh
    
        ' Clear selection
        ClearGridSelection
    End Sub
