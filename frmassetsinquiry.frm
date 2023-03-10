VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmassetsinquiry 
   Caption         =   "AC-Assets Inquiry"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8985
   Icon            =   "frmassetsinquiry.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdinqure 
      Caption         =   "Assets Inquiry"
      Height          =   375
      Left            =   480
      TabIndex        =   17
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdprocess 
      Caption         =   "Process Depreciation"
      Height          =   435
      Left            =   6120
      TabIndex        =   11
      Top             =   8040
      Width           =   1815
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "Load"
      Height          =   435
      Left            =   2160
      TabIndex        =   10
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdfinder 
      Height          =   285
      Left            =   3720
      Picture         =   "frmassetsinquiry.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Add New record"
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtNETREALIABLEVALUE 
      Height          =   285
      Left            =   6720
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin MSComctlLib.ListView Lvwassetsinquiry 
      Height          =   6495
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   11456
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtSERIALNO 
      Appearance      =   0  'Flat
      DataField       =   "assetserialno"
      Height          =   285
      Index           =   1
      Left            =   1980
      TabIndex        =   5
      Top             =   555
      Width           =   3975
   End
   Begin VB.TextBox txtASSETSNAME 
      Appearance      =   0  'Flat
      DataField       =   "assetsname"
      Height          =   285
      Index           =   1
      Left            =   1980
      TabIndex        =   1
      Top             =   870
      Width           =   3975
   End
   Begin VB.TextBox txtASSETSNO 
      DataField       =   "assetsno"
      Height          =   285
      Index           =   1
      Left            =   1980
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   8160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy"
      Format          =   108265475
      CurrentDate     =   40748
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   8160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMM"
      Format          =   108265475
      CurrentDate     =   40748
   End
   Begin MSComctlLib.ListView lvwasset 
      Height          =   6375
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   11245
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Asset Code"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Asset Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "AssetType"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Depreciation Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "GL Account No"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "PDate"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Purchase Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Current Value"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "SerialNo"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Month"
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   7920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Year"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   7920
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "NET REALISABLE VALUE"
      Height          =   255
      Left            =   6720
      TabIndex        =   7
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Asset Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   390
      TabIndex        =   4
      Top             =   900
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "RegNo/Serial No.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   390
      TabIndex        =   3
      Top             =   570
      Width           =   1590
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Asset No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   390
      TabIndex        =   2
      Top             =   240
      Width           =   840
   End
End
Attribute VB_Name = "frmassetsinquiry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As Connection

Private Sub cmdfinder_Click()
On Error Resume Next
frmsearchassets.Show vbModal
Dim Y As String
Y = sel
'm = False
If Y <> "" Then
     Dim cn As Connection
    Set cn = New ADODB.Connection
    
    cn.Open frmODBCLogon.cboDSNList, "atm", "atm"
sql = ""
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "SELECT     AssetsNo, AssetserialNo, AssetsName from assets where assetsno='" & Y & "' order by AssetsNo"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then

If Not IsNull(rs.Fields(0)) Then txtASSETSNO(1) = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtSERIALNO(1) = (rs.Fields(1))
If Not IsNull(rs.Fields(2)) Then txtASSETSNAME(1) = (rs.Fields(2))


sql = "Select * from assetstrans where assetcode='" & txtASSETSNO(1) & "'"
rs.Open sql, cn
If rs.EOF Then
'MsgBox "No data in the such data available"
Exit Sub
Else
    'If .RecordCount > 0 Then
    txtNETREALIABLEVALUE = rs!nrv
    'txtserialno(1) = rs!assetserialno
    End If
'Call cboname_p

End If
End If
End Sub

Private Sub cmdinqure_Click()
Dim lis As ListItem
Lvwassetsinquiry.ListItems.Clear
lvwasset.ListItems.Clear
lvwasset.Visible = False
Lvwassetsinquiry.Visible = True
Set rs = New Recordset
sql = "Select * from assets "
Set rs = oSaccoMaster.GetRecordset(sql)
If rs.EOF Then
'MsgBox "No data in the such data available"
Exit Sub
Else
With rs
        Do While Not .EOF
            Set lis = Lvwassetsinquiry.ListItems.Add(, , !AssetsNo & "")
            lis.ListSubItems.Add , , !assetserialno & ""
            lis.ListSubItems.Add , , !assetsname & ""
            lis.ListSubItems.Add , , !assettype & ""
            lis.ListSubItems.Add , , !datebought & ""
            lis.ListSubItems.Add , , !Unitno & ""
            lis.ListSubItems.Add , , !PurchasePrice & ""
            lis.ListSubItems.Add , , !depreciation & ""
            lis.ListSubItems.Add , , !CurrentValue & ""
            lis.ListSubItems.Add , , !notes & ""
            lis.ListSubItems.Add , , !transdate & ""
           
            
            .MoveNext
        Loop
End With
End If
Set rs = Nothing
End Sub

Private Sub cmdLoad_Click()
Dim rsassets As New Recordset
Lvwassetsinquiry.Visible = False
lvwasset.Visible = True
lvwasset.ListItems.Clear
sql = ""
sql = "select * from assets order by assetsid asc"
Set rsassets = oSaccoMaster.GetRecordset(sql)
With rsassets
    While Not .EOF
        Set li = lvwasset.ListItems.Add(, , IIf(IsNull(!AssetsNo), "", !AssetsNo))
        li.SubItems(1) = IIf(IsNull(!assetsname), "", !assetsname)
        li.SubItems(2) = IIf(IsNull(!assettype), "", !assettype)
        li.SubItems(3) = IIf(IsNull(!depreciation), 0, !depreciation)
        li.SubItems(4) = IIf(IsNull(!ACCNO), "", !ACCNO)
        li.SubItems(5) = IIf(IsNull(!datebought), "", !datebought)
        li.SubItems(6) = IIf(IsNull(!PurchasePrice), "", !PurchasePrice)
        li.SubItems(7) = IIf(IsNull(!CurrentValue), "", !CurrentValue)
        li.SubItems(8) = IIf(IsNull(!assetserialno), "", !assetserialno)
    .MoveNext
    Wend
End With
End Sub

Private Sub cmdprocess_Click()
Dim code As String
Dim name As String
Dim deprate, depam As Double
Dim Pprice As Double
Dim transdate As Date, amount As Double, DRaccno As String, Craccno As String, DocumentNo As String
Dim TransSource As String, User1 As String, ErrorMessage As String, transDescription As String, CashBook As Long, doc_posted As Integer, chequeno As String
Set rs = New Recordset

sql = ""
sql = "select  * from depreciation where mmonth=" & month(DTPicker1) & " and yyear=" & year(DTPicker2) & ""
Set rs2 = oSaccoMaster.GetRecordset(sql)
If Not rs2.EOF Then
    MsgBox "Depreciation for the period you choose has been processed", vbInformation
Else
  
    Startdate = DateSerial(year(DTPicker2), month(DTPicker1), 1)
    Enddate = DateSerial(year(DTPicker2), month(DTPicker1) + 1, 1 - 1)
    transdate = Format(Get_Server_Date, "dd/mm/yyyy")
    
    If transdate < Enddate Then MsgBox "You have not reached End month", vbInformation: Exit Sub


 For I = 1 To lvwasset.ListItems.Count
        code = lvwasset.ListItems(I).Text
        Craccno = lvwasset.ListItems(I).SubItems(4)
        amount = lvwasset.ListItems(I).SubItems(7)
        
        If amount > 0 Then
         deprate = lvwasset.ListItems(I).SubItems(3)
         depam = amount * deprate / 100 * 1 / 12
         depam = Round(depam, 2)
        End If

        If depam > 0 Then
                sql = ""
                sql = "INSERT INTO Depreciation  (AssetCode, mmonth, yyear, DepreciationAmt, uuser)"
                sql = sql & " VALUES     ('" & lvwasset.ListItems(I).Text & "', " & month(DTPicker1) & ", " & year(DTPicker2) & ", " & depam & ", '" & User & "')"
                oSaccoMaster.GetRecordset (sql)
                
                '*** Update Fixed Asset current value ***********
                sql = "Update assets Set CurrentValue= CurrentValue-" & depam & "  where AssetsNo='" & lvwasset.ListItems(I).Text & "'"
                oSaccoMaster.GetRecordset (sql)
                   
                    DRaccno = Trim(20409)
                    Craccno = Trim(Craccno)
                    DocumentNo = code
                          TransSource = code
                          User1 = User
                          transDescription = "Asset Depreciation"
                          CashBook = 1
                          doc_posted = 1
                          chequeno = code
                    If Not Save_GLTRANSACTION(Enddate, depam, DRaccno, Craccno, DocumentNo, _
                          TransSource, User1, ErrorMessage, transDescription, CashBook, doc_posted, chequeno, transactionNo) Then
                              If ErrorMessage <> "" Then
                                  MsgBox ErrorMessage, vbInformation, Me.Caption
                                  ErrorMessage = ""
                              End If
                          End If
                    End If
                
                
                
        amount = 0
        depam = 0
        
        
    Next I
    MsgBox "Depreciation processing  for the period Complete", vbInformation
    cmdLoad_Click
End If
    
End Sub

Private Sub Form_Load()
Set cn = New ADODB.Connection
'Set CSecurity = New PAYROLLSYSTEMS.CSecurity


With Lvwassetsinquiry
    .ColumnHeaders.Add , , "Asset No", 1000
    .ColumnHeaders.Add , , "Asset Serial No", 3000
    .ColumnHeaders.Add , , "Asset Name", 4000
    .ColumnHeaders.Add , , "Asset Type", 4000
    .ColumnHeaders.Add , , "Date Bought", 1500
    .ColumnHeaders.Add , , "Units", 1500
    .ColumnHeaders.Add , , "Purchase Price", 1500
    .ColumnHeaders.Add , , "Depreciation", 1000
    .ColumnHeaders.Add , , "Current Value", 1500
    .ColumnHeaders.Add , , "Notes", 1500
    
    .View = lvwReport
End With
End Sub

Private Sub txtASSETSNO_Change(index As Integer)

Dim lis As ListItem
Lvwassetsinquiry.ListItems.Clear

Set rs = New Recordset
Set cn = New Connection
cn.Open frmODBCLogon.cboDSNList, "atm", "atm"
Dim rst As Recordset
Set rst = New ADODB.Recordset
sql = ""
sql = "SELECT     * FROM         assetstrans where assetcode='" & txtASSETSNO(1) & "'"
rst.Open sql, cn, adOpenKeyset, adLockOptimistic
If Not rst.EOF Then
If Not IsNull(rst.Fields("NRV")) Then txtNETREALIABLEVALUE = rst.Fields("NRV")
End If

sql = "Select * from assets where assetsno='" & txtASSETSNO(1) & "'"
rs.Open sql, cn
If rs.EOF Then
'MsgBox "No data in the such data available"
Exit Sub
Else
With rs
    'If .RecordCount > 0 Then
    txtASSETSNAME(1) = !assetsname
    txtSERIALNO(1) = !assetserialno
        '.MoveFirst
        Do While Not .EOF
            Set lis = Lvwassetsinquiry.ListItems.Add(, , !AssetsNo & "")
            lis.ListSubItems.Add , , !assetserialno & ""
            lis.ListSubItems.Add , , !assetsname & ""
            lis.ListSubItems.Add , , !assettype & ""
            lis.ListSubItems.Add , , !datebought & ""
            lis.ListSubItems.Add , , !Unitno & ""
            lis.ListSubItems.Add , , !PurchasePrice & ""
            lis.ListSubItems.Add , , !depreciation & ""
            lis.ListSubItems.Add , , !CurrentValue & ""
            lis.ListSubItems.Add , , !notes & ""
            lis.ListSubItems.Add , , !transdate & ""
            'PamSet = True
            
            .MoveNext
        Loop
        .MoveFirst
   ' End If
End With
End If
Set rs = Nothing
End Sub
