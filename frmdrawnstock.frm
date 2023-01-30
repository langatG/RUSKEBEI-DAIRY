VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdrawnstock 
   Caption         =   "Drawn stock"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView Lvwdrawn 
      Height          =   3495
      Left            =   120
      TabIndex        =   28
      Top             =   4560
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Product code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Product name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "User name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Quantity"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Branch code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Drawnstock"
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin MSComCtl2.DTPicker DTPdateentered 
         Height          =   375
         Left            =   8640
         TabIndex        =   32
         Top             =   480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   108068865
         CurrentDate     =   42648
      End
      Begin VB.ComboBox Cbobranch 
         Height          =   315
         ItemData        =   "frmdrawnstock.frx":0000
         Left            =   1680
         List            =   "frmdrawnstock.frx":000D
         TabIndex        =   31
         Top             =   3720
         Width           =   2175
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   4440
         TabIndex        =   29
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtpcode 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtpname 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   1200
         Width           =   4335
      End
      Begin VB.TextBox txtdescription 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtquantity 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   8640
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtbalance 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   8640
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   17015
         Height          =   360
         Left            =   3600
         Picture         =   "frmdrawnstock.frx":0033
         ScaleHeight     =   360
         ScaleWidth      =   240
         TabIndex        =   13
         Top             =   600
         Width           =   240
      End
      Begin VB.Frame fra1 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   7320
         TabIndex        =   4
         Top             =   2520
         Width           =   4335
         Begin VB.PictureBox Picture1 
            Height          =   255
            Left            =   1320
            Picture         =   "frmdrawnstock.frx":01B5
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   8
            Top             =   600
            Width           =   255
         End
         Begin VB.PictureBox Picture3 
            Height          =   255
            Left            =   1320
            Picture         =   "frmdrawnstock.frx":0A7F
            ScaleHeight     =   195
            ScaleWidth      =   195
            TabIndex        =   7
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox txtdracc 
            Height          =   375
            Left            =   1680
            TabIndex        =   6
            Top             =   480
            Width           =   2535
         End
         Begin VB.TextBox txtcracc 
            Height          =   375
            Left            =   1680
            TabIndex        =   5
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label lbldracc 
            BackColor       =   &H8000000E&
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblcracc 
            BackColor       =   &H8000000E&
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "DrAccNo"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Craccno"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.TextBox txtpprice 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox txtsellingprice 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox txtprice 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   8640
         TabIndex        =   1
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Branch"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Product Code"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Product Name"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Description"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Quantity"
         Height          =   255
         Left            =   7560
         TabIndex        =   24
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Balance In Store"
         Height          =   255
         Left            =   7320
         TabIndex        =   23
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Date Entered"
         Height          =   255
         Left            =   7320
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Purchase Price "
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Selling Price "
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Price per goods"
         Height          =   255
         Left            =   7320
         TabIndex        =   19
         Top             =   2160
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmdrawnstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtpcode11_Change()
'//TWNG001
Provider = "MAZIWA"
Set cn = New ADODB.Connection
 cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierid,pprice, sprice from ag_products where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
 txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
If Not IsNull(rs.Fields(3)) Then txtBalance = (rs.Fields(3))
If Not IsNull(rs.Fields(4)) Then cbosupplier = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then txtpprice = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
If txtBalance <= 0 Then
MsgBox "Warning:Your stock is below zero please reorder", vbInformation
Else

End If
End If
'LstSearch.Refresh
 Lvwdrawn.ListItems.Clear
 'If chkshowallmembers = vbChecked Then
       Set rst = oSaccoMaster.GetRecordset("SELECT     DATE, DESCRIPTION, QUANTITY, TOTALAMOUNT, PRODUCTID, PRODUCTNAME, USERNAME, PRICEEACH, MONTH, YEAR, Branch, updated From DRAWNSTOCK where productid= '" & txtpcode & "'   order by productid")
    'End If
    With rst
        If Not .EOF Then
            While Not .EOF
                Set li = Lvwdrawn.ListItems.Add(, , !Date)
               ' li.SubItems(1) = IIf(IsNull(!productid), "", !productid)
                li.SubItems(1) = IIf(IsNull(!description), "", !description)
                li.SubItems(2) = IIf(IsNull(!PRODUCTID), "", !PRODUCTID)
                li.SubItems(3) = IIf(IsNull(!ProductName), "", !ProductName)
                li.SubItems(4) = IIf(IsNull(!username), "", !username)
                li.SubItems(5) = IIf(IsNull(!totalamount), "", !totalamount)
                li.SubItems(6) = IIf(IsNull(!PRICEEACH), "", !PRICEEACH)
                li.SubItems(7) = IIf(IsNull(!Quantity), "", !Quantity)
                li.SubItems(8) = IIf(IsNull(!Branch), "", !Branch)
                .MoveNext
            Wend
        End If
    End With
    TxtRecords = rst.RecordCount
'    cboSearchField.Text = cboSearchField.List(0)
'    cboCriteria.Text = cboCriteria.List(3)
'cmdMemberSearch_Click


'// check with serial no if it exist
End Sub


Private Sub cmdsave_Click()
Set rst = New Recordset
If lbldracc = "" Then MsgBox "select the account to Debit": Exit Sub

If lblcracc = "" Then MsgBox "select the account to credit": Exit Sub


'
Dim unsera As Integer
'Dim cn As Connection
If Trim(txtquantity) = "" Then
MsgBox "Quantity cannot be Zero", vbInformation
Exit Sub

End If
If Trim(txtBalance) = "" Then txtBalance = 0
'If chkserialrequired = vbChecked Then
'
'seria = 1
unsera = txtquantity

'// should only be one item
'If txtquantity > 1 Then
'MsgBox "Serialized items should only be added as one", vbCritical
'Exit Sub
'End If
'Else
'seria = 0
'unsera = 0
'End If

Provider = cn
Set cn = New ADODB.Connection
cn.Open Provider
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,qout,unserialized from ag_products where p_code='" & txtpcode & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If rs.EOF Then
'// insert into ag_products
'If txtserialno = "" Then txtserialno = 0
sql = ""
sql = "set dateformat dmy insert into  ag_products(p_code,p_name,s_no,qin,qout,date_entered,last_d_updated,user_id,audit_date,o_bal,supplierid,serialized,unserialized,seria,pprice,sprice )"
sql = sql & "  values('" & txtpcode.Text & "','" & txtpname.Text & "'," & txtSERIALNO.Text & "," & txtquantity.Text & "," & txtBalance.Text + txtquantity.Text & ",'" & DTPdateentered & "','" & DTPdateentered & "','Admin','" & Date & "'," & txtquantity.Text & ",'" & cbosupplier & "',0," & unsera & "," & seria & "," & txtpprice & "," & txtsellingprice & ")"
cn.Execute sql


If txtsellingprice = "" Then txtsellingprice = 0
If txtpprice = "" Then txtpprice = 0

'sql = ""
'sql = "set dateformat DMY INSERT INTO ag_stockbalance"
'sql = sql & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid,pprice,sprice,RLevel)"
'sql = sql & " VALUES     ('" & txtpcode.Text & "','" & txtpname & "', " & txtbalance & ", " & txtquantity & ", " & txtbalance.Text + txtquantity.Text & ", '" & txtdateenterered & "',1," & txtpprice & "," & txtsellingprice & "," & txtRLevel & ")"
'cn.Execute sql



Else
Dim d As Double
If Not IsNull(rs.Fields(2)) Then d = rs.Fields(2)
'sql = ""
'sql = "set dateformat DMY update ag_products set p_name='" & txtpname & "',qin=" & txtquantity.Text & ", qout= " & rs.Fields("qout") & ",o_bal=" & rs.Fields("qout") & ",last_d_updated='" & Date & "',user_id='" & User & "',audit_date='" & Date & "',unserialized=" & unsera - D & ",SERIA=" & unsera & ",pprice=" & txtpprice & ",sprice=" & txtsellingprice & " where p_code='" & txtpcode.Text & "'"
'cn.Execute sql

Dim rsst As Recordset
sql = ""
sql = "set dateformat DMY select top 1 * from ag_stockbalance where p_code='" & txtpcode & "' order by trackid desc "
Set rsst = New ADODB.Recordset
rsst.Open sql, cn
If Not rsst.EOF Then
'sql = ""
'sql = "set dateformat DMY INSERT INTO ag_stockbalance"
'sql = sql & " (p_code, productname, openningstock, changeinstock, stockbalance, transdate,companyid)"
'sql = sql & " VALUES     ('" & txtpcode & "', '" & txtpname & "', '" & txtbalance & "', '" & txtquantity & "', '" & txtquantity.Text - rs.Fields("qout") & "', '" & txtdateenterered & "',1)"
'cn.Execute sql
'Else
'sql = "Update ag_stockbalance"
'sql = sql & " SET              productname = '" & txtpname & "', openningstock = " & txtbalance & ", changeinstock = " & txtquantity & ", ag_stockbalance = " & txtquantity.Text + rs.Fields("qout") & ", transdate = '" & txtdateenterered & "'"
'sql = sql & " WHERE     (p_code = '" & txtpcode & "') AND trackid=" & rsst.Fields("trackid") & ""
'cn.Execute sql
End If
'// update serialno database

'' ///update gl


End If
If seria = 1 Then
Set rst = Nothing
    sql = ""
   sql = "select * from serialno where serialno='" & txtSERIALNO & "' AND P_CODE='" & txtpcode & "' and used=0"
   Set rst = New ADODB.Recordset
   rst.Open sql, cn, adOpenKeyset, adLockOptimistic

If rst.EOF Then
sql = ""
sql = "set dateformat DMY INSERT INTO serialno(serialno,p_code,used)"
sql = sql & " values('" & txtSERIALNO & "','" & txtpcode & "',0)"
cn.Execute sql
Else
MsgBox "Item is in place and not yet used", vbInformation
Exit Sub
End If
End If
sql = ""
sql = "Set dateformat dmy INSERT INTO DRAWNSTOCK(DATE, DESCRIPTION, QUANTITY, TOTALAMOUNT, PRODUCTID, PRODUCTNAME, USERNAME, PRICEEACH, MONTH, YEAR, Branch) VALUES     ('" & DTPdateentered & "','" & txtdescription & "', '" & txtquantity & "', '" & txtPrice * txtquantity & "', '" & txtpcode & "', '" & txtpname & "', '" & User & "', '" & txtPrice & "', '" & month(txtdateenterered) & "','" & year(txtdateenterered) & "','" & cboBranch & "')"
oSaccoMaster.ExecuteThis (sql)
sql = ""
sql = "set dateformat dmy insert into gltransactions(transdate,amount,draccno,craccno,documentno,source,transdescript,auditid,cash,doc_posted) values('" & DTPdateentered & "'," & txtquantity & " *" & txtpprice & ",'" & lbldracc & "','" & lblcracc & "','stock intake','" & cbosupplier & "' ,'stock intake','" & User & "',0,0)"
oSaccoMaster.ExecuteThis (sql)

txtBalance = ""
txtpcode = ""
txtpname = ""
txtSERIALNO = ""
txtquantity = ""
txtpprice = ""
txtsellingprice = ""
cbosupplier = ""
txtdescription = ""
cboBranch = ""
txtPrice = ""
MsgBox "Record Saved Successfully"
End Sub

Private Sub lblcracc_Change()
Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lblcracc & "'")
    If Not rst.EOF Then
    txtcracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub lblcracc_Click()
Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lblcracc & "'")
    If Not rst.EOF Then
    txtcracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub lbldracc_Change()
Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lbldracc & "'")
    If Not rst.EOF Then
    txtdracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub lbldracc_Click()
Set rst = oSaccoMaster.GetRecordset("select glaccname from glsetup where accno='" & lbldracc & "'")
    If Not rst.EOF Then
    txtdracc = rst.Fields("glaccname")
    End If
End Sub

Private Sub Picture1_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lbldracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub Picture2_Click()
frmSearch.Show vbModal
Dim Y As String
Y = sel

If Y <> "" Then

Provider = "MAZIWA"

Set cn = New ADODB.Connection
 cn.Open Provider, "atm", "atm"
'If reportpath = "" Then reportpath = GetSetting("payroll", "AppName", "rptPath", rptPath)
sql = "select P_CODE,P_NAME,S_NO,QOUT,supplierID,pprice,sprice from ag_products where p_code='" & Y & "'"
Set rs = New ADODB.Recordset
rs.Open sql, cn
If Not rs.EOF Then
If Not IsNull(rs.Fields(0)) Then txtpcode = (rs.Fields(0))
If Not IsNull(rs.Fields(1)) Then txtpname = (rs.Fields(1))
If Not IsNull(rs.Fields(4)) Then cbosupplier = (rs.Fields(4))
If Not IsNull(rs.Fields(5)) Then txtpprice = (rs.Fields(5))
If Not IsNull(rs.Fields(6)) Then txtsellingprice = (rs.Fields(6))
'If Not IsNull(rs.Fields(2)) Then txtserialno = (rs.Fields(2))
If Not IsNull(rs.Fields(3)) Then txtBalance = (rs.Fields(3))

If txtBalance <= 0 Then
MsgBox "Your stock is below zero please reorder", vbInformation
End If
'// check with serial no if it exist


End If
End If
End Sub

Private Sub Picture3_Click()
Me.MousePointer = vbHourglass
        frmSearchGLAccounts.Show vbModal, Me
    If Continue Then
        If SearchValue <> "" Then
            lblcracc = SearchValue
            SearchValue = ""
        End If
    End If
    Me.MousePointer = 0
End Sub

Private Sub txtpcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtpcode11_Change
Else
Exit Sub
End If
End Sub

Private Sub txtserialno_Change()

End Sub
