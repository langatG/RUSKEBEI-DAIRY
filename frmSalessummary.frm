VERSION 5.00
Begin VB.Form frmSalessummary 
   Caption         =   "Sales summary"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraSales 
      Caption         =   "Sales"
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.CommandButton Command2 
         Caption         =   "Invoice"
         Height          =   495
         Left            =   2160
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cash sales summary"
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSalessummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
reportname = "cashsalessum.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""
End Sub

Private Sub Command2_Click()


reportname = "Invoicesalessum.rpt"
    Show_Sales_Crystal_Report STRFORMULA, reportname, ""

End Sub
