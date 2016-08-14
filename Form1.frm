VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   5685
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   960
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2160
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim w As Object
On Error GoTo err
Set w = CreateObject("wscript.shell")
Text1.Text = w.regread("HKEY_LOCAL_MACHINE\SOFTWARE\TopDomain\e-learning Class Standard\1.00\UninstallPasswd")
Exit Sub
err:
err.Clear
MsgBox "找不到这个值"
Exit Sub
End Sub
