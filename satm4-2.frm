VERSION 5.00
Begin VB.Form frmTabl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Параметры стандартной атмосферы"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9435
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9435
   Begin VB.TextBox TxtTbSa 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   240
      Width           =   9015
   End
   Begin VB.CommandButton CmdExt 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   5520
      Width           =   1695
   End
End
Attribute VB_Name = "frmTabl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub CmdExt_Click()
frmSa.Show
Me.Hide
End Sub

Private Sub Form_Activate()
frmTabl.TxtTbSa.Text = frmSa.gstrP_sa
End Sub

