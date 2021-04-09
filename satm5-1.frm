VERSION 5.00
Begin VB.Form frmSa 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Исходные данные"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   3990
   Begin VB.TextBox TxtHk 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox TxtHpr 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox TxtHn 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton CmdTbl 
      Caption         =   "&Table"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MousePointer    =   12  'No Drop
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MousePointer    =   12  'No Drop
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton CmdHlp 
      Caption         =   "&Help"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MousePointer    =   12  'No Drop
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton CmdCnl 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MousePointer    =   12  'No Drop
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Конечная высота(м)"
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   2040
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Приращение высоты(м)"
      Height          =   195
      Left            =   1440
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Начальная высота(м)"
      Height          =   195
      Left            =   1440
      TabIndex        =   7
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "frmSa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim gsngSa() As Single
Public KeyAscii As Integer
Public gstrP_sa As String

Private Sub CmdCnl_Click()
End
End Sub

Private Sub CmdHlp_Click()
Dim msg, style
msg = "1. Допустимый диапазон высот от 0 до 15000 м"
msg = msg + vbCrLf + "2. Значение начальной высоты должно быть меньше значения конечной высоты"
msg = msg + vbCrLf + "3. Приращение высоты должно быть меньше значения конечной высоты"
msg = msg + vbCrLf + "4. Для запуска - нажать клавишу Enter"
style = vbOKOnly + vbDefaultButton1 + vbCritical
MsgBox msg, style, "Ввод исходных данных"
End Sub

Private Sub CmdOk_Click()
Rem Описание локальных переменных
Dim msg, style
Dim TextErVvod As Single
Dim gsngH_n As Single 'Начальное значение высоты
Dim gsngH_pr As Single 'Приращение значения высоты
Dim gsngH_k As Single 'Конечное значение высоты
Dim j As Integer
Dim k As Integer
Dim m As Integer
Dim i As Integer
Rem Присвоение и ввод переменных
gsngH_n = Val(TxtHn.Text)
gsngH_k = Val(TxtHk.Text)
gsngH_pr = Val(TxtHpr.Text)
Select Case gsngH_pr
Case Is >= gsngH_k
MsgBox "1.Значение конечной высоты должно быть больше " + "приращение высоты"
Case Is = 0
MsgBox "1.Значение приращения высоты должно быть больше 0" + vbCrLf + "2.Повторите ввод"
End Select
Select Case gsngH_k
Case Is <= gsngH_n
MsgBox "1.Значение конечной высоты должно быть больше 0 и меньше 15000" + vbCrLf + "2.Повторите ввод"
TxtHn.Text = "" 'Ввод пустой строки в поле ввода
TxtHk.Text = ""
TxtHpr.Text = ""
GoTo M1 'переход к концу процедуры Private Sub CmdOk_Click()
End Select

j = 1
If gsngH_n < 0 Or gsngH_n > 15000 Or gsngH_k < 0 Or gsngH_k > 15000 Then
Rem ввод сообщения об ошибке ввода высоты
TextErVvod = ErVvod_1()
Else
If gsngH_k <= gsngH_n Then
MsgBox "1.Значение конечной высоты меньше начальной"
TxtHn.Text = "" 'Ввод пустой строки в поле ввода
TxtHk.Text = ""
TxtHpr.Text = ""
Else
i = (gsngH_k - gsngH_n) / gsngH_pr + 1
ReDim gsngSa(5, i + 1)
gsngSa(1, j) = gsngH_n 'Начальное значение высоты присваивается элементу с индексом 1.
Rem Расчёт параметров атмосферы
Do While gsngSa(1, j) <= gsngH_k
If gsngSa(1, j) <= 11000 Then Call Sa_Tropo(j)
If gsngSa(1, j) > 11000 Then Call Sa_Strato(j)

Rem Расчёт параметров цикла
gsngSa(1, j + 1) = gsngSa(1, j) + gsngH_pr 'Определение высоты для нового цикла
j = j + 1 'Увелечение счёчиков количества циклов
Loop
End If
TxtHn.Text = ""
TxtHk.Text = ""
TxtHpr.Text = ""
Rem Формирование ввода результатов расчёта
For m = 1 To i
For k = 1 To 5
gstrP_sa = gstrP_sa + Format(gsngSa(k, m), "Scientific") + Space$(5)
Next k
gstrP_sa = gstrP_sa + vbCrLf
Next m
gstrP_sa = "Высота  " + " Давление" + " Плотность  " + "Температура " + " Cкорость звука" + vbCrLf + gstrP_sa
CmdTbl.Enabled = True
End If
M1:
End Sub

Private Sub CmdTbl_Click()
frmTabl.Show
frmSa.CmdTbl.Enabled = False
End Sub

Public Sub Sa_Tropo(j As Integer)
Const sngT0 As Single = 288.15 'Температура у земли
Const sngR0 As Single = 1.225 'Плотность у земли
Const sngP0 As Single = 101325 'Давление у земли
Dim sngGradTempH As Single
sngGradTempH = -0.0065
gsngSa(2, j) = sngP0 * (1 - gsngSa(1, j) / 44300) ^ 5.256
gsngSa(3, j) = sngR0 * (1 - gsngSa(1, j) / 44300) ^ 4.256
gsngSa(4, j) = sngT0 + sngGradTempH * (gsngSa(1, j))
gsngSa(5, j) = 20 * (gsngSa(4, j) ^ (1 / 2))
End Sub

Public Sub Sa_Strato(j As Integer)
Const sngCfnP As Single = 133.322 'Коэффицент перевода давления
Const sngCfnR As Single = 9.0665 'Коэффицент перевода плотности
Const sngTh11000 As Single = 216.7 'Температура на высоте 11000 м
Dim sngStpn As Single
sngStpn = -(gsngSa(1, j) - 11000) / 6340
gsngSa(2, j) = 169.4 * Exp(sngStpn) * sngCfnP
gsngSa(3, j) = 20 * (sngTh11000 ^ (1 / 2))
gsngSa(4, j) = sngTh11000
gsngSa(5, j) = 0.037 * Exp(sngStpn) * sngCfnR
End Sub

Private Function ErVvod_1()
MsgBox "1.Значение высоты или приращения выходят за допустимые пределы" + vbCrLf + "2.Введите значение высоты от 0 до 15000 метров"
End Function

Private Function Prov(KeyAscii)
Static DecPoint As Integer
Select Case KeyAscii
Case Asc("0") To Asc("9")
Case Asc(".")
If DecPoint Then
KeyAscii = 0: Beep
Else
DecPoint = True
End If
Case Asc(",")
KeyAscii = 0
Case Else
KeyAscii = 0: Beep
End Select
Prov = KeyAscii
End Function

Private Sub TxtHk_KeyPress(KeyAscii As Integer)
Dim Simv
Simv = Prov(KeyAscii)
End Sub

Private Sub TxtHn_KeyPress(KeyAscii As Integer)
Dim Simv
Simv = Prov(KeyAscii)
End Sub

Private Sub TxtHpr_KeyPress(KeyAscii As Integer)
Dim Simv
Simv = Prov(KeyAscii)
End Sub

