VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "����λ"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "����˵��"
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "һ������λ"
      Height          =   735
      Left            =   720
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim xlapp As Excel.Application 'Excel����
Dim xlbook As Excel.Workbook '������
Dim xlsheet As Excel.Worksheet '������
Set xlapp = CreateObject("Excel.Application")
Set xlbook = xlapp.Workbooks.Add
xlapp.Visible = True
Set xlsheet = xlbook.Worksheets(1)
Dim a(51) As String, i As Integer, j As Integer

a(1) = "10���@"
a(2) = "11��׿Ȼ"
a(3) = "1�����"
a(4) = "2����ɯ"
a(5) = "12���"
a(6) = "13�"
a(7) = "3����ͥ"
a(8) = "4ʷ����"
a(9) = "15��º�"
a(10) = "16���ܲ"
a(11) = "5������"
a(12) = "6ׯ���"
a(13) = "17�ε���"
a(14) = "18��о�"
a(15) = "7����֮"
a(16) = "8����"
a(17) = "21���"
a(18) = "23������"
a(19) = "9������"
a(20) = "47л����"
a(21) = "24�żһ�"
a(22) = "26�����"
a(23) = "19����"
a(24) = "20�����"
a(25) = "27��Դ��"
a(26) = "29�����"
a(27) = "22����ܰ"
a(28) = "25������"
a(29) = "31�׷���"
a(30) = "35������"
a(31) = "28���"
a(32) = "30������"
a(33) = "36��Զ��"
a(34) = "38��չ��"
a(35) = "32�ܻ���"
a(36) = "33֣ٳ�"
a(37) = "40������"
a(38) = "41����Դ"
a(39) = "49����˼"
a(40) = "34֣����"
a(41) = "42�����"
a(42) = "44���"
a(43) = "37����ӱ"
a(44) = "39������"
a(45) = "46�����"
a(46) = "48��׿�"
a(47) = "43Խʫ�"
a(48) = "45������"
a(49) = "50κΡ"
a(50) = ""

Dim k As Integer
Dim flag As Boolean, xb As Boolean, ty As Boolean
    Randomize
    k = Int(Rnd * (24)) + 1
    If k <> 25 Then
    k = 4 * Int(Rnd * (11)) + Int(Rnd) + 1
    a(50) = a(k)
    a(k) = a(49)
    a(49) = a(50)
    End If
For j = 1 To 2
    For i = 0 To 47
        flag = False
        xb = False
        ty = False
        k = Int(Rnd * 47)
        If i Mod 4 >= 2 Then
        flag = True
        End If
        If k Mod 4 >= 2 Then
        xb = True
        End If
        If xb <> flag Then
        i = i - 1
        ty = True
        End If
        If ty = False Then
        a(50) = a(i + 1)
        a(i + 1) = a(k + 1)
        a(k + 1) = a(50)
        End If
    Next i
Next j
    For i = 0 To 7
    xlsheet.Cells(Int(i / 8) + 2, (i Mod 8) + 4) = a(i + 1)
    xlsheet.Cells(Int(i / 8) + 4, (i Mod 8) + 4) = a(i + 1 + 16)
    xlsheet.Cells(Int(i / 8) + 6, (i Mod 8) + 4) = a(i + 1 + 32)
    xlsheet.Cells(Int(i / 8) + 3, 11 - (i Mod 8)) = a(i + 1 + 8)
    xlsheet.Cells(Int(i / 8) + 5, 11 - (i Mod 8)) = a(i + 1 + 24)
    xlsheet.Cells(Int(i / 8) + 7, 11 - (i Mod 8)) = a(i + 1 + 40)
    Next
xlsheet.Cells(1, 8) = a(49)
xlsheet.Cells(8, 7) = "��̨"
With xlsheet
End With
End Sub

Private Sub Command2_Click()
Dim a As Long
a = MsgBox("3.0�汾�����и��๦�ܴ�����", vbOKOnly, "����")
End Sub
