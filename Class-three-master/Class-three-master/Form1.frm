VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "排座位"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "更新说明"
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "一键排座位"
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
Dim xlapp As Excel.Application 'Excel对象
Dim xlbook As Excel.Workbook '工作簿
Dim xlsheet As Excel.Worksheet '工作表
Set xlapp = CreateObject("Excel.Application")
Set xlbook = xlapp.Workbooks.Add
xlapp.Visible = True
Set xlsheet = xlbook.Worksheets(1)
Dim a(51) As String, i As Integer, j As Integer

a(1) = "10李丰旲"
a(2) = "11李卓然"
a(3) = "1马天慧"
a(4) = "2王莉莎"
a(5) = "12李健雄"
a(6) = "13李骏"
a(7) = "3王楚庭"
a(8) = "4史宝琳"
a(9) = "15李德浩"
a(10) = "16杨晋懿"
a(11) = "5朱沁秋"
a(12) = "6庄礼嘉"
a(13) = "17何狄其"
a(14) = "18余承峻"
a(15) = "7刘逸之"
a(16) = "8汤妮"
a(17) = "21张宇辰"
a(18) = "23张烨兴"
a(19) = "9苏文琪"
a(20) = "47谢安安"
a(21) = "24张家华"
a(22) = "26陈宇睿"
a(23) = "19宋悦"
a(24) = "20张竹歆"
a(25) = "27陈源广"
a(26) = "29陈颢文"
a(27) = "22张怡馨"
a(28) = "25陈子期"
a(29) = "31易方博"
a(30) = "35钟蓝海"
a(31) = "28陈旖"
a(32) = "30邵宗琪"
a(33) = "36洪远境"
a(34) = "38郭展威"
a(35) = "32周慧卿"
a(36) = "33郑俪淇"
a(37) = "40崔文昱"
a(38) = "41康子源"
a(39) = "49黎若思"
a(40) = "34郑嘉怡"
a(41) = "42梁桢睿"
a(42) = "44程昊"
a(43) = "37贺文颖"
a(44) = "39黄子茵"
a(45) = "46舒马赫"
a(46) = "48雷卓宸"
a(47) = "43越诗瑜"
a(48) = "45傅惠玲"
a(49) = "50魏巍"
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
xlsheet.Cells(8, 7) = "讲台"
With xlsheet
End With
End Sub

Private Sub Command2_Click()
Dim a As Long
a = MsgBox("3.0版本，还有更多功能待开发", vbOKOnly, "帮助")
End Sub
