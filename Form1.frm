VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "迅雷挤号器"
   ClientHeight    =   8385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   5400
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "状态"
      Height          =   3735
      Left            =   120
      TabIndex        =   16
      Top             =   4320
      Width           =   5055
      Begin VB.ListBox lstStatus 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3390
         ItemData        =   "Form1.frx":030A
         Left            =   120
         List            =   "Form1.frx":0311
         TabIndex        =   17
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      Picture         =   "Form1.frx":0320
      ScaleHeight     =   945
      ScaleWidth      =   5265
      TabIndex        =   9
      Top             =   0
      Width           =   5295
      Begin VB.Label lblMinify 
         BackStyle       =   0  'Transparent
         Caption         =   "━"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblClose 
         BackStyle       =   0  'Transparent
         Caption         =   "╳"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4920
         TabIndex        =   12
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "迅雷挤号器"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "当前账号"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   5055
      Begin VB.ListBox lstAccout 
         Appearance      =   0  'Flat
         Height          =   1290
         ItemData        =   "Form1.frx":6D9B
         Left            =   120
         List            =   "Form1.frx":6DA2
         TabIndex        =   18
         Top             =   840
         Width           =   4815
      End
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "填入密码"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "填入账号"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "号"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "密"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "被挤次数："
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1455
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "次"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   4000
      Left            =   600
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   120
      Top             =   840
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "Form1.frx":6DB1
      ScaleHeight     =   345
      ScaleWidth      =   5265
      TabIndex        =   10
      Top             =   8040
      Width           =   5295
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "QQ：251121753"
         BeginProperty Font 
            Name            =   "楷体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   75
         Width           =   1695
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   1200
      Top             =   840
   End
   Begin VB.Label lblNewMovies 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "最新电影"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      ToolTipText     =   "单击的访问 http://www.dytt8.net（电影天堂)"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Index           =   1
      X1              =   20
      X2              =   20
      Y1              =   960
      Y2              =   8040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      Index           =   0
      X1              =   5260
      X2              =   5260
      Y1              =   960
      Y2              =   8040
   End
   Begin VB.Label lblFreeAccount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "免费账号   "
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   315
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "单击的访问智能提取，双击手动获取"
      Top             =   1320
      Width           =   1365
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private hwdThunder As Long
Private wshshell
Private mMainPageCode As String, mDetailPageCode As String, mItemCode As String
Private mLogPath As String

Public web As New webCode
Public file As New CFile

Private Sub Form_Load()
  Set wshshell = CreateObject("WScript.Shell")
  mLogPath = App.Path & "\log.txt"
  web.TimeOut = 40
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub lblFreeAccount_Click()
  mMainPageCode = ""
  mDetailPageCode = ""
  mItemCode = ""
  lstStatus.Clear
  lstAccout.Clear
  lstStatus.AddItem "正在打开www.521xunlei.com(爱密码迅雷帐号分享平台)并获取源码  " & Time
  mMainPageCode = web.GetHTMLCode("http://www.521xunlei.com")
  If file.FileExists(mLogPath) Then
    file.Delete mLogPath
  End If
  file.WriteLineToTextFile mLogPath, "[Page1]"
  file.WriteLineToTextFile mLogPath, mMainPageCode
  buzhou1
End Sub


Private Sub lblFreeAccount_DblClick()
  ShellExecute Me.hwnd, "open", "http://www.521xunlei.com/", "", "", 5
End Sub


Private Sub lblClose_Click()
  End
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblClose.ForeColor = vbRed
End Sub

Private Sub lblMinify_Click()
  Me.WindowState = 1
End Sub

Private Sub lblMinify_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblMinify.ForeColor = vbGreen
End Sub

Private Sub lblNewMovies_Click()
  ShellExecute Me.hwnd, "open", "http://www.dytt8.net//", "", "", 5
End Sub

Private Sub lstAccout_Click()
  Dim strz As String
  Dim i As Integer, i2 As Integer
  Dim str1 As String, str2 As String
  strz = lstAccout.List(lstAccout.ListIndex)
  For i = 1 To Len(strz)
    If Asc(Mid(strz, i, 1)) >= 48 And Asc(Mid(strz, i, 1)) <= 122 Then
      str1 = str1 & Mid(strz, i, 1)
      i2 = i
      On Error GoTo ExitSub
      If Asc(Mid(strz, i + 1, 1)) < 48 Or Asc(Mid(strz, i + 1, 1)) > 122 Then
        Exit For
      End If
    End If
  Next i
  For i = i2 + 1 To Len(strz)
    If Asc(Mid(strz, i, 1)) >= 48 And Asc(Mid(strz, i, 1)) <= 122 Then
      str2 = str2 & Mid(strz, i, 1)
    End If
  Next i
  txtNo.Text = Trim(str1)
  txtPass.Text = Trim(str2)
ExitSub:
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static FormX As Long, FormY As Long
  If Button = 1 Then
    Me.Move Me.Left - FormX + X, Me.Top - FormY + Y
  ElseIf Button = 0 Then
    FormX = X
    FormY = Y
  End If
  lblClose.ForeColor = vbWhite
  lblMinify.ForeColor = vbWhite
End Sub

Private Sub Text1_Change()
  Timer1.Enabled = False
  Timer2.Enabled = True
  Timer3.Enabled = False
End Sub

Private Sub txtNo_Click()
  Dim strz As String
  Dim i As Integer, i2 As Integer
  Dim str1 As String, str2 As String
  strz = Clipboard.GetText(vbCFText)
  For i = 1 To Len(strz)
    If Asc(Mid(strz, i, 1)) >= 48 And Asc(Mid(strz, i, 1)) <= 122 Then
      str1 = str1 & Mid(strz, i, 1)
      i2 = i
      On Error GoTo ExitSub
      If Asc(Mid(strz, i + 1, 1)) < 48 Or Asc(Mid(strz, i + 1, 1)) > 122 Then
        Exit For
      End If
    End If
  Next i
  For i = i2 + 1 To Len(strz)
    If Asc(Mid(strz, i, 1)) >= 48 And Asc(Mid(strz, i, 1)) <= 122 Then
      str2 = str2 & Mid(strz, i, 1)
    End If
  Next i
  txtNo.Text = str1
  txtPass.Text = str2
ExitSub:
End Sub

Private Sub Timer1_Timer()
  Static ci As Integer
  Static zong As Integer
  hwdThunder = FindWindow(vbNullString, "迅雷7登录") '得到登录界面句柄
  If hwdThunder <> 0 Then
    SetForegroundWindow hwdThunder  '让登录界面切换到前台
    Sleep 1000
    wshshell.SendKeys "{Tab 5}"
    Sleep 500
    wshshell.SendKeys "{enter}"
    Sleep 500
    hwdThunder = FindWindow(vbNullString, "迅雷7登录")
    If hwdThunder = 0 Then
      lstStatus.AddItem "于" & Time & "被挤下线，已成功帮您重新登录。"
      ci = ci + 1
      zong = zong + 1
      Label1.Caption = zong
    Else
      If ci = 2 Then
        lstStatus.AddItem "于" & Time & "被挤下线，自动登录失败，尝试备用帐号登录。。。"
        Timer1.Enabled = False
        Timer2.Enabled = True
        Timer3.Enabled = False
      End If
    End If
  Else
    ci = 0
  End If
End Sub

Private Sub Timer2_Timer()
  Static sj As Integer
  Static sb As Integer
  Static success As Boolean
  Static xs As Boolean
  hwdThunder = FindWindow(vbNullString, "迅雷7登录") '得到登录界面句柄
  If hwdThunder <> 0 And txtNo.Text <> "" Then
    SetForegroundWindow hwdThunder  '让登录界面切换到前台
    Sleep 1000
    If sb = 3 Then
      lstStatus.AddItem "登录失败，帐号或着密码错误，请重新找一个"
      Timer1.Enabled = False
      Timer2.Enabled = False
      Timer3.Enabled = True
      xs = True
      sb = 0
      Exit Sub
    End If
    If xs = True Then
      wshshell.SendKeys "{Tab}"
      Sleep 100
    End If
    xs = False
    Clipboard.Clear
    Clipboard.SetText txtNo.Text
    wshshell.SendKeys "{Backspace 20}"
    Sleep 500
    wshshell.SendKeys "^{v}"
    wshshell.SendKeys "{Tab}"
    Sleep 500
    Clipboard.Clear
    Clipboard.SetText txtPass.Text
    wshshell.SendKeys "^{v}"
    wshshell.SendKeys "{Tab}"
    wshshell.SendKeys "{enter}"
    wshshell.SendKeys "{Tab}"
    wshshell.SendKeys "{enter}"
    wshshell.SendKeys "{Tab 2}"
    wshshell.SendKeys "{enter}"
    success = True
    sj = 0
    sb = sb + 1
  Else
    sj = sj + 1
    If sj = 3 And success = True Then
      Timer1.Enabled = True
      Timer2.Enabled = False
      Timer3.Enabled = False
      lstStatus.AddItem "于" & Time & "登录成功。"
      sj = 0
      success = False
    End If
    If sj > 3 Then
      Timer1.Enabled = False
      Timer2.Enabled = False
      Timer3.Enabled = True
    End If
  End If
End Sub

Private Sub Timer3_Timer()
  hwdThunder = FindWindow(vbNullString, "迅雷7登录") '得到登录界面句柄
  If hwdThunder = 0 Then
  Timer3.Enabled = False
  Timer2.Enabled = True
  Timer1.Enabled = False
  End If
End Sub

Private Sub buzhou1() '获取最新帐号分享的网页源码，第二次
  Dim mubiao As String
  Dim url As String
  Dim beginPos As Long
  Dim myhtml As New HTMLDocument
  Dim ems As IHTMLElementCollection
  Dim item As IHTMLAnchorElement
  Dim i As Long
  Dim article As String
  
  lstStatus.AddItem "[" & Time & "]获取网页源码成功"
  lstStatus.AddItem "[" & Time & "]正在提取最新更新网址"
  
  If mMainPageCode <> "" Then
    mubiao = Format(Date, "yyyy-mm-dd")
    myhtml.body.innerHTML = mMainPageCode
    Set ems = myhtml.getElementsByTagName("em")
    If ems.length > 0 Then
      For i = 0 To ems.length - 1
        If ems(i).innerText = mubiao Then
          Set item = ems(i).parentNode.parentNode.childNodes(1)
          article = item.pathname
          Exit For
        End If
      Next i
      If i = ems.length Then
        Exit Sub
      End If
    End If
    
    url = "http://www.521xunlei.com/" & article
    lstStatus.AddItem "[" & Time & "]提取成功"
    lstStatus.AddItem url
    mDetailPageCode = web.GetHTMLCode(url)
    file.WriteLineToTextFile mLogPath, "[Page2]"
    file.WriteLineToTextFile mLogPath, mDetailPageCode
    buzhou2
  End If
End Sub

Private Sub buzhou2()
  Dim mubiao As String
  lstStatus.AddItem "[" & Time & "]正在提取帐号和密码  " & Time & "..."
  If mDetailPageCode <> "" Then
    mubiao = "独享不挤迅雷白金会员账号一年只要18元"
    mItemCode = Mid(mDetailPageCode, InStr(mDetailPageCode, mubiao))
    buzhou3
  End If
End Sub

Private Sub buzhou3()
  Dim v
  Dim i As Integer
  Dim lineCode As String
  lstStatus.AddItem "[" & Time & "]免费会员帐号和密码智能获取完毕"
  lstAccout.Clear
  If mItemCode <> "" Then
    v = Split(mItemCode, vbCrLf)
    Dim line As New HTMLDocument
    
    lstAccout.Clear
    For i = 8 To 20
    If Mid(v(i), 1, 5) = "<font" Then
      line.body.innerHTML = v(i)
      lineCode = line.body.innerText
      lstAccout.AddItem lineCode
    Else
      lstAccout.AddItem Replace(v(i), "<br />", "")
    End If
    Next i
  End If
  lstAccout.Selected(0) = True
End Sub
