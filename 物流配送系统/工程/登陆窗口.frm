VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "登陆窗口"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "华文行楷"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4485
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox LoginPicture 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "华文行楷"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   6
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "登陆"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
      End
      Begin MSForms.OptionButton Option2 
         Height          =   855
         Left            =   2640
         TabIndex        =   8
         Top             =   120
         Width           =   1455
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2566;1508"
         Value           =   "0"
         Caption         =   "管理员"
         FontName        =   "华文行楷"
         FontHeight      =   285
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin MSForms.OptionButton Option1 
         Height          =   855
         Left            =   960
         TabIndex        =   7
         Top             =   120
         Width           =   1335
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "2355;1508"
         Value           =   "0"
         Caption         =   "用户"
         FontName        =   "华文行楷"
         FontHeight      =   285
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Caption         =   "密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "帐号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'单击登陆按钮
Private Sub Command1_Click()
'加载数据库
'如果用户没有选择用户类型，则要求其选择
If Option1.Value = False And Option2.Value = False Then
MsgBox "登陆错误，请先选择用户类型", vbOKOnly, "登陆错误"
Exit Sub
End If
'定义数据集，保存查询语句的字符串等
Dim rs As New ADODB.Recordset
Dim SQLstr As String
Dim str1 As String
Dim str2 As String
Dim i As Integer
'用户类型为客户
If Option1.Value = True Then
'打开相应的权限表
SQLstr = "select * from 用户权限表"
rs.Open SQLstr, DBCn, adOpenStatic, adLockBatchOptimistic
rs.MoveFirst
'如果非空,则依次进行判断，如果帐号和密码匹配，则关闭数据集进入主窗体
Do While Not rs.EOF

str1 = rs.Fields("帐号").Value
str2 = rs.Fields("密码").Value

If Text1.Text = str1 And Text2.Text = str2 Then
Unload Me
Form2.Show '进入主窗体
rs.Close
Exit Sub
End If
'移至下一记录集
rs.MoveNext
Loop
End If
'用户类型为管理员
If Option2.Value = True Then
SQLstr = "select * from 管理员权限表"
rs.Open SQLstr, DBCn, adOpenStatic, adLockBatchOptimistic
rs.MoveFirst
Do While Not rs.EOF
str1 = rs.Fields("帐号").Value
str2 = rs.Fields("密码").Value
If Text1.Text = str1 And Text2.Text = str2 Then
Unload Me
Form2.Show '进入主窗体
rs.Close
Exit Sub
End If
rs.MoveNext
Loop
End If

'关闭数据集
rs.Close
'弹出错误信息，并且将帐号和密码文本框置空，同时让帐号文本框获得焦点
MsgBox "帐号或密码错误，请重新登陆", vbOKOnly, "登陆错误"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End Sub
'单击取消按钮
Private Sub Command2_Click()
End
End Sub

Private Sub Form_Activate()
Option1.Value = True
End Sub

'加载窗体
Private Sub Form_Load()
Text1.Text = "123"
Text2.Text = "456"
'加载图片
LoginPicture.Picture = LoadPicture(App.path + "\..\" + "图片图标\兰州大学.jpg  ")
DBstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.path + "\..\数据\CustomerInfo.mdb;Persist Security Info=False"
DBCn.Open DBstr
'将登陆窗体显示在屏幕的中心
Form1.Move Screen.Width / 2 - Form1.Width / 2, Screen.Height / 2 - Form1.Height / 2
End Sub


'单击学生单选框
Private Sub Option1_Click()
Text1.SetFocus '帐号文本框获得焦点
End Sub
'单击管理员单选框
Private Sub Option2_Click()
Text1.SetFocus '帐号文本框获得焦点
End Sub
'按下enter键后转成tab键
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{Tab}"
KeyAscii = 0
End If
End Sub
