VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��½����"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4485
   BeginProperty Font 
      Name            =   "�����п�"
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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox LoginPicture 
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "�����п�"
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
         Caption         =   "ȡ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��½"
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "����Ա"
         FontName        =   "�����п�"
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
         Caption         =   "�û�"
         FontName        =   "�����п�"
         FontHeight      =   285
         FontCharSet     =   134
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�ʺ�"
         BeginProperty Font 
            Name            =   "����"
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


'������½��ť
Private Sub Command1_Click()
'�������ݿ�
'����û�û��ѡ���û����ͣ���Ҫ����ѡ��
If Option1.Value = False And Option2.Value = False Then
MsgBox "��½��������ѡ���û�����", vbOKOnly, "��½����"
Exit Sub
End If
'�������ݼ��������ѯ�����ַ�����
Dim rs As New ADODB.Recordset
Dim SQLstr As String
Dim str1 As String
Dim str2 As String
Dim i As Integer
'�û�����Ϊ�ͻ�
If Option1.Value = True Then
'����Ӧ��Ȩ�ޱ�
SQLstr = "select * from �û�Ȩ�ޱ�"
rs.Open SQLstr, DBCn, adOpenStatic, adLockBatchOptimistic
rs.MoveFirst
'����ǿ�,�����ν����жϣ�����ʺź�����ƥ�䣬��ر����ݼ�����������
Do While Not rs.EOF

str1 = rs.Fields("�ʺ�").Value
str2 = rs.Fields("����").Value

If Text1.Text = str1 And Text2.Text = str2 Then
Unload Me
Form2.Show '����������
rs.Close
Exit Sub
End If
'������һ��¼��
rs.MoveNext
Loop
End If
'�û�����Ϊ����Ա
If Option2.Value = True Then
SQLstr = "select * from ����ԱȨ�ޱ�"
rs.Open SQLstr, DBCn, adOpenStatic, adLockBatchOptimistic
rs.MoveFirst
Do While Not rs.EOF
str1 = rs.Fields("�ʺ�").Value
str2 = rs.Fields("����").Value
If Text1.Text = str1 And Text2.Text = str2 Then
Unload Me
Form2.Show '����������
rs.Close
Exit Sub
End If
rs.MoveNext
Loop
End If

'�ر����ݼ�
rs.Close
'����������Ϣ�����ҽ��ʺź������ı����ÿգ�ͬʱ���ʺ��ı����ý���
MsgBox "�ʺŻ�������������µ�½", vbOKOnly, "��½����"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End Sub
'����ȡ����ť
Private Sub Command2_Click()
End
End Sub

Private Sub Form_Activate()
Option1.Value = True
End Sub

'���ش���
Private Sub Form_Load()
Text1.Text = "123"
Text2.Text = "456"
'����ͼƬ
LoginPicture.Picture = LoadPicture(App.path + "\..\" + "ͼƬͼ��\���ݴ�ѧ.jpg  ")
DBstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.path + "\..\����\CustomerInfo.mdb;Persist Security Info=False"
DBCn.Open DBstr
'����½������ʾ����Ļ������
Form1.Move Screen.Width / 2 - Form1.Width / 2, Screen.Height / 2 - Form1.Height / 2
End Sub


'����ѧ����ѡ��
Private Sub Option1_Click()
Text1.SetFocus '�ʺ��ı����ý���
End Sub
'��������Ա��ѡ��
Private Sub Option2_Click()
Text1.SetFocus '�ʺ��ı����ý���
End Sub
'����enter����ת��tab��
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{Tab}"
KeyAscii = 0
End If
End Sub
