VERSION 5.00
Object = "{9BD6A640-CE75-11D1-AF04-204C4F4F5020}#2.0#0"; "Mo20.ocx"
Object = "{6C20C089-0689-11D5-B2F8-000102D87123}#2.0#0"; "MO21ScaleBar.ocx"
Object = "{C7FC2F7C-0688-11D5-B2F8-000102D87123}#1.0#0"; "MO21Legend.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������ϵͳ"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11700
   StartUpPosition =   3  '����ȱʡ
   Begin MO21ScaleBar.ScaleBar ScaleBar1 
      Height          =   495
      Left            =   9360
      TabIndex        =   33
      Top             =   8520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      BackColor       =   12648384
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MapUnits        =   2
      ScaleBarUnits   =   3
      ScreenUnits     =   1
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11640
      TabIndex        =   32
      Top             =   8505
      Width           =   11700
   End
   Begin VB.Frame Frame1 
      Caption         =   "ӥ��"
      BeginProperty Font 
         Name            =   "�����п�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   6480
      Width           =   2175
      Begin MapObjects2.Map Map2 
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1935
         _Version        =   131072
         _ExtentX        =   3413
         _ExtentY        =   2778
         _StockProps     =   225
         BackColor       =   16777215
         BorderStyle     =   1
         Contents        =   "��������ϵͳ.frx":0000
      End
   End
   Begin VB.PictureBox Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   11640
      TabIndex        =   1
      Top             =   0
      Width           =   11700
      Begin VB.Timer Timer2 
         Interval        =   100
         Left            =   6240
         Top             =   0
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   5520
         Top             =   0
      End
      Begin VB.PictureBox ImageList1 
         BackColor       =   &H80000005&
         Height          =   480
         Left            =   3960
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   47
         Top             =   0
         Width           =   1200
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   10398
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      BackColor       =   16576
      TabCaption(0)   =   "��ͼ"
      TabPicture(0)   =   "��������ϵͳ.frx":001A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TabStrip1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "legend1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TreeView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "��ѯ"
      TabPicture(1)   =   "��������ϵͳ.frx":0036
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "����"
      TabPicture(2)   =   "��������ϵͳ.frx":0052
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame7 
         Caption         =   "����������"
         Height          =   5415
         Left            =   -74880
         TabIndex        =   35
         Top             =   120
         Width           =   2055
         Begin VB.CommandButton Command6 
            Caption         =   "���"
            Height          =   375
            Left            =   120
            TabIndex        =   46
            Top             =   4920
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   4080
            Width           =   1335
         End
         Begin VB.CommandButton Command5 
            Caption         =   "ȷ��"
            Height          =   375
            Left            =   1200
            TabIndex        =   41
            Top             =   4920
            Width           =   855
         End
         Begin VB.ComboBox Combo7 
            Height          =   300
            Left            =   120
            TabIndex        =   40
            Top             =   3420
            Width           =   1815
         End
         Begin VB.ComboBox Combo6 
            Height          =   300
            Left            =   120
            TabIndex        =   39
            Top             =   2640
            Width           =   1815
         End
         Begin VB.OptionButton Option5 
            Caption         =   "�滺��"
            Height          =   375
            Left            =   240
            TabIndex        =   38
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton Option4 
            Caption         =   "�߻���"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   1140
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "�㻺��"
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "��ѡ��ռ�ʵ��"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "�����뻺�����뾶/m"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   3840
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "��ѡ��ռ�ʵ�����"
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   2280
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4454
         Left            =   -74820
         TabIndex        =   11
         Top             =   1080
         Width           =   2055
         Begin VB.Frame Frame6 
            Caption         =   "���·����ѯ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2055
            Left            =   0
            TabIndex        =   21
            Top             =   2400
            Width           =   2055
            Begin VB.CommandButton Command4 
               Caption         =   "���"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   34
               Top             =   1560
               Width           =   735
            End
            Begin VB.CommandButton Command3 
               Caption         =   "��ѯ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1200
               TabIndex        =   27
               Top             =   1560
               Width           =   735
            End
            Begin VB.ComboBox Combo5 
               Height          =   300
               Left            =   240
               TabIndex        =   25
               Top             =   1200
               Width           =   1575
            End
            Begin VB.ComboBox Combo4 
               Height          =   300
               Left            =   240
               TabIndex        =   24
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label4 
               Caption         =   "��ѡ���յ�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   29
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label3 
               Caption         =   "��ѡ�����"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   28
               Top             =   360
               Width           =   1695
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "�ռ�ʵ���ѯ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   0
            TabIndex        =   20
            Top             =   120
            Width           =   2055
            Begin VB.CommandButton Command2 
               Caption         =   "��ѯ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1200
               TabIndex        =   26
               Top             =   1800
               Width           =   735
            End
            Begin VB.ComboBox Combo3 
               Height          =   300
               ItemData        =   "��������ϵͳ.frx":006E
               Left            =   120
               List            =   "��������ϵͳ.frx":0070
               TabIndex        =   23
               Top             =   1320
               Width           =   1695
            End
            Begin VB.ComboBox Combo2 
               Height          =   300
               Left            =   120
               TabIndex        =   22
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Label6 
               Caption         =   "��ѡ��ռ�ʵ��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label Label5 
               Caption         =   "��ѡ��ռ�ʵ������"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   360
               Width           =   1815
            End
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "�����п�"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   -74880
         TabIndex        =   8
         Top             =   0
         Width           =   2175
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4600
            Left            =   30
            TabIndex        =   10
            Top             =   600
            Width           =   2115
            Begin MSAdodcLib.Adodc Adodc1 
               Height          =   615
               Left            =   480
               Top             =   3480
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1085
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   "Adodc1"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin MSDataGridLib.DataGrid DataGrid1 
               Height          =   1695
               Left            =   0
               TabIndex        =   19
               Top             =   2880
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   2990
               _Version        =   393216
               HeadLines       =   1
               RowHeight       =   15
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   2
               BeginProperty Column00 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   ""
                  Caption         =   ""
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   2052
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  BeginProperty Column00 
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
            Begin VB.CommandButton Command1 
               Caption         =   "��ѯ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1320
               TabIndex        =   17
               Top             =   1920
               Width           =   735
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   16
               Top             =   1920
               Width           =   1095
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   14
               Top             =   1200
               Width           =   1815
            End
            Begin VB.OptionButton Option2 
               Caption         =   "��ȷ��ѯ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   240
               TabIndex        =   13
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton Option1 
               Caption         =   "ģ����ѯ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1080
               TabIndex        =   12
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label2 
               Caption         =   "�������ѯ��Ϣ"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   1560
               Width           =   1815
            End
            Begin VB.Label Label1 
               Caption         =   "��ѡ���ѯ�ֶ�"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   15
               Top             =   840
               Width           =   1815
            End
         End
         Begin VB.PictureBox TabStrip2 
            BeginProperty Font 
               Name            =   "�����п�"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5295
            Left            =   0
            ScaleHeight     =   5235
            ScaleWidth      =   2115
            TabIndex        =   9
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.PictureBox TreeView1 
         Height          =   4455
         Left            =   120
         ScaleHeight     =   4395
         ScaleWidth      =   1995
         TabIndex        =   7
         Top             =   600
         Width           =   2055
      End
      Begin MO21legend.legend legend1 
         Height          =   4455
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   7858
         BackColor       =   -2147483644
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox TabStrip1 
         BeginProperty Font 
            Name            =   "�����п�"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5535
         Left            =   0
         ScaleHeight     =   5475
         ScaleWidth      =   2235
         TabIndex        =   5
         Top             =   0
         Width           =   2295
      End
   End
   Begin MapObjects2.Map Map1 
      Height          =   7935
      Left            =   2400
      TabIndex        =   0
      Top             =   480
      Width           =   9255
      _Version        =   131072
      _ExtentX        =   16325
      _ExtentY        =   13996
      _StockProps     =   225
      BackColor       =   16777215
      BorderStyle     =   1
      ScrollBars      =   0   'False
      Contents        =   "��������ϵͳ.frx":0072
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rectangle(50) As mapobjects2.Rectangle '����Ŵ�ǰ�ľ�������
Dim cnt As Integer                          '����Ŵ�ľ��εĴ���
Dim dc As New DataConnection '����Դ
Dim Layer As MapLayer 'ͼ�����
Dim strscale As String '��¼�����ߣ�����ע���Ƿ���ʾ


'������Ҫ
Dim g_line As mapobjects2.line 'ֱ�߱���
Dim pts As mapobjects2.Points '�㼯����

'������ʾ��̾������Ҫ
Dim g_line1 As mapobjects2.line 'ֱ��
Dim pts1 As mapobjects2.Points '�㼯

'����״�����ƶ�
Dim m_ptsDemo As mapobjects2.Points
Dim m_nNum As Integer
Dim m_Evnt As GeoEvent
Dim m_blsMove As Boolean

'����������
Dim polygon1 As mapobjects2.polygon
Dim line1 As mapobjects2.line
Dim pt1 As mapobjects2.Point



'ͼ������ӹ���
Public Sub LayerLoad(map As map)
Set Layer = New MapLayer '��̬����һ��ͼ�����
Layer.GeoDataset = dc.FindGeoDataset("סլС��") '����һ��shp��
Layer.Symbol.Color = &H40C0&
map.Layers.Add Layer
Set Layer = New MapLayer '��̬����һ��ͼ�����
Layer.GeoDataset = dc.FindGeoDataset("��ͼ") '����һ��shp��
Layer.Symbol.Color = &H40C0&
map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("ѧУ") '����ڶ���shp��
Layer.Symbol.Color = &H80FF80
map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("ҽԺ") '���������shp��
Layer.Symbol.Color = &HFF8080
map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("��ҵ��") '������ĸ�shp��
Layer.Symbol.Color = &HC0C0FF
map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("��԰") '��������shp��
Layer.Symbol.Color = &H80FF&
map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("����") '���������shp��
Layer.Symbol.Color = &HC0C0&
map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("road") '����ڰ˸�shp��
Layer.Symbol.Color = &HFF8080

map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("��") '������߸�shp��
Layer.Symbol.Color = &HFF80FF
Layer.Visible = False '��ʼ�����õ�㲻��ʾ
map.Layers.Add Layer
map.Extent = Map1.FullExtent 'ȫ����ʾ
map.Refresh '����ͼ��

End Sub

'�Բ�ѯ�ֶν��г�ʼ����Ĭ���ֶ�Ϊ����
Private Sub combo1_Init()
Combo1.AddItem ("����")
Combo1.AddItem ("��ݵ���")
Combo1.AddItem ("��ϸ��ַ")
Combo1.AddItem ("סլС��")
Combo1.AddItem ("�绰")
Combo1.Text = "����"
End Sub
'���ѡ��������ͺ���Ӧ���ڿռ������Ͽ��м��ظ��ռ�ʵ��
Private Sub Combo2_Click()
Combo3.Clear '�����Ͽ������
Dim str1 As String
Dim str2 As String
str1 = Combo2.Text
Dim recs As mapobjects2.Recordset
Set recs = Map1.Layers(str1).Records '����ѡ��ĵ������ͼ�¼��
'����ѡ��ĸ��ռ�ʵ����ص���Ͽ���
Do While Not recs.EOF
    str2 = recs.Fields.Item("����")
    If str2 <> "" Then
        Combo3.AddItem str2
    End If
     recs.MoveNext
Loop



End Sub
'���ѡ��������ͺ���Ӧ���ڿռ������Ͽ��м��ظ��ռ�ʵ��
Private Sub Combo6_Click()
Combo7.Clear '�����Ͽ������
Dim str1 As String
Dim str2 As String
str1 = Combo6.Text
Dim recs As mapobjects2.Recordset
Set recs = Map1.Layers(str1).Records '����ѡ��ĵ������ͼ�¼��
'����ѡ��ĸ��ռ�ʵ����ص���Ͽ���
Do While Not recs.EOF
    str2 = recs.Fields.Item("����")
    If str2 <> "" Then
        Combo7.AddItem str2
    End If
     recs.MoveNext
Loop
End Sub

'��ʼ�������Ͽ�
Private Sub Combo4_Init()
    Dim recs As mapobjects2.Recordset
    Dim str As String
    Set recs = Map1.Layers("��").Records
    recs.MoveFirst
    Do While Not recs.EOF
        str = recs.Fields("����").ValueAsString
        Combo4.AddItem str
    recs.MoveNext
    Loop
    
End Sub
'��ʼ���յ���Ͽ�
Private Sub combo5_Init()
   Dim recs As mapobjects2.Recordset
    Dim str As String
    Set recs = Map1.Layers("��").Records
    recs.MoveFirst
    Do While Not recs.EOF
        str = recs.Fields("����").ValueAsString
        Combo5.AddItem str
    recs.MoveNext
    Loop
End Sub


'�����ѯ��ť,
Private Sub Command1_Click()
Dim str1 As String
Dim str2 As String
Dim remark As String
'�����ѯ�ֶκͲ�ѯ��Ϣ
If Text1.Text = "" Then
MsgBox "�������ѯ��Ϣ", vbOKOnly, "������ʾ"
End If
str1 = Combo1.Text
str2 = Text1.Text

'ѡ����Ǿ�ȷ��ѯ
If Option2.Value = True Then
Adodc1.RecordSource = "select * from �û���Ϣ where " & str1 & " like '" & str2 & "'"
'Debug.Print Adodc1.RecordSource
Set DataGrid1.DataSource = Adodc1
Adodc1.Refresh '����Դ����
End If
'ѡ�����ģ����ѯ
If Option1.Value = True Then
Adodc1.RecordSource = "select * from �û���Ϣ where " & str1 & " like '%" & str2 & "%'"
'Debug.Print Adodc1.RecordSource
Adodc1.Refresh '����Դ����
Set DataGrid1.DataSource = Adodc1
End If
End Sub

'����ռ�����ѯ��ť
Private Sub Command2_Click()
If Combo2.Text = "" Then
MsgBox "��ѡ��Ҫ��ѯ�Ŀռ�ʵ������", vbOKOnly, "û��ѡ������"
Exit Sub
End If

If Combo3.Text = "" Then
MsgBox "��ѡ��Ҫ��ѯ�Ŀռ�ʵ��", vbOKOnly, "û��ѡ��ռ�ʵ��"
Exit Sub
End If

'������в�ѯ
Dim recs As mapobjects2.Recordset
Set recs = Map1.Layers(Combo2.Text).Records '����Ӧ�ռ�ʵ�����͵ļ�¼��
Do While Not recs.EOF
    If recs.Fields("����") Like Combo3.Text Then '���ƥ�䣬����˸����
    '����õ����ڵ�ͼ֮��,��һֱͼ�����ļ�����ʾ
    Dim polygon As mapobjects2.polygon
    Set polygon = recs.Fields("shape").Value
    Dim pt As New mapobjects2.Point
    Set pt = polygon.Extent.Center
    If Not Map1.Extent.IsPointIn(pt) Then
        Map1.CenterAt pt.x, pt.y
    End If
    Map1.FlashShape recs.Fields("shape").Value, 3
    Exit Sub
    End If
    recs.MoveNext
    Loop
End Sub

'������·����ѯ��ť
Private Sub Command3_Click()
'�ҵ���ʼ�����ֹ������Ӧ������Ļ���
Dim Recs2 As mapobjects2.Recordset
Set Recs2 = Map1.Layers("��").Records
'������ȡ��ʼ��
Dim pt3 As New mapobjects2.Point
Dim pt4 As New mapobjects2.Point
 '���ó�ʼֵ�������ж���ѡ�����ʼ�����ֹ���ܷ���Ч
 pt3.x = -555
 pt3.y = -555
 pt4.x = -555
 pt4.y = -555

'���û��ѡ����㣬����ʾ
If Combo4.Text = "" Then
MsgBox "��ѡ�����", vbOKOnly, "������ʾ"
Exit Sub
End If

'���û��ѡ���յ㣬����ʾ
If Combo5.Text = "" Then
MsgBox "��ѡ���յ�", vbOKOnly, "������ʾ"
Exit Sub
End If

If Combo4.Text = Combo5.Text Then
MsgBox "�����յ���ͬ��������ѡ��", vbOKOnly, "����"
Exit Sub
End If
'������Ͽ���ѡ���������ȡ������
Do While Not Recs2.EOF
    If Recs2.Fields("����").ValueAsString = Combo4.Text Then
        Set pt3 = Recs2.Fields("shape").Value
    End If
    If Recs2.Fields("����").ValueAsString = Combo5.Text Then
        Set pt4 = Recs2.Fields("shape").Value
    End If
    Recs2.MoveNext
Loop

If pt3.x = -555 And pt3.y = -555 Then
    MsgBox "��ѡ��������Ч", vbOKOnly, "����"
    Exit Sub
End If

If pt4.x = -555 And pt4.y = -555 Then
    MsgBox "��ѡ����յ���Ч", vbOKOnly, "����"
    Exit Sub
End If
'�õ����ٽ����ε�id��
Dim m1 As Integer
Dim m2 As Integer
m1 = GetLineId(pt3.x, pt3.y)
m2 = GetLineId(pt4.x, pt4.y)

Dim fnode1 As Integer
Dim fnode2 As Integer
Dim tnode1 As Integer
Dim tnode2 As Integer
Dim length1 As Double
Dim length2 As Double
'��ȡ���α�,��¼�����յ�����ٽ����ε�ǰ�ڵ�ͺ���

Set Recs2 = Map1.Layers("road").Records
Dim line3 As mapobjects2.line
Dim ptt3 As mapobjects2.Points
Dim line4 As mapobjects2.line
Dim ptt4 As mapobjects2.Points
Dim ptt As mapobjects2.Point
Recs2.MoveFirst
Do While Not Recs2.EOF
If Recs2.Fields("featureId").Value = m1 Then
fnode1 = Recs2.Fields("FNODE_").Value
tnode1 = Recs2.Fields("TNODE_").Value
length1 = Recs2.Fields("LENGTH").Value
Set line3 = Recs2.Fields("shape").Value
End If
If Recs2.Fields("featureId").Value = m2 Then
fnode2 = Recs2.Fields("FNODE_").Value
tnode2 = Recs2.Fields("TNODE_").Value
length2 = Recs2.Fields("LENGTH").Value
Set line4 = Recs2.Fields("shape").Value
End If
Recs2.MoveNext
Loop

Dim pt5 As mapobjects2.Point
Dim pt6 As mapobjects2.Point
Dim length3 As Double
Dim length4 As Double
length3 = 0
length4 = 0
GetNearestPoint pt3, m1, length3, pt5
GetNearestPoint pt4, m2, length4, pt6
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'������·��������¼�
    Set g_line1 = Nothing
    Set pts1 = Nothing
    Map1.TrackingLayer.ClearEvents
    Map1.Refresh
'��֤��Ϊ0��Ϊ0������Ϊ-1��������㷨������ͻ����Ȼ��ɵ�Ӱ���С�������Ǳ��������������⣬�������������ͬ��
If length3 = 0 Then
    length3 = 0.1
End If
If length4 = 0 Then
    length4 = 0.1
End If
'�������·����ѯ

'�����������飬һά�����Ž�㣬��ά�����Ž����ڽӾ���
Dim array1(1 To 550) As Integer
Dim array2(1 To 550, 1 To 550) As Double

'��������г�ʼ��
Dim i As Integer, j As Integer
For i = 1 To 550
    array1(i) = i
Next i
    
For i = 1 To 550
For j = 1 To 550
    '��������ͬ�������Ϊ0
    If i = j Then
    array2(i, j) = 0
    '��������Ϊ99999������������
    Else
    array2(i, j) = 99999
    End If
Next j
Next i

'��ȡ���α������ڽӾ���
Dim recs As mapobjects2.Recordset
Set recs = Map1.Layers("road").Records
Do While Not recs.EOF
    i = recs.Fields("FNODE_").Value
    j = recs.Fields("TNODE_").Value
    If i <> j Then
    array2(i, j) = recs.Fields("LENGTH").Value '����
    array2(j, i) = array2(i, j)
    End If
    recs.MoveNext
Loop

'˼��Ϊͨ��Ѱ����ʼ�����ֹ��������ε����ٽ��ĵ㣬������һ��Ϊ2�����¹���������
array2(540, tnode1) = length3
array2(tnode1, 540) = length3
array2(540, fnode1) = length1 - length3
If array2(540, fnode1) <= 0 Then
    array2(540, fnode1) = 0.1
End If
array2(fnode1, 540) = length1 - length3
If array2(fnode1, 540) <= 0 Then
    array2(fnode1, 540) = 0.1
End If
'��ԭ���Ļ��ε�Ȩ��Ϊ�����
array2(tnode1, fnode1) = 99999
array2(fnode1, tnode1) = 99999

array2(550, tnode2) = length4
array2(tnode2, 550) = length4
array2(550, fnode2) = length2 - length4
If array2(550, fnode2) <= 0 Then
    array2(550, fnode2) = 0.1
End If
array2(fnode2, 550) = length2 - length4
If array2(fnode2, 550) <= 0 Then
    array2(550, fnode2) = 0.1
 End If
'��ԭ�����ε�Ȩ��Ϊ�����
array2(fnode2, tnode2) = 99999
array2(tnode2, fnode2) = 99999

'����������н���ڽӾ���Ĵ���
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim weight(1 To 550, 1 To 550) As Double
Dim path(1 To 550, 1 To 550) As Integer
Dim k As Integer

'��ʼ��
For i = 1 To 550
    For j = 1 To 550
        weight(i, j) = array2(i, j)
        If i <> j And weight(i, j) < 99999 Then
        path(i, j) = i
        Else
        path(i, j) = -1
        End If
    Next j
Next i
'n�ε���
For k = 1 To 550
    For i = 1 To 550
        For j = 1 To 550
            '�õ��µ����·������
            If (weight(i, j) > weight(i, k) + weight(k, j)) Then
                weight(i, j) = weight(i, k) + weight(k, j)
                '�õ����·�������Ľ�����
                path(i, j) = path(k, j)
            End If
        Next j
    Next i
Next k
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
i = 540
j = 550
'������ͨ·
If path(i, j) = -1 Then
MsgBox "��ͼ���ڣ���ѡ����֮����ͨ·", vbOKOnly, "������ʾ"
Exit Sub
End If

Dim MinRoad(1 To 550) As Integer '���·��
Dim count As Integer '���·�������Ľ��ĸ���
'���������·�������Ľ��
count = 1
MinRoad(count) = j
k = j
Do While path(i, k) <> i
    count = count + 1
    k = path(i, k)
    MinRoad(count) = k
Loop
count = count + 1
MinRoad(count) = i
'���洢�����·������
For i = 1 To count \ 2
    j = MinRoad(i)
    MinRoad(i) = MinRoad(count + 1 - i)
    MinRoad(count + 1 - i) = j
Next i

'Ѱ�����·�������Ļ��Σ���������˸��ʾ���������еĻ������ϳ�һ����
If g_line1 Is Nothing Then
Set g_line1 = New mapobjects2.line
End If
If pts1 Is Nothing Then
Set pts1 = New mapobjects2.Points
End If
Set recs = Map1.Layers("road").Records

Dim n As Integer
Dim g_line2 As mapobjects2.line '������·��
Dim g_pt As mapobjects2.Point
Dim g_pts2 As mapobjects2.Points
'��֯���͵ڶ����ڵ㣬����shape�ļ���ֱ��ֻ�������㣬Ϊ�˴ﵽ�Ϻõ�Ч������ֻ���м�ĵ���Ϊ�������ʼ�������ֹ����
'�ٽ��ĵ㣬�˴��д�����
Dim flag As Integer
flag = 0
 Set ptt3 = line3.Parts.Item(0)
    
    If ptt3.count = 2 Then
        Dim p1 As New mapobjects2.Point
        p1.x = (ptt3.Item(0).x + ptt3.Item(1).x) / 2
        p1.y = (ptt3.Item(0).y + ptt3.Item(1).y) / 2
        pts1.Add p1
        flag = 3
    End If
If fnode1 = MinRoad(2) Then
   If flag = 3 Then
    pts1.Add ptt3.Item(0)
   Else
     
                    For n = ptt3.count - 1 To 0 Step -1
                        Set g_pt = New mapobjects2.Point
                        Set g_pt = ptt3.Item(n)
                        If g_pt.x = pt5.x And g_pt.y = pt5.y Then
                           flag = 1
                        End If
                        If flag = 1 Then
                        pts1.Add g_pt
                        End If
                    Next
  End If
  
End If

If tnode1 = MinRoad(2) Then
    If flag = 3 Then
        pts1.Add p1
        pts1.Add ptt3.Item(1)
    Else
    
    flag = 0
    For Each g_pt In line3.Parts(0)
        If g_pt.x = pt5.x And g_pt.y = pt5.y Then
        flag = 1
        End If
        If flag = 1 Then
        pts1.Add g_pt
        End If
    Next g_pt
    End If
End If

'Ѱ����Ӧ�Ļ���
For k = 2 To count - 2
        recs.MoveFirst
           Do While Not recs.EOF
           '��¼�������ҽڵ��ֵ
           i = recs.Fields("FNODE_").Value
           j = recs.Fields("TNODE_").Value
              '�жϻ���
              '''''''''''''''''''''''''''''''''''''''''
              '��������򻡶Σ���ֱ�Ӷ�ȡ��
                If i = MinRoad(k) And j = MinRoad(k + 1) Then
                    Map1.FlashShape recs.Fields("shape").Value, 3
                    Set g_line2 = recs.Fields("shape").Value
                    For Each g_pt In g_line2.Parts(0)
                        pts1.Add g_pt
                    Next g_pt
                End If
                '��������򻡶Σ��������ȡ��
                If i = MinRoad(k + 1) And j = MinRoad(k) Then
                    Map1.FlashShape recs.Fields("shape").Value, 3
                    Set g_line2 = recs.Fields("shape").Value
                    Set g_pts2 = g_line2.Parts.Item(0)
                    For n = g_pts2.count - 1 To 0 Step -1
                        Set g_pt = New mapobjects2.Point
                        Set g_pt = g_pts2.Item(n)
                        pts1.Add g_pt
                    Next
                End If
           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           recs.MoveNext
           Loop
Next k

'��֯�յ�͵����ڶ����ڵ�,����shape�ļ���ֱ��ֻ�������㣬Ϊ�˴ﵽ�Ϻõ�Ч������ֻ���м�ĵ���Ϊ�������ʼ�������ֹ����
'�ٽ��ĵ㣬�˴��д�����
flag = 0
Set ptt4 = line4.Parts.Item(0)
If ptt4.count = 2 Then
    Dim p21 As New mapobjects2.Point
        p21.x = (ptt4.Item(0).x + ptt4.Item(1).x) / 2
        p21.y = (ptt4.Item(0).y + ptt4.Item(1).y) / 2
        pts1.Add p21
        flag = 3
End If
If tnode2 = MinRoad(count - 1) Then
    
    If flag = 3 Then
        
    Else
                    For n = ptt4.count - 1 To 0 Step -1
                        Set g_pt = New mapobjects2.Point
                        Set g_pt = ptt4.Item(n)
                        If g_pt.x = pt6.x And g_pt.y = pt6.y Then
                           Exit For
                        End If
                        pts1.Add g_pt
                   Next
                   pts1.Add pt6
    End If
End If

If fnode2 = MinRoad(count - 1) Then
    If flag = 3 Then
    Else
    For Each g_pt In line4.Parts(0)
        If g_pt.x = pt6.x And g_pt.y = pt6.y Then
         Exit For
        End If
        pts1.Add g_pt
    Next g_pt
    pts1.Add pt6
    End If
    
End If
'������·����������ʾ
If pts1.count <= 1 Then
    MsgBox "��ѡ��ͬ�������յ�", vbOKOnly, "����"
    Exit Sub
End If
        g_line1.Parts.Add pts1
        Set pts1 = g_line1.Parts(0)

Map1.Refresh '�˴�������Map1_AfterLayerDraw
Dim recs3 As mapobjects2.Recordset
Dim str5 As String
str5 = ""
Set recs3 = Map1.Layers("��").Records
Dim pt9 As mapobjects2.Point
'��¼�����·��ֱ��С��30�ĵ㣬��Ϊ��;������վ��
Do While Not recs3.EOF
    Set pt9 = recs3.Fields("shape").Value
    If g_line1.DistanceTo(pt9) < 30 Then
        str5 = str5 + recs3.Fields("����").ValueAsString + "  "
    End If
    recs3.MoveNext
Loop
'������·��
If MsgBox("���·������:" + Format(str(weight(540, 550)), "###,###,###") & "m" + Chr(13) + "��;��������վ��:" + str5, vbOKOnly, "���·��") = vbOK Then
             
End If
    
'����״�����ƽ�
Dim p2 As mapobjects2.Point
Dim line As mapobjects2.line
Set p2 = Map1.ToMapPoint(100, 100)
'���ԭ���¼�������ʱ��ؼ�����Ŀ�������г�ʼ��
Timer2.Interval = 0
Map1.TrackingLayer.ClearEvents
 m_nNum = 0
'���μ�¼��Ҫ����ֱ�ߵĸ����ڵ�
 Set m_ptsDemo = New mapobjects2.Points
 Set line = g_line1
 Dim pt As mapobjects2.Point
  For Each pt In line.Parts(0)
    m_ptsDemo.Add pt
 Next pt
'�������������¼��У�ʱ���interval��Ϊ500ms
 p2.x = m_ptsDemo.Item(0).x
 p2.y = m_ptsDemo.Item(0).y
    Set m_Evnt = Map1.TrackingLayer.AddEvent(p2, 0)
    m_blsMove = True
    Timer2.Interval = 500
    
End Sub


'������·�����¼������ҽ���ʼ�����ֹ����Ͽ��ÿ�
Private Sub Command4_Click()
    Set g_line1 = Nothing
    Set pts1 = Nothing
    Combo4.Text = ""
    Combo5.Text = ""
    Map1.TrackingLayer.ClearEvents
    Map1.Refresh

End Sub


'����������
Private Sub Command5_Click()
    If Combo6.Text = "" Then
        MsgBox "û��ѡ��ռ�ʵ������", vbInformation
        Exit Sub
    End If
    If Combo7.Text = "" Then
        MsgBox "û��ѡ��ռ�ʵ��", vbInformation
        Exit Sub
    End If
    If Text2.Text = "" Then
        MsgBox "û�����뻺�����뾶", vbInformation
        Exit Sub
    End If
    Dim recs As mapobjects2.Recordset
    
    ' �滺��
    If Option5.Value = True Then
        Set recs = Map1.Layers(Combo6.Text).Records
        Do While Not recs.EOF
            If recs.Fields("����").Value Like Combo7.Text Then
                Dim polygon As mapobjects2.polygon
                Set polygon = recs.Fields("shape").Value
                Set polygon1 = polygon.Buffer(Val(Text2.Text), Map1.FullExtent)
                Map1.Refresh
                Exit Sub
            End If
            recs.MoveNext
        Loop
    End If
    '�߻���
    If Option4.Value = 1 Then
    End If
    '�㻺��
    If Option3.Value = 1 Then
    End If
    
    
    
End Sub




'���������
Private Sub Command6_Click()
    Set polygon1 = Nothing
    Map1.Refresh
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim recs As mapobjects2.Recordset
    Set recs = Map1.Layers("����").Records
   
    If DataGrid1.Columns.count < 10 Then
    Exit Sub
    End If
     '�ж��Ƿ�Ϊ��
    If DataGrid1.Columns(0).Text = "" Then
    Exit Sub
    End If
    recs.MoveFirst
    Do While Not recs.EOF
        If recs.Fields("����").Value Like DataGrid1.Columns(10).Text Then
        Map1.FlashShape recs.Fields("shape").Value, 3
        Exit Sub
        End If
        recs.MoveNext
    Loop
End Sub

'�������
Private Sub Form_Load()
'��ʼ����������ʾ��λ��
Form2.Move 300, 300
'��������Դ
dc.Database = App.path + "\..\" + "����" '����Դָ�����ļ����ļ���,���ݺ�Ӧ�ó��������һ��
If Not dc.Connect Then
MsgBox "���ļ����²����������������ļ�"
End
End If

LayerLoad Map1 '��map1�м���ͼ��
LayerLoad Map2 '��map2�м���ͼ��
cnt = 0         '���ó�ʼcntΪ0
legend1.setMapSource Map1
legend1.LoadLegend True
Map1.Refresh
'��һЩ�ؼ��Ƿ���ʾ��������
TreeView1.Visible = False
Frame3.Visible = False

Adodc1.ConnectionString = DBstr
Adodc1.Visible = False

'Ĭ�ϲ�ѯ��ʽΪ��ȷ��ѯ
Option2.Value = True
'�Բ�ѯ�ֶν��г�ʼ��,Ĭ��Ϊ����
combo1_Init
'�Ե�ͼ���������г�ʼ��
TreeView1_Init
'��Ҫѡ��Ŀռ�������ͽ��г�ʼ��
combo2_Init
'�Ա����߿ؼ����г�ʼ��
strscale = "109202"
Call refreshScale
'����ע�ǽ��г�ʼ��
Call renderer
'�������Ͽ���г�ʼ��
Combo4_Init
'���յ���Ͽ�ܽ��г�ʼ��
combo5_Init

'��ʼ���û�����Ϊ�滺��

Option5.Value = 1

Option3.Enabled = False
Option4.Enabled = False
combo6_Init


'���ø�����ʽ
Map1.TrackingLayer.SymbolCount = 2
With Map1.TrackingLayer.Symbol(0)
    .SymbolType = moPointSymbol
    .Color = moBlue
    .Style = 0
    .Size = 8
End With

With Map1.TrackingLayer.Symbol(1)
    .SymbolType = moPointSymbol
    .Color = moBlue
    .Style = 0
    .Size = 5
End With
    
m_blsMove = True
m_nNum = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form3
End Sub

Private Sub legend1_AfterSetLayerVisible(index As Integer, isVisible As Boolean)
Map1.Refresh
End Sub



Private Sub Map1_AfterLayerDraw(ByVal index As Integer, ByVal canceled As Boolean, ByVal hDC As stdole.OLE_HANDLE)
Map2.TrackingLayer.Refresh True '��map1��map2��������
refreshScale 'ˢ�¿ؼ���ʾ
renderer '��������ע���Ƿ���ʾ
'�������ʱ����
If Not pts Is Nothing Then
    If pts.count >= 1 Then
        Dim sym2 As New Symbol
        sym2.Color = moBlue
        sym2.SymbolType = moPointSymbol
        sym2.Size = 5
        Map1.DrawShape pts, sym2
    End If
    
End If
'�������ʱ����
If Not g_line Is Nothing Then
    If pts.count > 1 Then
        Dim sym As New Symbol
        sym.Color = moRed
        sym.SymbolType = moLineSymbol
        sym.Size = 2
        Map1.DrawShape g_line, sym
    End If
End If

'��̾���ʱ������ʾ
If Not g_line1 Is Nothing Then
    If pts1.count > 1 Then
        Dim sym3 As New Symbol
        sym3.Color = &HFFFF00
        sym3.SymbolType = moLineSymbol
        sym3.Size = 4
        Map1.DrawShape g_line1, sym3
    End If
End If

'���ɻ�����
If Not polygon1 Is Nothing Then
    
        Dim sym4 As New Symbol
        sym4.Color = &HFFFF00
        sym4.SymbolType = moFillSymbol
        sym4.Size = 4
        Map1.DrawShape polygon1, sym4
   
End If


End Sub
'��Map1�ϰ������ִ�еĲ���
Private Sub Map1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'���û�е���Ŵ�ť����cnt����
If Toolbar1.Buttons(1).Value = 0 Then
cnt = 0
End If

If Toolbar1.Buttons(1).Value = 1 Then 'ѡ����ǷŴ����,
If Button = vbLeftButton Then 'ѡ������������ִ�������п�Ŵ�
    Set Rectangle(cnt) = Map1.Extent '����Ŵ�ǰ��ͼ������
    cnt = cnt + 1                   '������һ
    Set Map1.Extent = Map1.TrackRectangle
    ElseIf Button = vbRightButton Then '���ѡ������Ҽ�,��ָ����ϴηŴ�ǰ��ͼ��
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If cnt > 0 Then     '֮ǰ���й��Ŵ����
        Map1.Extent = Rectangle(cnt - 1) '�ָ��Ŵ�ǰ�ľ���
        cnt = cnt - 1           '������һ
    End If
        Exit Sub
End If
ElseIf Toolbar1.Buttons(4).Value = 1 Then  'ѡ�����ƽ�Ʋ���
Map1.Pan
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(2).Value = 1 Then 'ѡ�������С��������ִ�������п���С�����������ο�����Ϊͼ���ģ���С1.8��
    Dim r As mapobjects2.Rectangle
    Set r = Map1.TrackRectangle
    Map1.CenterAt r.Center.x, r.Center.y
    Set r = Map1.Extent
    r.ScaleRectangle 1.8
    Map1.Extent = r
'ѡ����Ǳ�־���������ҿռ�ʵ�������
ElseIf Toolbar1.Buttons(6).Value = 1 Then
    Dim p As mapobjects2.Point
    Set p = Map1.ToMapPoint(x, y)
    
    '�����û��Զ���Ŀռ�ʵ��ʶ����
    Call Form3.Identify(x, y)
    Form3.ZOrder 0
    Form3.Show
'ѡ��ִ�в����������
ElseIf Toolbar1.Buttons(7).Value = 1 Then
    '����������ʵ�ֻ�ֱ���������
    If Button = vbLeftButton Then
    '������
    Dim p1 As mapobjects2.Point
        If g_line Is Nothing Then
        Set g_line = New mapobjects2.line
        End If
        If pts Is Nothing Then
        Set pts = New mapobjects2.Points
        End If
        Set p1 = Map1.ToMapPoint(x, y)
        pts.Add p1
        If pts.count = 1 Then
            g_line.Parts.Add pts
            Set pts = g_line.Parts(0)
        End If
        Map1.Refresh
    End If
    '����Ҽ�
    
    
    If Button = vbRightButton Then
        '��ʼ������Ϊ0
        Dim dis As Double
        Dim i As Integer
        dis = 0
        '�������
        If pts.count > 1 Then
            For i = 0 To pts.count - 2
            dis = dis + Math.Sqr((pts(i + 1).x - pts(i).x) * (pts(i + 1).x - pts(i).x) + (pts(i + 1).y - pts(i).y) * (pts(i + 1).y - pts(i).y))
            Next i
       End If
        '����ʾ�����ľ��룬���ȷ��������ʱ����ֱ��ɾ��
        If MsgBox("����:" + Format(str(dis), "###,###,###") + "m", vbOKOnly, "���") = vbOK Then
            Set g_line = Nothing
            Set pts = Nothing
            Map1.Refresh
        End If
    End If
End If
End Sub
'��״̬������ʾ��ǰ����λ��
Private Sub Map1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim pt As New mapobjects2.Point
Set pt = Map1.ToMapPoint(x, y)
StatusBar1.Panels(2).Text = "x=" & pt.x
StatusBar1.Panels(3).Text = "y=" & pt.y

End Sub

'��map2�ϻ���ɫָʾ��
Private Sub Map2_AfterTrackingLayerDraw(ByVal hDC As stdole.OLE_HANDLE)
Dim sym As New Symbol
sym.OutlineColor = moBlue
sym.Size = 2
sym.Style = moTransparentFill
Map2.DrawShape Map1.Extent, sym
End Sub


Private Sub Map2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim r As mapobjects2.Rectangle
Dim pt As New mapobjects2.Point
'�����ο�ı�map1�Ĵ�С
Set r = Map2.TrackRectangle
Set Map1.Extent = r
'����ı�map1��λ��
Set pt = Map2.ToMapPoint(x, y)
Map1.CenterAt pt.x, pt.y

End Sub



'�����ͬ��ѡ�ʱ����һЩ�ؼ��Ƿ���ʾ
Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.index = 2 Then
legend1.Visible = False
TreeView1.Visible = True
ElseIf TabStrip1.SelectedItem.index = 1 Then
TreeView1.Visible = False
legend1.Visible = True
End If
End Sub
'�����ͬ��ѡ�ʱ����frame�ؼ��Ƿ���ʾ���������ڸ�frame�������ؼ���ʾ���ͬframe
Private Sub TabStrip2_Click()
If TabStrip2.SelectedItem.index = 1 Then
Frame3.Visible = False
Frame4.Visible = True
ElseIf TabStrip2.SelectedItem.index = 2 Then
Frame4.Visible = False
Frame3.Visible = True
End If
End Sub

'��ʱ���ؼ���ÿ��1�����»�ȡϵͳ��ʱ�䣬������ʾ��״̬����
Private Sub Timer1_Timer()
   StatusBar1.Panels(4).Text = Format(Now, "yyyy-mm-dd hh:mm:ss")
End Sub

Private Sub Timer2_Timer()
   Dim count As Integer
   Dim x, y As Double
   '�����¼�������ȷ���ƶ�
   If m_blsMove And Map1.TrackingLayer.EventCount > 0 Then
        count = m_nNum
        If count > m_ptsDemo.count Then
        Timer2.Interval = 0
        End If
    Set m_Evnt = Map1.TrackingLayer.Event(0)
     m_nNum = m_nNum + 1
    '���δ�ֱ���ظ����ڵ������ʾ����ͼ��
    If count < m_ptsDemo.count Then
    x = m_ptsDemo.Item(count).x
    y = m_ptsDemo.Item(count).y
    m_Evnt.MoveTo x, y
    End If
    End If
End Sub

'�����������ϵİ�ť
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Toolbar1.Buttons(1).Value = 1 Then '����Ŵ�ͼ��
Map1.MousePointer = moZoomIn
'''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(2).Value = 1 Then '�����Сͼ��
Map1.MousePointer = moZoomOut
''''''''''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(4).Value = 1 Then '���ƽ��ͼ��
Map1.MousePointer = moPan
'''''''''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(3).Value = 1 Then '����ָ�ͼ��
Map1.Extent = Map1.FullExtent
Map1.MousePointer = moArrow
''''''''''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(5).Value = 1 Then '���ѡ��ͼ��
Map1.MousePointer = moArrow
'''''''''''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(6).Value = 1 Then '���ʶ��ͼ��
Map1.MousePointer = moIdentify
'''''''''''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(7).Value = 1 Then '������ͼ��
Map1.MousePointer = moCross


End If
End Sub

Private Sub TreeView1_Init()
Dim nodx As Node
Dim cnt As Integer
TreeView1.Style = tvwTreelinesPlusMinusPictureText '��״��۰���ȫ��Ԫ��
'��������Ϊ"�����гǹ���"�ĸ��ڵ�,     '''''ѡ������Ϊ1��ͼ��
Set nodx = TreeView1.Nodes.Add(, , "3��", "�����гǹ�����ͼ")
        
    '��"�����гǹ���"���ڵ��½����ӽڵ�,
   
        Set nodx = TreeView1.Nodes.Add("3��", tvwChild, "child31", "סլС��")
            cnt = 310
            addlist "child31", "סլС��", "����", cnt
        Set nodx = TreeView1.Nodes.Add("3��", tvwChild, "child33", "ѧУ")
            cnt = 330
            addlist "child33", "ѧУ", "����", cnt
        Set nodx = TreeView1.Nodes.Add("3��", tvwChild, "child34", "ҽԺ")
            cnt = 340
            addlist "child34", "ҽԺ", "����", cnt
        Set nodx = TreeView1.Nodes.Add("3��", tvwChild, "child35", "��ҵ��")
            cnt = 350
            addlist "child35", "��ҵ��", "����", cnt
        Set nodx = TreeView1.Nodes.Add("3��", tvwChild, "child36", "��԰")
            cnt = 360
            addlist "child36", "��԰", "����", cnt
    
               

End Sub
'���Ӹ��ռ�ʵ�壬�����ֱ������㣬ͼ������ͼ�����Ա��ֶΣ��������ɹؼ���
Private Sub addlist(str1 As String, str2 As String, str3 As String, cnt As Integer)
        Dim nodx As Node
        Dim str10 As String
        Dim recs As mapobjects2.Recordset
        Set recs = Map1.Layers(str2).Records '�򿪼�¼��
        Do While Not recs.EOF
        cnt = cnt + 1
        str10 = "child"
        str10 = str2 + str(cnt)
        If recs.Fields.Item(str3) <> "" Then '�����Լ�¼�ǿ�ʱ������ӵ����οؼ���
             Set nodx = TreeView1.Nodes.Add(str1, tvwChild, str10, recs.Fields.Item(str3))
        End If
        recs.MoveNext
        Loop
End Sub
'��combo2����ӿռ��������
Private Sub combo2_Init()
Combo2.AddItem ("��԰")
Combo2.AddItem ("ѧУ")
Combo2.AddItem ("��ҵ��")
Combo2.AddItem ("סլС��")
Combo2.AddItem ("ҽԺ")
Combo2.AddItem ("��԰")

 
End Sub
'�����������ñȽϼ��Ĵ�����д����Ҫ������ͬ����֯��κ���ʹ�����ݵĲ�һ������ɵģ��˴�����Լ���д�����ݣ�û��ͨ����
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
 Dim recs As mapobjects2.Recordset
'���ѡ��Ľڵ�Ϊĩβ�ڵ㣬��
If Node.Children <= 0 Then
    '���ѡ���������
    If Node.Parent.Text = "����" Then
    Set recs = Map1.Layers(Node.Text).Records
     Do While Not recs.EOF
        If recs.Fields("����") Like Node.Text Then '���ƥ�䣬����˸����
        '����õ����ڵ�ͼ֮��,��һֱͼ�����ļ�����ʾ
        Dim polygon As mapobjects2.polygon
        Dim pt As mapobjects2.Point
        Set polygon = recs.Fields("shape").Value
        Set pt = polygon.Extent.Center
        If Not Map1.Extent.IsPointIn(pt) Then
            Map1.CenterAt pt.x, pt.y
         End If
        Map1.FlashShape recs.Fields("shape").Value, 3
        Exit Sub
        End If
        recs.MoveNext
        Loop
    End If
    
    
    Set recs = Map1.Layers(Node.Parent.Text).Records '����Ӧ�ռ�ʵ�����͵ļ�¼��
    Do While Not recs.EOF
        If recs.Fields("����") Like Node.Text Then '���ƥ�䣬����˸����
        '����õ����ڵ�ͼ֮��,��һֱͼ�����ļ�����ʾ
         Set polygon = recs.Fields("shape").Value
        Set pt = polygon.Extent.Center
        If Not Map1.Extent.IsPointIn(pt) Then
            Map1.CenterAt pt.x, pt.y
         End If
        Map1.FlashShape recs.Fields("shape").Value, 3
        Exit Sub
        End If
        recs.MoveNext
        Loop
End If
End Sub




'�����߿ؼ��ĳ�ʼ��
Private Sub refreshScale()
    ScaleBar1.MapExtent.MaxX = Map1.Extent.Right
    ScaleBar1.MapExtent.MinX = Map1.Extent.Left
    ScaleBar1.MapExtent.MaxY = Map1.Extent.Bottom
    ScaleBar1.MapExtent.MinY = Map1.Extent.Top
    
    ScaleBar1.PageExtent.MaxX = (Map1.Left + Map1.Width) / Screen.TwipsPerPixelX
    ScaleBar1.PageExtent.MinX = Map1.Left / Screen.TwipsPerPixelX
    ScaleBar1.PageExtent.MaxY = (Map1.Top + Map1.Height) / Screen.TwipsPerPixelY
    ScaleBar1.PageExtent.MinY = Map1.Top / Screen.TwipsPerPixelY
    ScaleBar1.Refresh
    '��д�����ߵ���ʽ
    StatusBar1.Panels(1).Text = "�����ߣ� 1��" & Format(ScaleBar1.RFScale, "###,###,###,###,###")
    '��¼��ǰ�ı����ߣ����ھ���ע�ǵ���ʾ���
    strscale = ScaleBar1.RFScale
  End Sub
'��ǣ���1��30000Ϊ�ֽ��߾����Ƿ���
Private Sub renderer()
Dim i As Integer
If Val(strscale) > 30000 Then
    For i = 0 To Map1.Layers.count - 1
    Set Map1.Layers(i).renderer = Nothing
    Next
Exit Sub
End If

If Val(strscale) < 5300 Then
    For i = 0 To Map1.Layers.count - 1
        Set Map1.Layers(i).renderer = New LabelRenderer
            Map1.Layers(i).renderer.Field = "����"
    Next
End If
End Sub
'''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''
'���ȸ��ݸ����ĵ�Ѱ������Ļ��ε�id��
Public Function GetLineId(x As Double, y As Double) As Integer

Dim rs As mapobjects2.Recordset
Dim pt3 As New mapobjects2.Point
pt3.x = x
pt3.y = y
'�ҵ����εļ�¼��
Set rs = Map1.Layers("road").Records
Dim dis As Double
dis = 99999
Dim id As Integer
id = -1
rs.MoveFirst
'�Ի������ν����жϣ��ҳ�����õ�����Ļ���
Do While Not rs.EOF
    Dim line1 As mapobjects2.line
    Dim d As Double
    Set line1 = rs.Fields("shape").Value
    d = line1.DistanceTo(pt3)
    
    If dis > d Then
        dis = d
        id = rs.Fields("featureId").Value
    End If
    rs.MoveNext
Loop
'���ظû��ε�id��
If id <> -1 Then
    GetLineId = id
End If
End Function

'��㵽������ε�����ĵ�������ţ����뻡��ǰ�ڵ�ľ��룬
Public Function GetNearestPoint(pt1 As mapobjects2.Point, id As Integer, ByRef length1 As Double, ByRef pt2 As mapobjects2.Point)
    Dim rs As mapobjects2.Recordset
    Dim pt As mapobjects2.Point
    Set rs = Map1.Layers("road").Records
    rs.MoveFirst
    Do While Not rs.EOF
        Dim line1 As mapobjects2.line
        If rs.Fields("featureId").Value = id Then
        Set line1 = rs.Fields("shape").Value
        GoTo continue
        End If
        rs.MoveNext
    Loop
    
continue:
    Dim pts As mapobjects2.Points
    Set pts = line1.Parts(0)
    Dim dis As Double
    Dim d As Double
    dis = 99999
    For Each pt In line1.Parts(0)
        d = Math.Sqr((pt.x - pt1.x) * (pt.x - pt1.x) + (pt.y - pt1.y) * (pt.y - pt1.y))
        If dis > d Then
            dis = d
           Set pt2 = pt
        End If
    Next pt
    length1 = 0
    Dim i As Integer
    For i = 0 To pts.count - 1
        If pts.Item(i).x = pt2.x And pts.Item(i).y = pt2.y Then
            Exit Function
        End If
        length1 = length1 + Math.Sqr((pts.Item(i + 1).x - pts.Item(i).x) * (pts.Item(i + 1).x - pts.Item(i).x) + (pts.Item(i + 1).y - pts.Item(i).y) * (pts.Item(i + 1).y - pts.Item(i).y))
    Next i
    
End Function

'��combo6����ӿռ��������
Private Sub combo6_Init()
Combo6.AddItem ("סլС��")
Combo6.AddItem ("ѧУ")
Combo6.AddItem ("ҽԺ")
Combo6.AddItem ("��ҵ��")
Combo6.AddItem ("����")
Combo6.AddItem ("��԰")

End Sub

