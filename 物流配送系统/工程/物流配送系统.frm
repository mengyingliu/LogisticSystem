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
   Caption         =   "物流配送系统"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "宋体"
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
   StartUpPosition =   3  '窗口缺省
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
      Caption         =   "鹰眼"
      BeginProperty Font 
         Name            =   "华文行楷"
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
         Contents        =   "物流配送系统.frx":0000
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
      TabCaption(0)   =   "地图"
      TabPicture(0)   =   "物流配送系统.frx":001A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TabStrip1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "legend1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TreeView1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "查询"
      TabPicture(1)   =   "物流配送系统.frx":0036
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "分析"
      TabPicture(2)   =   "物流配送系统.frx":0052
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame7 
         Caption         =   "缓冲区分析"
         Height          =   5415
         Left            =   -74880
         TabIndex        =   35
         Top             =   120
         Width           =   2055
         Begin VB.CommandButton Command6 
            Caption         =   "清除"
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
            Caption         =   "确定"
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
            Caption         =   "面缓冲"
            Height          =   375
            Left            =   240
            TabIndex        =   38
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton Option4 
            Caption         =   "线缓冲"
            Height          =   375
            Left            =   240
            TabIndex        =   37
            Top             =   1140
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "点缓冲"
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "请选择空间实体"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "请输入缓冲区半径/m"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   3840
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "请选择空间实体类别"
            Height          =   375
            Left            =   120
            TabIndex        =   42
            Top             =   2280
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "宋体"
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
            Caption         =   "最短路径查询"
            BeginProperty Font 
               Name            =   "宋体"
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
               Caption         =   "清除"
               BeginProperty Font 
                  Name            =   "宋体"
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
               Caption         =   "查询"
               BeginProperty Font 
                  Name            =   "宋体"
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
               Caption         =   "请选择终点"
               BeginProperty Font 
                  Name            =   "宋体"
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
               Caption         =   "请选择起点"
               BeginProperty Font 
                  Name            =   "宋体"
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
            Caption         =   "空间实体查询"
            BeginProperty Font 
               Name            =   "宋体"
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
               Caption         =   "查询"
               BeginProperty Font 
                  Name            =   "宋体"
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
               ItemData        =   "物流配送系统.frx":006E
               Left            =   120
               List            =   "物流配送系统.frx":0070
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
               Caption         =   "请选择空间实体"
               BeginProperty Font 
                  Name            =   "宋体"
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
               Caption         =   "请选择空间实体类型"
               BeginProperty Font 
                  Name            =   "宋体"
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
            Name            =   "华文行楷"
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
               Name            =   "宋体"
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
                  Name            =   "宋体"
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
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
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
               Caption         =   "查询"
               BeginProperty Font 
                  Name            =   "宋体"
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
                  Name            =   "宋体"
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
                  Name            =   "宋体"
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
               Caption         =   "精确查询"
               BeginProperty Font 
                  Name            =   "宋体"
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
               Caption         =   "模糊查询"
               BeginProperty Font 
                  Name            =   "宋体"
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
               Caption         =   "请输入查询信息"
               BeginProperty Font 
                  Name            =   "宋体"
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
               Caption         =   "请选择查询字段"
               BeginProperty Font 
                  Name            =   "宋体"
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
               Name            =   "华文行楷"
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
            Name            =   "华文行楷"
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
      Contents        =   "物流配送系统.frx":0072
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rectangle(50) As mapobjects2.Rectangle '保存放大前的矩形数组
Dim cnt As Integer                          '保存放大的矩形的次数
Dim dc As New DataConnection '数据源
Dim Layer As MapLayer '图层变量
Dim strscale As String '记录比例尺，决定注记是否显示


'测距的需要
Dim g_line As mapobjects2.line '直线变量
Dim pts As mapobjects2.Points '点集变量

'高亮显示最短距离的需要
Dim g_line1 As mapobjects2.line '直线
Dim pts1 As mapobjects2.Points '点集

'沿线状特征移动
Dim m_ptsDemo As mapobjects2.Points
Dim m_nNum As Integer
Dim m_Evnt As GeoEvent
Dim m_blsMove As Boolean

'缓冲区变量
Dim polygon1 As mapobjects2.polygon
Dim line1 As mapobjects2.line
Dim pt1 As mapobjects2.Point



'图层加载子过程
Public Sub LayerLoad(map As map)
Set Layer = New MapLayer '动态申请一个图层变量
Layer.GeoDataset = dc.FindGeoDataset("住宅小区") '加入一个shp层
Layer.Symbol.Color = &H40C0&
map.Layers.Add Layer
Set Layer = New MapLayer '动态申请一个图层变量
Layer.GeoDataset = dc.FindGeoDataset("底图") '加入一个shp层
Layer.Symbol.Color = &H40C0&
map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("学校") '加入第二个shp层
Layer.Symbol.Color = &H80FF80
map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("医院") '加入第三个shp层
Layer.Symbol.Color = &HFF8080
map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("商业区") '加入第四个shp层
Layer.Symbol.Color = &HC0C0FF
map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("公园") '加入第五个shp层
Layer.Symbol.Color = &H80FF&
map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("其他") '加入第六个shp层
Layer.Symbol.Color = &HC0C0&
map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("road") '加入第八个shp层
Layer.Symbol.Color = &HFF8080

map.Layers.Add Layer
Set Layer = New MapLayer
Layer.GeoDataset = dc.FindGeoDataset("点") '加入第七个shp层
Layer.Symbol.Color = &HFF80FF
Layer.Visible = False '初始化设置点层不显示
map.Layers.Add Layer
map.Extent = Map1.FullExtent '全屏显示
map.Refresh '更新图层

End Sub

'对查询字段进行初始化，默认字段为姓名
Private Sub combo1_Init()
Combo1.AddItem ("姓名")
Combo1.AddItem ("快递单号")
Combo1.AddItem ("详细地址")
Combo1.AddItem ("住宅小区")
Combo1.AddItem ("电话")
Combo1.Text = "姓名"
End Sub
'点击选择地物类型后，相应的在空间地物组合框中加载各空间实体
Private Sub Combo2_Click()
Combo3.Clear '清空组合框的内容
Dim str1 As String
Dim str2 As String
str1 = Combo2.Text
Dim recs As mapobjects2.Recordset
Set recs = Map1.Layers(str1).Records '打开所选择的地物类型记录集
'将所选择的各空间实体加载到组合框中
Do While Not recs.EOF
    str2 = recs.Fields.Item("名称")
    If str2 <> "" Then
        Combo3.AddItem str2
    End If
     recs.MoveNext
Loop



End Sub
'点击选择地物类型后，相应的在空间地物组合框中加载各空间实体
Private Sub Combo6_Click()
Combo7.Clear '清空组合框的内容
Dim str1 As String
Dim str2 As String
str1 = Combo6.Text
Dim recs As mapobjects2.Recordset
Set recs = Map1.Layers(str1).Records '打开所选择的地物类型记录集
'将所选择的各空间实体加载到组合框中
Do While Not recs.EOF
    str2 = recs.Fields.Item("名称")
    If str2 <> "" Then
        Combo7.AddItem str2
    End If
     recs.MoveNext
Loop
End Sub

'初始化起点组合框
Private Sub Combo4_Init()
    Dim recs As mapobjects2.Recordset
    Dim str As String
    Set recs = Map1.Layers("点").Records
    recs.MoveFirst
    Do While Not recs.EOF
        str = recs.Fields("区名").ValueAsString
        Combo4.AddItem str
    recs.MoveNext
    Loop
    
End Sub
'初始化终点组合框
Private Sub combo5_Init()
   Dim recs As mapobjects2.Recordset
    Dim str As String
    Set recs = Map1.Layers("点").Records
    recs.MoveFirst
    Do While Not recs.EOF
        str = recs.Fields("区名").ValueAsString
        Combo5.AddItem str
    recs.MoveNext
    Loop
End Sub


'点击查询按钮,
Private Sub Command1_Click()
Dim str1 As String
Dim str2 As String
Dim remark As String
'保存查询字段和查询信息
If Text1.Text = "" Then
MsgBox "请输入查询信息", vbOKOnly, "错误提示"
End If
str1 = Combo1.Text
str2 = Text1.Text

'选择的是精确查询
If Option2.Value = True Then
Adodc1.RecordSource = "select * from 用户信息 where " & str1 & " like '" & str2 & "'"
'Debug.Print Adodc1.RecordSource
Set DataGrid1.DataSource = Adodc1
Adodc1.Refresh '数据源更新
End If
'选择的是模糊查询
If Option1.Value = True Then
Adodc1.RecordSource = "select * from 用户信息 where " & str1 & " like '%" & str2 & "%'"
'Debug.Print Adodc1.RecordSource
Adodc1.Refresh '数据源更新
Set DataGrid1.DataSource = Adodc1
End If
End Sub

'点击空间地物查询按钮
Private Sub Command2_Click()
If Combo2.Text = "" Then
MsgBox "请选择要查询的空间实体类型", vbOKOnly, "没有选择类型"
Exit Sub
End If

If Combo3.Text = "" Then
MsgBox "请选择要查询的空间实体", vbOKOnly, "没有选择空间实体"
Exit Sub
End If

'否则进行查询
Dim recs As mapobjects2.Recordset
Set recs = Map1.Layers(Combo2.Text).Records '打开相应空间实体类型的记录集
Do While Not recs.EOF
    If recs.Fields("名称") Like Combo3.Text Then '如果匹配，则闪烁三次
    '如果该地物在地图之外,则一直图的中心加以显示
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

'点击最短路径查询按钮
Private Sub Command3_Click()
'找到起始点和终止点所对应的最近的弧段
Dim Recs2 As mapobjects2.Recordset
Set Recs2 = Map1.Layers("点").Records
'用于提取起始点
Dim pt3 As New mapobjects2.Point
Dim pt4 As New mapobjects2.Point
 '设置初始值，用于判断所选择的起始点和终止点受否有效
 pt3.x = -555
 pt3.y = -555
 pt4.x = -555
 pt4.y = -555

'如果没有选择起点，则提示
If Combo4.Text = "" Then
MsgBox "请选择起点", vbOKOnly, "错误提示"
Exit Sub
End If

'如果没有选择终点，则提示
If Combo5.Text = "" Then
MsgBox "请选择终点", vbOKOnly, "错误提示"
Exit Sub
End If

If Combo4.Text = Combo5.Text Then
MsgBox "起点和终点相同，请重新选择", vbOKOnly, "警告"
Exit Sub
End If
'根据组合框所选择的内容提取各个点
Do While Not Recs2.EOF
    If Recs2.Fields("区名").ValueAsString = Combo4.Text Then
        Set pt3 = Recs2.Fields("shape").Value
    End If
    If Recs2.Fields("区名").ValueAsString = Combo5.Text Then
        Set pt4 = Recs2.Fields("shape").Value
    End If
    Recs2.MoveNext
Loop

If pt3.x = -555 And pt3.y = -555 Then
    MsgBox "所选择的起点无效", vbOKOnly, "警告"
    Exit Sub
End If

If pt4.x = -555 And pt4.y = -555 Then
    MsgBox "所选择的终点无效", vbOKOnly, "警告"
    Exit Sub
End If
'得到最临近弧段的id号
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
'读取弧段表,记录起点和终点的最临近弧段的前节点和后结点

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
'清除最短路径，清除事件
    Set g_line1 = Nothing
    Set pts1 = Nothing
    Map1.TrackingLayer.ClearEvents
    Map1.Refresh
'保证不为0，为0则设置为-1，否则和算法产生冲突。虽然造成的影响很小，但这是本程序带解决的问题，下面的设置问题同此
If length3 = 0 Then
    length3 = 0.1
End If
If length4 = 0 Then
    length4 = 0.1
End If
'进行最短路径查询

'定义两个数组，一维数组存放结点，二维数组存放结点的邻接矩阵
Dim array1(1 To 550) As Integer
Dim array2(1 To 550, 1 To 550) As Double

'对数组进行初始化
Dim i As Integer, j As Integer
For i = 1 To 550
    array1(i) = i
Next i
    
For i = 1 To 550
For j = 1 To 550
    '如果结点相同，则距离为0
    If i = j Then
    array2(i, j) = 0
    '否则先设为99999，代表正无穷
    Else
    array2(i, j) = 99999
    End If
Next j
Next i

'读取弧段表，更新邻接矩阵
Dim recs As mapobjects2.Recordset
Set recs = Map1.Layers("road").Records
Do While Not recs.EOF
    i = recs.Fields("FNODE_").Value
    j = recs.Fields("TNODE_").Value
    If i <> j Then
    array2(i, j) = recs.Fields("LENGTH").Value '距离
    array2(j, i) = array2(i, j)
    End If
    recs.MoveNext
Loop

'思想为通过寻找起始点和终止点最近弧段的最临近的点，将弧段一分为2，重新构建理解矩阵
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
'将原来的弧段的权设为无穷大
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
'将原来弧段的权设为无穷大
array2(fnode2, tnode2) = 99999
array2(tnode2, fnode2) = 99999

'至此完成所有结点邻接矩阵的创建
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim weight(1 To 550, 1 To 550) As Double
Dim path(1 To 550, 1 To 550) As Integer
Dim k As Integer

'初始化
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
'n次递推
For k = 1 To 550
    For i = 1 To 550
        For j = 1 To 550
            '得到新的最短路径长度
            If (weight(i, j) > weight(i, k) + weight(k, j)) Then
                weight(i, j) = weight(i, k) + weight(k, j)
                '得到最短路径经过的结点序号
                path(i, j) = path(k, j)
            End If
        Next j
    Next i
Next k
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
i = 540
j = 550
'不存在通路
If path(i, j) = -1 Then
MsgBox "该图幅内，所选两点之间无通路", vbOKOnly, "错误提示"
Exit Sub
End If

Dim MinRoad(1 To 550) As Integer '最短路径
Dim count As Integer '最短路径包含的结点的个数
'逆序存放最短路径经过的结点
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
'将存储的最短路径逆序
For i = 1 To count \ 2
    j = MinRoad(i)
    MinRoad(i) = MinRoad(count + 1 - i)
    MinRoad(count + 1 - i) = j
Next i

'寻找最短路径经过的弧段，并依次闪烁显示，并将所有的弧段整合成一条线
If g_line1 Is Nothing Then
Set g_line1 = New mapobjects2.line
End If
If pts1 Is Nothing Then
Set pts1 = New mapobjects2.Points
End If
Set recs = Map1.Layers("road").Records

Dim n As Integer
Dim g_line2 As mapobjects2.line '存放最短路径
Dim g_pt As mapobjects2.Point
Dim g_pts2 As mapobjects2.Points
'组织起点和第二个节点，鉴于shape文件的直线只有两个点，为了达到较好的效果，当只有中间的点作为最后与起始点或者终止点最
'临近的点，此处有待改善
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

'寻找相应的弧段
For k = 2 To count - 2
        recs.MoveFirst
           Do While Not recs.EOF
           '记录弧段左右节点的值
           i = recs.Fields("FNODE_").Value
           j = recs.Fields("TNODE_").Value
              '判断弧段
              '''''''''''''''''''''''''''''''''''''''''
              '如果是正向弧段，则直接读取点
                If i = MinRoad(k) And j = MinRoad(k + 1) Then
                    Map1.FlashShape recs.Fields("shape").Value, 3
                    Set g_line2 = recs.Fields("shape").Value
                    For Each g_pt In g_line2.Parts(0)
                        pts1.Add g_pt
                    Next g_pt
                End If
                '如果是逆向弧段，则逆向读取点
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

'组织终点和倒数第二个节点,鉴于shape文件的直线只有两个点，为了达到较好的效果，当只有中间的点作为最后与起始点或者终止点最
'临近的点，此处有待改善
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
'输出最短路径，高亮显示
If pts1.count <= 1 Then
    MsgBox "请选择不同的起点和终点", vbOKOnly, "警告"
    Exit Sub
End If
        g_line1.Parts.Add pts1
        Set pts1 = g_line1.Parts(0)

Map1.Refresh '此处调用了Map1_AfterLayerDraw
Dim recs3 As mapobjects2.Recordset
Dim str5 As String
str5 = ""
Set recs3 = Map1.Layers("点").Records
Dim pt9 As mapobjects2.Point
'记录和最短路径直线小于30的点，视为沿途经过的站点
Do While Not recs3.EOF
    Set pt9 = recs3.Fields("shape").Value
    If g_line1.DistanceTo(pt9) < 30 Then
        str5 = str5 + recs3.Fields("区名").ValueAsString + "  "
    End If
    recs3.MoveNext
Loop
'输出最短路径
If MsgBox("最短路径长度:" + Format(str(weight(540, 550)), "###,###,###") & "m" + Chr(13) + "沿途所经过的站点:" + str5, vbOKOnly, "最短路径") = vbOK Then
             
End If
    
'沿线状特征推进
Dim p2 As mapobjects2.Point
Dim line As mapobjects2.line
Set p2 = Map1.ToMapPoint(100, 100)
'清除原有事件，并对时间控件和数目变量进行初始化
Timer2.Interval = 0
Map1.TrackingLayer.ClearEvents
 m_nNum = 0
'依次记录所要跟踪直线的各个节点
 Set m_ptsDemo = New mapobjects2.Points
 Set line = g_line1
 Dim pt As mapobjects2.Point
  For Each pt In line.Parts(0)
    m_ptsDemo.Add pt
 Next pt
'将各个结点加入事件中，时间的interval设为500ms
 p2.x = m_ptsDemo.Item(0).x
 p2.y = m_ptsDemo.Item(0).y
    Set m_Evnt = Map1.TrackingLayer.AddEvent(p2, 0)
    m_blsMove = True
    Timer2.Interval = 500
    
End Sub


'清楚最短路径和事件，并且将起始点和终止点组合框置空
Private Sub Command4_Click()
    Set g_line1 = Nothing
    Set pts1 = Nothing
    Combo4.Text = ""
    Combo5.Text = ""
    Map1.TrackingLayer.ClearEvents
    Map1.Refresh

End Sub


'缓冲区生成
Private Sub Command5_Click()
    If Combo6.Text = "" Then
        MsgBox "没有选择空间实体类型", vbInformation
        Exit Sub
    End If
    If Combo7.Text = "" Then
        MsgBox "没有选择空间实体", vbInformation
        Exit Sub
    End If
    If Text2.Text = "" Then
        MsgBox "没有输入缓冲区半径", vbInformation
        Exit Sub
    End If
    Dim recs As mapobjects2.Recordset
    
    ' 面缓冲
    If Option5.Value = True Then
        Set recs = Map1.Layers(Combo6.Text).Records
        Do While Not recs.EOF
            If recs.Fields("名称").Value Like Combo7.Text Then
                Dim polygon As mapobjects2.polygon
                Set polygon = recs.Fields("shape").Value
                Set polygon1 = polygon.Buffer(Val(Text2.Text), Map1.FullExtent)
                Map1.Refresh
                Exit Sub
            End If
            recs.MoveNext
        Loop
    End If
    '线缓冲
    If Option4.Value = 1 Then
    End If
    '点缓冲
    If Option3.Value = 1 Then
    End If
    
    
    
End Sub




'清除缓冲区
Private Sub Command6_Click()
    Set polygon1 = Nothing
    Map1.Refresh
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Dim recs As mapobjects2.Recordset
    Set recs = Map1.Layers("其他").Records
   
    If DataGrid1.Columns.count < 10 Then
    Exit Sub
    End If
     '判断是否为空
    If DataGrid1.Columns(0).Text = "" Then
    Exit Sub
    End If
    recs.MoveFirst
    Do While Not recs.EOF
        If recs.Fields("名称").Value Like DataGrid1.Columns(10).Text Then
        Map1.FlashShape recs.Fields("shape").Value, 3
        Exit Sub
        End If
        recs.MoveNext
    Loop
End Sub

'窗体加载
Private Sub Form_Load()
'初始化主窗体显示的位置
Form2.Move 300, 300
'关联数据源
dc.Database = App.path + "\..\" + "数据" '数据源指向存放文件的文件夹,数据和应用程序关联在一起
If Not dc.Connect Then
MsgBox "该文件夹下不存在满足条件的文件"
End
End If

LayerLoad Map1 '向map1中加载图层
LayerLoad Map2 '向map2中加载图层
cnt = 0         '设置初始cnt为0
legend1.setMapSource Map1
legend1.LoadLegend True
Map1.Refresh
'对一些控件是否显示加以设置
TreeView1.Visible = False
Frame3.Visible = False

Adodc1.ConnectionString = DBstr
Adodc1.Visible = False

'默认查询方式为精确查询
Option2.Value = True
'对查询字段进行初始化,默认为姓名
combo1_Init
'对地图索引树进行初始化
TreeView1_Init
'对要选择的空间地物类型进行初始化
combo2_Init
'对比例尺控件进行初始化
strscale = "109202"
Call refreshScale
'对是注记进行初始化
Call renderer
'对起点组合框进行初始化
Combo4_Init
'对终点组合框架进行初始化
combo5_Init

'初始设置缓冲区为面缓冲

Option5.Value = 1

Option3.Enabled = False
Option4.Enabled = False
combo6_Init


'设置跟踪样式
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
Map2.TrackingLayer.Refresh True '将map1和map2关联起来
refreshScale '刷新控件显示
renderer '重新设置注记是否显示
'量测距离时画点
If Not pts Is Nothing Then
    If pts.count >= 1 Then
        Dim sym2 As New Symbol
        sym2.Color = moBlue
        sym2.SymbolType = moPointSymbol
        sym2.Size = 5
        Map1.DrawShape pts, sym2
    End If
    
End If
'量测距离时画线
If Not g_line Is Nothing Then
    If pts.count > 1 Then
        Dim sym As New Symbol
        sym.Color = moRed
        sym.SymbolType = moLineSymbol
        sym.Size = 2
        Map1.DrawShape g_line, sym
    End If
End If

'最短距离时高亮显示
If Not g_line1 Is Nothing Then
    If pts1.count > 1 Then
        Dim sym3 As New Symbol
        sym3.Color = &HFFFF00
        sym3.SymbolType = moLineSymbol
        sym3.Size = 4
        Map1.DrawShape g_line1, sym3
    End If
End If

'生成缓冲区
If Not polygon1 Is Nothing Then
    
        Dim sym4 As New Symbol
        sym4.Color = &HFFFF00
        sym4.SymbolType = moFillSymbol
        sym4.Size = 4
        Map1.DrawShape polygon1, sym4
   
End If


End Sub
'在Map1上按下鼠标执行的操作
Private Sub Map1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'如果没有点击放大按钮，则将cnt归零
If Toolbar1.Buttons(1).Value = 0 Then
cnt = 0
End If

If Toolbar1.Buttons(1).Value = 1 Then '选择的是放大操作,
If Button = vbLeftButton Then '选择的是左键，则执行拉进行框放大
    Set Rectangle(cnt) = Map1.Extent '保存放大前的图幅矩形
    cnt = cnt + 1                   '次数加一
    Set Map1.Extent = Map1.TrackRectangle
    ElseIf Button = vbRightButton Then '如果选择的是右键,则恢复到上次放大前的图幅
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If cnt > 0 Then     '之前进行过放大操作
        Map1.Extent = Rectangle(cnt - 1) '恢复放大前的矩形
        cnt = cnt - 1           '次数减一
    End If
        Exit Sub
End If
ElseIf Toolbar1.Buttons(4).Value = 1 Then  '选择的是平移操作
Map1.Pan
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(2).Value = 1 Then '选择的是缩小操作，则执行拉举行框缩小，以所拉矩形框中心为图中心，缩小1.8倍
    Dim r As mapobjects2.Rectangle
    Set r = Map1.TrackRectangle
    Map1.CenterAt r.Center.x, r.Center.y
    Set r = Map1.Extent
    r.ScaleRectangle 1.8
    Map1.Extent = r
'选择的是标志操作，查找空间实体的属性
ElseIf Toolbar1.Buttons(6).Value = 1 Then
    Dim p As mapobjects2.Point
    Set p = Map1.ToMapPoint(x, y)
    
    '调用用户自定义的空间实体识别函数
    Call Form3.Identify(x, y)
    Form3.ZOrder 0
    Form3.Show
'选择执行测量距离操作
ElseIf Toolbar1.Buttons(7).Value = 1 Then
    '点击左键，则实现画直线量测距离
    If Button = vbLeftButton Then
    '量距离
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
    '点击右键
    
    
    If Button = vbRightButton Then
        '初始距离设为0
        Dim dis As Double
        Dim i As Integer
        dis = 0
        '计算距离
        If pts.count > 1 Then
            For i = 0 To pts.count - 2
            dis = dis + Math.Sqr((pts(i + 1).x - pts(i).x) * (pts(i + 1).x - pts(i).x) + (pts(i + 1).y - pts(i).y) * (pts(i + 1).y - pts(i).y))
            Next i
       End If
        '则显示测量的距离，点击确定后将量测时画的直线删除
        If MsgBox("距离:" + Format(str(dis), "###,###,###") + "m", vbOKOnly, "测距") = vbOK Then
            Set g_line = Nothing
            Set pts = Nothing
            Map1.Refresh
        End If
    End If
End If
End Sub
'在状态栏上显示当前鼠标的位置
Private Sub Map1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim pt As New mapobjects2.Point
Set pt = Map1.ToMapPoint(x, y)
StatusBar1.Panels(2).Text = "x=" & pt.x
StatusBar1.Panels(3).Text = "y=" & pt.y

End Sub

'在map2上画蓝色指示框
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
'画矩形框改变map1的大小
Set r = Map2.TrackRectangle
Set Map1.Extent = r
'点击改变map1的位置
Set pt = Map2.ToMapPoint(x, y)
Map1.CenterAt pt.x, pt.y

End Sub



'点击不同的选项卡时设置一些控件是否显示
Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.index = 2 Then
legend1.Visible = False
TreeView1.Visible = True
ElseIf TabStrip1.SelectedItem.index = 1 Then
TreeView1.Visible = False
legend1.Visible = True
End If
End Sub
'点击不同的选项卡时设置frame控件是否显示，并且属于该frame的其他控件显示与否同frame
Private Sub TabStrip2_Click()
If TabStrip2.SelectedItem.index = 1 Then
Frame3.Visible = False
Frame4.Visible = True
ElseIf TabStrip2.SelectedItem.index = 2 Then
Frame4.Visible = False
Frame3.Visible = True
End If
End Sub

'定时器控件，每隔1秒重新获取系统的时间，并且显示到状态栏中
Private Sub Timer1_Timer()
   StatusBar1.Panels(4).Text = Format(Now, "yyyy-mm-dd hh:mm:ss")
End Sub

Private Sub Timer2_Timer()
   Dim count As Integer
   Dim x, y As Double
   '存在事件，并且确认移动
   If m_blsMove And Map1.TrackingLayer.EventCount > 0 Then
        count = m_nNum
        If count > m_ptsDemo.count Then
        Timer2.Interval = 0
        End If
    Set m_Evnt = Map1.TrackingLayer.Event(0)
     m_nNum = m_nNum + 1
    '依次从直线沿各个节点跟踪显示各个图层
    If count < m_ptsDemo.count Then
    x = m_ptsDemo.Item(count).x
    y = m_ptsDemo.Item(count).y
    m_Evnt.MoveTo x, y
    End If
    End If
End Sub

'单击工具栏上的按钮
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Toolbar1.Buttons(1).Value = 1 Then '点击放大图标
Map1.MousePointer = moZoomIn
'''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(2).Value = 1 Then '点击缩小图标
Map1.MousePointer = moZoomOut
''''''''''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(4).Value = 1 Then '点击平移图标
Map1.MousePointer = moPan
'''''''''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(3).Value = 1 Then '点击恢复图标
Map1.Extent = Map1.FullExtent
Map1.MousePointer = moArrow
''''''''''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(5).Value = 1 Then '点击选择图标
Map1.MousePointer = moArrow
'''''''''''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(6).Value = 1 Then '点击识别图标
Map1.MousePointer = moIdentify
'''''''''''''''''''''''''''''''''''''''''''''''
ElseIf Toolbar1.Buttons(7).Value = 1 Then '点击测距图标
Map1.MousePointer = moCross


End If
End Sub

Private Sub TreeView1_Init()
Dim nodx As Node
Dim cnt As Integer
TreeView1.Style = tvwTreelinesPlusMinusPictureText '树状外观包含全部元素
'建立名称为"兰州市城关区"的父节点,     '''''选择索引为1的图像
Set nodx = TreeView1.Nodes.Add(, , "3区", "兰州市城关区地图")
        
    '在"兰州市城关区"父节点下建立子节点,
   
        Set nodx = TreeView1.Nodes.Add("3区", tvwChild, "child31", "住宅小区")
            cnt = 310
            addlist "child31", "住宅小区", "名称", cnt
        Set nodx = TreeView1.Nodes.Add("3区", tvwChild, "child33", "学校")
            cnt = 330
            addlist "child33", "学校", "名称", cnt
        Set nodx = TreeView1.Nodes.Add("3区", tvwChild, "child34", "医院")
            cnt = 340
            addlist "child34", "医院", "名称", cnt
        Set nodx = TreeView1.Nodes.Add("3区", tvwChild, "child35", "商业区")
            cnt = 350
            addlist "child35", "商业区", "名称", cnt
        Set nodx = TreeView1.Nodes.Add("3区", tvwChild, "child36", "公园")
            cnt = 360
            addlist "child36", "公园", "名称", cnt
    
               

End Sub
'增加各空间实体，参数分别代表父结点，图层名，图层属性表字段，用于生成关键字
Private Sub addlist(str1 As String, str2 As String, str3 As String, cnt As Integer)
        Dim nodx As Node
        Dim str10 As String
        Dim recs As mapobjects2.Recordset
        Set recs = Map1.Layers(str2).Records '打开记录集
        Do While Not recs.EOF
        cnt = cnt + 1
        str10 = "child"
        str10 = str2 + str(cnt)
        If recs.Fields.Item(str3) <> "" Then '当属性记录非空时，则添加到树形控件中
             Set nodx = TreeView1.Nodes.Add(str1, tvwChild, str10, recs.Fields.Item(str3))
        End If
        recs.MoveNext
        Loop
End Sub
'向combo2中添加空间地物类型
Private Sub combo2_Init()
Combo2.AddItem ("公园")
Combo2.AddItem ("学校")
Combo2.AddItem ("商业区")
Combo2.AddItem ("住宅小区")
Combo2.AddItem ("医院")
Combo2.AddItem ("公园")

 
End Sub
'本来可以利用比较简洁的代码书写，主要是树不同的组织层次和所使用数据的不一致性造成的，此处针对自己所写的数据，没有通用性
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
 Dim recs As mapobjects2.Recordset
'如果选择的节点为末尾节点，则
If Node.Children <= 0 Then
    '如果选择的是其他
    If Node.Parent.Text = "其他" Then
    Set recs = Map1.Layers(Node.Text).Records
     Do While Not recs.EOF
        If recs.Fields("名称") Like Node.Text Then '如果匹配，则闪烁三次
        '如果该地物在地图之外,则一直图的中心加以显示
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
    
    
    Set recs = Map1.Layers(Node.Parent.Text).Records '打开相应空间实体类型的记录集
    Do While Not recs.EOF
        If recs.Fields("名称") Like Node.Text Then '如果匹配，则闪烁三次
        '如果该地物在地图之外,则一直图的中心加以显示
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




'比例尺控件的初始化
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
    '书写比例尺的形式
    StatusBar1.Panels(1).Text = "比例尺： 1：" & Format(ScaleBar1.RFScale, "###,###,###,###,###")
    '记录当前的比例尺，用于决定注记的显示与否
    strscale = ScaleBar1.RFScale
  End Sub
'标记，以1：30000为分界线决定是否标记
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
            Map1.Layers(i).renderer.Field = "名称"
    Next
End If
End Sub
'''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''
'首先根据给定的点寻找最近的弧段的id号
Public Function GetLineId(x As Double, y As Double) As Integer

Dim rs As mapobjects2.Recordset
Dim pt3 As New mapobjects2.Point
pt3.x = x
pt3.y = y
'找到弧段的记录集
Set rs = Map1.Layers("road").Records
Dim dis As Double
dis = 99999
Dim id As Integer
id = -1
rs.MoveFirst
'对弧段依次进行判断，找出距离该点最近的弧段
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
'返回该弧段的id号
If id <> -1 Then
    GetLineId = id
End If
End Function

'求点到最近弧段的最近的点的数组编号，距离弧段前节点的距离，
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

'向combo6中添加空间地物类型
Private Sub combo6_Init()
Combo6.AddItem ("住宅小区")
Combo6.AddItem ("学校")
Combo6.AddItem ("医院")
Combo6.AddItem ("商业区")
Combo6.AddItem ("其他")
Combo6.AddItem ("公园")

End Sub

