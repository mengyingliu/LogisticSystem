VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "属性窗口"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2445
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   2445
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   1680
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label location 
      Caption         =   "location"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label shape 
      Caption         =   "shapetype"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3900
      Width           =   2055
   End
   Begin VB.Label Layer 
      Caption         =   "Layer"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "属性"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label findNum 
      Caption         =   "Label1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const SEARCHTOLPIXELS = 3
Dim Loc As New Point
Dim Recs2() As mapobjects2.Recordset
Dim strlayerName() As String
Dim layerNum() As Long

Sub Identify(x As Single, y As Single)
   
   Dim curCount As Long, layerCount As Long, layer_c As Long
   Dim theTol As Double
   Dim featCount As Long, fCount As Long
   Dim objaLayer As Object
   Dim recs As mapobjects2.Recordset
   Dim straName As String, strtheItem As String
   Dim objaField As Object
   
   layer_c = Form2.Map1.Layers.Count
   ReDim strlayerName(layer_c)
   ReDim Recs2(layer_c)
   
   Screen.MousePointer = 11
   
   Combo1.Clear
   List1.Clear
    Set Loc = Form2.Map1.ToMapPoint(x, y)
   
   Dim xStr As String, yStr As String
   
   If Loc.x > 1000 Or Loc.y > 1000 Then
      xStr = Int(Loc.x)
      yStr = Int(Loc.y)
   Else
      xStr = Loc.x
      yStr = Loc.y
   End If
   
   location.Caption = "位置坐标：(" & xStr & "," & yStr & ")"
   featCount = 0
   layerCount = -1
   theTol = Form2.Map1.ToMapDistance(SEARCHTOLPIXELS * Screen.TwipsPerPixelX)
   
   For Each objaLayer In Form2.Map1.Layers
      If objaLayer.Visible And objaLayer.LayerType = moMapLayer Then
         Set recs = objaLayer.SearchByDistance(Loc, theTol, "")
         layerCount = layerCount + 1
         strlayerName(layerCount) = objaLayer.Name
         Set Recs2(layerCount) = recs
         curCount = -1
        
         If recs.Count <> 0 Then
            straName = "Featureid"
            
            '有点疑问？定义的不是数组，怎么保留每个值
            For Each objaField In recs.Fields
               If objaField.Type = moString Then
                  straName = objaField.Name
                  Exit For
               End If
            Next
            
            While Not recs.EOF
               ReDim Preserve layerNum(2, featCount + 1)
               curCount = curCount + 1
               layerNum(1, featCount) = layerCount
               layerNum(2, featCount) = curCount
               featCount = featCount + 1
               strtheItem = recs(straName).ValueAsString
               If strtheItem = "" Then
                  Combo1.AddItem recs("FeatureId").ValueAsString
               Else
                  Combo1.AddItem strtheItem
               End If
               recs.MoveNext
            Wend
         End If
      End If
   
   Next objaLayer
   
   If featCount = 0 Then
   findNum.Caption = "没有找到任何空间实体"
   Else
   findNum.Caption = Str(featCount) + "个对象被找到"
   End If
   
      
   If featCount > 0 Then
      Combo1.ListIndex = 0
      Call Identify_list
   End If
   
   Screen.MousePointer = 0
   
End Sub

Sub Identify_list()
   
   Dim curRec As mapobjects2.Recordset
   Dim curIndex As Long, aIndex As Long, aRec As Long, i As Long
   Dim objaField As Object
   Dim straName As String
   Dim obj As Object
   
   curIndex = Combo1.ListIndex
   
   If IsNull(Combo1.List(aIndex)) Then
      MsgBox "未选中目标", vbOKOnly, "图层"
      Exit Sub
   End If
   
   'ReDim layerNum(UBound(layerNum, 1), UBound(layerNum, 2))
   aIndex = layerNum(1, curIndex)
   aRec = layerNum(2, curIndex)
   straName = strlayerName(aIndex)
   
   Set curRec = Recs2(aIndex)
   curRec.MoveFirst
   
   If aRec > 0 Then
      For i = 1 To aRec
         curRec.MoveNext
      Next i
   End If
   
   Form2.Map1.FlashShape curRec("shape").Value, 2
    
    Layer.Caption = "所属图层：" + straName
    List1.Clear
  For Each objaField In curRec.Fields
    Select Case objaField.Type
    Case moString
      List1.AddItem objaField.Name + " = " + objaField.Value
    Case moPoint
      shape.Caption = "空间实体类型：点"
    Case moLine
      shape.Caption = "空间实体类型：线"
    Case moPolygon
      shape.Caption = "空间实体类型：面"
    Case Else
      List1.AddItem objaField.Name + " = " + objaField.ValueAsString
    End Select
  Next objaField
   
End Sub
'现则复选框的一项时，在list中显示其属性
Private Sub combo1_Click()
    Identify_list
End Sub



'窗体加载时属性窗口的位置
Private Sub Form_Load()

   Form3.Move Form2.Left + Form2.Width, Form2.Top
      If (Form3.Left + Form3.Width) > Screen.Width Then
   Form3.Left = Screen.Width - Form3.Width
   End If

End Sub

