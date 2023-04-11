VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Dim oAcadApp As Object
    Dim oAcadDoc As Object
    Dim strTiffFullName As String, strDwgHead As String, strDwgFullName As String
    Dim PrnDrvName As String
    Dim PrnSize As String
   Dim PrnRotate As String
   Dim blnret As String
    PrnDrvName = "Bullzip PDF Printer"
    
     On Error Resume Next
     
     Set oAcadApp = GetObject(, "AutoCAD.Application")
   If Err Then
      Set oAcadApp = CreateObject("AutoCAD.Application.18")
      oAcadApp.WindowState = 2
      If Err = 0 Then
         'Call subSetLogData("Create Autocad")
         'Call subSetErrMsg("AutoCAD不能访问")
         'Call subSetErrMsg("应用程序服务器的AutoCAD停止了吗？请检查")
         Exit Sub
      End If
   End If
   oAcadApp.Visible = True
   
   strDwgFullName = "\\10.105.10.80\DwgServer\MhsAcadMmf\JOB\ZEA1338\CP004\ZEA1338_CP004_0101.dwg"
   strTiffFullName = "C:\Users\040283\Desktop\test.tif"
   
   Set oAcadDoc = oAcadApp.Documents.Open(strDwgFullName, True)
   Call acfPlotLayout(oAcadDoc, PrnDrvName, PrnSize, PrnRotate)
   oAcadDoc.Activate
   With oAcadDoc.Plot
        .QuietErrorMode = True
        blnret = .PlotToDevice
        'Sleep 10000
                    
   End With
End Sub

   
Public Sub acfPlotLayout( _
  ByVal oAcadDoc As Object, ByVal PrnDrvName As String, _
  Optional PrnSize As String, Optional PrnRotate As String _
)
'------------------------------------------------------------
'[Param]
' Document:    AcadDocument
' PrnDrvName:   Printer driver name
' PrnSize:      drawing size (option)
' PrnRotate:    Rotation (option)
'------------------------------------------------------------
'[Subroutines]
' acfGetDSize, acfGetDScale
'------------------------------------------------------------
'============================================================
  Dim DSize As String, Dscale As String
  Dim vMinPoint, vMaxPoint

 'Dwg frame size and Dwg scale settings
  DSize = acfGetDSize(oAcadDoc)
  If Right$(DSize, 1) Like "[1234]" Then
    'DScale = acfGetDScale(acdDoc, DSize)
    'Dscale = acfGetMecTitScale(acdDoc)            'Change by F/ST
    Dscale = oAcadDoc.GetVariable("DIMSCALE")       'Change by F/ST
  End If

 'Plot Settings
  With oAcadDoc.ModelSpace.Layout
    .RefreshPlotDeviceInfo
    .ConfigName = PrnDrvName
    If DSize Like "*1" Then     '仧A1
      .CanonicalMediaName = IIf(PrnSize <> "", PrnSize, "A3")
      .PlotRotation = IIf(PrnRotate <> "", PrnRotate, 1) '1 -> ac90degrees
      '.PlotType = 2      'acLimits
      '.StandardScale = 1 'acVpCustomScale
      .PlotType = 1      'acExtents
      .StandardScale = 0 'acScaleToFit
'      If DScale < 0 Then
'        .SetCustomScale -DScale * 2, 1
'      Else
'        .SetCustomScale 1, DScale * 2
'      End If
'      If DScale < 0 Then
'        .SetCustomScale -DScale, 1
'      Else
'        .SetCustomScale 1, DScale
'      End If
    ElseIf DSize Like "*2" Then '仧A2
      .CanonicalMediaName = IIf(PrnSize <> "", PrnSize, "A3")
      .PlotRotation = IIf(PrnRotate <> "", PrnRotate, 1) '1 -> ac90degrees
'      .PlotType = 2      'acLimits
'      .StandardScale = 0 'acScaleToFit
      .PlotType = 1      'acExtents
      .StandardScale = 0 'acScaleToFit
    ElseIf DSize Like "*3" Then '仧A3
      .CanonicalMediaName = IIf(PrnSize <> "", PrnSize, "A3")
      .PlotRotation = IIf(PrnRotate <> "", PrnRotate, 1) '1 -> ac90degrees
'      .PlotType = 2      'acLimits
'      .StandardScale = 1 'acVpCustomScale
      .PlotType = 1      'acExtents
      .StandardScale = 0 'acScaleToFit
'      If DScale < 0 Then
'        .SetCustomScale -DScale, 1
'      Else
'        .SetCustomScale 1, DScale
'      End If
    ElseIf DSize Like "*4" Then '仧A4
      .CanonicalMediaName = IIf(PrnSize <> "", PrnSize, "A4")
      .PlotRotation = IIf(PrnRotate <> "", PrnRotate, 0) '0 -> ac0degrees
'      .PlotType = 2      'acLimits
'      .StandardScale = 1 'acVpCustomScale
      .PlotType = 1      'acExtents
      .StandardScale = 0 'acScaleToFit
'      If DScale < 0 Then
'        .SetCustomScale -DScale, 1
'      Else
'        .SetCustomScale 1, DScale
'      End If
    Else                         ' Unidentifiable
  
      vMinPoint = oAcadDoc.GetVariable("EXTMIN")
      vMaxPoint = oAcadDoc.GetVariable("EXTMAX")
   
      If (vMaxPoint(0) - vMinPoint(0)) > _
         (vMaxPoint(1) - vMinPoint(1)) Then ' Horizontal
        .CanonicalMediaName = "A3"
        .PlotRotation = 0  'ac0degrees
      Else                                  ' Vertical
        .CanonicalMediaName = "A4"
        .PlotRotation = 1  'ac90degrees
      End If
      .PlotType = 1        'acExtents
      .StandardScale = 0   'acScaleToFit
    End If
    .CenterPlot = True
    '.StyleSheet = g_PltStyle '(->from INI file)
    .PlotWithPlotStyles = True
    .PaperUnits = 1        'acMillimeters
  End With
  

  oAcadDoc.Regen 1  'acAllViewports

End Sub

Public Function acfGetDSize(oAcadDoc As Object) As String
'------------------------------------------------------------
'[Param]
' acdDoc:   恾柺僆僽僕僃僋僩 : AcadDocument
' RetVal:   恾柺榞柤 : Drawing size
'============================================================
  
  Dim i As Long
  Dim dPX#(1 To 4), dPY#(1 To 4)
  Dim dLX#, dLY#
  Dim vLimits
  Dim Dscale As String
   
 '仭恾柺広搙庢摼
 'Get Dwg Scale
    'Dscale = acfGetMecTitScale(acdDoc)            'Change by F/ST
    Dscale = oAcadDoc.GetVariable("DIMSCALE")      'Change by F/ST
  
 '仭昗弨梡巻僒僀僘
 'Standard paper size
  dPX#(1) = 841# * Dscale: dPY#(1) = 594# * Dscale  'A1 Horizontal
  dPX#(2) = 594# * Dscale: dPY#(2) = 420# * Dscale  'A2 Horizontal
  dPX#(3) = 409# * Dscale: dPY#(3) = 282# * Dscale  'A3 Horizontal
  dPX#(4) = 196# * Dscale: dPY#(4) = 286# * Dscale  'A4 Vertical
  'dPX#(4) = 297# * DScale: dPY#(4) = 210# * DScale  'A4 Horizontal
  
 '仭弌椡斖埻偺庢摼
 'Get output range
  vLimits = oAcadDoc.Limits
  dLX# = vLimits(2) - vLimits(0)
  dLY# = vLimits(3) - vLimits(1)
  
 '仭梡巻僒僀僘偺摿掕
 'Determine the paper size
  If dLX# >= dPX#(1) And dLY# >= dPY#(1) Then       '仦A1
    acfGetDSize = "A1"
  ElseIf dLX# >= dPX#(2) And dLY# >= dPY#(2) Then   '仦A2
    acfGetDSize = "A2"
  ElseIf dLX# >= dPX#(3) And dLY# >= dPY#(3) Then   '仦A3
    acfGetDSize = "A3"
  ElseIf dLX# >= dPX#(4) And dLY# >= dPY#(4) Then   '仦A4
    acfGetDSize = "A4"
  Else                                              '仦敾掕晄壜 Unidentifiable
    acfGetDSize = ""
  End If
  'ReDetermine the paper size
  If acfGetDSize = "" Then
    If Round(dLX# - dPX#(1)) = 0 And Round(dLY# - dPY#(1)) = 0 Then      'A1
        acfGetDSize = "A1"
    ElseIf Round(dLX# - dPX#(2)) = 0 And Round(dLY# - dPY#(2)) = 0 Then  'A2
        acfGetDSize = "A2"
    ElseIf Round(dLX# - dPX#(3)) = 0 And Round(dLY# - dPY#(3)) = 0 Then  '仦A3
        acfGetDSize = "A3"
    ElseIf Round(dLX# - dPX#(4)) = 0 And Round(dLY# - dPY#(4)) = 0 Then  '仦A4
        acfGetDSize = "A4"
    Else                                                                 '仦敾掕晄壜 Unidentifiable
        acfGetDSize = ""
    End If
  End If
   
End Function

   
     
    
    
