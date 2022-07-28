VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UITranspEdge 
   Caption         =   "Transparent Edge :: to fix displacement assign an outline"
   ClientHeight    =   675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580.001
   OleObjectBlob   =   "UITranspEdge.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UITranspEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const sTranspEdge$ = "TransparentEdges"
Private bUndo%, bWorking%, hUI&

Const TM_LIVE& = 61378, TM_CLOSE& = 61379


Friend Sub onTimer(ByVal hWnd&, ByVal idEvent&)
   If bWorking Then bWorking = 2: Exit Sub
   KillTimer hWnd, idEvent
   Select Case idEvent
      Case TM_LIVE:  ApplyTransparentEdge txFeather
      Case TM_CLOSE: Unload Me
   End Select
End Sub

Private Sub btnApply_Click()
   SaveSetting sCorelDRAW, sTranspEdge, "Feather", txFeather
   ApplyTransparentEdge txFeather
   End Sub

Private Sub btnCancel_Click()
   If bWorking Then bWorking = 2: SetTimer hUI, TM_CLOSE, 50, AddressOf TimerTransparentEdge _
               Else hUI = 0: Unload Me
   End Sub

Private Sub btnFix_Click()
   Dim Shape As Shape, clr As New Color
   If ActiveShape Is Nothing Then Beep: Exit Sub
   On Error Resume Next
   Optimization = True
   EventsEnabled = False
   ActiveDocument.SaveSettings
   ActiveDocument.PreserveSelection = False
   ActiveDocument.BeginCommandGroup sTranspEdge & ": FIX"
   
   For Each Shape In ActiveSelectionRange
      If Shape.Outline.Type = cdrNoOutline Then
         Select Case Shape.Fill.Type
            Case cdrUniformFill:    clr.CopyAssign Shape.Fill.UniformColor
            Case cdrFountainFill:   clr.CopyAssign Shape.Fill.Fountain.EndColor
            Case Else:              clr.CMYKAssign 0, 0, 0, 100
         End Select
         Shape.Outline.Type = cdrOutline
         Shape.Outline.SetProperties ActiveDocument.ToUnits(0.001, cdrMillimeter), , clr
      End If
   Next
   
   ActiveDocument.EndCommandGroup
   ActiveDocument.PreserveSelection = True
   ActiveDocument.RestoreSettings
   EventsEnabled = True
   Optimization = False
   Application.CorelScript.RedrawScreen
   End Sub

Private Sub btnOK_Click()
   btnApply_Click
   Unload Me
   End Sub

Private Sub cbDPI_Change()
   cbDPI.Text = CLng(cbDPI.Text)
   cbDPI.Text = IIf(cbDPI.Text <= 0 Or cbDPI.Text > 5000, 300, cbDPI.Text)
   spinD.Value = cbDPI.Text
   If chkAuto Then SetTimer hUI, TM_LIVE, 250, AddressOf TimerTransparentEdge
   End Sub

Private Sub chkAuto_Change()
   If chkAuto Then ApplyTransparentEdge txFeather
   End Sub

Private Sub spinD_Change(): cbDPI = spinD.Value: End Sub
Private Sub spinF_Change(): txFeather = spinF.Value: End Sub

Private Sub txFeather_Change()
   On Error Resume Next
   txFeather = CLng(IIf(txFeather < 0 Or txFeather > 99, 5, txFeather))
   spinF.Value = txFeather
   If chkAuto Then SetTimer hUI, TM_LIVE, 250, AddressOf TimerTransparentEdge
   End Sub

Private Sub UserForm_Activate()
   DoEvents
   chkUndo = (Trim$(GetSetting(sCorelDRAW, sTranspEdge, "AutoUndo", "1")) = "1")
   chkAuto = (Trim$(GetSetting(sCorelDRAW, sTranspEdge, "AutoApply", "1")) = "1")
   End Sub

Private Sub UserForm_Initialize()
   hUI = UIPositionFromRegistry(Me, "TransparentEdges")
   
   txFeather = Trim$(GetSetting(sCorelDRAW, sTranspEdge, "Feather", "5"))
   txFeather = IIf(txFeather < 0 Or txFeather > 99, 5, txFeather)
   txFeather.SelStart = 0: txFeather.SelLength = Len(txFeather)
   
   cbDPI.List = Split("72 96 120 150 200 240 300 350 400 600 1200", " ")
   cbDPI.Text = GetSetting(sCorelDRAW, sTranspEdge, "dpi", "300")
   cbDPI_Change
   txFeather_Change
   'If chkAuto Then ApplyTransparentEdge txFeather
   End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   If bWorking Then bWorking = 2 Else hUI = 0
   SaveSetting sCorelDRAW, sTranspEdge, "WindowPos", Left & " " & Top
   SaveSetting sCorelDRAW, sTranspEdge, "AutoApply", IIf(chkAuto, "1", "0")
   SaveSetting sCorelDRAW, sTranspEdge, "AutoUndo", IIf(chkUndo, "1", "0")
   End Sub
            
Sub ApplyTransparentEdge(ByVal feather&)
   Dim origSR As New ShapeRange, Shapes As ShapeRange, Shape As Shape, sh2 As Shape, toSel As New ShapeRange
   Dim eff As Effect, tmp$, s$, cnt&, idx&, q#, step#, dpi&
   
   If bWorking Then Exit Sub
   bWorking = True
   
   Set origSR = ActiveSelectionRange
   If origSR.Count = 0 Then Exit Sub
   
   On Error Resume Next
   
   If feather < 0 Or feather > 99 Then Exit Sub
   dpi = cbDPI
   dpi = IIf(dpi <= 0 Or dpi > 5000, 300, dpi)
   
   If chkUndo And bUndo Then ActiveDocument.unDo
   
   BoostStart sTranspEdge & ": " & feather
   
   Application.Status.BeginProgress sTranspEdge, True
   
   cnt = origSR.Count: step = 100# / CDbl(cnt)
   For Each Shape In origSR
      q = 100# * CDbl(idx) / CDbl(cnt)
      idx = idx + 1
      s = sTranspEdge & idx & " / " & cnt & ": "
      
      If feather = 0 Then
         Shape.Transparency.ApplyNoTransparency
         GoTo NextShape
         End If
     
      If VersionMajor > 12 _
         Then Set eff = Shape.CreateDropShadow(cdrDropShadowFlat, 100, feather, 0#, 0#, CreateCMYKColor(0, 0, 0, 100), cdrFeatherInside, cdrEdgeLinear, IIf(Shape.Type = cdrBitmapShape, 9, cdrMergeNormal)) _
         Else Set eff = Shape.CreateDropShadow(cdrDropShadowFlat, 100, feather, 0#, 0#, CreateCMYKColor(0, 0, 0, 100), cdrFeatherInside, cdrEdgeLinear)
      DoEvents: If bWorking = 2 Then Exit For
      
      Set Shapes = eff.Separate: Set eff = Nothing: DoEvents: If bWorking = 2 Then Exit For
      
      With Application.Status: .message = s & "1/4": .Progress = q + 0.15 * step: If .Aborted Then Exit For
      End With
      
      Set sh2 = Shapes(1).ConvertToBitmap(8, True, , False, dpi, cdrNormalAntiAliasing, False)
         DoEvents: If bWorking = 2 Then Exit For

      With Application.Status: .message = s & "2/4": .Progress = q + 0.32 * step: If .Aborted Then Exit For
      End With
      
      sh2.ApplyEffectInvert
         DoEvents: If bWorking = 2 Then Exit For

      tmp = Environ$("temp"): If Right$(tmp, 1) <> "\" Then tmp = tmp + "\"
      tmp = tmp + Hex(Timer) + ".tif"
      
      sh2.Bitmap.SaveAs(tmp, cdrTIFF, cdrCompressionLZW).Finish
         DoEvents: If bWorking = 2 Then Exit For

      sh2.Delete
      Err.Clear: FileSystem.GetAttr tmp: If Err.Number Then Debug.Print "Error: " + tmp: GoTo NextShape
      
      With Shapes(2).Transparency.ApplyPatternTransparency(cdrBitmapPattern, tmp, 0, 0, 100, True)
         Application.Status.message = s & "3/4": Application.Status.Progress = q + 0.5 * step:: If Application.Status.Aborted Then Exit For
         DoEvents: If bWorking = 2 Then Exit For
         
         .OriginX = 0#
         .OriginY = 0#
         .TileWidth = Shape.SizeWidth
         
         Application.Status.message = s & "4/4": Application.Status.Progress = q + 0.8 * step:: If Application.Status.Aborted Then Exit For
         
         .TileHeight = Shape.SizeHeight
      End With
         
      DoEvents: If bWorking = 2 Then Exit For

      With Shapes(2).Transparency
         If VersionMajor = 13 Then .MergeMode = cdrMergeNormal
         .AppliedTo = cdrApplyToFillAndOutline
      End With
      toSel.Add Shapes(2)
      FileSystem.Kill tmp

NextShape:
   Next Shape
   
   Application.Status.EndProgress
   
   toSel.CreateSelection
   BoostFinish EndUndoGroup:=True
   bWorking = 0
   bUndo = True
   End Sub

