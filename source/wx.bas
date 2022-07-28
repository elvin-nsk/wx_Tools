Attribute VB_Name = "wx"
'===============================================================================
'   Макрос          : wOxxOm's Tools
'   Версия          : 2022.07.28
'   Сайты           : https://vk.com/elvin_macro/wx_Tools
'                     https://github.com/elvin-nsk/wx_Tools
'   Автор оригинала : wOxxOm
'   Поддерживается  : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = False

'===============================================================================

Sub BitmapsDownsample()
    Static dpi&
    Dim Shapes As ShapeRange, Shape As Shape
    Dim sr2 As ShapeRange, SelectedShapes As ShapeRange, Page As Page
    Dim dpi0&, cnt&, cvt&(), i&, j&, s$, bStat%, t0!, x#, y#, w#, h#
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveDocument Is Nothing Then Exit Sub
    If dpi <= 0 Then dpi = 300
    s = InputBox( _
        "Resolution DPI, " & vbCr _
      & "applied if Source dpi > 120%", "downsample bitmaps", _
        IIf(dpi <= 0, 300, dpi) _
    )
    If Len(s) = 0 Then Exit Sub
    dpi = Val(s): dpi0 = (dpi * 120) \ 100
    
    Set Shapes = New ShapeRange
    Set sr2 = New ShapeRange
    Set SelectedShapes = New ShapeRange
    Shapes.AddRange ActiveSelection.Shapes.FindShapes
    If Shapes.Count = 0 Then
        For Each Page In ActiveDocument.Pages
            Shapes.AddRange Page.FindShapes
        Next Page
    End If
    cnt = Shapes.Count: ReDim cvt(ActiveDocument.Pages.Count)
    
    BoostStart "Downsample bitmaps", RELEASE
    
    ActiveDocument.ReferencePoint = cdrCenter
    bStat = (Application.VersionMajor > 11)
    If bStat Then Application.Status.BeginProgress CanAbort:=True
    Do
        For Each Shape In Shapes
        i = i + 1
        If bStat Then
            If Timer - t0 > 0.1! Then
                t0 = Timer
                Application.Status.Progress = i / cnt * 100
            End If
            If Application.Status.Aborted Then Exit For
        End If
        If Shape.Type = cdrBitmapShape Then
            With Shape.Bitmap
                If (.ResolutionX > dpi0) Or (.ResolutionY > dpi0) Then
                    If .ResolutionX * .ResolutionY > 0 Then
                        Shape.GetPosition x, y: Shape.GetSize w, h
                        .Resample _
                            Round(dpi / .ResolutionX * .SizeWidth + 0.49999), _
                            Round(dpi / .ResolutionY * .SizeHeight + 0.49999), _
                            True, dpi, dpi                              'X3
                        Shape.SetSize w, h: Shape.SetPosition x, y
                        j = Shape.Layer.Page.Index And &HFFFF
                        cvt(j) = cvt(j) + 1
                        SelectedShapes.Add Shape
                    End If
                End If
            End With
        End If
        If Not Shape.PowerClip Is Nothing Then
            With Shape.PowerClip.Shapes.FindShapes
                cnt = cnt + .Count: sr2.AddRange .All
            End With
        End If
        Next Shape
        Shapes.RemoveAll
        Shapes.AddRange sr2
        sr2.RemoveAll
        If Application.Status.Aborted Then Exit Do
    Loop Until Shapes.Count = 0
    
    If SelectedShapes.Count Then
        SelectedShapes.CreateSelection
    Else
        ActiveDocument.ClearSelection
    End If
    If bStat Then Application.Status.EndProgress
    
    j = 0
    For i = 1 To UBound(cvt)
        If cvt(i) Then
            j = j + cvt(i): s = s & "Page." & i & ": " & cvt(i) & ", " & vbTab
        End If
    Next
    MsgBox IIf(Len(s), s, "No bitmaps processed"), , _
           "Downsampled " & j & " bitmaps to " & dpi & " dpi"
    
Finally:
    BoostFinish EndUndoGroup:=True
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub BitmapsResizer()
    Static sz$
    Dim Shape As Shape, Shapes As ShapeRange, sr2 As ShapeRange
    Dim s$, x#, y#, Page As Page, bAll%, tPClip&
    
    If RELEASE Then On Error GoTo Catch
    
    If Len(sz) = 0 Then
        sz = GetSetting(sCorelDRAW, "BitmapsResizer", "Settings")
    End If
    s = Trim$(InputBox(Replace( _
       "SizeX SizeY [AP]|0 = proportional|A = all pages|P = check powerclips||EXAMPLE: 100 0 ap", "|", vbCr & vbTab), _
       "wx.BitmapsResizer (no resample)", sz))
    If Len(s) = 0 Then Exit Sub
    x = Left$(s, InStr(s, " "))
    y = Split(LTrim$(Mid$(s, InStr(s, " "))), " ")(0)
    bAll = InStr(1, s, "a", vbTextCompare) <> 0
    tPClip = IIf(InStr(1, s, "Page", vbTextCompare) <> 0, 0, cdrBitmapShape)
    If Err.Number Or x < 0 Or y < 0 Then
        MsgBox "Bad number", vbExclamation: Exit Sub
    End If
    SaveSetting sCorelDRAW, "BitmapsResizer", "Settings", s
    sz = s
    
    Set Shapes = ActiveSelection.Shapes.FindShapes(, tPClip)
    If tPClip = 0 Then Set sr2 = New ShapeRange
       
    ActiveDocument.Unit = ActiveDocument.Rulers.HUnits
    
    BoostStart "resize bitmaps: " & x & "x" & y & UnitName(ActiveDocument.Unit) _
             & IIf(bAll, " All pages", vbNullString) _
             & IIf(tPClip = 0, " (also PowerClips)", vbNullString), _
               RELEASE
    
    For Each Page In ActiveDocument.Pages
        If bAll Or Page Is ActivePage Then
            If bAll Or Shapes.Count = 0 Then
                Set Shapes = Page.FindShapes(, tPClip)
            End If
            Do
                For Each Shape In Shapes
                    If Shape.Type = cdrBitmapShape Then
                        If x = 0 Then
                            If Abs(Shape.SizeHeight - y) > 0.000001 Then _
                            Shape.SetSize _
                                Shape.SizeWidth * y / Shape.SizeHeight, y
                        ElseIf y = 0 Then
                            If Abs(Shape.SizeWidth - x) > 0.000001 Then _
                                Shape.SetSize _
                                x, Shape.SizeHeight * x / Shape.SizeWidth
                        ElseIf Abs(Shape.SizeWidth - x) > 0.000001 _
                            Or Abs(Shape.SizeHeight - y) > 0.000001 Then
                            Shape.SetSize x, y
                        End If
                    End If
                    If tPClip = 0 Then
                        If Not Shape.PowerClip Is Nothing Then
                            sr2.AddRange Shape.PowerClip.Shapes.FindShapes
                        End If
                    End If
                Next Shape
                If tPClip <> 0 Then Exit Do
                Shapes.RemoveAll: Shapes.AddRange sr2: sr2.RemoveAll
            Loop Until Shapes.Count = 0
        End If
    Next Page
    
    ActiveDocument.ClearSelection
    
Finally:
    BoostFinish EndUndoGroup:=True
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub
                           
Sub BitmapsSetDPI()
    Static dpi&
    Dim Shapes As ShapeRange, Shape As Shape, sr2 As ShapeRange
    Dim SelectedShapes As ShapeRange, Page As Page
    Dim dpi0&, cnt&, cvt&(), i&, j&, s$, bStat%, t0!
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveDocument Is Nothing Then Exit Sub
    
    s = InputBox( _
            "Set resolution DPI to" & vbCr & vbCr & _
            "DRAW may display wrong dpi on some bitmaps especially from ai/eps/pdf. Ignore it. Corel's bug", _
            "Change dpi (NO resample)", _
            IIf(dpi <= 0, 300, dpi) _
        )
    If Len(s) = 0 Then Exit Sub
    dpi = Val(s)
    
    Set Shapes = New ShapeRange
    Set sr2 = New ShapeRange
    Set SelectedShapes = New ShapeRange
    Shapes.AddRange ActiveSelection.Shapes.FindShapes
    If Shapes.Count = 0 Then
        For Each Page In ActiveDocument.Pages
            Shapes.AddRange Page.FindShapes
        Next Page
    End If
    cnt = Shapes.Count: ReDim cvt(ActiveDocument.Pages.Count)
    
    BoostStart "Bitmaps set dpi", RELEASE
    
    ActiveDocument.ReferencePoint = cdrCenter
    ActiveDocument.Unit = cdrInch
    bStat = (Application.VersionMajor > 11)
    If bStat Then Application.Status.BeginProgress CanAbort:=True
    Do
        For Each Shape In Shapes
            i = i + 1
            If bStat Then
                If Timer - t0 > 0.1! Then
                    t0 = Timer
                    Application.Status.Progress = i / cnt * 100
                    If Application.Status.Aborted Then Exit For
                End If
            End If
            If Shape.Type = cdrBitmapShape Then
               With Shape.Bitmap
                  If (.ResolutionX <> dpi) Or (.ResolutionY <> dpi) Then
                     If .ResolutionX * .ResolutionY > 0 Then
                        Shape.SetSize _
                            Shape.Bitmap.SizeWidth / dpi, _
                            Shape.Bitmap.SizeHeight / dpi
                        j = Shape.Layer.Page.Index And &HFFFF
                        cvt(j) = cvt(j) + 1
                        SelectedShapes.Add Shape
                     End If
                  End If
               End With
            End If
            If Not Shape.PowerClip Is Nothing Then
                With Shape.PowerClip.Shapes.FindShapes
                    cnt = cnt + .Count
                    sr2.AddRange .All
                End With
            End If
        Next Shape
        Shapes.RemoveAll
        Shapes.AddRange sr2
        sr2.RemoveAll
        If Application.Status.Aborted Then Exit Do
    Loop Until Shapes.Count = 0
    If SelectedShapes.Count Then
        SelectedShapes.CreateSelection
    Else
        ActiveDocument.ClearSelection
    End If
    If bStat Then Application.Status.EndProgress
    
    j = 0: s = ""
    For i = 1 To UBound(cvt)
        If cvt(i) Then
            j = j + cvt(i)
            s = s & "Page." & i & ": " & cvt(i) & ", " & vbTab
        End If
    Next
    MsgBox IIf(Len(s), s, "No bitmaps processed"), , _
           "Set dpi " & dpi & " for " & j & " bitmaps"
    
Finally:
    BoostFinish EndUndoGroup:=True
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub
                           
Sub BitmapsToPowerclips()
    Static Bleed#
    Dim Shapes As ShapeRange, Shape As Shape, r As Shape
    Dim Page As Page, cvt&(), i&, j&, s$, Bleed2#, SelectedShapes As ShapeRange
    If ActiveDocument Is Nothing Then Exit Sub
    
    If RELEASE Then On Error GoTo Catch
    
    s = InputBox( _
            "Bleed, " & UnitName(ActiveDocument.Rulers.HUnits) & vbCr & vbCr & _
            "tip: macro is using default graphic style, change it before running", "Bitmaps to powerclip", Bleed _
        )
    If Len(s) = 0 Then Exit Sub
    Bleed = Val(s)
    If CDbl(s) <> Bleed Then
        If CDbl(s) <> 0 Then Bleed = CDbl(s)
    End If
    Bleed2 = Bleed * 2
    
    Set Shapes = New ShapeRange
    Set SelectedShapes = New ShapeRange
    If ActiveSelectionRange.Count Then
        Shapes.AddRange _
            ActiveSelection.Shapes.FindShapes(, cdrBitmapShape, False)
    Else
        For Each Page In ActiveDocument.Pages
            Shapes.AddRange Page.FindShapes(, cdrBitmapShape, False)
        Next Page
    End If
    ReDim cvt(ActiveDocument.Pages.Count)
    
    BoostStart "Bitmaps to powerclips", RELEASE
    
    ActiveDocument.ReferencePoint = cdrBottomLeft
    ActiveDocument.Unit = ActiveDocument.Rulers.HUnits
    For Each Shape In Shapes
        Set r = Shape.Layer.CreateRectangle2( _
                    Shape.PositionX + Bleed, _
                    Shape.PositionY + Bleed, _
                    Shape.SizeWidth - Bleed2, _
                    Shape.SizeHeight - Bleed2 _
                )
        r.OrderFrontOf Shape
        Shape.AddToPowerClip r
        SelectedShapes.Add r
        i = r.Layer.Page.Index And &HFFFF: cvt(i) = cvt(i) + 1
    Next
    SelectedShapes.CreateSelection
    
    For i = 1 To UBound(cvt)
        If cvt(i) Then
            j = j + cvt(i): s = s & "Page." & i & ": " & cvt(i) & ", " & vbTab
        End If
    Next
    MsgBox IIf(Len(s), s, "No bitmaps processed"), , _
        "Powerclipped " & j & " bitmaps. Bleed=" & Bleed & " " _
      & UnitName(ActiveDocument.Rulers.HUnits)
       
Finally:
    BoostFinish EndUndoGroup:=True
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub BlendSplit() 'X3                           MsgBox "Works only in DrawX3"
    Dim sel0 As Shape, t0!, a#
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveShape Is Nothing Then Beep: Exit Sub
    FrameWork.Automation.Invoke "6dd9cba5-ae47-48e6-9abf-1dbd683da2c7"
    SendKeys " {TAB} "
    DoEvents
    Set sel0 = ActiveShape
    
    BoostStart "fix blend control shape angle", RELEASE
    
    mouse_event 2, 0, 0, 0, 0 'MOUSEEVENTF_LEFTDOWN = &H2
    mouse_event 4, 0, 0, 0, 0 'MOUSEEVENTF_LEFTUP = &H4
    t0 = Timer
    Do While Timer - t0 < 1
        DoEvents
        If Not ActiveShape Is sel0 Then Exit Do
        Loop
    If ActiveShape Is Nothing Then
        MsgBoxEx "Split: resultant control shape cannot be Found, timeout 1s", _
                 vbExclamation, "wx.BlendSplit", , True
    Else
        a = ActiveShape.RotationAngle
        ActiveShape.RotationAngle = ActiveShape.RotationAngle + 1#
        ActiveShape.RotationAngle = a
    End If
    
Finally:
    BoostFinish True
    UIrefresh
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub ConvertShapesToCMYK()
    Dim idx&, Changed&, Bitmaps&, cnt&, res$, ch0&, b0&, s$
    idx = 0: Changed = 0
    Dim Page As Page, oldP As Page, Shapes As ShapeRange
    Dim OriginalSelection As New ShapeRange
    
    If ActiveDocument Is Nothing Then Exit Sub
    OriginalSelection.AddRange ActiveSelectionRange
    
    If RELEASE Then On Error GoTo Catch
    
    Set Shapes = ActiveSelection.Shapes.FindShapes: Set oldP = ActivePage
    
    BoostStart "Convert to CMYK", RELEASE
    
    If VersionMajor > 11 Then Application.Status.BeginProgress CanAbort:=True
    If Shapes.Count = 0 Then
        For Each Page In ActiveDocument.Pages
            Page.Activate: DoEvents
            z_ConvertShapesToCMYKiterate _
                Page.FindShapes, idx, Changed, Bitmaps, cnt
            If Changed - ch0 Or Bitmaps - b0 Then
                res = res & "Page" & Page.Index & IIf(Len(Page.Name), " <" & Page.Name & ">", "") & vbTab & _
                      IIf(Changed - ch0, "color Shapes: " & (Changed - ch0), vbTab) & vbTab & _
                      IIf(Bitmaps - b0, "bitmaps: " & (Bitmaps - b0), vbNullString) & vbCr
            End If
            ch0 = Changed: b0 = Bitmaps
        Next Page
        oldP.Activate
    Else
       z_ConvertShapesToCMYKiterate Shapes, idx, Changed, Bitmaps, cnt
       Shapes.CreateSelection
       If Bitmaps Then _
          res = IIf(Changed, "color Shapes: " & Changed, "") & vbTab & _
                IIf(Bitmaps, "bitmaps: " & Bitmaps, "") & vbCr
    End If
    If VersionMajor > 11 Then Application.Status.EndProgress
    
    MsgBox getCMSprofiles & vbCr & vbCr _
         & "Converted: " & (Changed + Bitmaps) & vbCr & res, , _
           "RGB to CMYK: " & IIf(Shapes.Count, "Selection", "Whole document") & " (" & idx & ")"
    If Shapes.Count = 0 Then ActiveDocument.ClearSelection _
        Else Shapes.CreateSelection
    If OriginalSelection.Count Then
        OriginalSelection.CreateSelection
    Else
        ActiveDocument.ClearSelection
    End If

Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub CreateFountain()
    Const Pi# = 3.14159265358979
    Dim Shape As Shape, Swatches As ShapeRange, Shapes As ShapeRange
    Dim Ccnt&, x#, y#, cx#, cy#, shift&, pos!, atPos&, prev&, prevSpan#
    Dim i&, span#
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveDocument Is Nothing Then Exit Sub
    'get selected Source Swatches
    Set Shapes = ActiveSelection.Shapes.FindShapes
    Ccnt = Shapes.Count
    If Ccnt < 2 Then Beep: Exit Sub
    If Ccnt > 100 Then
        MsgBox "Max colors=100", , "Create Fountain fill"
        Exit Sub
    End If
    'ask user to select target shape
    If ActiveDocument.GetUserClick(x, y, shift, -1, True, 309) Then Exit Sub
    With ActivePage.SelectShapesAtPoint(x, y, True).Shapes
       If .Count = 0 Then Beep: Exit Sub
       Set Shape = .Item(.Count)
    End With
    
    'sort by distance:  (1) find center
    ActiveDocument.SaveSettings
    ActiveDocument.ReferencePoint = cdrCenter
    cx = Shapes.PositionX:  cy = Shapes.PositionY: span = 0#
    'sort by distance:  (2) find the most remote object rel. to [cx,cy]
    For i = 1 To Ccnt
       x = Shapes(i).PositionX: y = Shapes(i).PositionY
       pos = Sqr((x - cx) ^ 2 + (y - cy) ^ 2)
       If pos >= span Then prev = atPos: atPos = i: span = pos _
          Else If pos > prevSpan Then prev = i: prevSpan = pos
    Next
    If prev Then
        If Shapes(prev).PositionX < cx And Shapes(atPos).PositionX > cx Then
            atPos = prev
        Else
            If Shapes(prev).PositionY > cy _
           And Shapes(atPos).PositionY < cy _
          Then atPos = prev
        End If
    End If
    'sort by distance:  (3) sort relative to Found item
    Set Swatches = New ShapeRange
    Swatches.Add Shapes(atPos): Shapes.Remove atPos
    x = Swatches(1).PositionX: y = Swatches(1).PositionY
    Do While Shapes.Count
       span = 3E+38
       For i = 1 To Shapes.Count
            pos = _
                Sqr((Shapes(i).PositionX - x) ^ 2 + (Shapes(i).PositionY - y) ^ 2)
            If pos <= span Then atPos = i: span = pos
       Next i
       Swatches.Add Shapes(atPos): Shapes.Remove atPos
    Loop
    If Swatches(Ccnt).PositionX < Swatches(1).PositionX Then
        If VersionMajor > 12 Then
            Set Swatches = Swatches.ReverseRange
        Else
            Shapes.RemoveAll
            For i = Swatches.Count To 1 Step -1
                Shapes.Add Swatches(i)
            Next i
            Set Swatches = Shapes
        End If
    End If
    
    ActiveDocument.BeginCommandGroup "Create fountain fill"
    
    pos = 100 / (Ccnt - 1)
    x = Swatches(1).PositionX: y = Swatches(1).PositionY
    span = Sqr( _
               (Swatches(Ccnt).PositionX - x) ^ 2 _
             + (Swatches(Ccnt).PositionY - y) ^ 2 _
           )
    With Shape.Fill.ApplyFountainFill( _
             Swatches(1).Fill.UniformColor, Swatches(Ccnt).Fill.UniformColor _
         )
    If Swatches(Ccnt).PositionX = x Then
        .SetAngle Sgn(Swatches(Ccnt).PositionY - Swatches(1).PositionY) * 90
    Else
        cx = 180# / Pi * Atn( _
                             (Swatches(Ccnt).PositionY - y) _
                           / (Swatches(Ccnt).PositionX - x) _
                         )
        .SetAngle cx
        'Switch(cX = 0, 0, cX > 0, cX, cX < 0, 180 + cX)
    End If
    For i = 2 To Ccnt - 1
        If span > 0 Then
            atPos = Sqr( _
                        (Swatches(i).PositionX - x) ^ 2 _
                      + (Swatches(i).PositionY - y) ^ 2 _
                    ) / span * 100
        Else
            atPos = pos * (i - 1)
        End If
        .Colors.Add Swatches(i).Fill.UniformColor, atPos
    Next
    .SetEdgePad 0
    End With
    
Finally:
    ActiveDocument.RestoreSettings
    ActiveDocument.EndCommandGroup
    If Not Swatches Is Nothing Then Swatches.CreateSelection
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub DupOnNextPage()
    Dim Shapes As ShapeRange, Page As Page, cnt&, i&, PClipShapes As Shapes
    
    If RELEASE Then On Error GoTo Catch
    
    If ActivePage Is Nothing Then Exit Sub
    Set Shapes = ActiveSelectionRange
    If Shapes.Count = 0 Then Set Shapes = ActivePage.Shapes.All.ReverseRange
    If Shapes.Count = 0 Then Exit Sub
    Set Page = ActivePage
    
    BoostStart "dupOnNextPage " _
             & (Page.Index And &HFFFF) _
             & " -> " _
             & (Page.Index + 1 And &HFFFF), _
               RELEASE
    
    If Page.Next Is Nothing Then
        ActiveDocument.InsertPages 1, False, Page.Index
    End If
    Shapes.CopyToLayer Page.Next.ActiveLayer
    
    Page.Next.Activate
    Set PClipShapes = Page.Next.ActiveLayer.Shapes
    cnt = Shapes.Count: Shapes.RemoveAll
    For i = 1 To cnt: Shapes.Add PClipShapes(i): Next
    Shapes.CreateSelection
    
Finally:
    BoostFinish True
    UIrefresh
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub ForEach()
    Dim Shapes As ShapeRange, s As Shape, cnt&, i&, bDraw11%
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveDocument Is Nothing Then Exit Sub
    Set Shapes = ActiveSelectionRange: cnt = Shapes.Count
    EventsEnabled = False
     
    bDraw11 = (VersionMajor = 11)
    If Not bDraw11 Then Application.Status.BeginProgress CanAbort:=True
    For Each s In Shapes
        i = i + 1
        s.CreateSelection
        ActiveDocument.Repeat
        
        If Not bDraw11 Then
            Application.Status.Progress = i / cnt * 100
            Application.Status.SetProgressMessage _
                "Repeating... Page.s. there'll be no undo for whole op." _
              & i & " / " & cnt
            If Application.Status.Aborted Then
                MsgBox "Command repeated " & i & " times"
                Exit For
            End If
        End If
    Next
    
    Shapes.CreateSelection
    
Finally:
    EventsEnabled = True
    Application.Status.EndProgress
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub GuideHorizontal():  Guide_ 1, 0: End Sub
Sub GuideVertical():       Guide_ 0, 1: End Sub
Private Sub Guide_(ByVal dx&, ByVal dy&)
    Dim Page As wPOINT, x#, y#, L As Layer
    GetCursorPos Page
    On Error Resume Next
    ActiveWindow.ScreenToDocument Page.x, Page.y, x, y
    Set L = ActiveLayer
    ActiveDocument.Pages(0).Layers(sLayerGuides) _
        .CreateGuide x, y, x + dx, y + dy
    L.Activate
End Sub

Sub InvertSelection()
    Dim Shapes As ShapeRange, SelectedShapes As ShapeRange
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveDocument Is Nothing Then Exit Sub
    Set Shapes = ActiveSelectionRange
    Set SelectedShapes = ActivePage.Shapes.All
    SelectedShapes.AddRange _
        ActiveDocument.Pages(0).Layers(sLayerDesktop).Shapes.All
    SelectedShapes.RemoveRange Shapes
    SelectedShapes.RemoveRange _
        ActiveDocument.Pages(0).Layers(sLayerGuides).Shapes.All
    
    If SelectedShapes.Count Then SelectedShapes.CreateSelection _
        Else ActiveDocument.ClearSelection
    
Finally:
    UIrefresh
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub OutlineBehind()
    On Error Resume Next
    ActiveSelectionRange.SetOutlineProperties , , , , , _
        (GetKeyState(vbKeyShift) And &H8000) = 0 _
    And (GetKeyState(vbKeyControl) And &H8000) = 0
End Sub

Sub OutlineCorners()
    On Error Resume Next
    ActiveSelectionRange.SetOutlineProperties , , , , , , , , _
       IIf(GetKeyState(vbKeyControl) And &H8000, cdrOutlineMiterLineJoin, _
          IIf(GetKeyState(vbKeyShift) And &H8000, cdrOutlineBevelLineJoin, _
             cdrOutlineRoundLineJoin))
End Sub

Sub OutlineDecrease(): OutlineSet -1: End Sub
Sub OutlineIncrease(): OutlineSet 1:  End Sub
Private Sub OutlineSet(ByVal iDirection%)
    Dim Shape As Shape, Shapes As ShapeRange, outlineColor As Color, outlineStep#, o#, threshold#
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveSelectionRange.Count = 0 Then Beep: Exit Sub
    outlineStep = _
        iDirection * IIf((GetKeyState(VK_SCROLL) And 1) = 0, 0.1, 0.01)
    
    Set Shapes = ActiveSelectionRange
    Set outlineColor = CreateCMYKColor(0, 0, 0, 100)
    
    BoostStart "set outline " & IIf(iDirection > 0, "+", "") _
             & outlineStep & UnitName(ActiveDocument.Rulers.HUnits), RELEASE
             
    outlineStep = ActiveDocument.ToUnits(outlineStep, ActiveDocument.Rulers.HUnits)
    threshold = outlineStep / 2#
    For Each Shape In ActiveSelection.Shapes.FindShapes
        If Shape.Type <> cdrGroupShape Then
            If iDirection > 0 Then
                If Shape.Outline.Type = cdrOutline Then _
                   o = Shape.Outline.Width Else o = 0#
                Shape.Outline.Width = o + outlineStep
                Shape.Outline.Type = cdrOutline
            Else
                o = Shape.Outline.Width + outlineStep
                If o < threshold Then
                    Shape.Outline.Type = cdrNoOutline
                Else
                    Shape.Outline.Width = o
                End If
            End If
        End If
    Next Shape
    Shapes.CreateSelection
    
Finally:
    BoostFinish EndUndoGroup:=True
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub OutlinesToQ_KillEmpty()
    Dim Shape As Shape, shO As Shape
    Dim Shapes As ShapeRange, srO As ShapeRange, toDel As ShapeRange
    Dim bStatus%, i&, j&, cnt&, bNoFill%, t0!
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveDocument Is Nothing Then Exit Sub
    Set Shapes = ActiveSelection.Shapes.FindShapes: cnt = Shapes.Count: If cnt = 0 Then Exit Sub
    Set toDel = New ShapeRange: Set srO = New ShapeRange
    
    BoostStart "Convert outline to curves++", RELEASE
    
    bStatus = (VersionMajor > 11)
    If bStatus Then _
        Application.Status.BeginProgress "Convert outline to curves++", True
    
    For Each Shape In Shapes
        If Shape.Outline.Type <> cdrNoOutline Then
            Set shO = Shape.Outline.ConvertToObject
            If Not shO Is Nothing Then
                srO.Add shO
                If Shape.Fill.Type = cdrNoFill Then
                   If Shape.PowerClip Is Nothing Then
                      If Shape.WrapText = cdrWrapNone Then
                         If Shape.Effects.Count = 0 Then
                            If Shape.Type <> 20 Then 'cdrMeshFillShape
                               toDel.Add Shape
                            End If
                         End If
                      End If
                   End If
                End If
            End If
            If bStatus Then i = i + 1
            If Timer - t0 > 0.1 Then
                Application.Status.Progress = i / cnt * 100
                If Application.Status.Aborted Then Exit For Else t0 = Timer
            End If
        End If
    Next
    toDel.Delete
    srO.CreateSelection
    If bStatus Then Application.Status.EndProgress
    
Finally:
    BoostFinish EndUndoGroup:=True
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub OutlineEqualsFill()
    Dim Shapes As ShapeRange, Shape As Shape
    
    If RELEASE Then On Error GoTo Catch
    
    Set Shapes = New ShapeRange
    If ActiveSelectionRange.Count = 0 Then Beep: Exit Sub
    
    BoostStart "Outline color = fill color", RELEASE
    
    For Each Shape In ActiveSelection.Shapes.FindShapes
        Err.Clear
        If Shape.Fill.Type = cdrUniformFill Then
            If Shape.Outline.Type = cdrNoOutline Then
                Shape.Outline.Type = cdrOutline
            End If
            Shape.Outline.Color.CopyAssign Shape.Fill.UniformColor
            If Err.Number = 0 Then Shapes.Add Shape
        End If
    Next Shape
    If Shapes.Count Then
        Shapes.CreateSelection
    Else
        ActiveDocument.ClearSelection
    End If
    
Finally:
    BoostFinish EndUndoGroup:=True
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub OverprintsRemove()
    Dim Shape As Shape, Shapes As ShapeRange, sr2 As ShapeRange, SelectedShapes As ShapeRange
    Dim cnt&, ob&(), oF&(), oO&(), cB&, cF&, cO&, i&, Page&, res$, p0 As Page, bStat%, t0!
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveDocument Is Nothing Then Exit Sub
    Set Shapes = New ShapeRange: Set sr2 = New ShapeRange
    Set SelectedShapes = New ShapeRange
    
    Set p0 = ActivePage
    Shapes.AddRange ActiveSelection.Shapes.FindShapes: cnt = Shapes.Count
    If cnt = 0 Then Page = 1 Else Page = p0.Index And &HFFFF
    i = ActiveDocument.Pages.Count: ReDim ob(i), oF(i), oO(i)
    
    i = 0
    bStat = (Application.VersionMajor > 11)
    If bStat Then Application.Status.BeginProgress CanAbort:=True
    
    BoostStart "RemoveOverprints", RELEASE
    
    Do
        If Shapes.Count = 0 Then
            Shapes.AddRange ActiveDocument.Pages(Page).FindShapes
            cnt = Shapes.Count
        End If
        For Each Shape In Shapes
            i = i + 1
            If bStat Then
                If Timer - t0 > 0.1! Then
                    t0 = Timer
                    Application.Status.Progress = i / cnt * 100
                    If Application.Status.Aborted Then Exit For
                End If
            End If
            If Shape.Type = cdrBitmapShape Then
                If VersionMajor >= 13 Then
                    If Shape.OverprintBitmap Then
                        ob(Page) = ob(Page) + 1
                        Shape.OverprintBitmap = False
                        SelectedShapes.Add Shape
                    End If
                End If
            End If
            If Shape.OverprintFill Then
                oF(Page) = oF(Page) + 1
                Shape.OverprintFill = False
                SelectedShapes.Add Shape
            End If
            If Shape.OverprintOutline Then
                oO(Page) = oO(Page) + 1
                Shape.OverprintOutline = False
                SelectedShapes.Add Shape
            End If
            
            If Not Shape.PowerClip Is Nothing Then
                sr2.AddRange Shape.PowerClip.Shapes.FindShapes
                cnt = cnt + sr2.Count
            End If
        Next
        If Application.Status.Aborted Then Exit Do
        Shapes.RemoveAll: Shapes.AddRange sr2: sr2.RemoveAll
        Page = Page + 1: ActiveDocument.Pages(Page).Activate
    Loop Until Shapes.Count = 0
    
    p0.Activate
    If bStat Then Application.Status.EndProgress
    If SelectedShapes.Count Then
        SelectedShapes.CreateSelection
    Else
        ActiveDocument.ClearSelection
    End If
       
    For i = 0 To UBound(ob)
       cB = cB + ob(i): cF = cF + oF(i): cO = cO + oO(i)
       If ob(i) Or oF(i) Or oO(i) Then _
          res = res & "Page " & i & ": " & _
             IIf(ob(i), ob(i) & " bitmaps   ", "") & _
             IIf(oF(i), oF(i) & " fills   ", "") & _
             IIf(oO(i), oO(i) & " outlines", "") & vbCr
    Next
    MsgBox "Total: " & (cB + cF + cO) & " overprints removed" & vbCr _
         & res, , "Remove overprints"
    
Finally:
    BoostFinish EndUndoGroup:=True
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub PageNamesAsNumbers()
    Dim Page As Page
    On Error Resume Next
    For Each Page In ActiveDocument.Pages
       Page.Name = " "
    Next
End Sub

Sub PasteAtMouse()
    Const gap& = 20
    If CountClipboardFormats = 0 Then Beep: Exit Sub
    If ActiveDocument Is Nothing Then Application.CreateDocument
    Dim cp As wPOINT, r As wRECT, x#, y#, cnt&
    GetCursorPos cp
    GetWindowRect ActiveWindow.Handle, r
    ActiveWindow.ScreenToDocument _
        vSqueezeL(cp.x, r.L + gap, r.r - gap), _
        vSqueezeL(cp.y, r.T + gap, r.b - gap), _
        x, y
    
    cnt = ActiveLayer.Shapes.Count
    ActiveDocument.BeginCommandGroup "Paste  at mouse"
    ActiveLayer.Paste
    If cnt <> ActiveLayer.Shapes.Count Then
        Dim i&, Shapes As ShapeRange
        Set Shapes = New ShapeRange
        For i = 1 To ActiveLayer.Shapes.Count - cnt
           Shapes.Add ActiveLayer.Shapes(i)
        Next
        Shapes.SetPosition x, y
        End If
    ActiveDocument.EndCommandGroup
End Sub

Sub PClipPick()
    Dim cp As wPOINT, Shape As Shape, x#, y#, i&, PClipShapes As Shapes
    Dim OriginalSelection As ShapeRange
    
    If ActiveDocument Is Nothing Then Exit Sub
    
    If RELEASE Then On Error GoTo Catch
    BoostStart , RELEASE
    
    Set Shape = ActiveShape
    Set OriginalSelection = ActiveSelectionRange
    GetCursorPos cp: ActiveWindow.ScreenToDocument cp.x, cp.y, x, y
    
    Do
        If Not Shape Is Nothing Then _
            If Not Shape.PowerClip Is Nothing Then _
                If Shape.IsOnShape(x, y) <> cdrOutsideShape Then _
                    Exit Do
        
        Set Shape = ActivePage.SelectShapesAtPoint(x, y, 0)
        If Shape.Shapes.Count = 0 Then Beep: Exit Do
        
        For i = Shape.Shapes.Count To 1 Step -1
            If Not Shape.Shapes(i).PowerClip Is Nothing Then
                Set Shape = Shape.Shapes(i)
                Exit For
            End If
        Next i
    Loop Until True
    
    Do
        If Shape.PowerClip Is Nothing Then Exit Do
        
        ActiveDocument.ClearSelection
        Set PClipShapes = Shape.PowerClip.Shapes
        For i = 1 To PClipShapes.Count
            If PClipShapes(i).IsOnShape(x, y) <> cdrOutsideShape Then
                If (GetKeyState(vbKeyShift) And &H8000) = 0 Then _
                   OriginalSelection.RemoveAll
                OriginalSelection.Add PClipShapes(i)
                OriginalSelection.CreateSelection
                Exit For
            End If
        Next
    Loop Until True
    
Finally:
    BoostFinish False
    UIrefresh
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub RectangleFixer() 'X3                           MsgBox "Works only in DrawX3"
    Dim sRect As Shape, Shapes As ShapeRange, newSel As New ShapeRange
    Dim bFixed%, bRound%, sDup As Shape, det#, x#, y#, sx#, sy#
    Dim r1#, r2#, r3#, r4#, d11#, d12#, d21#, d22#, i11#, i12#, i21#, i22#
    Dim PClip As ShapeRange, lWrap&, dWrapOffs#
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveShape Is Nothing Then Exit Sub
    Set Shapes = ActiveSelection.Shapes.FindShapes(Type:=cdrRectangleShape)
    
    BoostStart "Fix rounded rectangles", RELEASE
    
    For Each sRect In Shapes
        ' Check if the rectangle was actually stretched
        sx = sRect.AbsoluteHScale: sy = sRect.AbsoluteVScale
        r1 = sRect.Rectangle.CornerUpperLeft
        r2 = sRect.Rectangle.CornerUpperRight
        r3 = sRect.Rectangle.CornerLowerLeft
        r4 = sRect.Rectangle.CornerLowerRight
        bRound = (r1 <> 0#) Or (r2 <> 0#) Or (r2 <> 0#) Or (r2 <> 0#)
        
        If (Abs(sx - 1) > 0.00001 Or Abs(sy - 1) > 0.00001) And bRound Then
           
            If Not sRect.PowerClip Is Nothing Then _
               Set PClip = sRect.PowerClip.ExtractShapes
            lWrap = sRect.WrapText: dWrapOffs = sRect.TextWrapOffset
            
            ' Make a temporary copy of the rectangle
            Set sDup = sRect.TreeNode.GetCopy().VirtualShape
            sDup.GetMatrix d11, d12, d21, d22, x, y
            ' Remove skew and rotation from the object temporarily
            d11 = d11 / sx:    d12 = d12 / sy
            d21 = d21 / sx:    d22 = d22 / sy
            det = d11 * d22 - d12 * d21
            i11 = d22 / det:   i12 = -d12 / det
            i21 = -d21 / det:  i22 = d11 / det
            sDup.AffineTransform i11, i12, i21, i22, 0, 0
            ' Get the unrotated/unskewed size
            sDup.GetBoundingBox x, y, sx, sy
            sDup.Delete
            
            Set sDup = ActiveVirtualLayer.CreateRectangle( _
                           x, y, x + sx, y + sy, r1, r2, r3, r4 _
                       )
            sDup.Fill.CopyAssign sRect.Fill
            sDup.Outline.CopyAssign sRect.Outline
            sDup.AffineTransform d11, d12, d21, d22, 0, 0
            
            sRect.ReplaceWith sDup
            
            sRect.WrapText = lWrap: sRect.TextWrapOffset = dWrapOffs
            If Not PClip Is Nothing Then
                PClip.AddToPowerClip sRect, CenterInContainer:=cdrFalse
            End If
            
            newSel.Add sRect
        End If
    Next sRect
    newSel.CreateSelection
    
Finally:
    BoostFinish True
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub ScrollScreenDown():   ScrollScreen 0, 1:   End Sub
Sub ScrollScreenUp():   ScrollScreen 0, -1:   End Sub
Sub ScrollScreenLeft():   ScrollScreen -1, 0:   End Sub
Sub ScrollScreenRight():   ScrollScreen 1, 0:   End Sub
Sub ScrollScreenDownRight():   ScrollScreen 1, 1:   End Sub
Sub ScrollScreenUpRight():   ScrollScreen 1, -1:   End Sub
Sub ScrollScreenDownLeft():   ScrollScreen -1, 1:   End Sub
Sub ScrollScreenUpLeft():   ScrollScreen -1, -1:   End Sub
Sub ScrollScreen(dx#, dy#) ' dx,dy = -1 or 1 or 0
    Dim w#, h#
    If ActiveDocument Is Nothing Then Exit Sub
    If ActiveWindow.WindowState = cdrWindowMaximized Then
        ActiveWindow.ScreenToDocument _
            AppWindow.ClientWidth, AppWindow.ClientHeight, w, h
    Else
        ActiveWindow.ScreenToDocument _
            ActiveWindow.Width, ActiveWindow.Height, w, h
    End If
    
    Dim SCROLLpercent: SCROLLpercent = 0.8  '80%
    With ActiveWindow.ActiveView
        .SetViewPoint .OriginX - dx * (.OriginX - w) * 2 * SCROLLpercent, _
                      .OriginY - dy * (.OriginY - h) * 2 * SCROLLpercent
    End With
End Sub

Sub SelectComplexCurves()
    Dim Shapes As ShapeRange, Shape As Shape, sr2 As ShapeRange, lThreshold&
    Dim SelectedShapes As ShapeRange
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveDocument Is Nothing Then Exit Sub
    lThreshold = Val(InputBox( _
                         "Node count threshold", _
                         "Select complex curves", _
                         "1000" _
                     ))
    If lThreshold <= 0 Then Exit Sub
    Set Shapes = New ShapeRange
    Set sr2 = New ShapeRange
    Set SelectedShapes = New ShapeRange
    Shapes.AddRange ActiveSelection.Shapes.FindShapes
    If Shapes.Count = 0 Then Shapes.AddRange ActivePage.FindShapes
    Do
        For Each Shape In Shapes
            If Shape.Type = cdrCurveShape Then
                If Shape.Curve.Nodes.Count > lThreshold Then
                    SelectedShapes.Add Shape
                End If
            End If
            If Not Shape.PowerClip Is Nothing Then _
                sr2.AddRange Shape.PowerClip.Shapes.FindShapes
        Next
        Shapes.RemoveAll
        Shapes.AddRange sr2
        sr2.RemoveAll
    Loop Until Shapes.Count = 0
    If SelectedShapes.Count Then
        SelectedShapes.CreateSelection
    Else
        ActiveDocument.ClearSelection
    End If
    
Finally:
    CorelScript.RedrawScreen
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub SelectSmallObjects()
    Dim ShapesToDelete As ShapeRange, Shape As Shape, Source As ShapeRange
    Dim MaxSize#, bDoDelete%, i&, cnt&, bStatus%, t0!
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveDocument Is Nothing Then Exit Sub
    Dim InputReturn As Variant
    InputReturn = InputBox( _
                      "Size limit (" & UnitName(ActiveDocument.Rulers.HUnits) _
                    & ")" & vbCr & vbCr _
                    & "[ Shift+OK/Enter to delete immediately ]", _
                      "Select SMALL objects", "1" _
                  )
    If Not IsNumeric(InputReturn) Then Exit Sub
    MaxSize = CDbl(InputReturn)
    If MaxSize <= 0# Then Beep: Exit Sub
    bDoDelete = (GetKeyState(vbKeyShift) And &H8000) <> 0
    
    BoostStart IIf(bDoDelete, "DELETE small objects", vbNullString), RELEASE
    
    bStatus = (VersionMajor > 11)
    If bStatus Then Application.Status.BeginProgress CanAbort:=True
    
    ActiveDocument.Unit = ActiveDocument.Rulers.HUnits
    Set ShapesToDelete = New ShapeRange
    Set Source = ActiveSelection.Shapes.FindShapes: cnt = Source.Count
    If cnt = 0 Then Set Source = ActivePage.Shapes.FindShapes
    cnt = Source.Count
    If cnt = 0 Then Beep: Exit Sub
    For Each Shape In Source
        i = i + 1
        If bStatus Then
            If Timer - t0 > 0.1! Then _
                Application.Status.Progress = i / cnt * 100
            If Application.Status.Aborted Then Exit For Else t0 = Timer
        End If
        If Shape.SizeHeight <= MaxSize Then
            If Shape.SizeWidth <= MaxSize Then ShapesToDelete.Add Shape
        End If
    Next Shape
    If bDoDelete Then ShapesToDelete.Delete _
       Else: ShapesToDelete.CreateSelection
    If bStatus Then Application.Status.EndProgress
    
Finally:
    BoostFinish EndUndoGroup:=bDoDelete
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub SelectInside()
    Dim OriginalShapes As ShapeRange, Shapes As ShapeRange
    Dim ShapesInside As ShapeRange
    Dim ShapesToDelete As ShapeRange, Shape As Shape, ShapeInside As Shape
    Dim w#, h#, x#, y#, big As Shape, bBigCopy%
    Dim i&, cnt&, bStatus%, t0!
    Dim wndOldView As cdrViewType, wndZoom#, wndOX#, wndOY#
    Dim minSh As Shape, minS#
    
    If RELEASE Then On Error GoTo Catch
    
    Set Shapes = ActiveSelection.Shapes.All
    If Shapes.Count = 0 Then _
        MsgBox "Select Shapes before running me", , "wx.SelectInside": Exit Sub
    Set OriginalShapes = New ShapeRange: OriginalShapes.AddRange Shapes
    
    If ActiveDocument.GetUserClick(x, y, i, -1, True, 313) Then Exit Sub
    With ActivePage.SelectShapesAtPoint(x, y, True, -1).Shapes
       If .Count = 0 Then Beep: Exit Sub
       Set big = .Item(.Count)
    End With
    
    bStatus = (VersionMajor > 11)
    
    BoostStart "Select inside " & (Shapes.Count - 1) & " Shapes", RELEASE
    
    If bStatus Then Application.Status.BeginProgress CanAbort:=True
    ActiveDocument.ReferencePoint = cdrBottomLeft
    
    If bStatus Then Application.Status.SetProgressMessage "Analyzing"
    Application.Status.Progress = 3
    '   Set ShapesToDelete = New ShapeRange
    '   ShapesToDelete.AddRange Shapes
    '   ShapesToDelete.RemoveRange ActiveSelection.Shapes.FindShapes(, cdrGroupShape, recursive:=False)
    '   For Each Shape In ShapesToDelete
    '       Shape.GetSize w, h: w = w * h
    '   If w > maxS Then maxS = w: Set big = Shape
    '   Next
    
    Do
    '   If big Is Nothing Then MsgBox "No surround shape Found": Exit Do
       
        If bStatus Then Application.Status.SetProgressMessage _
                            "Close curve and assign fill to surrounding shape"
        
        Shapes.Remove Shapes.IndexOf(big)
        ShapesToDelete.RemoveAll
        
        If big.Type = cdrCurveShape Then _
           If Not big.Curve.Closed Then _
              bBigCopy = True
              Set big = big.DuplicateAsRange()(1)
              big.Curve.Closed = True
           
        Err.Clear
        If big.Fill.Type = cdrNoFill Then
           If Err.Number = 0 Then
              If Not bBigCopy Then bBigCopy = True
              Set big = big.DuplicateAsRange()(1)
              big.Fill.ApplyUniformFill CreateCMYKColor(0, 0, 0, 100)
           End If
           End If
        
        big.GetPosition x, y:  big.GetSize w, h
        If bStatus Then _
            Application.Status.SetProgressMessage _
                "Sorting: eliminate Shapes outside frame's bounding box": _
                Application.Status.Progress = 6
        Set ShapesToDelete = ActivePage.Shapes.All
        If bStatus Then Application.Status.Progress = 9
        If Application.Status.Aborted Then Exit Do
        ShapesToDelete.Add _
            ActiveDocument.Pages(0).Layers(z.sLayerDesktop).Shapes.All
        If bStatus Then Application.Status.Progress = 12
        If Application.Status.Aborted Then Exit Do
        ShapesToDelete.RemoveRange _
            ActivePage.SelectShapesFromRectangle( _
                x, y, x + w, y + h, True _
            ).Shapes.All
            If bStatus Then Application.Status.Progress = 15
            If Application.Status.Aborted Then Exit Do
        ShapesToDelete.RemoveRange _
            ActiveDocument.Pages(0).Layers(z.sLayerDesktop) _
                .SelectShapesFromRectangle(x, y, x + w, y + h, True).Shapes.All
            If bStatus Then Application.Status.Progress = 18
            If Application.Status.Aborted Then Exit Do
        Shapes.RemoveRange ShapesToDelete
        If bStatus Then Application.Status.Progress = 21
        If Application.Status.Aborted Then Exit Do
        
        cnt = Shapes.Count
        ShapesToDelete.RemoveAll: Set ShapesInside = New ShapeRange
        If bStatus Then Application.Status.SetProgressMessage _
                            "Sorting: determine inner Shapes"
        
        'get minimum sized shape to set the zoom level -
        'to trick bug with IsOnShape on low zoom
        minS = w * h: t0 = Timer: i = 0
        For Each Shape In Shapes: Shape.GetSize w, h: w = w * h
            If w > 0 Then If w < minS Then minS = w: Set minSh = Shape
            i = i + 1
            If (i And 255) = 0 Then If Timer - t0 > 0.1 Then Exit For
        Next
           
        With ActiveWindow.ActiveView
            If Not minSh Is Nothing Then _
                wndZoom = .Zoom: wndOX = .OriginX: wndOY = .OriginY
                .ToFitShape minSh
            wndOldView = .Type
            If wndOldView <= cdrWireframeView Then _
                .Type = 2 Else wndOldView = -1
        End With
           
        For Each Shape In Shapes
            i = i + 1
            If bStatus Then
                If Timer - t0 > 0.1! Then
                    Application.Status.Progress = 21 + i / cnt * 79
                    If Application.Status.Aborted Then Exit For Else t0 = Timer
                End If
            End If
            Shape.GetPosition x, y
            Do
                Select Case big.IsOnShape(x, y, -1)
                   Case cdrInsideShape
                      Shape.GetSize w, h
                      If big.IsOnShape(x + w, y, -1) Then _
                         If big.IsOnShape(x, y + h, -1) Then _
                            If big.IsOnShape(x + w, y + h, -1) Then _
                                ShapesInside.Add Shape: Exit Do
                   Case cdrOutsideShape
                      Shape.GetSize w, h
                      If big.IsOnShape(x + w, y, -1) = 0 Then _
                         If big.IsOnShape(x, y + h, -1) = 0 Then _
                            If big.IsOnShape(x + w, y + h, -1) = 0 Then Exit Do
                End Select
                   
                Set ShapeInside = _
                    big.Trim(Shape, LeaveSource:=True, LeaveTarget:=True)
                If ShapeInside Is Nothing Then ShapesInside.Add Shape: Exit Do
                ShapesToDelete.Add ShapeInside
                Set ShapeInside = Nothing
            Loop Until True
        Next
        
        If ShapesToDelete.Count Then ShapesToDelete.Delete
        If bBigCopy Then big.Delete
        
        If wndOldView >= 0 Then ActiveWindow.ActiveView.Type = wndOldView
        If wndZoom <> 0# Then _
            ActiveWindow.ActiveView.SetViewPoint wndOX, wndOY, wndZoom
        
        If GetKeyState(VK_SCROLL) And 1 Then
            OriginalShapes.RemoveRange ShapesInside
            OriginalShapes.CreateSelection
        Else
            ShapesInside.CreateSelection
        End If
       
    Loop Until True
    If bStatus Then Application.Status.EndProgress
    
Finally:
    BoostFinish EndUndoGroup:=True
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub SelectSameFillColor():       z_selectSameColor True, False: End Sub
Sub SelectSameFillAndOutline():  z_selectSameColor True, True:  End Sub
Sub SelectSameOutline():         z_selectSameColor False, True: End Sub
Sub SelectSameDialog(): UIColorSel.Show: End Sub

Sub SizePagetoFIT()
    Static pad#, b2ndRun%
    Dim Shapes As ShapeRange, SRlock As ShapeRange, Shape As Shape
    Dim s$, shift%, w#, h#, halfUnit#, RoundDigits&
    
    If RELEASE Then On Error GoTo Catch
    
    If ActiveDocument Is Nothing Then Exit Sub
       
    With ActiveDocument
        If ActiveShape Is Nothing Then
           Set Shapes = ActivePage.Shapes.All
           Shapes.RemoveRange .Pages(0).Layers(sLayerGuides).FindShapes
           Shapes.RemoveRange ActivePage.FindShapes(, cdrGuidelineShape)
        Else
           Set Shapes = ActiveSelectionRange
        End If
        
        s = InputBox( _
                "Page size increase by (" & _
                UnitName(.Rulers.HUnits) & ")" & vbCr & vbCr & _
                "Shift+OK/Enter - use rounded numbers", _
                "page size = <" & IIf( _
                                      ActiveShape Is Nothing, _
                                      "ALL", _
                                      "SELECTED" _
                                  ) & _
                " objects> + <pad>", IIf(b2ndRun, pad, 1) _
            )
        If Len(s) = 0 Then Exit Sub Else pad = Val(s)
        shift = (GetKeyState(vbKeyShift) And &H8000) <> 0
        
        BoostStart "sizePagetoFIT + " & s & UnitName(.Rulers.HUnits), RELEASE
        
        .Unit = .Rulers.HUnits
        If shift Then
           halfUnit = 0.49999: RoundDigits = 0
           Select Case .Unit
              Case cdrCentimeter: halfUnit = halfUnit / 10: RoundDigits = 1
              Case cdrInch, cdrMeter: halfUnit = halfUnit / 100: RoundDigits = 2
              Case cdrMile, cdrKilometer: halfUnit = halfUnit / 1000: RoundDigits = 3
           End Select
           End If
        
        Set SRlock = New ShapeRange
        For Each Shape In ActivePage.FindShapes
           If Shape.Locked Then SRlock.Add Shape: Shape.Locked = False
           Next
        
        Shapes.GetSize w, h
        If shift Then
            .Pages(0).SetSize Round(w + pad + halfUnit, RoundDigits), _
                              Round(h + pad + halfUnit, RoundDigits)
        Else
            .Pages(0).SetSize w + pad, h + pad
         End If
        .ReferencePoint = cdrCenter
        With ActivePage.Shapes.All
           .SetPosition _
              ActivePage.CenterX + .PositionX - Shapes.PositionX, _
              ActivePage.CenterY + .PositionY - Shapes.PositionY
           End With
        
        SRlock.Lock
        
        halfUnit = pad + IIf(w > h, w, h) * 0.05
        w = w + halfUnit: h = h + halfUnit
        ActiveWindow.ActiveView.SetViewArea _
            ActivePage.CenterX - w / 2, ActivePage.CenterY - h / 2, w, h
    End With
    b2ndRun = True
    
Finally:
    BoostFinish EndUndoGroup:=True
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub TextToCurves()
    Dim srQ As ShapeRange, Shapes As ShapeRange, sr2 As ShapeRange
    Dim Shape As Shape, i&, curP As Page, bAll%, bDigPClip%
        
    If RELEASE Then On Error GoTo Catch
    
    If ActiveDocument Is Nothing Then Exit Sub
    Set curP = ActivePage
    Set Shapes = New ShapeRange
    Set sr2 = New ShapeRange
    Set srQ = New ShapeRange
    bAll = (ActiveSelectionRange.Count = 0)
    bDigPClip = (VersionMajor > 11)
    For i = 1 To ActiveDocument.Pages.Count
        With ActiveDocument.Pages(i)
            If bAll Or .Index = curP.Index Then
                If bDigPClip Then
                    If bAll Then
                        Shapes.AddRange .FindShapes
                    Else
                        Shapes.AddRange ActiveSelection.Shapes.FindShapes
                    End If
                    Do
                       For Each Shape In Shapes
                          If Shape.Type = cdrTextShape Then srQ.Add Shape
                          If Not Shape.PowerClip Is Nothing Then
                              sr2.AddRange Shape.PowerClip.Shapes.FindShapes
                          End If
                       Next
                       Shapes.RemoveAll: Shapes.AddRange sr2: sr2.RemoveAll
                    Loop Until Shapes.Count = 0
                Else
                    If bAll Then
                        srQ.AddRange .FindShapes(, cdrTextShape, True)
                    Else
                        srQ.AddRange _
                            ActiveSelection.Shapes.FindShapes( _
                                , cdrTextShape, True _
                            )
                    End If
                End If
            End If
        End With
    Next
    srQ.ConvertToCurves
    
Finally:
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub

Sub TransparentEdge(): UITranspEdge.Show: End Sub

Sub ZoomFullPage()
    Const MARGIN = 5    ' 5% is by default
    If ActiveWindow Is Nothing Then Beep: Exit Sub
    On Error Resume Next
    With ActivePage
        Dim w#, h#
        w = .SizeWidth * (100 + MARGIN) / 100
        h = .SizeHeight * (100 + MARGIN) / 100
        ActiveWindow.ActiveView.SetViewArea .CenterX - w / 2, .CenterY - h / 2, w, h
    End With
End Sub

Sub ZoomOutBack()
    On Error Resume Next
    ActiveWindow.ActiveView.ZoomOut
End Sub

'===============================================================================
                           
Private Sub z_ConvertShapesToCMYKiterate( _
                Scope As ShapeRange, idx&, Changed&, Bitmaps&, cnt& _
            )
    Dim Shape As Shape, fc As FountainColor, notCmyk As Boolean
    Dim bmCvt%, bDoStatus%
    
    On Error Resume Next
    
    cnt = cnt + Scope.Count
    bDoStatus = (VersionMajor > 11)
    For Each Shape In Scope
       idx = idx + 1
       If Shape.Type <> cdrGroupShape Then
          notCmyk = False
          If bDoStatus Then Application.Status.Progress = idx / cnt * 100
          bmCvt = 0
          If Shape.Type = cdrBitmapShape Then
             Select Case Shape.Bitmap.Mode
                Case cdrBlackAndWhiteImage: bmCvt = 0
                Case cdrCMYKColorImage, 8: bmCvt = 1 'cdrCMYKMultiChannelImage=8
                Case Else: bmCvt = 2
             End Select
             End If
          If bmCvt = 2 Then
             Shape.Bitmap.ConvertTo cdrCMYKColorImage: Bitmaps = Bitmaps + 1
          ElseIf bmCvt = 0 Then
             With Shape.Fill
                Select Case Shape.Fill.Type
                   Case cdrUniformFill
                       notCmyk = z_ConvertColorToCMYK(.UniformColor)
                   Case cdrFountainFill:
                       For Each fc In .Fountain.Colors
                           notCmyk = notCmyk Or z_ConvertColorToCMYK(fc.Color)
                       Next
                   Case cdrPatternFill:
                      notCmyk = z_ConvertColorToCMYK(.Pattern.BackColor) _
                             Or z_ConvertColorToCMYK(.Pattern.FrontColor)
                End Select
             End With
          End If
          If Shape.Outline.Type = cdrOutline Then _
             notCmyk = notCmyk Or z_ConvertColorToCMYK(Shape.Outline.Color)
          If notCmyk Then Changed = Changed + 1
       End If
       If Not Shape.PowerClip Is Nothing Then _
           z_ConvertShapesToCMYKiterate _
               Shape.PowerClip.Shapes.FindShapes, idx, Changed, Bitmaps, cnt
    Next
End Sub

Private Function z_ConvertColorToCMYK(c As Color) As Boolean
    Select Case c.Type
       Case cdrColorCMYK 'nothing, it's OK
       Case cdrColorBlackAndWhite: z_ConvertColorToCMYK = True
          c.CMYKAssign 0, 0, 0, IIf(c.IsWhite, 0, 100)
       Case cdrColorGray: z_ConvertColorToCMYK = True
          c.CMYKAssign 0, 0, 0, (255 - c.Gray) / 255 * 100
       Case Else: c.ConvertToCMYK: z_ConvertColorToCMYK = True
    End Select
End Function

Sub z_selectSameColor( _
        Optional checkFill% = True, _
        Optional checkOutline% = False, _
        Optional sens_ As Variant _
    )
   Static sens&, bNotFirstRun%
   Dim Shapes As ShapeRange, Shape As Shape, clr As New Color, srg As ShapeRange
   Dim seekFillColor As New Color, seekOutlineColor As New Color
   Dim Found As New ShapeRange, s$, selectInGroups%, diff&
   Dim seekShapeType&, seekFillType&, seekOutlineType&, _
       passFill%, passOutline%

   If ActiveShape Is Nothing Then Beep: Exit Sub

   If Not IsMissing(sens_) Then
      s = sens_
   Else
      s = sens
      If bNotFirstRun Then
          s = Trim$(GetSetting(sCorelDRAW, "sameColorSel", "Sens", "-1"))
          bNotFirstRun = True
      End If
      If (GetKeyState(VK_SCROLL) And 1) <> 0 Then
         s = Trim$(InputBox("Difference allowed (0-100%):" & vbCrLf & "(Negative = select in groups)", "Select same color", sens))
         If s = "" Then Exit Sub
      End If
   End If
   sens = Int(Val(s)): selectInGroups = (Left$(s, 1) = "-"): diff = Abs(sens)
   
   seekShapeType = ActiveShape.Type: seekFillType = -1: seekOutlineType = -1
   Select Case seekShapeType
      Case cdrCurveShape, 21, cdrEllipseShape, 26, _
           cdrPolygonShape, cdrRectangleShape, cdrTextShape 'cdrCustomShape=21, cdrPerfectShape=26
         If checkFill Then
            seekFillType = ActiveShape.Fill.Type
            Select Case seekFillType
               Case cdrNoFill: '
               Case cdrUniformFill
                  seekFillColor.CopyAssign ActiveShape.Fill.UniformColor
                  If diff > 0 Then seekFillColor.ConvertToCMYK
               Case cdrFountainFill:
               Case Else
                  MsgBox "Unsupported type of fill.": Exit Sub
            End Select
         End If
         If checkOutline Then
            seekOutlineType = ActiveShape.Outline.Type
            If seekOutlineType <> cdrNoOutline Then
               seekOutlineColor.CopyAssign ActiveShape.Outline.Color
               If diff > 0 Then seekOutlineColor.ConvertToCMYK
            End If
         End If
   End Select
   
   If RELEASE Then On Error GoTo Catch
   
   BoostStart , RELEASE
   
   If Not selectInGroups Then
      Set Shapes = ActivePage.Shapes.All
      Shapes.RemoveRange ActivePage.FindShapes(, cdrGroupShape, False)
   Else
      Set Shapes = ActivePage.FindShapes
   End If
      
   Shapes.AddRange _
       ActiveDocument.Pages(0).Layers(sLayerDesktop) _
       .FindShapes(, , selectInGroups)
   
   For Each Shape In Shapes
      Select Case Shape.Type
      Case cdrCurveShape, 21, cdrEllipseShape, 26, _
           cdrPolygonShape, cdrRectangleShape, cdrTextShape 'cdrCustomShape=21, cdrPerfectShape=26
         passFill = False: passOutline = False
         If checkFill And Shape.Fill.Type = seekFillType Then
            Select Case seekFillType
               Case cdrNoFill
                  passFill = True
               Case cdrUniformFill
                  If diff = 0 Then
                     passFill = Shape.Fill.UniformColor.IsSame(seekFillColor)
                  Else
                     clr.CopyAssign Shape.Fill.UniformColor
                     If clr.Type <> cdrColorCMYK Then clr.ConvertToCMYK
                     passFill = diff >= _
                         Sqr((clr.CMYKCyan - seekFillColor.CMYKCyan) ^ 2 _
                       + (clr.CMYKMagenta - seekFillColor.CMYKMagenta) ^ 2 _
                       + (clr.CMYKYellow - seekFillColor.CMYKYellow) ^ 2 _
                       + (clr.CMYKBlack - seekFillColor.CMYKBlack) ^ 2)
                  End If
            End Select
         End If
         If checkOutline And Shape.Outline.Type = seekOutlineType Then
            Select Case seekOutlineType
               Case cdrNoOutline
                  passOutline = True
               Case cdrOutline
                  If diff = 0 Then
                     passOutline = Shape.Outline.Color.IsSame(seekOutlineColor)
                  Else
                     clr.CopyAssign Shape.Outline.Color
                     If clr.Type <> cdrColorCMYK Then clr.ConvertToCMYK
                     passOutline = diff >= _
                         Sqr((clr.CMYKCyan - seekOutlineColor.CMYKCyan) ^ 2 _
                       + (clr.CMYKMagenta - seekOutlineColor.CMYKMagenta) ^ 2 _
                       + (clr.CMYKYellow - seekOutlineColor.CMYKYellow) ^ 2 _
                       + (clr.CMYKBlack - seekOutlineColor.CMYKBlack) ^ 2)
                  End If
            End Select
         End If
         If IIf(checkFill, passFill, True) _
        And IIf(checkOutline, passOutline, True) Then
             Found.Add Shape
         End If
      Case Else
         If Shape.Type = seekShapeType Then Found.Add Shape
      End Select
   Next Shape
   If Found Is Nothing _
      Then ActiveDocument.ClearSelection _
      Else Found.CreateSelection
  
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
End Sub
