VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UIColorSel 
   Caption         =   "Same color select"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   OleObjectBlob   =   "UIColorSel.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UIColorSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit
Private Sub btnCancel_Click(): Unload Me: End Sub
Private Sub btnFill_Click():     UISave: wx.z_selectSameColor True, 0, IIf(chGroups, "-", "") & txSens.Text:  End Sub
Private Sub btnOutline_Click():  UISave: wx.z_selectSameColor 0, True, IIf(chGroups, "-", "") & txSens:       End Sub
Private Sub btnBoth_Click():     UISave: wx.z_selectSameColor True, True, IIf(chGroups, "-", "") & txSens:    End Sub
Private Sub spin_Change():  txSens = spin.Value: End Sub
Private Sub txSens_Change(): On Error Resume Next: txSens = CLng(txSens): spin.Value = txSens: End Sub
Private Sub UISave(): SaveSetting sCorelDRAW, "sameColorSel", "Sens", IIf(chGroups, "-", "") & txSens: End Sub
Private Sub UserForm_Activate(): With txSens: .SetFocus: .SelStart = 0: .SelLength = Len(.Text): End With: End Sub

Private Sub UserForm_Initialize()
   Dim s$
   s = Trim$(GetSetting(sCorelDRAW, "sameColorSel", "Sens", "-0"))
   chGroups = (Left$(s, 1) = "-")
   txSens = Mid$(s, IIf(chGroups, 2, 1))
   UIPositionFromRegistry Me, "sameColorSel"
   End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
   SaveSetting sCorelDRAW, "sameColorSel", "WindowPos", Left & " " & Top
   End Sub
