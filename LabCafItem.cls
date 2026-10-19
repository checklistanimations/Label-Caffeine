VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LabCafItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents frm As MSForms.UserForm
Attribute frm.VB_VarHelpID = -1
Private WithEvents lbl As MSForms.Label
Attribute lbl.VB_VarHelpID = -1

Public ctrl As MSForms.Control
Public parentForm As MSForms.UserForm

'boolean checking properties
Public useBack As Boolean
Public useOpacity As Boolean
Public useBorder As Boolean
Public useOutline As Boolean
Public useFont As Boolean
Public useFore As Boolean
Public useImage As Boolean
Public useSpecial As Boolean
Public useGlass As Boolean
Public useExpandTo As Boolean
Public useStyleFrom As Boolean
Public useAction As Boolean
Public useGroup As Boolean
Public useTooltip As Boolean
Public forceSoloClickable As Boolean
Public hasError As Boolean

'base properties
Public styleFrom As String
Public expandTo As String
Public action As String
Public groupName As String
Public groupItems As String
Public propItems As String
Public tooltip As String
Public flag As String
Public errors As String

'original Properties
Public name As String
Public orLeft As Single
Public orTop As Single
Public orHeight As Single
Public orWidth As Single
Public orBack As Long
Public orOpacity As Byte
Public orBorder As Long
Public orOutline As Byte
Public orFontName As String
Public orFontBold As Boolean
Public orFontItalic As Boolean
Public orFontSize As Single
Public orFontUnder As Boolean
Public orFontStrike As Boolean
Public orFore As Long
Public orImage As StdPicture
Public orSpecial As Byte

'additional Properties
Public clickableLeft As Single
Public clickableRight As Single
Public clickableUp As Single
Public clickableDown As Single





Public Sub drink(block As Variant, pf As UserForm)
  'wiring
  'block is an array, only 1 control at a time
  
  '0=c.Name, 1=c.type, 2=c.tag, 3=useprops, 4=styleFrom
  '5=group name, 6=group items, 7=action, 8=expandTo
  '9=tooltip, 10=errors
  Set ctrl = pf.Controls(block(0))
  Set parentForm = pf
  If TypeName(ctrl) = "UserForm" Then Set frm = ctrl
  If TypeName(ctrl) = "Label" Then Set lbl = ctrl
  Call BruteForceSetBaseProps(block)
  Call SetUseProps
End Sub

Private Sub BruteForceSetBaseProps(block As Variant)
  On Error Resume Next
  'these are all always set
  Me.name = block(0)
  Me.propItems = block(3)
  Me.styleFrom = block(4)
  Me.groupName = block(5)
  Me.groupItems = block(6)
  Me.action = block(7)
  Me.expandTo = block(8)
  Me.tooltip = block(9)
  Me.orLeft = Me.ctrl.Left
  Me.orTop = Me.ctrl.Top
  Me.orHeight = Me.ctrl.Height
  Me.orWidth = Me.ctrl.Width
  Me.orBack = CLng(Me.ctrl.BackColor)
  Me.orOpacity = CByte(Me.ctrl.BackStyle)
  Me.orBorder = CLng(Me.ctrl.Border.Color)
  Me.orOutline = CByte(Me.ctrl.Border.LineStyle)
  Me.orFontBold = Me.ctrl.Font.Bold
  Me.orFontItalic = Me.ctrl.Font.Italic
  Me.orFontName = Me.ctrl.Font.name
  Me.orFontSize = Me.ctrl.Font.Size
  Me.orFontStrike = Me.ctrl.Font.Strikethrough
  Me.orFontUnder = Me.ctrl.Font.Underline
  Me.orFore = CLng(Me.ctrl.ForeColor)
  Set Me.orImage = Me.ctrl.Picture
  Me.orSpecial = CByte(Me.ctrl.SpecialEffect)
  Me.hasError = False
  Me.errors = ""
  On Error GoTo 0
End Sub

Private Sub SetUseProps()
  Dim s As String
  s = Me.propItems
  If InStr(1, s, "useback=1") Then Me.useBack = True
  If InStr(1, s, "useopacity=1") Then Me.useOpacity = True
  If InStr(1, s, "useborder=1") Then Me.useBorder = True
  If InStr(1, s, "useoutline=1") Then Me.useOutline = True
  If InStr(1, s, "usefont=1") Then Me.useFont = True
  If InStr(1, s, "usefore=1") Then Me.useFore = True
  If InStr(1, s, "useimage=1") Then Me.useImage = True
  If InStr(1, s, "usespecial=1") Then Me.useSpecial = True
  If InStr(1, s, "useglass=1") Then Me.useGlass = True
  If InStr(1, s, "forceSoloClickable=1") Then Me.forceSoloClickable = True
  If Me.expandTo <> "" Then Me.useExpandTo = True
  If Me.styleFrom <> "" Then Me.useStyleFrom = True
  If Me.action <> "" Then Me.useAction = True
  If Me.groupName <> "" Then Me.useGroup = True
  If Me.tooltip <> "" Then Me.useTooltip = True
End Sub



'<START-------------------------Control Events------------------------->'
Private Sub lbl_Click()
  Exit Sub
End Sub

Private Sub lbl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Exit Sub
End Sub

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call LabCafCore.LabCafMouseMove(Me)
End Sub

Private Sub Label1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub



Private Sub Label1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub


'add mousedown (dont run when out of bounds)
'add mouseup
'add mouse move

Private Sub frm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Call LabCafCore.LabCafFormMouseMove
End Sub

Private Sub frm_QueryClose(Cancel As Integer, CloseMode As Integer)

End Sub


'<CLOSE-------------------------Control Events------------------------->'
