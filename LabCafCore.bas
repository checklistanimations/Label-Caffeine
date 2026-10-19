Attribute VB_Name = "LabCafCore"
'Label Caffeine QUICK START
'To begin, paste this line in your Form_Initialize:
'    setupLabelCaffeine Me

'This sends your form to the framework and scans controls
'for tags using the LC_prefix.
'Adjust the prefix and toggle validation in the user subs
'below, and map your custom actions as needed.
Dim lastHovered As String
Dim lastGroup As String
Dim LC_prefix As String
Dim devMode As Boolean
Dim tooltipsState As Byte
Dim failLevel As Integer
Dim LabCafItems As Collection

'<START---------------------USER ADJUSTMENTS----------------------->'
Sub LabCafUserActions(LCaction As String, frm As Object)
  LCaction = LCase(LCaction)
  Select Case act
    Case LCaction = "your_custom_action_string_here1" 'call routine to run
    Case LCaction = "your_custom_action_string_here2" 'call routine to run
    Case LCaction = "your_custom_action_string_here3" 'call routine to run
    Case Else
      MsgBox "No Label Caffeine action mapped in USER_LabCafAction"
  End Select
End Sub

Sub LabCafUserSettings()
  LC_prefix = "LC=="
  devMode = True
  tooltipsState = 0 'set to 99 to completely bypass LC tooltip engine
End Sub
'<CLOSE---------------------USER ADJUSTMENTS----------------------->'




'<START---------------------Feedback Subs----------------------->'
Sub LabCafMouseDown(lc As LabCafItem)
  
End Sub

Sub LabCafMouseMove(lc As LabCafItem)
  Dim s As String
  Dim block As Variant, v As Variant
  If lc.name = currentlyHovered Then Exit Sub
  If lc.groupName = lastGroup Then Exit Sub
  If lc.hasError Then Exit Sub
  'unhover controls first
  s = lc.name
  If lc.useGroup = True Then s = s & ";" & lc.groupItems
  block = Split(s, ";")
  For Each v In block
    Call LabCafUnhover(LabCafItems(v))
  Next v
  For Each v In block
    LabCafHover LabCafItems(v)
    If LabCafItems(v).hasError Then
      Call LabCafUnhover(LabCafItems(v))
      If devMode Then
        MsgBox LabCafItems(v).Error & vbLf & "Program will end now"
        Call ExitLabelCaffeine
        Exit Sub
      End If
    End If
  Next v
  currentlyHovered = lc.name
End Sub

Sub LabCafFormMouseMove()
  If currentlyHovered = "labCafUserForm" Then Exit Sub
  LabCafUnhover LabCafItems(v)
  currentlyHovered = "labCafUserForm"
End Sub

Sub LabCafHover(lc As LabCafItem)
  Dim sc As Control
  Dim es As String, s As String
  On Error Resume Next
  Set sc = LabCafItems(lc.styleFrom)
  If devMode Then
    If lc.useBack Then lc.ctrl.BackColor = sc.BackColor
    If Err.Number <> 0 Then es = es & ";Backcolor": Err.Clear
    If lc.useBorder Then lc.ctrl.BorderColor = sc.BorderColor
    If Err.Number <> 0 Then es = es & ";BorderColor": Err.Clear
    If lc.useFont Then
      lc.ctrl.Font.Bold = sc.Font.Bold
      lc.ctrl.Font.Italic = sc.Font.Italic
      lc.ctrl.Font.name = sc.Font.name
      lc.ctrl.Font.Size = sc.Font.Size
      lc.ctrl.Font.Underline = sc.Font.Underline
    End If
    If Err.Number <> 0 Then es = es & ";Font": Err.Clear
    If lc.useFore Then lc.ctrl.ForeColor = sc.ForeColor
    If Err.Number <> 0 Then es = es & ";Forecolor": Err.Clear
    If lc.useImage Then lc.ctrl.Picture = sc.Picture
    If Err.Number <> 0 Then es = es & ";Image": Err.Clear
    If lc.useOpacity Then lc.ctrl.BackStyle = sc.BackStyle
    If Err.Number <> 0 Then es = es & ";Opacity": Err.Clear
    If lc.useOutline Then lc.ctrl.BorderStyle = sc.BorderStyle
    If Err.Number <> 0 Then es = es & ";Outline": Err.Clear
    If lc.useSpecial Then lc.ctrl.SpecialEffect = sc.SpecialEffect
    If Err.Number <> 0 Then es = es & ";Special": Err.Clear
    If es <> "" Then
      lc.hasError = True
      s = lc.name & " control crashed on hover for the following "
      s = s & "properties: " & es
      lc.errors = s
  Else
    If lc.useBack Then lc.ctrl.BackColor = sc.BackColor
    If lc.useBorder Then lc.ctrl.BorderColor = sc.BorderColor
    If lc.useFont Then
      lc.ctrl.Font.Bold = sc.Font.Bold
      lc.ctrl.Font.Italic = sc.Font.Italic
      lc.ctrl.Font.name = sc.Font.name
      lc.ctrl.Font.Size = sc.Font.Size
      lc.ctrl.Font.Underline = sc.Font.Underline
    End If
    If lc.useFore Then lc.ctrl.ForeColor = sc.ForeColor
    If lc.useImage Then lc.ctrl.Picture = sc.Picture
    If lc.useOpacity Then lc.ctrl.BackStyle = sc.BackStyle
    If lc.useOutline Then lc.ctrl.BorderStyle = sc.BorderStyle
    If lc.useSpecial Then lc.ctrl.SpecialEffect = sc.SpecialEffect
    If Err.Number <> 0 Then lc.hasError = True
  End If
End Sub

Sub LabCafUnhover(lc As LabCafItem)
  On Error Resume Next
  If lc.useBack Then lc.ctrl.BackColor = lc.orBack
  If lc.useBorder Then lc.ctrl.BorderColor = lc.orBorder
  If lc.useFont Then
    lc.ctrl.Font.Bold = lc.orFontBold
    lc.ctrl.Font.Italic = lc.orFontItalic
    lc.ctrl.Font.name = lc.orFontName
    lc.ctrl.Font.Size = lc.orFontSize
    lc.ctrl.Font.Underline = lc.orFontUnder
  End If
  If lc.useFore Then lc.ctrl.ForeColor = lc.orFore
  If lc.useImage Then lc.ctrl.Picture = lc.orImage
  If lc.useOpacity Then lc.ctrl.BackStyle = lc.orOpacity
  If lc.useOutline Then lc.ctrl.BorderStyle = lc.orOutline
  If lc.useSpecial Then lc.ctrl.SpecialEffect = lc.orSpecial
  Err.Clear
  On Error GoTo 0
End Sub


'<CLOSE---------------------Feedback Subs----------------------->'



'!===!===!===!===!===!===!===!<<CAUTION>>!===!===!===!===!===!===!===!
'                          IMPORTANT  NOTICE
'  The code below is core Label Caffeine framework logic.
'  Editing it may break control routing, validation, or action
'  handling. Only experienced developers should continue.
'  All intended user adjustments are already above this block.
'!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!=!



Sub LabCafSetup(frm As Object)
  Dim lcca As Collection
  Dim c As Variant
  Dim lcItem As LabCafItem
  
  If TypeName(frm) <> "userform" Then failLevel = 100
  If failLevel <> 0 Then EndLabCafSetup lcca, frm: Exit Sub
  Set lcca = New Collection
  Call LabCafUserSettings
  Call LabCafcreateLCCA(frm, lcca)
  If failLevel <> 0 Then EndLabCafSetup lcca, frm: Exit Sub
  Call LabCafSetGroups(lcca)
  If runValidation Then
    'verification steps
  End If
  If failLevel <> 0 Then EndLabCafSetup lcca, frm: Exit Sub
  'build the objects
  For Each c In lcca
    Set lcItem = New LabCafItem
    lcItem.drink c, frm
    LabCafItems.Add lcItem, c(0)
  Next c
  Set lcca = Nothing
  
End Sub


Sub LabCafcreateLCCA(frm As Object, lcca As Collection)
  '0=c.Name, 1=c.type, 2=c.tag, 3=useprops, 4=styleFrom
  '5=group name, 6=group items, 7=action, 8=expandTo
  '9=tooltip, 10=errors
  Dim c As Control
  Dim tg As String, tgPrevix As String, s As String
  Dim plen As Integer
  pfxlen = Len(LC_prefix)
    
  For Each c In frm.Controls
    tg = c.Tag
    If tg <> "" Then
      tgPrefix = Left(tg, pfxlen)
      If LCase(tgPrevix) = LCase(LC_prefix) Then
        Dim block(0 To 10) As Variant
        block(2) = Mid(tg, pfxlen + 1) 'put tag in array
        block(0) = c.name
        block(1) = TypeName(c)
        Call LabCafFillArray(block)
        'block is now full
        lcca.Add block, c.name
      End If
    End If
  Next c
  If lcca.Count <> 0 Then
    'process groups
    Call LabCafSetGroups(lcca)
  End If
  
  
End Sub

Sub LabCafFillArray(block As Variant)
  'takes the array that has the tag string and parses it
  Dim pair As Variant, temp As Variant, t As Variant
  Dim ky As String, vl As String, props As String
  Dim pass As Boolean
  temp = Split(block(2), ";")
  props = ""
  For Each t In temp
    pair = Split(t, "=")
    If UBound(pair) = 2 Then
      ky = Trim(LCase(pair(0)))
      vl = Trim(pair(1))
      pass = True
      If Len(ky) < 4 Then pass = False
      If vl = "0" Or vl = "" Then pass = False
      If pass = True Then
        If Left(ky, 3) = "use" Then
          props = props & ";" & ky & "=" & vl
        Else
          Select Case ky
            Case ky = "stylefrom": block(4) = vl
            Case ky = "group": block(5) = vl
            Case ky = "action": block(7) = vl
            Case ky = "expandto": block(8) = vl
            Case ky = "tooltip": block(9) = vl
          End Select
        End If
      End If
    End If
  Next t
  If Len(props) <> 0 Then block(3) = Mid(props, 2)
End Sub

Sub LabCafSetGroups(lcca As Collection)
  Dim groupIndexes As Collection
  Dim X As Long, i As Long
  Dim groupItems() As String
  Dim c As Variant
  Dim gname As String, s As String
  Set groupIndexes = New Collection
  'get all control names per group
  ReDim groupItems(1 To lcca.Count) As String
  X = 1
  For Each c In lcca
    gname = c(5)
    If gname <> "" Then
      If LabCafInCollection(groupIndexes, gname) = False Then
        groupIndexes.Add X, gname
        groupItems(X) = "" 'array
        X = X + 1
      End If
      i = CLng(groupIndexes(gname))
      s = groupItems(i)
      s = s & ";" & c(0)
      groupItems(i) = s
    End If
  Next c
  
  'add them to the correct index in lcca
  For Each c In lcca
    gname = c(5)
    If gname <> "" Then
      i = CLng(groupIndexes(gname))
      s = groupItems(i) & ";"
      s = Replace(s, ";" & c(0) & ";", ";") 'take out the main control name
      c(6) = Mid(s, 2, Len(s) - 2) 'take off leading encapsulating semicolons
    End If
  Next c
  Set groupIndexes = Nothing
End Sub

'we will send a packet.  not doing long parameter lists
Sub runPipeline(lcca As Collection, frm As UserForm, failLevel As Integer)

End Sub

Sub EndLabCafSetup(lcca As Collection, frm As Object)
  Dim s As String
  Dim c As Variant
  s = ""
  Select Case failLevel
    Case 100: s = "Userform not recognized. Cant continue"
  End Select
  

  'print each error
  If failLevel < 100 Then
    If lcca.Count <> 0 Then
      s = s & vbCrLf & "Please see all errors in the immediate window"
      Debug.Print vbCrLf & vbCrLf
      Debug.Print "vvvvvvvvvvvvvvvv" & vbCrLf & s & " -- START BLOCK"
      For Each c In lcca
        Debug.Print lcca(10)
      Next c
      Debug.Print s & " -- END BLOCK" & vbCrLf & "^^^^^^^^^^^^^^^^"
    End If
  End If
  
  'destroy objects
  Set lcca = Nothing
  Set frm = Nothing
  
  If s <> "" Then
    MsgBox s, vbCritical
  End If
End Sub

Sub ExitLabelCaffeine()
  Set LabCafItems = Nothing
End Sub


'<START--------------helper functions-------------->'
Function LabCafInCollection(recs As Collection, ky As String) As Boolean
  On Error Resume Next
  isInCollection = Not recs(ky) Is Nothing
  On Error GoTo 0
End Function



'<CLOSE--------------helper functions-------------->'


