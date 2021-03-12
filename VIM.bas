Attribute VB_Name = "VIM"
Option Explicit

Declare PtrSafe Sub keybd_event Lib "User32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Declare PtrSafe Function GetKeyState Lib "User32.dll" (ByVal vKey As Long) As Long

Const KEYUP = &H2
Global v_mod As Boolean
Global tail_cell As Range
Public Sub toggle_v_mode()
    If v_mod = False Then
        Set tail_cell = ActiveCell
        v_mod = True
        Call VShortcuts
        Call select_mode
    Else
        v_mod = False
        Call NShortcuts
    End If
    Debug.Print "v_mod: " & v_mod
End Sub
Sub VShortcuts()
    ' movement
    Application.OnKey "{h}", "move_left_tail"
    Application.OnKey "{l}", "move_right_tail"
    Application.OnKey "{j}", "move_down_tail"
    Application.OnKey "{k}", "move_up_tail"
End Sub
Sub NShortcuts()
    ' movement
    Application.OnKey "{h}", "move_left"
    Application.OnKey "{l}", "move_right"
    Application.OnKey "{j}", "move_down"
    Application.OnKey "{k}", "move_up"
End Sub

Sub CreateShortcut()
    Application.OnKey "{i}", "insert_mode"
    Application.OnKey "{v}", "toggle_v_mode"
    ' increment
    Application.OnKey "^{a}", "C_a"
    ' append bottom value
    Application.OnKey "{J}", "J"
End Sub
Sub DeleteShortcut()
    'Application.OnKey "{i}"
End Sub
Sub move_right()
    On Error GoTo ErrorHandler
    ActiveCell.Offset(0, 1).Activate
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & vbNewLine & Err.Description
End Sub
Sub move_left()
    On Error GoTo ErrorHandler
    ActiveCell.Offset(0, -1).Activate
Exit Sub
ErrorHandler:
    Debug.Print Err.Number & vbNewLine & Err.Description
End Sub
Sub move_up()
    On Error GoTo ErrorHandler
    ActiveCell.Offset(-1, 0).Activate
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & vbNewLine & Err.Description
End Sub
Sub move_down()
    On Error GoTo ErrorHandler
    ActiveCell.Offset(1, 0).Activate
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & vbNewLine & Err.Description
End Sub
Sub move_right_tail()
    On Error GoTo ErrorHandler
    Set tail_cell = tail_cell.Offset(0, 1)
    Call select_mode
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & vbNewLine & Err.Description
End Sub
Sub move_left_tail()
    On Error GoTo ErrorHandler
    Set tail_cell = tail_cell.Offset(0, -1)
    Call select_mode
Exit Sub
ErrorHandler:
    Debug.Print Err.Number & vbNewLine & Err.Description
End Sub
Sub move_up_tail()
    On Error GoTo ErrorHandler
    Set tail_cell = tail_cell.Offset(-1, 0)
    Call select_mode
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & vbNewLine & Err.Description
End Sub
Sub move_down_tail()
    On Error GoTo ErrorHandler
    Set tail_cell = tail_cell.Offset(1, 0)
    Call select_mode
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & vbNewLine & Err.Description
End Sub
Sub insert_mode()
    On Error GoTo ErrorHandler
    keybd_event vbKeyF2, 0, 0, 0
    keybd_event vbKeyF2, 0, KEYUP, 0
    Exit Sub
ErrorHandler:
    Debug.Print Err.Number & vbNewLine & Err.Description
End Sub
Sub C_a()
    ' increment action
    
    ' get selection string
    ' cut rightmost number
    ' number+=1
    ' glue number back
    
    ' aaa1a -> aaa2a
    ' a3aa9a -> a3aa10a
    ' -00003 -> -00004
    ' -1 -> 0
    ' -5 -> -4
    ' -100 -> -99 [fail] now -099
    
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    Dim matchobj ' as
    Dim replace_str As String

    With regex
        .Pattern = "((-(?!0))?\d+)(?!.*\d)"
        .Global = True
        .MultiLine = False
    End With

    Set matchobj = regex.Execute(ActiveCell.Value)
    'matchobj.count
    'Dim mtc As Match
    'Set mtc = matchobj(0)
    If regex.test(ActiveCell.Value) Then
        replace_str = matchobj(0).Value
    End If
    
    Dim digits As Integer
    digits = Len(replace_str)
    
    ' if negative number cut "-" position from "000" text format
    If InStr(1, replace_str, "-") = 1 Then
        digits = digits - 1
    End If
   
    If regex.test(ActiveCell.Value) Then
        Dim fmt As String
        fmt = String$(digits, "0")
        Dim return_int As Integer
        return_int = CInt(replace_str) + 1 'TODO -1 for C_x
        Dim return_str As String
        return_str = Format(CStr(return_int), fmt)
        ActiveCell.Value = regex.Replace(ActiveCell.Value, return_str)
    End If
End Sub
Sub J()
    ' activecell value += bottomcell value
    ' clear bottomcell value
    If ActiveCell.Offset(1, 0).Value <> "" Then
        ActiveCell.Value = ActiveCell.Value & " " & ActiveCell.Offset(1, 0).Value
        ActiveCell.Offset(1, 0).ClearContents
    End If
End Sub

Sub select_mode()

    Dim active_cell As Range
    Set active_cell = ActiveCell

    Range(active_cell, tail_cell).Select 'active cell lose position
    active_cell.Activate 'again
    
End Sub
