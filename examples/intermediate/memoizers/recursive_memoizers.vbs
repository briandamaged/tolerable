Option Explicit

Require "t\core\memoizer.vbs"



Private Function Fibb_recursive(ByVal memo, ByVal x)
    WScript.Echo "Computing fibb(" & x & ")"
    If x = 0 Then
        Fibb_recursive = 1
    ElseIf x = 1 Then
        Fibb_recursive = 1
    Else
        Fibb_recursive = memo(x - 1) + memo(x - 2)
    End If
End Function






Dim temp : Set temp = Lambda(1, "Invoke = stored(0)(stored(1), arg0)")

Dim fibb : Set fibb = Memoizer(1, temp)
temp.Stored = Array(GetRef("Fibb_Recursive"), fibb)


WScript.Echo "Fibb(3) = " & fibb(3)
WScript.Echo "Fibb(5) = " & fibb(5)