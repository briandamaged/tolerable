Option Explicit

Require "t\core\memoizer.vbs"

' Everyone knows that you can't demonstrate a memoizer
' without implementing an efficient Fibonacci number
' function.  So, without further ado...



Private Function Fibb__(ByVal x)
    WScript.Echo "Computing Fibb(" & x & ")"
    If x = 0 Then
        Fibb__ = 1
    ElseIf x = 1 Then
        Fibb__ = 1
    Else
        ' Notice that the return value is based upon
        ' Fibb(...) instead of Fibb__(...).  Fibb is
        ' the name that we will give to the Memoizer
        ' that will wrap the Private function Fibb(...)
        Fibb__ = Fibb(x - 1) + Fibb(x - 2)
    End If
End Function

' Now we'll define our Memoizer.  As promised, we have
' named it Fibb(...)
Dim Fibb : Set Fibb = Memoizer(1, GetRef("Fibb__"))


' Let's check whether or not our implementation is
' working correctly.  When we compute Fibb(3), we
' should only see the values of Fibb(2), Fibb(1),
' and Fibb(0) being computed once.
WScript.Echo "Fibb(3) is " & Fibb(3)


' Likewise, when we compute Fibb(5), we should only
' see that the value of Fibb(4) still needs to be
' computed.  All other values have already been cached.
WScript.Echo "Fibb(5) is " & Fibb(5)



' However, there is a flaw w/ this approach: it introduces
' unnecessary coupling between Fibb(...) and Fibb__(...).
' In particular, Fibb__(...) must know the exact name of
' the Memoizer.  Hence, it would be impossible to create
' an anonymous Memoizer using this approach.

' I'll create an Intermediate or Advanced example that
' shows how to fix this.
