Option Explicit

Require "t\core\closures.vbs"


' The most basic form of Lambda will just compile
' any String that you pass it into a function.
' For example, here is a function that returns a
' Closure that increments its argument by one.
Public Function Create_IncrementBy1()
    Set Create_IncrementBy1 = Lambda(1, "Invoke = arg0 + 1")
End Function

Dim f : Set f = Create_IncrementBy1()
WScript.Echo "f(5) = " & f(5)



' Of course, that really wasn't that exciting, and
' it begs the question: why did we bother using a
' Lambda here at all?  In reality, this simple form
' of Lambda is really only useful when you are
' implementing a Strategy pattern (see the Enumerator
' class for many examples of this).  In most other
' cases, we'd like to change the way our Closure
' behaves based upon the input we supply to the
' Function that created the Closure.  This can be
' achieved by calling the .Store(...) method on
' the Closure.

' This second version of the Function will create a
' Closure that increments values by a specified amount:
Public Function Create_IncrementBy(ByVal value)
    Set Create_IncrementBy = Lambda(1, "Invoke = arg0 + stored").Store(value)
End Function

Dim g : Set g = Create_IncrementBy(3)
WScript.Echo "g(5) = " & g(5)



' Furthermore, the values that you store in the Closure
' don't need to be treated as Constants.  In fact, you can
' actually write Closures that update their internal state
' each time they are called.  For instance, this Function
' will return a Closure that increments its value each
' time it is invoked:
Public Function Create_Counter()
    Set Create_Counter = Lambda(0, "Invoke = stored : stored = stored + 1").Store(0)
End Function

Dim h : Set h = Create_Counter()
WScript.Echo "h() = " & h()
WScript.Echo "h() = " & h()
WScript.Echo "h() = " & h()



