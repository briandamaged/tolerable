Option Explicit

Require "t\core\closures.vbs"

' We'll start out by defining a few ordinary functions.
' Nothing too surprising here.
Public Function f(ByVal x)
    f = x + 2
End Function

Public Function g(ByVal x)
    g = x * 2
End Function


' Now we're define the Compose function.  This function accepts
' two functions as arguments, and it returns their composition.
' In other words:  Compose(X, Y) => X(Y(a))
Public Function Compose(ByVal f1, ByVal f2)
    Set Compose = Lambda(1, "Invoke = stored(0)(stored(1)(arg0))").Store(Array(f1, f2))
End Function


' Let's see if it worked.  We'll use the Compose function to
' create a new function X.  Invoking X(a) will be equivalent
' to invoking f(g(a))
Dim X : Set X = Compose(GetRef("f"), GetRef("g"))
WScript.Echo "f(g(3)) = " & f(g(3))
WScript.Echo "X(3) = " & X(3)

' We'll compose the functions in the opposite order
' to create Y
Dim Y : Set Y = Compose(GetRef("g"), GetRef("f"))
WScript.Echo "g(f(3)) = " & g(f(3))
WScript.Echo "Y(3) = " & Y(3)



' Hopefully it will come as no surprise to learn that
' Compose(...) works equally well when its arguments
' are Closures.  In fact, the only real difference
' between Functions and Closures is that you need to
' use the annoying GetRef(...) function to get a
' "pointer" to the function.  In contrast, you can just
' refer to Closures directly.  Thus, no calls to
' GetRef(...) are needed below:
Dim Z : Set Z = Compose(X, Y)
WScript.Echo "f(g(g(f(3)))) = " & f(g(g(f(3))))
WScript.Echo "X(Y(3)) = " & X(Y(3))
WScript.Echo "Z(3) = " & Z(3)

