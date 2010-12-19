Option Explicit

Require "t\core\memoizer.vbs"


' It's often handy to have a function that checks the Type of
' an object.  Rather than writing a new function each time we
' need a different type checker, we can write one function that
' writes the type checkers for us:
Public Function TypeChecker1(ByVal t)
    WScript.Echo "Creating type checker for " & t
    Set TypeChecker1 = Lambda(1, "Invoke = (TypeName(arg0) = stored)").Store(t)
End Function

Dim is_string  : Set is_string  = TypeChecker1("String")
Dim is_integer : Set is_integer = TypeChecker1("Integer")

WScript.Echo "Is ""bob"" a String?"   & vbCR & is_string("bob")
WScript.Echo "Is ""bob"" an Integer?" & vbCR & is_integer("bob")



' The problem w/ this approach is that TypeChecker1(...) is not
' smart enough to recognize when it has already constructed a
' particular type checker for us.  It will run exactly the same
' code again and construct another Lambda that does the same
' thing.
'
' In other words, when we run this statement, we will see the
' "Creating type checker" message again:
Dim is_string2 : Set is_string2 = TypeChecker1("String")



' We can prevent this from happening by wrapper TypeChecker1(...)
' with a Memoizer.
Dim TypeChecker2 : Set TypeChecker2 = Memoizer(1, GetRef("TypeChecker1"))


' The first time we invoke the Memoizer, we'll see the same message.
' Remember: the first 2 times we created the String type checker,
' we did it without the aid of the Memoizer, so it doesn't know
' any better!
Dim is_string3 : Set is_string3 = TypeChecker2("String")


' This time, however, the Memoizer won't bother constructing a new
' type checker.  It will realize that it has already constructed a
' String type checker and return its cached copy:
WScript.Echo "This time, we'll use the cached copy rather than running the code again"
Dim is_string4 : Set is_string4 = TypeChecker2("String")
WScript.Echo "Hopefully the ""Creating type checker"" message didn't print this time!"