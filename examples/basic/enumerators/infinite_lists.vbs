Option Explicit

Require "t\core\enumerator.vbs"

' Enumerators use lazy evaluation for just about everything.
' Because of this, they can operate on infinite lists of data
' without any difficulty.

' For instance, I can create an infinite list of even integers
' as follows:
Dim x : Set x = eCounter.Discard("arg0 mod 2 = 0")


' Of course, we all know that it would require an infinite number of
' operations to construct an infinite list.  Hence, if we constructed
' the whole list upfront, then our program would never finish.
WScript.Echo "If this prints, then the program has not gone into an infinite loop!"


' Phew.  Looks like the program's still running.  So, let's grab the
' first 5 values and make sure everything looks okay.
WScript.Echo x.Pop
WScript.Echo x.Pop
WScript.Echo x.Pop
WScript.Echo x.Pop
WScript.Echo x.Pop



' Of course, you can always turn an infinite list into a finite list.
' Let's just grab the next 5 values and call it quits.
Set x = x.Limit(5)

' Since x is now a finite list, we can store its entire contents into
' an Array.  So, let's create an Array and use VBScript's Join(...)
' function to print everything at once
WScript.Echo Join(x.TO_Array, ", ")