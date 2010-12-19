Option Explicit

Require "t\core\enumerator.vbs"

' Enumerators provide a number of methods that allow you to specify
' operations that will be performed on your data in the future.  For
' instance, let's start by creating a list of integers from 0 to 50:
Dim x : Set x = eRange(0, 50)


' For some reason, I decided I only like odd integers today.  So, we'll
' use the Retain(...) method to keep only the odd integers:
Set x = x.Retain("arg0 mod 2 = 1")


' I also hate numbers that are divisible by 5, so let's discard these:
Set x = x.Discard("arg0 mod 5 = 0")


' Perfect squares are cool.  Let's square every value in x:
Set x = x.Map("arg0 * arg0")


' Alright, let's print the result into a comma-separated list:
WScript.Echo "Basic Enumerator operations example:" & vbCR & Join(X.TO_Array, ", ")




' Of course, these are fluent methods, so we could have done everything
' on one line:
WScript.Echo "Fluent Methods example:" & vbCR & Join(eRange(0, 50).Retain("arg0 mod 2 = 1").Discard("arg0 mod 5 = 0").Map("arg0 * arg0").TO_Array, ", ")



' Here is the equivalent code using vanilla VBScript.
ReDim temp(50)
Dim index : index = -1
Dim i
For i = 0 To 50                      ' eRange(0, 50)
    If i mod 2 = 1 Then              ' Retain("arg0 mod 2 = 1")
        If i mod 5 <> 0 Then         ' Discard("arg0 mod 5 = 0")
            index       = index + 1
            temp(index) = i * i      ' Map("arg0 * arg0")
        End If
    End If
Next
ReDim Preserve temp(index)
WScript.Echo "Vanilla VBScript example:" & vbCR & Join(temp, ", ")


' Obviously, this code is quite a bit longer and more difficult
' to read.  What might not be so obvious is that the meaning of
' this code is very rigid.  The only thing that this code does
' is print a comma-separated list of integers, and it does this
' by constructing a temporary array.
'
' What if we wanted to print one number at a time instead?  We
' could iterate over the temporary Array using a For Each loop,
' but this begs the question: Why did we create the temporary
' Array in the first place?  Couldn't we just print each number
' rather than putting it into an Array?
'
' Likewise, what if we later decide that we don't want any perfect
' squares that are greater than 50?  With the vanilla VBScript
' code, we would have wasted a number of cycles computing values
' that we didn't need.  Since Enumerators are lazily evaluated,
' they will only compute as many values as necessary.  Therefore,
' we can add the condition "Stop before the result is greater than
' 50" to the end of our Enumerator specification, and the Enumerator
' will take care of the rest of the details.



' To demonstrate these points, let's create a function that returns
' the Enumerator we have been using:
Public Function Example()
    Set Example = eRange(0, 50).Retain("arg0 mod 2 = 1").Discard("arg0 mod 5 = 0").Map("arg0 * arg0")
End Function


' We'll start w/ the original example again: Creating a comma-separated string.
WScript.Echo "Basic function returning Enumerator:" & vbCR & Join(Example.TO_Array, ", ")


' This time, we'll only print values one at a time, and we'll stop before
' we get to a value that is greater than 50:
WScript.Echo "Now we'll print the values one at a time"
Example.StopBefore("arg0 > 50").DoEach("WScript.Echo arg0")



