
Require "t\core\closures.vbs"
Require "t\core\memoizer.vbs"

Dim L_Type_Singleton__
Public Function L_Type(ByVal t)
    If IsEmpty(L_Type_Singleton__) Then
        Set L_Type_Singleton__ = Memoizer(1, Q("Lambda(1, 'Invoke = (TypeName(arg0) = stored)').Store(arg0)"))
    End If
    Set L_Type = L_Type_Singleton__(t)
End Function


Private Function L_IsArray_Wrapper(ByVal x)
    L_IsArray_Wrapper = IsArray(x)
End Function

Public Function L_IsArray()
    Set L_IsArray = GetRef("L_IsArray_Wrapper")
End Function

