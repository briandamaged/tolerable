Option Explicit
'
' functional.vbs
' Copyright (C) Brian Lauber 2010 <constructible.truth@gmail.com>
'
' Tolerable is free software: you can redistribute it and/or modify it
' under the terms of the GNU Lesser General Public License as published
' by the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' Tolerable is distributed in the hope that it will be useful, but
' WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
' See the GNU Lesser General Public License for more details.
'
' You should have received a copy of the GNU Lesser General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'


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

