Option Explicit
'
' assignment.vbs
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




' This subroutine allows us to ignore the difference
' between object and primitive assignments.  This is
' essential for many parts of the engine.
Public Sub Assign(ByRef var, ByVal val)
   If IsObject(val) Then
      Set var = val
   Else
      var = val
   End If
End Sub

' This is similar to the   ? :   operator of other languages.
' Unfortunately, both the   if_true   and    if_false    "branches"
' will be evalauted before the condition is even checked.  So,
' you'll only want to use this for simple expressions.
Public Function Choice(ByVal cond, ByVal if_true, ByVal if_false)
   If cond Then
      Assign Choice, if_true
   Else
      Assign Choice, if_false
   End If
End Function
