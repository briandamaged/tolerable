Option Explicit
'
' iterator.vbs
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



Require "t\core\assignment.vbs"
Require "t\core\enumerator.vbs"

' This is just here for reference
Class Iterator_Class
    Public Function HasNext()
    End Function
    
    Public Function PeekNext()
    End Function
    
    Public Function GetNext()
    End Function
    
    
    Public Function HasPrev()
    End Function
    
    Public Function PeekPrev()
    End Function
    
    Public Function GetPrev()
    End Function
End Class


Class Enumerator_Source_Iterator_Class

    Private m_iter
    
    Public Sub Initialize(ByVal iter)
        Set m_iter = iter
    End Sub
    
    Private Sub Class_Terminate()
        Set m_iter = Nothing
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        If m_iter.HasNext Then
            Assign retval, m_iter.GetNext
            successful = True
        Else
            successful = False
        End If
    End Sub
End Class


Public Function Enumerator_Iterator(ByVal iter)
    Dim retval
    Set retval = New Enumerator_Source_Iterator_Class
    retval.Initialize iter
    Set Enumerator_Iterator = Scanner(retval)
End Function

