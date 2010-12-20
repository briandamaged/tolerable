Option Explicit
'
' linked_list.vbs
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

Class LinkedList_Node_Class
    Public  m_prev
    Public  m_next
    Public  m_value

    Private Sub Class_Initialize()
        Set m_prev = Nothing
        Set m_next = Nothing
    End Sub
    
    Private Sub Class_Terminate()
        Set m_prev  = Nothing
        Set m_next  = Nothing
        Set m_value = Nothing
    End Sub
    
    Public Sub SetValue(ByVal value)
        Assign m_value, value
    End Sub

End Class


Class Iterator_LinkedList_Class
    Private m_left
    Private m_right
    
    Public Sub Initialize(ByVal r)
        Set m_left  = Nothing
        Set m_right = r
    End Sub
    
    Private Sub Class_Terminate()
        Set m_Left  = Nothing
        Set m_Right = Nothing
    End Sub
    
    Public Function HasNext()
        HasNext = Not(m_right Is Nothing)
    End Function
    
    Public Function PeekNext()
        Assign PeekNext, m_right.m_value
    End Function
    
    Public Function GetNext()
        Assign GetNext, m_right.m_value
        Set m_left  = m_right
        Set m_right = m_right.m_next
    End Function
    
    Public Function HasPrev()
        HasPrev = Not(m_left Is Nothing)
    End Function
    
    Public Function PeekPrev()
        Assign PeekPrev, m_left.m_value
    End Function
    
    Public Function GetPrev()
        Assign GetPrev, m_left.m_value
        Set m_right = m_left
        Set m_left  = m_left.m_prev
    End Function
    
End Class




Class LinkedList_Class
    Private m_first
    Private m_last
    Private m_size
    
    Private Sub Class_Initialize()
        Me.Reset
    End Sub
    
    Private Sub Class_Terminate()
        Me.Reset
    End Sub
    
    Public Function Clear()
        Set m_first = Nothing
        Set m_last  = Nothing
        m_size      = 0
        Set Clear   = Me
    End Function
    
    Private Function NewNode(ByVal value)
        Dim retval
        Set retval = New LinkedList_Node_Class
        retval.SetValue value
        Set NewNode = retval
    End Function
    
    Public Sub Reset()
        Set m_first = Nothing
        Set m_last  = Nothing
        m_size      = 0
    End Sub
    
    Public Function IsEmpty()
        IsEmpty = (m_last Is Nothing)
    End Function
    
    Public Property Get Count
        Count = m_size
    End Property
    
    Public Function Iterator()
        Dim retval
        Set retval = New Iterator_LinkedList_Class
        retval.Initialize m_first
        Set Iterator = retval
    End Function
    
    Public Function Push(ByVal value)
        Dim temp
        Set temp = NewNode(value)
        If Me.IsEmpty Then
            Set m_first = temp
            Set m_last  = temp
        Else
            Set temp.m_prev   = m_last
            Set m_last.m_next = temp
            Set m_last        = temp
        End If
        m_size = m_size + 1
        Set Push = Me
    End Function


    Public Function Peek()
        ' TODO: Error handling
        Assign Peek, m_last.m_value
    End Function

    ' Alias for Peek
    Public Function Back()
        ' TODO: Error handling
        Assign Back, m_last.m_value
    End Function
    
    Public Function Pop()
        Dim temp
        
        ' TODO: Error Handling
        Assign Pop, m_last.m_value

        Set temp          = m_last
        Set m_last        = temp.m_prev
        Set temp.m_prev   = Nothing
        If m_last Is Nothing Then
            Set m_first = Nothing
        Else
            Set m_last.m_next = Nothing
        End If
        m_size = m_size - 1
    End Function



    Public Function Unshift(ByVal value)
        Dim temp
        Set temp = NewNode(value)
        If Me.IsEmpty Then
            Set m_first = temp
            Set m_last  = temp
        Else
            Set temp.m_next    = m_first
            Set m_first.m_prev = temp
            Set m_first        = temp
        End If
        m_size = m_size + 1
        Set Unshift = Me
    End Function



    ' Alias for Peek
    Public Function Front()
        ' TODO: Error handling
        Assign Front, m_first.m_value
    End Function
    
    Public Function Shift()
        Dim temp
        
        ' TODO: Error Handling
        Assign Shift, m_first.m_value

        Set temp          = m_first
        Set m_first       = temp.m_next
        Set temp.m_next   = Nothing
        If m_first Is Nothing Then
            Set m_last = Nothing
        Else
            Set m_first.m_prev = Nothing
        End If
        m_size = m_size - 1
    End Function
    
    
    
    Public Function TO_Array()
        Dim i, iter
        
        ReDim retval(Me.Count - 1)
        i = 0
        Set iter = Me.Iterator
        While iter.HasNext
            retval(i) = iter.GetNext
            i = i + 1
        Wend
        TO_Array = retval
    End Function
    
    Public Function TO_En()
        Set TO_En = Enumerator_Iterator(Iterator)
    End Function


End Class


Public Function LinkedList
    Set LinkedList = New LinkedList_Class
End Function


Converter_Enumerator.Given(L_Type("LinkedList_Class")).Dispatch("arg0.TO_En")

