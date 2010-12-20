Option Explicit
'
' dynamic_array.vbs
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


Require "t\core\iterator.vbs"

Class DynamicArray_Class
    Private m_data
    Private m_size
    

    
    Public Sub Initialize(ByVal d, ByVal s)
        m_data = d
        m_size = s
    End Sub
    
    Private Sub Class_Terminate()
        Set m_data = Nothing
    End Sub
    
    
    Public Property Get Capacity
        Capacity = UBOUND(m_data) + 1
    End Property
    
    Public Property Get Count
        Count = m_size
    End Property
    
    ' Alias for Count
    Public Property Get Size
        Size = m_size
    End Property
    
    Public Function IsEmpty()
        IsEmpty = (m_size = 0)
    End Function
    
    Public Function Clear()
        m_size    = 0
        Set Clear = Me
    End Function
    
    Private Sub Grow
        ' TODO: There's probably a better way to
        '       do this.  Doubling might be excessive
        ReDim Preserve m_data(UBOUND(m_data) * 2)
    End Sub
    
    Public Function Push(ByVal val)
        If m_size >= UBOUND(m_data) Then
            Grow
        End If
        Assign m_data(m_size), val
        m_size = m_size + 1
        Set Push = Me
    End Function
    
    
    ' Look at the last element
    Public Function Peek()
        Assign Peek, m_data(m_size - 1)
    End Function
    
    ' Look at the last element and
    ' pop it off of the list
    Public Function Pop()
        Assign Pop, m_data(m_size - 1)
        m_size = m_size - 1
    End Function
    
    
    ' If pseudo_index < 0, then we assume we're counting
    ' from the back of the Array.
    Private Function CalculateIndex(ByVal pseudo_index)
        If pseudo_index >= 0 Then
            CalculateIndex = pseudo_index
        Else
            CalculateIndex = m_size + pseudo_index
        End If
    End Function
    
    Public Default Function Item(ByVal i)
        Assign Item, m_data(CalculateIndex(i))
    End Function
    
    
    ' This does not treat negative indices as wrap-around.
    ' Thus, it is slightly faster.
    Public Function FastItem(ByVal i)
        Assign FastItem, m_data(i)
    End Function


    Public Function Slice(ByVal s, ByVal e)
        s = CalculateIndex(s)
        e = CalculateIndex(e)
        If e < s Then
            Set Slice = DynamicArray()
        Else
            ReDim retval(e - s)
            Dim i, j
            j = 0
            For i = s to e
                Assign retval(j), m_data(i)
                j = j + 1
            Next
            Set Slice = DynamicArray1(retval)
        End If
    End Function
    
    
    Public Function Iterator()
        Dim retval
        Set retval = New Iterator_DynamicArray_Class
        retval.Initialize Me
        Set Iterator = retval
    End Function
    
    Public Function TO_En()
        Set TO_En = Scanner_Iterator(Me.Iterator)
    End Function
    
    Public Function TO_Array()
        Dim i
        ReDim retval(m_size - 1)
        For i = 0 to UBOUND(retval)
            Assign retval(i), m_data(i)
        Next
        TO_Array = retval
    End Function

End Class


Public Function DynamicArray()
    ReDim data(3)
    Set DynamicArray = DynamicArray2(data, 0)
End Function

Public Function DynamicArray1(ByVal data)
    Set DynamicArray1 = DynamicArray2(data, UBOUND(data) + 1)
End Function

Private Function DynamicArray2(ByVal data, ByVal size)
    Dim retval
    Set retval = New DynamicArray_Class
    retval.Initialize data, size
    Set DynamicArray2 = retval
End Function







Class Iterator_DynamicArray_Class
    Private m_dynamic_array
    Private m_index
    
    Public Sub Initialize(ByVal dynamic_array)
        Set m_dynamic_array = dynamic_array
        m_index = 0
    End Sub
    
    Private Sub Class_Terminate
        Set m_dynamic_array = Nothing
    End Sub
    
    Public Function HasNext()
        HasNext = (m_index < m_dynamic_array.Size)
    End Function
    
    Public Function PeekNext()
        Assign PeekNext, m_dynamic_array.FastItem(m_index)
    End Function
    
    Public Function GetNext()
        Assign GetNext, m_dynamic_array.FastItem(m_index)
        m_index = m_index + 1
    End Function
    
    
    
    Public Function HasPrev()
        HasPrev = (m_index > 0)
    End Function
    
    Public Function PeekPrev()
        Assign PeekPrev, m_dynamic_array.FastItem(m_index - 1)
    End Function
    
    Public Function GetPrev()
        Assign GetPrev, m_dynamic_array.FastItem(m_index - 1)
        m_index = m_index - 1
    End Function
End Class




Converter_Enumerator.Given(L_Type("DynamicArray_Class")).Dispatch("arg0.TO_En")

