Option Explicit
'
' enumerator.vbs
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



' In many ways, Enumerators are really just Iterators.
' Unlike iterators, however, Enumerators provide a number
' of methods that facilitate "lazy" data processing.
' This means that Enumerators can help you define complex
' operations at one point in your code yet evaluate
' these operations at a later time.

Require "t\core\assignment.vbs"
Require "t\core\closures.vbs"

Require "t\core\linked_list.vbs"



Class Enumerator_Class

    Private m_source

    Private m_ready
    Private m_completed
    
    Private m_current

    Public Sub Initialize(ByVal source)
        Set m_source = source
        m_ready      = False
        m_completed  = False
    End Sub
    
    Public Function HasNext()
        If m_completed Then
            HasNext = False
        ElseIf m_ready Then
            HasNext = True
        Else
            m_source.GetNext m_current, m_ready
            If Not m_ready Then
                m_completed = True
            End If
            HasNext = m_ready
        End If
    End Function
    
    Public Function IsEmpty()
        IsEmpty = Not Me.HasNext
    End Function
    
    
    Public Function Peek()
        If Me.HasNext Then
            Assign Peek, m_current
        Else
            ' TODO: Errors are no longer handled in this manner
            ERR.RAISE ERROR_Enumerator_EMPTY
        End If
    End Function

    
    Public Function Pop()
        If Me.HasNext Then
            Assign Pop, m_current
            m_ready = False
        Else
            ' TODO: Errors are no longer handled in this manner
            ERR.RAISE ERROR_Enumerator_EMPTY
        End If
    End Function

    Public Function Retain(ByVal cond)
        Dim retval
        Set retval = New Enumerator_Source_Retain_Class
        retval.Initialize cond, Me
        Set Retain = Enumerator(retval)
    End Function
    
    Public Function Discard(ByVal cond)
        Dim retval
        Set retval = New Enumerator_Source_Discard_Class
        retval.Initialize cond, Me
        Set Discard = Enumerator(retval)
    End Function
    
    
    ' Returns the first element that makes the
    ' condition True as well as all subsequent
    ' elements.
    Public Function StartWhen(ByVal cond)
        Dim retval
        Set retval = New Enumerator_Source_StartWhen_Class
        retval.Initialize cond, Me
        Set StartWhen = Enumerator(retval)
    End Function
    
    
    ' Returns all elements after the first element
    ' that makes the condition True.
    Public Function StartAfter(ByVal cond)
        Dim retval
        Set retval = New Enumerator_Source_StartAfter_Class
        retval.Initialize cond, Me
        Set StartAfter = Enumerator(retval)
    End Function
    
    
    ' Stops as soon as the condition becomes True.
    ' Does not return the elements that made the
    ' condition True.
    Public Function DoUntil(ByVal cond)
        Set DoUntil = StopBefore(cond)
    End Function
    
    ' This is the same as DoUntil
    Public Function StopBefore(ByVal cond)
        Dim retval
        Set retval = New Enumerator_Source_Stop_Class
        retval.Initialize cond, True, Me
        Set StopBefore = Enumerator(retval)
    End Function
    
    
    ' Stops when the condition becomes True.  This
    ' will also include the element that made the
    ' condition True.
    Public Function StopWhen(ByVal cond)
        Dim retval
        Set retval = New Enumerator_Source_Stop_Class
        retval.Initialize cond, False, Me
        Set StopWhen = Enumerator(retval)
    End Function
    
    
    ' Returns elements as long as the condition remains True
    Public Function DoWhile(ByVal cond)
        Dim retval
        Set retval = New Enumerator_Source_DoWhile_Class
        retval.Initialize cond, Me
        Set DoWhile = Enumerator(retval)
    End Function
    
    Public Function Map(ByVal func)
        Dim retval
        Set retval = New Enumerator_Source_Map_Class
        retval.Initialize func, Me
        Set Map = Enumerator(retval)
    End Function
    
    Public Function WithIndex()
        Dim retval
        Set retval = New Enumerator_Source_WithIndex_Class
        retval.Initialize Me
        Set WithIndex = Enumerator(retval)
    End Function
    
    Public Sub DoEach(ByVal func)
        Set func = TO_Func(1, func)
        While Me.HasNext
            func Me.Pop
        Wend
    End Sub
    
    Public Function Fold2(ByVal func, ByVal memory)
        Set func = TO_Expr(2, func)
        While Me.HasNext
            Assign memory, func(memory, Me.Pop)
        Wend
        Assign Fold2, memory
    End Function
    
    Public Function Fold(ByVal func)
        Assign Fold, Me.Fold2(func, Me.Pop)
    End Function
    
    
    Public Function Cons()
        Set Cons = eCons(Me)
    End Function
    
    
    Public Function Limit(ByVal number)
        Dim retval
        Set retval = New Enumerator_Source_Limit_Class
        retval.Initialize Me, number
        Set Limit = Enumerator(retval)
    End Function
    
    
    ' Slightly faster than calling the normal
    ' TO_LinkedList dispatcher function
    Public Function TO_LinkedList()
        Dim retval
        Set retval = LinkedList
        While Me.HasNext
            retval.Push Me.Pop
        Wend
        Set TO_LinkedList = retval
    End Function
    
    ' This should be slightly faster than calling
    ' the normal TO_Array dispatcher function.
    Public Function TO_Array()
        TO_Array = Me.TO_LinkedList.TO_Array
    End Function

End Class

Public Function Enumerator(ByVal source)
    Dim retval
    Set retval = New Enumerator_Class
    retval.Initialize source
    Set Enumerator = retval
End Function





Class Enumerator_Source_Map_Class
    Private m_map
    Private m_en
    
    Public Sub Initialize(ByVal mapr, ByVal enumr)
        Set m_map = TO_Expr(1, mapr)
        Set m_en  = enumr
    End Sub
    
    Private Sub Class_Terminate()
        Set m_map     = Nothing
        Set m_en = Nothing
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        If m_en.HasNext Then
            Assign retval, m_map(m_en.Pop)
            successful = True
        Else
            successful = False
        End If
    End Sub
End Class








Class Enumerator_Source_Retain_Class
    Private m_cond
    Private m_en
    
    Public Sub Initialize(ByVal cond, ByVal enumr)
        Set m_cond  = TO_Expr(1, cond)
        Set m_en    = enumr
    End Sub
    
    Private Sub Class_Terminate()
        Set m_cond    = Nothing
        Set m_en = Nothing
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        While m_en.HasNext
            If m_cond(m_en.Peek) Then
                Assign retval, m_en.Pop
                successful = True
                Exit Sub
            Else
                m_en.Pop
            End If
        Wend
        successful = False
    End Sub
End Class




Class Enumerator_Source_Discard_Class
    Private m_cond
    Private m_en
    
    Public Sub Initialize(ByVal cond, ByVal enumr)
        Set m_cond  = TO_Expr(1, cond)
        Set m_en    = enumr
    End Sub
    
    Private Sub Class_Terminate()
        Set m_cond    = Nothing
        Set m_en = Nothing
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        Do While m_en.HasNext
            If Not m_cond(m_en.Peek) Then
                Assign retval, m_en.Pop
                successful = True
                Exit Sub
            Else
                m_en.Pop
            End If
        Loop
        successful = False
    End Sub
End Class







Class Enumerator_Source_StartWhen_Class
    Private m_cond
    Private m_started
    Private m_en
    
    Public Sub Initialize(ByVal cond, ByVal enumr)
        Set m_cond    = TO_Expr(1, cond)
        Set m_en = enumr
        
        m_started     = False
    End Sub
    
    Private Sub Class_Terminate()
        Set m_cond    = Nothing
        Set m_en = Nothing
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        If m_started Then
            If m_en.HasNext Then
                Assign retval, m_en.Pop
                successful = True
            Else
                successful = False
            End If
        Else
            Do While m_en.HasNext
                Assign retval, m_en.Pop
                If m_cond(retval) Then
                    successful = True
                    m_started  = True
                    Exit Sub
                End If
            Loop
            successful = False
        End If
    End Sub
End Class





Class Enumerator_Source_StartAfter_Class
    Private m_cond
    Private m_started
    Private m_en
    
    Public Sub Initialize(ByVal cond, ByVal enumr)
        Set m_cond    = TO_Expr(1, cond)
        Set m_en = enumr
        
        m_started     = False
    End Sub
    
    Private Sub Class_Terminate()
        Set m_cond    = Nothing
        Set m_en = Nothing
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        If m_started Then
            If m_en.HasNext Then
                Assign retval, m_en.Pop
                successful = True
            Else
                successful = False
            End If
        Else
            While m_en.HasNext
                If m_cond(m_en.Pop) Then
                    m_started  = True
                    ' Slightly faster than calling
                    ' myself recursively
                    If m_en.HasNext Then
                        Assign retval, m_en.Pop
                        successful = True
                    Else
                        successful = False
                    End If
                    Exit Sub
                End If
            Wend
            successful = False
        End If
    End Sub
End Class





Class Enumerator_Source_Stop_Class
    Private m_cond
    Private m_stopped
    Private m_en
    Private m_stop_immediately
    
    Public Sub Initialize(ByVal cond, ByVal stop_immediately, ByVal enumr)
        Set m_cond         = TO_Expr(1, cond)
        Set m_en      = enumr
        m_stop_immediately = stop_immediately
        m_stopped          = False
    End Sub
    
    Private Sub Class_Terminate()
        Set m_cond    = Nothing
        Set m_en = Nothing
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        If m_stopped Then
            successful = False
        Else
            If m_en.HasNext Then
                Assign retval, m_en.Pop
                If m_cond(retval) Then
                    m_stopped  = True
                    successful = (Not m_stop_immediately)
                Else
                    successful = True
                End If
            Else
                m_stopped  = True
                successful = False
            End If
        End If
    End Sub
End Class



Class Enumerator_Source_DoWhile_Class
    Private m_cond
    Private m_stopped
    Private m_en
    
    Public Sub Initialize(ByVal cond, ByVal enumr)
        Set m_cond    = TO_Expr(1, cond)
        Set m_en      = enumr
        m_stopped     = False
    End Sub
    
    Private Sub Class_Terminate()
        Set m_cond    = Nothing
        Set m_en = Nothing
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        If m_stopped Then
            successful = False
        Else
            If m_en.HasNext Then
                Assign retval, m_en.Pop
                If m_cond(retval) Then
                    successful = True
                Else
                    m_stopped  = True
                    successful = False
                End If
            Else
                m_stopped  = True
                successful = False
            End If
        End If
    End Sub
End Class





Class Enumerator_Source_Zip_Class
    Private m_ens
    Private m_empty
    
    Public Sub Initialize(ByVal ens)
        m_ens   = ens
        m_empty = False
    End Sub
    
    Private Sub Class_Terminate()
        Set m_ens = Nothing
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        Dim i
        ReDim temp(UBOUND(m_ens))
        For i = 0 to UBOUND(temp)
            If m_ens(i).HasNext Then
                Assign temp(i), m_ens(i).Pop
            Else
                successful = False
                Exit Sub
            End If
        Next
        successful = True
        retval     = temp
    End Sub
End Class


Public Function eZip(ByVal ens)
    Dim retval
    Set retval = New Enumerator_Source_Zip_Class
    retval.Initialize ens
    Set eZip = Enumerator(retval)
End Function

Public Function eZip2(ByVal arg0, ByVal arg1)
    Set eZip2 = eZip(Array(arg0, arg1))
End Function

Public Function eZip3(ByVal arg0, ByVal arg1, arg2)
    Set eZip3 = eZip(Array(arg0, arg1, arg2))
End Function




' Creates an Enumerator that contains exactly 1 element.
Class Enumerator_Source_Trivial_Class
    Private m_value
    Private m_empty
    
    Public Sub Initialize(ByVal value)
        Assign m_value, value
        m_empty = False
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        If m_empty Then
            successful = False
        Else
            Assign retval, m_value
            Set m_value = Nothing
            successful  = True
            m_empty     = True
        End If
    End Sub
End Class


Public Function sTrivial(ByVal value)
    Dim retval
    Set retval   = New Enumerator_Source_Trivial_Class
    retval.Initialize value
    Set sTrivial = Enumerator(retval)
End Function





' Default EnumeratorSource implementation.  Never
' yields any item.
Class Enumerator_Source_Empty_Class
    Public Sub GetNext(ByRef retval, ByRef successful)
        successful = False
    End Sub
End Class



Public Function sEmpty()
    Dim retval
    Set retval = New Enumerator_Source_Empty_Class
    Set sEmpty = Enumerator(retval)
End Function









Class Enumerator_Source_Collection_Class
	Private m_collection
	Private m_index

	Public Sub Initialize(ByVal c)
		Set m_collection = c
		m_index = 0
	End Sub
	
	Private Sub Class_Terminate()
		Set m_collection = Nothing
	End Sub
	
	Public Sub GetNext(ByRef retval, ByRef successful)
		If m_index >= m_collection.Count Then
			successful = False
		Else
			successful = True
			Assign retval, m_collection(m_index)
			m_index = m_index + 1
		End If
	End Sub
	
End Class

Public Function Enumerator_Collection(ByVal c)
	Dim retval
	Set retval = New Enumerator_Source_Collection_Class
	retval.Initialize c
	Set Enumerator_Collection = Enumerator(retval)
End Function





Class Enumerator_Source_Counter_Class
    Private m_value
    Private m_stepsize
    
    Public Sub Initialize(ByVal value, ByVal stepsize)
        m_value    = value
        m_stepsize = stepsize
    End Sub

    Public Sub GetNext(ByRef retval, ByRef successful)
        retval     = m_value
        m_value    = m_value + m_stepsize
        successful = True
    End Sub
End Class

Public Function eCounter()
    Set eCounter = eCounter2(0, 1)
End Function

Public Function eCounter1(ByVal start)
    Set eCounter1 = eCounter2(start, 1)
End Function

Public Function eCounter2(ByVal start, ByVal stepsize)
    Dim retval
    Set retval = New Enumerator_Source_Counter_Class
    retval.Initialize start, stepsize
    Set eCounter2 = Enumerator(retval)
End Function





Public Function eRange3(ByVal start, ByVal finish, ByVal stepsize)
    If stepsize >= 0 Then
        Set eRange3 = eCounter2(start, stepsize).DoWhile(Lambda(1, "Invoke = (arg0 <= stored)").Store(finish))
    Else
        Set eRange3 = eCounter2(start, stepsize).DoWhile(Lambda(1, "Invoke = (arg0 >= stored)").Store(finish))
    End If
End Function



Public Function eRange(ByVal start, ByVal finish)
    Set eRange = eRange3(start, finish, 1)
End Function









Class Enumerator_Struct_WithIndex_Class
    Private m_index
    Private m_item
    
    Public Sub Initialize(ByVal index, ByVal item)
        m_index = index
        Assign m_item, item
    End Sub
    
    Private Sub Class_Terminate()
        Set m_item = Nothing
    End Sub
    
    Public Property Get Index
        Index = m_index
    End Property
    
    Public Property Get Item
        Assign Item, m_item
    End Property
End Class


Class Enumerator_Source_WithIndex_Class
    Private m_enumr
    Private m_counter
    
    Public Sub Initialize(ByVal enumr)
        Set m_enumr   = enumr
        Set m_counter = eCounter
    End Sub
    
    Private Sub Class_Terminate()
        Set m_enumr   = Nothing
        Set m_counter = Nothing
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        If m_enumr.HasNext Then
            Set retval = New Enumerator_Struct_WithIndex_Class
            retval.Initialize m_counter.Pop, m_enumr.Pop
            successful = True
        Else
            successful = False
        End If
    End Sub
End Class


Class Enumerator_Source_Limit_Class
    Private m_en
    Private m_limit
    Private m_i
    
    Public Sub Initialize(ByVal enumr, ByVal lim)
        Set m_en  = enumr
        m_limit   = lim
        m_i       = 0
    End Sub
    
    Private Sub Class_Terminate()
        Set m_en = Nothing
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        If m_i < m_limit Then
            If m_en.HasNext Then
                Assign retval, m_en.Pop
                successful = True
                m_i        = m_i + 1
            Else
                successful = False
            End If
        Else
            successful = False
        End If
    End Sub
    
End Class




Class Enumerator_Source_Cons_Class
    Private m_enums
    
    Public Sub Initialize(ByVal enums)
        ' TODO: Eventually, this should accept any
        '       Enumerable datatype.
        Set m_enums = enums
    End Sub
    
    
    Private Sub Class_Terminate()
        Set m_enums = Nothing
    End Sub
    
    Public Sub GetNext(ByRef retval, ByRef successful)
        Do While m_enums.HasNext
            If m_enums.Peek.HasNext Then
                Assign retval, m_enums.Peek.Pop
                successful = True
                Exit Sub
            End If
            m_enums.Pop
        Loop
        successful = False
    End Sub

End Class


' Concatenates the output of multiple enumerators
' into one giant enumerator.
Public Function eCons(ByVal enumerators)
    Dim retval : Set retval = New Enumerator_Source_Cons_Class
    retval.Initialize enumerators
    Set eCons = Enumerator(retval)
End Function


