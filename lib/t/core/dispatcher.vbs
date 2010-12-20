Option Explicit
'
' dispatcher.vbs
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

Require "t\core\linked_list.vbs"
Require "t\core\memoizer.vbs"


Private Function L_ArgListMatcher_Code(ByVal m, ByVal arg_count)
    If arg_count = 0 Then
        L_ArgListMatcher_Code = "Invoke = True"
    Else
        Dim n      : n      = arg_count - 1
        Dim arg    : arg    = ClosureFactory.ArgName(n)
        Dim stored : stored = ClosureFactory.StoredName(n)

        Dim code : code = _
            "If " & stored & "(" & arg & ") Then" & vbCR &_
            m(n) & vbCR &_
            "Else" & vbCR &_
            "Invoke = False" & vbCR &_
            "End If"
        L_ArgListMatcher_Code = code
    End If
End Function



Dim L_ArgListMatcherFactory_Singleton__
Public Function L_ArgListMatcher(ByVal conditions)
    If IsEmpty(L_ArgListMatcherFactory_Singleton__) Then
        Dim temp : Set temp = Lambda(1, "Assign Invoke, stored(0)(stored(1), arg0)")
        Set L_ArgListMatcherFactory_Singleton__ = Memoizer(1, temp)
        temp.Stored = Array(GetRef("L_ArgListMatcher_Code"), L_ArgListMatcherFactory_Singleton__)
    End If
    
    Dim i
    For i = 0 TO UBOUND(conditions)
        Set conditions(i) = TO_Expr(1, conditions(i))
    Next
    
    Dim size   : size = UBOUND(conditions) + 1
    Dim retval : Set retval = Lambda(size, L_ArgListMatcherFactory_Singleton__(size))
    retval.Stored = conditions
    Set L_ArgListMatcher = retval
End Function









Public Function DispatcherFactory_CreateDispatcherConstructor(ByVal arg_count)
    ExecuteGlobal _
        "Class Dispatcher_Rule_Class" & arg_count & vbCR &_
        "    Public Condition" & vbCR &_
        "    Public Action" & vbCR &_
        "    Public Sub Initialize(ByVal c)" & vbCR &_
        "        Set Condition = TO_Expr(" & arg_count & ", c)" & vbCR &_
        "        Set Action    = Nothing" & vbCR &_
        "    End Sub" & vbCR &_
        "    Private Sub Class_Terminate()" & vbCR &_
        "        Set Condition = Nothing" & vbCR &_
        "        Set Action    = Nothing" & vbCR &_
        "    End Sub" & vbCR &_
        "    Public Sub Dispatch(ByVal func)" & vbCR &_
        "        Set Action = TO_Expr(" & arg_count & ", func)" & vbCR &_
        "    End Sub" & vbCR &_
        "End Class" & vbCR &_
        vbCR &_
        "Class Dispatcher_Class" & arg_count & vbCR &_
        "    Private m_rules" & vbCR &_
        "    Private m_otherwise" & vbCR &_
        "    Private Sub Class_Initialize()" & vbCR &_
        "        Set m_rules     = LinkedList()" & vbCR &_
        "    End Sub" & vbCR &_
        "    Private Sub Class_Terminate()" & vbCR &_
        "        Set m_rules = Nothing" & vbCR &_
        "    End Sub" & vbCR &_
        "    Public Function Given(" & ClosureFactory.ByValArgs(arg_count) & ")" & vbCR &_
        "        Set Given = When(L_ArgListMatcher(Array(" & ClosureFactory.InvokeArgs(arg_count) & ")))" & vbCR &_
        "    End Function" & vbCR &_
        "    Public Function When(ByVal cond)" & vbCR &_
        "        Dim retval : Set retval = New Dispatcher_Rule_Class" & arg_count & vbCR &_
        "        retval.Initialize(cond)" & vbCR &_
        "        m_rules.Push retval" & vbCR &_
        "        Set When = retval" & vbCR &_
        "    End Function" & vbCR &_
        "    Public Function Otherwise()" & vbCR &_
        "        Set m_otherwise = New Dispatcher_Rule_Class" & arg_count & vbCR &_
        "        Set Otherwise = m_otherwise" & vbCR &_
        "    End Function" & vbCR &_
        "    Public Default Function Invoke(" & ClosureFactory.ByRefArgs(arg_count) & ")" & vbCR &_
        "        Dim i : Set i = m_rules.Iterator" & vbCR &_
        "        Dim r" & vbCR &_
        "        Do While i.HasNext" & vbCR &_
        "            Set r = i.GetNext" & vbCR &_
        "            If r.Condition(" & ClosureFactory.InvokeArgs(arg_count) & ") Then" & vbCR &_
        "                Assign Invoke, r.Action(" & ClosureFactory.InvokeArgs(arg_count) & ")" & vbCR &_
        "                Exit Function" & vbCR &_
        "            End If" & vbCR &_
        "        Loop" & vbCR &_
        "        Assign Invoke, m_otherwise.Action(" & ClosureFactory.InvokeArgs(arg_count) & ")" & vbCR &_
        "    End Function" & vbCR &_
        "End Class"

    Set DispatcherFactory_CreateDispatcherConstructor = Lambda(0, "Set Invoke = New Dispatcher_Class" & arg_count)

End Function


Class DispatcherFactory_Class

    Private m_dispatcher_constructor

    Private Sub Class_Initialize()
        Set m_dispatcher_constructor = Memoizer(1, GetRef("DispatcherFactory_CreateDispatcherConstructor"))
    End Sub
    
    Private Sub Class_Terminate()
        Set m_dispatcher_constructor = Nothing
    End Sub

    Public Function Create(ByVal arg_count)
        Dim c : Set c = m_dispatcher_constructor(arg_count)
        Set Create = c()
    End Function

End Class



Dim DispatcherFactory_Singleton__
Public Function DispatcherFactory()
    If IsEmpty(DispatcherFactory_Singleton__) Then
        Set DispatcherFactory_Singleton__ = New DispatcherFactory_Class
    End If
    Set DispatcherFactory = DispatcherFactory_Singleton__
End Function



Public Function Dispatcher(ByVal arg_count)
    Set Dispatcher = DispatcherFactory.Create(arg_count)
End Function

