Option Explicit

Require "t\core\linked_list.vbs"
Require "t\core\memoizer.vbs"





Private Function CHECK_ArgList_Code(ByVal m, ByVal arg_count)
    If arg_count = 0 Then
        CHECK_ArgList_Code = "Invoke = True"
    Else
        Dim n      : n      = arg_count - 1
        Dim arg    : arg    = ClosureFactory.ArgName(n)
        Dim stored : stored = ClosureFactory.StoredName(n)
        
        CHECK_ArgList_Code = _
            "If " & stored & "(" & arg & ") Then" & vbCR &_
            m(n) & vbCR &_
            "Else" & vbCR &_
            "Invoke = False" & vbCR &_
            "End If"
    End If
End Function




Class CHECK_ArgListFactory_Class

    Private m_CHECK_ArgList_code
    
    Private Sub Class_Initialize()
        Dim temp : Set temp = Lambda(1, "Invoke = stored(0)(stored(1), arg0)")
        Set m_CHECK_ArgList_code = Memoizer(1, temp)
        temp.Stored = Array(GetRef("CHECK_ArgList_Code"), m_CHECK_ArgList_code)
    End Sub
    
    Private Sub Class_Terminate()
        Set m_CHECK_ArgList_code = Nothing
    End Sub
    
    Public Function Create(ByVal arg_count)
        Set Create = Lambda(arg_count, m_CHECK_ArgList_code(arg_count))
    End Function

End Class



Dim CHECK_ArgListFactory_Singleton__
Public Function CHECK_ArgList(ByVal conditions)
    If IsEmpty(CHECK_ArgListFactory_Singleton__) Then
        Set CHECK_ArgListFactory_Singleton__ = New CHECK_ArgListFactory_Class
    End If
    
    
    Dim i
    For i = 0 TO UBOUND(conditions)
        Set conditions(i) = TO_Expr(1, conditions(i))
    Next
    
    Dim retval : Set retval = CHECK_ArgListFactory_Singleton__.Create(UBOUND(conditions) + 1)
    retval.Stored = conditions
    Set CHECK_ArgList = retval
End Function







Class Dispatcher_Rule_Class2
    Public Condition
    Public Action
    
    Public Sub Initialize(ByVal c)
        Set Condition = TO_Expr(2, c)
        Set Action    = Nothing
    End Sub
    
    Private Sub Class_Terminate()
        Set Condition = Nothing
        Set Action    = Nothing
    End Sub
    
    Public Sub Dispatch(ByVal func)
        Set Action = TO_Expr(2, func)
    End Sub
End Class




Class Dispatcher_Class2
    Private m_rules

    Private Sub Class_Initialize()
        Set m_rules     = LinkedList()
    End Sub
    
    Private Sub Class_Terminate()
        Set m_rules = Nothing
    End Sub

    Public Function Given(ByVal arg0, ByVal arg1)
        Set Given = When(CHECK_ArgList(Array(arg0, arg1)))
    End Function
    
    Public Function When(ByVal cond)
        Dim retval : Set retval = New Dispatcher_Rule_Class2
        retval.Initialize(TO_Expr(2, cond))
        m_rules.Unshift retval
        Set When = retval
    End Function

    Public Default Function Invoke(ByVal arg0, ByVal arg1)
        Dim i : Set i = m_rules.Iterator
        Dim r
        Do While i.HasNext
            Set r = i.GetNext
            If r.Condition(arg0, arg1) Then
                Assign Invoke, r.Action(arg0, arg1)
                Exit Function
            End If
        Loop
        ' TODO
    End Function
End Class








Public Function DispatcherFactory_DispatcherCode(ByVal arg_count)

    DispatcherFactory_DispatcherCode = _
        "Class Dispatcher_Class" & arg_count & vbCR &_
        "    Private m_rules" & vbCR &_
        "    Private Sub Class_Initialize()" & vbCR &_
        "        Set m_rules     = LinkedList()" & vbCR &_
        "    End Sub" & vbCR &_
        "    Private Sub Class_Terminate()" & vbCR &_
        "        Set m_rules = Nothing" & vbCR &_
        "    End Sub" & vbCR &_
        "    Public Function Given(" & ClosureFactory.ByValArgs(arg_count) & ")" & vbCR &_
        "        Set Given = When(CHECK_ArgList(Array(" & ClosureFactory.InvokeArgs(arg_count) & ")))" & vbCR &_
        "    End Function" & vbCR &_
        "    Public Function When(ByVal cond)" & vbCR &_
        "        Dim retval : Set retval = New Dispatcher_Rule_Class" & arg_count & vbCR &_
        "        retval.Initialize(cond)" & vbCR &_
        "        m_rules.Unshift retval" & vbCR &_
        "        Set When = retval" & vbCR &_
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
        "    End Function" & vbCR &_
        "End Class"


End Function


Class DispatcherFactory_Class
End Class



Dim DispatcherFactory_Singleton__
Public Function DispatcherFactory()
    If IsEmpty(DispatcherFactory_Singleton__) Then
        Set DispatcherFactory_Singleton__ = New DispatcherFactory_Class
    End If
    Set DispatcherFactory = DispatcherFactory_Singleton__
End Function



Public Function Dispatcher(ByVal arg_count)
    Set Dispatcher = New Dispatcher_Class2
End Function

