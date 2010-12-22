
Require "t\core\closures.vbs"
Require "t\core\memoizer.vbs"





Class T_MethodLookup_Class
    Private m_self
    Private m_class
    Private m_name
    
    Public Sub Initialize(ByVal self, ByVal clazz, ByVal name)
        Set m_self  = self
        Set m_class = clazz
        m_name      = name
    End Sub
    
    Private Sub Class_Terminate()
        Set m_self  = Nothing
        Set m_class = Nothing
    End Sub

End Class


Class T_UnboundMethod_Retriever_Fixed_Class
    Private m_unbound_method
    
    Public Sub Initialize(ByVal unbound_method)
        Set m_unbound_method = unbound_method
    End Sub
    
    Private Sub Class_Terminate()
        Set m_unbound_method = Nothing
    End Sub
    
    Public Default Function Invoke(ByVal self, ByVal name)
        Set Invoke = m_unbound_method
    End Function
End Class

Private Function T_UnboundMethod_Retriever_Fixed(ByVal unbound_method)
    Dim retval : Set retval = New T_UnboundMethod_Retriever_Fixed_Class
    retval.Initialize unbound_method
    Set T_UnboundMethod_Retriever_Fixed = retval
End Function










Private Function T_Method_ClassFactory__(ByVal arg_count)
    Dim invoke_args
    If arg_count = 0 Then
        invoke_args = ""
    Else
        invoke_args = ", " & ClosureFactory.InvokeArgs(arg_count)
    End If

    
    ExecuteGlobal _
        "Class T_UnboundMethod_Class" & arg_count & vbCR &_
        "    Private m_name" & vbCR &_
        "    Private m_func" & vbCR &_
        "    Public Sub Initialize(ByVal name, ByVal func)" & vbCR &_
        "        m_name = name" & vbCR &_
        "        Set m_func = TO_Func(" & arg_count + 1 & ", func)" & vbCR &_
        "    End Sub" & vbCR &_
        "    Private Sub Class_Terminate()" & vbCR &_
        "        Set m_func = Nothing" & vbCR &_
        "    End Sub" & vbCR &_
        "    Public Property Get Name" & vbCR &_
        "        Name = m_name" & vbCR &_
        "    End Property" & vbCR &_
        "    Public Property Get ArgCount" & vbCR &_
        "        ArgCount = " & arg_count & vbCR &_
        "    End Property" & vbCR &_
        "    Public Property Get Func" & vbCR &_
        "        Set Func = m_func" & vbCR &_
        "    End Property" & vbCR &_
        "    Public Function Bind(ByVal self)" & vbCR &_
        "        Dim retval : Set retval = New T_Method_Class" & arg_count & vbCR &_
        "        retval.Initialize self, m_name, m_func" & vbCR &_
        "        Set Bind = retval" & vbCR &_
        "    End Function" & vbCR &_
        "End Class" & vbCR &_
        vbCR &_
        "Class T_Method_Class" & arg_count & vbCR &_
        "    Private m_self" & vbCR &_
        "    Private m_name" & vbCR &_
        "    Private m_func" & vbCR &_
        "    Public Sub Initialize(ByVal self, ByVal name, ByVal func)" & vbCR &_
        "        Set m_self = self" & vbCR &_
        "            m_name = name" & vbCR &_
        "        Set m_func = func" & vbCR &_
        "    End Sub" & vbCR &_
        "    Private Sub Class_Terminate()" & vbCR &_
        "        Set m_self = Nothing" & vbCR &_
        "        Set m_func = Nothing" & vbCR &_
        "    End Sub" & vbCR &_
        "    Public Property Get Name" & vbCR &_
        "        Name = m_name" & vbCR &_
        "    End Property" & vbCR &_
        "    Public Property Get ArgCount" & vbCR &_
        "        ArgCount = " & arg_count & vbCR &_
        "    End Property" & vbCR &_
        "    Public Property Get Func" & vbCR &_
        "        Set Func = m_func" & vbCR &_
        "    End Property" & vbCR &_
        "    Public Function Unbind()" & vbCR &_
        "        Dim retval : Set retval = New T_UnboundMethod_Class" & arg_count & vbCR &_
        "        retval.Initialize m_name, m_func" & vbCR &_
        "        Set Unbind = retval" & vbCR &_
        "    End Function" & vbCR &_
        "    Public Function X(" & ClosureFactory.ByRefArgs(arg_count) & ")" & vbCR &_
        "        Assign X, m_func(m_self" & invoke_args & ")" & vbCR &_
        "    End Function" & vbCR &_
        "End Class"

    Set T_Method_ClassFactory__ = Lambda(0, "Set Invoke = New T_UnboundMethod_Class" & arg_count)
End Function


Dim T_UnboundMethod_Factory_Singleton__
Public Function T_UnboundMethod(ByVal name, ByVal arg_count, ByVal func)
    If IsEmpty(T_UnboundMethod_Factory_Singleton__) Then
        Set T_UnboundMethod_Factory_Singleton__ = Memoizer(1, GetRef("T_Method_ClassFactory__"))
    End If

    Dim c      : Set c = T_UnboundMethod_Factory_Singleton__(arg_count)
    Dim retval : Set retval = c()
    retval.Initialize name, func
    
    Set T_UnboundMethod = retval
End Function

