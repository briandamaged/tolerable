Option Explicit

Require "t\oop\methods.vbs"



Class T_Object_Class

    Private m_fields
    Private m_unbound_methods
    
    Private Sub Class_Initialize()
        Set m_fields          = CreateObject("Scripting.Dictionary")
        Set m_unbound_methods = CreateObject("Scripting.Dictionary")
    End Sub
    
    Private Sub Class_Terminate()
        Set m_fields          = Nothing
        Set m_unbound_methods = Nothing
    End Sub

    
    ' I can't figure out any good way to make this field
    ' private.  So, I'll make the name annoying enough so
    ' that no one will ever want to access it directly
    Public Property Get Fields__()
        Set Fields__ = m_fields
    End Property
    
    Public Function GetF(ByVal name)
        Assign GetF, Fields__.Item(name)
    End Function
    
    Public Function SetF(ByVal name, ByVal value)
        If Fields__.Exists(name) Then
            Fields.Remove(name)
        End If
        Fields.Add name, value
        Assign SetF, value
    End Function
    
    Public Property Get UnboundMethods__()
        Set UnboundMethods__ = m_unbound_methods
    End Property

    Public Function M(ByVal name)
        If m_unbound_methods.Exists(name) Then
            Set M = m_unbound_methods.Item(name).Bind(Me)
        Else
            Dim clazz : Set clazz = GetF("__class")
            Dim i
            Do Until clazz Is Nothing
                Set i = clazz.Fields__.Item("__unbound_instance_methods")
                If i.Exists(name) Then
                    Set M = i.Item(name).Bind(Me)
                    Exit Function
                End If
                Set clazz = clazz.Fields__.Item("__superclass")
            Loop
        End If
    End Function

End Class






Dim T_Object_Singleton__
Public Function T_Object()
    If IsEmpty(T_Object_Singleton__) Then
        Set T_Object_Singleton__ = New T_Object_Class
        Dim f : Set f = T_Object_Singleton__.Fields__
        
        f.Add "__class",      T_Class
        f.Add "__superclass", Nothing
        f.Add "__unbound_instance_methods", CreateObject("Scripting.Dictionary")
        
        Dim s : Set s = T_Object_Singleton__
        ' s.M("def").X "sdef", 3, GetRef("T_Object_M_SDef")
    End If
    Set T_Object = T_Object_Singleton__
End Function

Private Function T_Object_M_SDef(ByVal self, ByVal name, ByVal arg_count, ByVal func)
    self.UnboundMethods__.Add name, T_UnboundMethod(name, arg_count, func)
End Function


Dim T_Module_Singleton__
Public Function T_Module()
    If IsEmpty(T_Module_Singleton__) Then
        Set T_Module_Singleton__ = New T_Object_Class
        Dim f : Set f = T_Module_Singleton__.Fields__
        Dim i : Set i = CreateObject("Scripting.Dictionary")
        
        f.Add "__unbound_instance_methods", i
        i.Add "def", T_UnboundMethod("def", 3, GetRef("T_Module_M_Def"))


        f.Add "__class",      T_Class
        f.Add "__superclass", T_Object        
    End If
    Set T_Module = T_Module_Singleton__
End Function

Private Function T_Module_M_Def(ByVal self, ByVal name, ByVal arg_count, ByVal func)
    self.GetF("__unbound_instance_methods").Add name, T_UnboundMethod(name, arg_count, func)
End Function



Dim T_Class_Singleton__
Public Function T_Class()
    If IsEmpty(T_Class_Singleton__) Then
        Set T_Class_Singleton__ = New T_Object_Class
        Dim f : Set f = T_Class_Singleton__.Fields__

        f.Add "__class",      T_Class_Singleton__
        f.Add "__superclass", T_Module
        f.Add "__unbound_instance_methods", CreateObject("Scripting.Dictionary")
        
        Dim s : Set s = T_Class_Singleton__
        s.M("def").X "new",    0, GetRef("T_Class_M_New")
        s.M("def").X "extend", 0, GetRef("T_Class_M_Extend")
    End If
    Set T_Class = T_Class_Singleton__
End Function


Private Function T_Class_M_New(ByVal self)
    Dim retval : Set retval = New T_Object_Class
    retval.Fields__.Add "__class", self

    Set T_Class_M_New = retval
End Function

Private Function T_Class_M_Extend(ByVal self)
    Dim retval : Set retval = New T_Object_Class
    retval.Fields__.Add "__class", T_Class
    retval.Fields__.Add "__superclass", self

    Set T_Class_M_Extend = retval
End Function
