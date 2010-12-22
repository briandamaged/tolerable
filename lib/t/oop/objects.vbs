Option Explicit

Require "t\oop\methods.vbs"



Class T_Object_Class

    Private m_fields
    Private m_unbound_methods
    
    Private Sub Class_Initialize()
        Set m_fields                  = CreateObject("Scripting.Dictionary")
        m_fields.CompareMode          = 1
        Set m_unbound_methods         = CreateObject("Scripting.Dictionary")
        m_unbound_methods.CompareMode = 1
    End Sub
    
    Private Sub Class_Terminate()
        Set m_fields          = Nothing
        Set m_unbound_methods = Nothing
    End Sub

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
        Fields__.Add name, value
        Assign SetF, value
    End Function
    
    Public Property Get UnboundMethods__()
        Set UnboundMethods__ = m_unbound_methods
    End Property

    Public Function M(ByVal name)
        If m_unbound_methods.Exists(name) Then
            Set M = m_unbound_methods.Item(name)(Me, name).Bind(Me)
        Else
            Dim clazz : Set clazz = GetF("__class")
            Dim i
            Do Until clazz Is Nothing
                Set i = clazz.Fields__.Item("__unbound_instance_methods")
                If i.Exists(name) Then
                    Set M = i.Item(name)(Me, name).Bind(Me)
                    Exit Function
                End If
                
                Set clazz = clazz.Fields__.Item("__superclass")
            Loop
        End If
    End Function

End Class








Private Sub T_Objects_Initialize()
    Set T_Object_Singleton__ = New T_Object_Class
    Set T_Module_Singleton__ = New T_Object_Class
    Set T_Class_Singleton__  = New T_Object_Class
    
    ' Shortcuts
    Dim  o : Set  o = T_Object_Singleton__
    Dim  m : Set  m = T_Module_Singleton__
    Dim  c : Set  c = T_Class_Singleton__
    
    
    ' All three top-level classes are of type T_Class
    o.Fields__.Add "__class", T_Class_Singleton__
    m.Fields__.Add "__class", T_Class_Singleton__
    c.Fields__.Add "__class", T_Class_Singleton__
    
    ' And that's where the similarities end
    o.Fields__.Add "__superclass", Nothing
    m.Fields__.Add "__superclass", T_Object_Singleton__
    c.Fields__.Add "__superclass", T_Module_Singleton__
    
    ' Call constructors explicitly.
    T_Module_M_Init o
    T_Module_M_Init m
    T_Module_M_Init c
    
    ' Bootstrap the "def" method so we can easily define more methods
    m.Fields__.Item("__unbound_instance_methods").Add "Def", T_UnboundMethod_Retriever_Fixed(T_UnboundMethod("def", 3, GetRef("T_Module_M_Def")))

    o.M("def").X "Init",    0, GetRef("T_Object_M_Init")
    o.M("def").X "SDef",    3, GetRef("T_Object_M_SDef")

    m.M("def").X "Init",    0, GetRef("T_Module_M_Init")
    m.M("def").X "InstanceMethods", 0, GetRef("T_Module_M_InstanceMethods")
    m.M("def").X "InstanceMethod",  1, GetRef("T_Module_M_InstanceMethod")
    
    c.Fields__.Item("__unbound_instance_methods").Add "New", T_UnboundMethod_Retriever_New()
    c.M("sdef").X "New",     0, GetRef("T_Class_SM_New")

    c.M("def").X "Extends", 0, GetRef("T_Class_M_Extend")
End Sub



Dim T_Object_Singleton__
Public Function T_Object()
    If IsEmpty(T_Object_Singleton__) Then
        T_Modules_Initialize
    End If
    Set T_Object = T_Object_Singleton__
End Function

Private Function T_Object_M_Init(ByVal self)
End Function

Private Function T_Object_M_SDef(ByVal self, ByVal name, ByVal arg_count, ByVal func)
    self.UnboundMethods__.Add name, T_UnboundMethod_Retriever_Fixed(T_UnboundMethod(name, arg_count, func))
End Function




Dim T_Module_Singleton__
Public Function T_Module()
    If IsEmpty(T_Module_Singleton__) Then
        T_Objects_Initialize
    End If
    Set T_Module = T_Module_Singleton__
End Function

Private Function T_Module_M_Init(ByVal self)
    Dim i : Set i = CreateObject("Scripting.Dictionary")
    i.CompareMode = 1
    self.Fields__.Add "__unbound_instance_methods", i
End Function

Private Function T_Module_M_Def(ByVal self, ByVal name, ByVal arg_count, ByVal func)
    self.Fields__.Item("__unbound_instance_methods").Add name, T_UnboundMethod_Retriever_Fixed(T_UnboundMethod(name, arg_count, func))
End Function


Private Function T_Module_M_InstanceMethods(ByVal self)
    T_Module_M_InstanceMethods = self.Fields__.Item("__unbound_instance_methods").Keys
End Function

Private Function T_Module_M_InstanceMethod(ByVal self, ByVal name)
    Set T_Module_M_InstanceMethod = self.Fields__.Item("__unbound_instance_methods").Item(name)(self, name)
End Function



Dim T_Class_Singleton__
Public Function T_Class()
    If IsEmpty(T_Class_Singleton__) Then
        T_Objects_Initialize
    End If
    Set T_Class = T_Class_Singleton__
End Function





Private Function T_Class_M_New_ClassGenerator(ByVal arg_count)
    Dim args : args = Join(eRange(1, arg_count).Map("ClosureFactory.ArgName(arg0)").TO_Array, ", ")

    ' TODO: Need to update the Closure code so that the ClosureFactory
    '       returns a function that produces an instance of a Closure.
    Set T_Class_M_New_ClassGenerator = Lambda(arg_count + 1, _
        "Dim retval : Set retval = New T_Object_Class" & vbCR &_
        "retval.Fields__.Add ""__class"", arg0" & vbCR &_
        "retval.M(""Init"").X " & args & vbCR &_
        "Set Invoke = retval")
End Function

Dim T_Class_M_New_Factory_Singleton__
Private Function T_Class_M_New_Factory(Byval arg_count)
    If IsEmpty(T_Class_M_New_Factory_Singleton__) Then
        Set T_Class_M_New_Factory_Singleton__ = Memoizer(1, GetRef("T_Class_M_New_ClassGenerator"))
    End If
    Set T_Class_M_New_Factory = T_Class_M_New_Factory_Singleton__(arg_count)
End Function

Class T_UnboundMethod_Retriever_New_Class
    Public Default Function Invoke(ByVal self, ByVal name)
        Dim clazz : Set clazz = self
        Dim ivs
        Dim init
        Do Until clazz Is Nothing
            Set ivs = clazz.Fields__.Item("__unbound_instance_methods")
            If ivs.Exists("Init") Then
                ' HACK: This should be bound to the actual instance object!
                '       Should be fixable once Closures are enhanced...
                Set init = ivs.Item("Init")(Nothing, "Init")
                Exit Do
            End If
            Set clazz = clazz.Fields__.Item("__superclass")
        Loop
        Set Invoke = T_UnboundMethod("New", init.ArgCount, T_Class_M_New_Factory(init.ArgCount))
    End Function
End Class

Private Function T_UnboundMethod_Retriever_New()
    Set T_UnboundMethod_Retriever_New = New T_UnboundMethod_Retriever_New_Class
End Function






Private Function T_Class_M_Init(ByVal self, ByVal superclass)
    self.Fields__.Add "__superclass", superclass
End Function

' Creating a new instance of Class is equivalent to
' extending Object.
Private Function T_Class_SM_New(ByVal self)
    Set T_Class_SM_New = T_Class_M_Extend(T_Object)
End Function

Private Function T_Class_M_Extend(ByVal self)
    Dim retval : Set retval = New T_Object_Class
    retval.Fields__.Add "__class",      T_Class
    retval.Fields__.Add "__superclass", self

    retval.M("Init").X
    
    Set T_Class_M_Extend = retval
End Function
