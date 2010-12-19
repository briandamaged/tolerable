Option Explicit
'
' memoizer.vbs
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


Require "t\core\closures.vbs"




' The 0-argument Memoizer works a little differently than
' the others.  So, we'll define a separate class for it.
Class Memoizer_Class0
    Private m_func
    Private m_value
    Private m_need_eval
    
    Public Sub Initialize(ByVal func)
        Set m_func  = TO_Expr(0, func)
        m_need_eval = False
    End Sub
    
    Private Sub Class_Terminate()
        Set m_func = Nothing
    End Sub
    
    Public Default Function Invoke()
        If m_need_eval Then
            Assign m_value, m_func()
        End If
        Assign Invoke, m_value
    End Function
End Class


Private Function Memoizer_CreateMemoizer0__()
    Set Memoizer_CreateMemoizer0__ = New Memoizer_Class0
End Function




' Bootstrapping: To simplify the Closure implementation, we'll
'                initialize the MemoizerFactory with the 1-argument
'                Memoizer class already defined.  This way, the
'                ClosureFactory can leverage the Memoizer to simplify
'                its task.
Private Function Memoizer_CreateMemoizer1__()
    Set Memoizer_CreateMemoizer1__ = New Memoizer_Class1
End Function





Private Function MemoizerFactory_ArgLookupItr(ByVal i)
    MemoizerFactory_ArgLookupItr = _
        "        If Not d.Exists(arg" & i & ") Then" & vbCR &_
        "            d.Add arg" & i & ", CreateObject(""Scripting.Dictionary"")" & vbCR &_
        "        End If" & vbCR &_
        "        Set d = d.Item(arg" & i & ")" & vbCR
End Function


' This will create Memoizer classes as necessary
Class MemoizerFactory_Class

    Private m_constructors
    
    Private Sub Class_Initialize()
        Set m_constructors = CreateObject("Scripting.Dictionary")
        
        ' Bootstrapping: Predefine the 0 and 1 argument Memoizers
        m_constructors.Add 0, GetRef("Memoizer_CreateMemoizer0__")
        
        ExecuteGlobal GenerateClass__(1, "arg0", "ByRef arg0", "")
        m_constructors.Add 1, GetRef("Memoizer_CreateMemoizer1__")
    End Sub
    
    Private Sub Class_Terminate()
        Set m_constructors = Nothing
    End Sub

    
    Private Function GenerateClass(ByVal arg_count)
        Dim args : args = ClosureFactory.GetArgList(arg_count).Args
        Dim brs  : brs  = ClosureFactory.GetArgList(arg_count).ByValArgs

        GenerateClass   = GenerateClass__(arg_count, args, brs, "")
    End Function

    Private Function GenerateClass__(ByVal arg_count, ByVal args, ByVal brs, ByVal itr)
        Dim last : last = arg_count - 1

        GenerateClass__ = _
            "Class Memoizer_Class" & arg_count & vbCR &_
            "    Private m_memo" & vbCR &_
            "    Private m_func" & vbCR &_
            "    Public Sub Initialize(ByVal func)" & vbCR &_
            "        Set m_memo = CreateObject(""Scripting.Dictionary"")" & vbCR &_
            "        Set m_func = TO_Expr(" & arg_count & ", func)" & vbCR &_
            "    End Sub" & vbCR &_
            "    Private Sub Class_Terminate()" & vbCR &_
            "        Set m_memo = Nothing" & vbCR &_
            "    End Sub" & vbCR &_
            "    Public Default Function Invoke(" & brs &")" & vbCR &_
            "        Dim d    : Set d    = m_memo" & vbCR &_
            itr &_
            "        If Not d.Exists(arg" & last & ") Then" & vbCR &_
            "            d.Add arg" & last & ", m_func(" & args & ")" & vbCR &_
            "        End If" & vbCR &_
            "        Assign Invoke, d.Item(arg" & last & ")" & vbCR &_
            "    End Function" & vbCR &_
            "End Class"
    End Function

    Public Function Create(ByVal arg_count, ByVal func)
        If Not m_constructors.Exists(arg_count) Then
            ExecuteGlobal GenerateClass(arg_count)
            m_constructors.Add arg_count, Lambda(0, Nothing, "Set Invoke = New Memoizer_Class" & arg_count)
        End If
        
        Dim c, retval
        Set c = m_constructors.Item(arg_count)
        
        Set retval = c()
        retval.Initialize func
        Set Create = retval
    End Function

End Class


Dim MemoizerFactory_Singleton__
Public Function MemoizerFactory()
    If IsEmpty(MemoizerFactory_Singleton__) Then
        Set MemoizerFactory_Singleton__ = New MemoizerFactory_Class
    End If
    Set MemoizerFactory = MemoizerFactory_Singleton__
End Function





Public Function Memoizer(ByVal arg_count, ByVal func)
    Set Memoizer = MemoizerFactory.Create(arg_count, func)
End Function
