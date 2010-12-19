Option Explicit
'
' paths.vbs
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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TODO: I have not finished implementing this module yet, and I may end
'       up doing a full redesign.  My goal is to create a nice overlay
'       to VBScript's standard FileSystem model that also plays nicely
'       w/ Enumerators.


Class PathUtil_Class

    Private m_fso
    Private m_absolute_regex
    
    Private Sub Class_Initialize()
        Set m_fso            = CreateObject("Scripting.FileSystemObject")
        Set m_absolute_regex = New RegExp

        m_absolute_regex.Pattern = "^(\\\\|\w:)"
    End Sub
    
    Private Sub Class_Terminate()
        Set m_fso            = Nothing
        Set m_absolute_regex = Nothing
    End Sub

    Public Function Exists(ByVal p)
        If m_fso.FileExists(p) Then
            Exists = True
        ElseIf m_fso.FolderExists(p) Then
            Exists = True
        Else
            Exists = False
        End If
    End Function
    
    
    Public Function GetAbsolute(ByVal p)
        GetAbsolute = m_fso.GetAbsolutePathName(p)
    End Function
    
    Public Function IsAbsolute(ByVal p)
        IsAbsolute = m_absolute_regex.Test(p)
    End Function
    
    Public Function IsRelative(ByVal p)
        IsRelative = Not IsAbsolute(p)
    End Function


End Class


Dim PathUtil_Singleton__
Private Function PathUtil
    If IsEmpty(PathUtil_Singleton__) Then
        Set PathUtil_Singleton__ = New PathUtil_Class
    End If
    Set PathUtil = PathUtil_Singleton__
End Function




Class Path_Class
    Private m_value
    Private m_absolute
    
    Public Sub Initialize(ByVal p)
        m_value    = p
        m_absolute = Empty
    End Sub
    
    Private Sub Class_Terminate()
        Set m_value = Nothing
    End Sub
    
    Public Property Get Value
        Value = m_value
    End Property
    
    
    Public Property Get Exists
        Exists = PathUtil.Exists(Value)
    End Property
    
    
    ' If I'm already an absolute path, then just return myself.
    ' Otherwise, create an absolute path based on myself.
    Public Function GetAbsolute()
        If IsAbsolute Then
            Set GetAbsolute = Me
        Else
            Set GetAbsolute = Path(PathUtil.GetAbsolute(Value))
        End If
    End Function
    
    Public Property Get IsAbsolute
        If IsEmpty(m_absolute) Then
            m_absolute = PathUtil.IsAbsolute(Value)
        End If
        IsAbsolute = m_absolute
    End Property
    
    Public Property Get IsRelative
        IsRelative = Not IsAbsolute
    End Property

End Class

Public Function Path(ByVal p)
    Dim retval
    Set retval = New Path_Class
    retval.Initialize p
    Set Path = retval
End Function


Public Function TO_Path(ByVal p)
    If TypeName(p) = "Path_Class" Then
        Set TO_Path = p
    Else
        Set TO_Path = Path(p)
    End If
End Function