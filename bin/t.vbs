Option Explicit
'
' t.vbs
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






' This is a bare-bones implementation of a LibraryManager.
' It will allow us to pull in the core libraries so that
' we can then create the ACTUAL LibraryManager.
Class LibraryManager_Bootstrapper_Class

    Private m_tolerable_lib
    Private m_fso
    
    Private m_loaded
    
    Private m_abs_path_regex
    
    Public Sub Initialize(ByVal tolerable_lib)
        m_tolerable_lib             = tolerable_lib
        
        Set m_fso                   = CreateObject("Scripting.FilesystemObject")
        Set m_loaded                = CreateObject("Scripting.Dictionary")

        ' TODO: I don't like the fact that this will be
        '       redefined elsewhere in the code.  Oh well.
        Set m_abs_path_regex        = new RegExp        
        m_abs_path_regex.Pattern    = "^(\\\\|\w:)"
        m_abs_path_regex.IgnoreCase = True
    End Sub

    Private Sub Class_Terminate()
        Set m_tolerable_lib  = Nothing
        Set m_fso            = Nothing
        Set m_loaded         = Nothing
        
        Set m_abs_path_regex = Nothing
    End Sub
    
    Private Function IsAbsolutePath(ByVal p)
        IsAbsolutePath = m_abs_path_regex.Test(p)
    End Function
    
    Public Function GetAbsolutePath(ByVal p)
        GetAbsolutePath = m_fso.GetAbsolutePathName(p)
    End Function
    
    
    Private Function IsLoaded(ByVal lib)
        lib = m_fso.GetAbsolutePathName(lib)
        IsLoaded = m_loaded.Exists(lib)
    End Function
    
    Public Property Get Loaded
        Loaded = m_loaded.keys
    End Property
    
    
    Public Function MarkAsLoaded(ByVal lib)
        lib = m_fso.GetAbsolutePathName(lib)
        If IsLoaded(lib) Then
            MarkAsLoaded = False
        Else
            m_loaded.Add lib, Nothing
            MarkAsLoaded = True
        End If
    End Function
    
    
    Private Function Locate(ByVal lib)
        If IsAbsolutePath(lib) Then
            Locate = lib
        Else
            Locate = m_tolerable_lib & lib
        End If
    End Function
    
    Public Function Import(ByVal lib)
        lib = Locate(lib)
        If IsLoaded(lib) Then
            Import = False
        Else
            On Error Resume Next
            Err.Clear
            Dim fin
            Set fin = m_fso.OpenTextFile(lib, 1)
            If Err.Number <> 0 Then
                Dim source : source = Err.Source
                On Error Goto 0
                Err.Raise 53, source, "File not found: " & lib
            End If
            On Error Goto 0
            MarkAsLoaded lib
            ExecuteGlobal fin.ReadAll
            fin.Close
            Import = True
        End If
    End Function

End Class








Dim LibraryManager_Singleton__
Public Function LibraryManager
    If IsEmpty(LibraryManager_Singleton__) Then
        Dim t_lib : t_lib = WScript.ScriptFullName
        
        ' 9 is the number of characters in  "bin\t.vbs"
        t_lib = LEFT(t_lib, LEN(t_lib) - 9) & "lib\" 
        
        ' Instantiate the LibraryManager bootstrapper.
        Set LibraryManager_Singleton__ = New LibraryManager_Bootstrapper_Class
        LibraryManager_Singleton__.Initialize t_lib

        
        ' Now we will bootstrap the actual LibraryManager
        Require "t\core\library_manager.vbs"
        
        '''' TODO
        ' Dim temp : Set temp = New LibraryManager_Class
        ' temp.Paths.Unshift t_lib
        
        
        ' Set LibraryManager_Singleton__ = temp
    End If
    Set LibraryManager = LibraryManager_Singleton__
End Function

Public Function Require(ByVal lib)
    LibraryManager.Import(lib)
End Function






' Place the bootstrapping logic inside
' sub so that we don't needlessly create
' global variables.
Private Sub Tolerable_Launch()
    If WScript.Arguments.Count = 0 Then
        WScript.Echo "Usage: t [script] {arg1 {arg2 {...}}}"
    Else
        LibraryManager.Import LibraryManager.GetAbsolutePath(WScript.Arguments.Item(0))
    End If

End Sub



Tolerable_Launch
