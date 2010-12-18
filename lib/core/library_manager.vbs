Option Explicit
'
' library_manager.vbs
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


' TODO: I have not finished the implementation of this class yet,
'       and I may in fact remove this class altogether.  I'm trying
'       to figure out the best way to provide seamless integration
'       for QTP (in my previous work, the QTP library manager worked
'       very differently than the standard VBScript one).


Require "t\core\enumerator.vbs"




Class LibraryManager_Class

    Private m_paths
    Private m_loaded
    
    Private Sub Class_Initialize()
        Set m_paths = LinkedList
        Set m_loaded = CreateObject("Scripting.Dictionary")
    End Sub
    
    Private Sub Class_Terminate()
        Set m_paths  = Nothing
        Set m_loaded = Nothing
    End Sub
    
    Public Property Get Paths
        Set Paths = m_paths
    End Property


End Class




