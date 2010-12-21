Option Explicit

Require "t\core\enumerator.vbs"

' If you understand this code, then
' you understand Enumerators.
WScript.Echo Join(eRange(1, 10).Map(Q("arg0 & ' little'")) _
                               .GroupsOf(3) _
                               .Map(Q("Join(arg0) & ' programmer'")) _
                               .TO_Array,
                  vbCR) & " nerds"