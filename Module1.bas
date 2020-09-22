Attribute VB_Name = "Module1"
Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)

'Ïðèñâîåíèå p2 ê p1 ñ ïðîâåðêîé ðåçóëüòàòà íà ðàâåíñòâî íóëþ
Public Function b(p1 As Long, p2 As Long) As Boolean
    p1 = p2
    b = CBool(p1)
End Function

'Write lng to pos in byte array
Public Sub lngToArr(ByRef arrTo() As Byte, ByRef lng As Long, ByRef pos As Long)
    CopyMemory arrTo(pos), ByVal VarPtr(lng), 4
    pos = pos + 4
End Sub
