VERSION 5.00
Begin VB.UserControl TextEdit 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "TextEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private RuntimeMode As Boolean
Private Offset As Long
Private Declare Function SetClassLong Lib "user32.dll" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function SetFocus Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Dim m_Font As IFont
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateBrushIndirect Lib "gdi32.dll" (ByRef lpLogBrush As LOGBRUSH) As Long
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type
Private m_BackColor As Long
Private m_BackColorColor As Long
Private m_ForeColor As Long

Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, _
                        ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, _
                        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, _
                        ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC As Long = -4
Private Declare Function DefWindowProc Lib "user32.dll" Alias "DefWindowProcW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WS_EX_WINDOWEDGE As Long = &H100&
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_EX_OVERLAPPEDWINDOW As Long = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const ES_AUTOHSCROLL As Long = &H80&
Private Const ES_NOHIDESEL As Long = &H100&
Private Const ES_MULTILINE As Long = &H4&       'Multiline
Private Const ES_AUTOVSCROLL As Long = &H40&    '
Private Const WS_VSCROLL As Long = &H200000     '
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Const AsmMain As String = "558BEC83C4FC8D45FC50FF7514FF7510FF750CFF75086800000000B800000000FFD08B45FCC9C21000"
Private ASMArr() As Byte
Private editproc As Long
Private editwnd As Long
Private prtproc As Long
Private prtwnd As Long
Private multilined As Boolean
Private text_ As String

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_LIMITTEXT As Long = &HC5
Private Const EM_SETSEL As Long = &HB1
Private Const EM_LINELENGTH As Long = &HC1
Private Const EM_GETLINE As Long = &HC4
Private Const WM_SETFONT As Long = &H30
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINEINDEX As Long = &HBB

'Private Const WM_NCPAINT = &H85
'Private Const WM_ERASEBKGND = &H14
Private Const WM_PAINT = &HF

'Private Const WM_NCHITTEST = &H84
'Private Const WM_SETCURSOR = &H20   'if the mouse causes the cursor to move within a window and mouse input is not captured
'Private Const WM_NCMOUSEMOVE = &HA0 'MouseMove nonclient area
'Private Const WM_MOUSEMOVE = &H200  'MouseMove client area

Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONUP = &H205
'Private Const WM_CHAR As Long = &H102
'Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_DESTROY As Long = &H2
'Private Const WM_IME_SETCONTEXT As Long = &H281
'Private Const WM_IME_NOTIFY As Long = &H282
'Private Const WM_NCDESTROY As Long = &H82
Private Const WM_CTLCOLOREDIT As Long = &H133
Private Const WM_SETTEXT As Long = &HC

Private Const WM_COMMAND As Long = &H111
Private Const EN_CHANGE As Long = &H300

Public Event KeyDown(ByRef KeyCode As Long, ByRef lParam As Long)
Public Event KeyUp(ByRef KeyCode As Long, ByRef lParam As Long)
Public Event MouseDown(ByVal Button As Long)
Public Event MouseUp(ByVal Button As Long)
Public Event Changed()
Public Event EditLostFocus()

Public Function EditWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
    Case WM_KILLFOCUS: RaiseEvent EditLostFocus
    Case WM_KEYDOWN: RaiseEvent KeyDown(wParam, lParam)
    Case WM_KEYUP: RaiseEvent KeyUp(wParam, lParam)
    Case WM_LBUTTONDOWN: RaiseEvent MouseDown(vbLeftButton)
    Case WM_RBUTTONDOWN: RaiseEvent MouseDown(vbRightButton)
    Case WM_LBUTTONUP: RaiseEvent MouseUp(vbLeftButton)
    Case WM_RBUTTONUP: RaiseEvent MouseUp(vbRightButton)
    Case WM_CTLCOLOREDIT:
        SetBkColor wParam, m_BackColorColor
        SetTextColor wParam, m_ForeColor
        EditWindowProc = m_BackColor
        Exit Function
    Case WM_COMMAND:
        Select Case wParam \ &H10000 'HiWord
        Case EN_CHANGE:
            RaiseEvent Changed
        End Select
'    Case WM_PAINT:  'Áåç ýòîãî ìîæåò ïðåêðàòèòüñÿ ïåðåðèñîâêà ðîäèòåëüñêîãî îêíà
'        EditWindowProc = CallWindowProc(editproc, hwnd, uMsg, wParam, lParam)   '???
'        EditWindowProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
'        Exit Function
    End Select
    If hwnd = editwnd Then _
        EditWindowProc = CallWindowProc(editproc, hwnd, uMsg, wParam, lParam) _
    Else EditWindowProc = CallWindowProc(prtproc, hwnd, uMsg, wParam, lParam)
End Function

Private Sub StartSubclass(ByRef ASM() As Byte, ByVal hwnd As Long, ByRef OldWndProc As Long, Optional ByVal ProcNumber As Long)                 ' Ñàáêëàññèíã ñ ïîì. ASM (àâòîðà íå çíàþ...)
    Dim lng As Long, tPtr As Long
    lng = Len(AsmMain) \ 2&
    ReDim ASM(0 To lng - 1)
    For lng = 0 To lng - 1
        ASM(lng) = Val("&H" & Mid$(AsmMain, (lng) * 2& + 1, 2&))
    Next lng
    Call CopyMemory(tPtr, ByVal ObjPtr(Me), 4&)
    Call CopyMemory(ASM(23), ObjPtr(Me), 4&)
    Call CopyMemory(ASM(28), ByVal tPtr + &H7A4 + (4& * ProcNumber), 4&)
    OldWndProc = SetWindowLong(hwnd, &HFFFC, VarPtr(ASM(0)))
End Sub

Private Sub StopSubclass(ByVal hwnd As Long, ByVal OldWndProc As Long)
    Call SetWindowLong(hwnd, &HFFFC, OldWndProc)
End Sub

Public Sub init(Optional ByVal multi As Boolean = False)
With UserControl
    multilined = multi
    If prtwnd = 0 Then
        prtwnd = CreateWindowEx(0, StrPtr("STATIC"), StrPtr(""), _
                    WS_VISIBLE Or WS_CHILD, _
                    Offset, Offset, .ScaleWidth - Offset * 2, .ScaleHeight - Offset * 2, hwnd, 0&, App.hInstance, ByVal 0&)
        Call StartSubclass(ASMArr, prtwnd, prtproc)
    End If
    If editwnd = 0 Then 'Âîîáùå, åñëè âûçûâàåòñÿ init, ýòî óñëîâèå íóæíî âûïîëíèòü
        editwnd = CreateWindowEx(0, StrPtr("EDIT"), StrPtr(""), _
                    WS_VISIBLE Or WS_CHILD Or ES_AUTOHSCROLL Or ES_NOHIDESEL Or IIf(multi, ES_MULTILINE Or ES_AUTOVSCROLL Or WS_VSCROLL, 0), _
                    0, 0, .ScaleWidth - Offset * 2, .ScaleHeight - Offset * 2, prtwnd, 0, App.hInstance, ByVal 0&)
        Call StartSubclass(ASMArr, editwnd, editproc)
        text = text_
    End If
    If m_BackColor = 0 Then
        Dim c1 As Long, c2 As Long
        If OleTranslateColor(vbWindowBackground, 0, c1) Then BackColor = vbBlack Else BackColor = c1
        If OleTranslateColor(vbWindowText, 0, c2) Then ForeColor = vbWhite Else ForeColor = c2
    End If
    If m_Font Is Nothing Then
        Set m_Font = New StdFont
        m_Font.Name = "Verdana"
    End If
    SendMessage editwnd, WM_SETFONT, m_Font.hFont, ByVal 1
End With
End Sub

Public Sub SelectText(Optional ByVal pStart As Long = 0, Optional ByVal pFinish As Long = -1)
    SendMessage editwnd, EM_SETSEL, 0, ByVal -1
End Sub

Public Property Let TextLimit(ByVal maxlen As Long)
    SendMessage editwnd, EM_LIMITTEXT, maxlen, ByVal 0
End Property

Public Property Let Font(ByVal fontname As String)
    m_Font.Name = fontname
    SendMessage editwnd, WM_SETFONT, m_Font.hFont, ByVal 1
End Property

Public Property Get length() As Long
    length = SendMessage(editwnd, EM_LINELENGTH, 0, ByVal 0)
End Property

Public Property Get text() As String
If RuntimeMode Then
    Dim textlen As Long, copied As Long
    If multilined Then
        Dim lineCount As Long, firstchar As Long, i As Long
        lineCount = SendMessage(editwnd, EM_GETLINECOUNT, 0, 0&)
        ReDim res(lineCount - 1) As String
        For i = 0 To lineCount - 1
            firstchar = SendMessage(editwnd, EM_LINEINDEX, i, ByVal 0)
            If b(textlen, SendMessage(editwnd, EM_LINELENGTH, firstchar, ByVal 0)) Then
                ReDim buf(textlen * 2) As Byte
                lngToArr buf, textlen, 0
                copied = SendMessage(editwnd, EM_GETLINE, i, buf(0))
                res(i) = MidB(buf, 1, copied * 2)
            End If
        Next i
        text = Join$(res, vbNewLine)
    Else
        If b(textlen, length) Then
            ReDim buf(textlen * 2) As Byte
            lngToArr buf, textlen, 0
            copied = SendMessage(editwnd, EM_GETLINE, 0, buf(0))
            text = MidB(buf, 1, copied * 2)
        End If
    End If
Else
    text = text_
End If
End Property

Public Property Let text(ByRef txt As String)
If RuntimeMode Then
    Call SendMessage(editwnd, WM_SETTEXT, 0, ByVal StrPtr(txt))
Else
    text_ = txt
    PropertyChanged "text_"
End If
End Property

Public Property Let BackColor(ByVal Color As OLE_COLOR)
    Dim lb As LOGBRUSH
    lb.lbColor = Color
    If m_BackColor <> 0 Then DeleteObject m_BackColor
    m_BackColor = CreateBrushIndirect(lb) 'brush with new background color of editbox
    m_BackColorColor = Color
    InvalidateRect editwnd, 0, True
End Property

Public Property Let ForeColor(ByVal Color As Long)
    m_ForeColor = Color
    InvalidateRect editwnd, 0, True
End Property

Public Property Let Multiline(ByVal bool As Boolean)
    If bool = multilined Then Exit Property 'Íåçà÷åì ïåðåñîçäàâàòü îêíî è ìåíÿòü ñâîéñòâî
If RuntimeMode Then
    Dim txt As String
    txt = text
    If editproc Then Call StopSubclass(editwnd, editproc): editproc = 0
    If editwnd Then DestroyWindow editwnd: editwnd = 0
    init bool
    text = txt
End If
    multilined = bool
End Property

Public Property Get Multiline() As Boolean
    Multiline = multilined
End Property

Private Sub UserControl_GotFocus()
    SetFocus editwnd
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    RuntimeMode = Ambient.UserMode
    Multiline = PropBag.ReadProperty("multiline_", False)
'    Offset = PropBag.ReadProperty("offset_", 0)
    text_ = PropBag.ReadProperty("text_")
    If RuntimeMode Then
        Call init(Multiline)
        text = text_
    End If
End Sub

Private Sub UserControl_Resize()
With UserControl
    MoveWindow prtwnd, Offset, Offset, .ScaleWidth - Offset * 2, .ScaleHeight - Offset * 2, 1
    MoveWindow editwnd, 0, 0, .ScaleWidth - Offset * 2, .ScaleHeight - Offset * 2, 1
End With
End Sub

Private Sub UserControl_Terminate()
    If editwnd Then
        If editproc Then Call StopSubclass(editwnd, editproc): editproc = 0
        If prtproc Then Call StopSubclass(prtwnd, prtproc): prtproc = 0
        If editwnd Then DestroyWindow editwnd: editwnd = 0
        If prtwnd Then DestroyWindow prtwnd: prtwnd = 0
        If m_BackColor Then DeleteObject m_BackColor: m_BackColor = 0
        Call m_Font.ReleaseHfont(m_Font.hFont): Set m_Font = Nothing
    End If
End Sub

'Ýìóëèðóåò ðàìêó çàäàííîé òîëùèíû è öâåòà
Public Property Let Frame(ByVal Color As Long, ByVal Size As Long)
    UserControl.BackColor = Color
    Offset = Size
    Call UserControl_Resize
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("multiline_", Multiline)
    Call PropBag.WriteProperty("text_", text_)
'    Call PropBag.WriteProperty("offset_", Offset)
End Sub

Public Property Get wnd() As Boolean
    wnd = hwnd
End Property
