Attribute VB_Name = "Module1"
  Option Explicit
  ' demo project showing how to use the API to manipulate a messagebox
  ' by Bryan Stafford of New Vision Software® - newvision@imt.net
  ' this demo is released into the public domain "as is" without
  ' warranty or guaranty of any kind.  In other words, use at
  ' your own risk.
  
 Public Const MSG_TITLE = "Question"
  
  ' the max length of a path for the system (usually 260 or there abouts)
  ' this is used to size the buffer string for retrieving the class name of the active window below
  Public Const MAX_PATH As Long = 260&

  Public Const API_TRUE As Long = 1&
  Public Const API_FALSE As Long = 0&
  
  ' font *borrowed* from the form used to replace MessageBox font
  Public g_hBoldFont As Long
  
  Public Const MSGBOXTEXT As String = "Have you ever seen a standard message box with a different font than all the others on the system?"
  Public Const WM_SETFONT As Long = &H30

  ' made up constants for setting our timer
  Public Const NV_CLOSEMSGBOX As Long = &H5000&
  Public Const NV_MOVEMSGBOX As Long = &H5001&
  Public Const NV_MSGBOXCHNGFONT As Long = &H5002&

  ' MessageBox() Flags
  Public Const MB_ICONQUESTION As Long = &H20&
  Public Const MB_TASKMODAL As Long = &H2000&

  ' SetWindowPos Flags
  Public Const SWP_NOSIZE As Long = &H1&
  Public Const SWP_NOZORDER As Long = &H4&
  Public Const HWND_TOP As Long = 0&

  Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  ' API declares
  Public Declare Function LockWindowUpdate& Lib "user32" (ByVal hwndLock&)
  
  Public Declare Function GetActiveWindow& Lib "user32" ()
  
  Public Declare Function GetDesktopWindow& Lib "user32" ()
  
  Public Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, _
                                                                        ByVal lpWindowName$)

  Public Declare Function FindWindowEx& Lib "user32" Alias "FindWindowExA" (ByVal hWndParent&, _
                              ByVal hWndChildAfter&, ByVal lpClassName$, ByVal lpWindowName$)

  Public Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal _
                                                        wMsg&, ByVal wParam&, lParam As Any)

  Public Declare Function MoveWindow& Lib "user32" (ByVal hWnd&, ByVal x&, ByVal y&, _
                                              ByVal nWidth&, ByVal nHeight&, ByVal bRepaint&)

  Public Declare Function ScreenToClientLong& Lib "user32" Alias "ScreenToClient" (ByVal hWnd&, _
                                                                                    lpPoint&)
  
  Public Declare Function GetDC& Lib "user32" (ByVal hWnd&)
  Public Declare Function ReleaseDC& Lib "user32" (ByVal hWnd&, ByVal hDC&)

  ' drawtext flags
  Public Const DT_WORDBREAK As Long = &H10&
  Public Const DT_CALCRECT As Long = &H400&
  Public Const DT_EDITCONTROL As Long = &H2000&
  Public Const DT_END_ELLIPSIS As Long = &H8000&
  Public Const DT_MODIFYSTRING As Long = &H10000
  Public Const DT_PATH_ELLIPSIS As Long = &H4000&
  Public Const DT_RTLREADING As Long = &H20000
  Public Const DT_WORD_ELLIPSIS As Long = &H40000
  
  Public Declare Function DrawText& Lib "user32" Alias "DrawTextA" (ByVal hDC&, ByVal lpsz$, _
                                          ByVal cchText&, lpRect As RECT, ByVal dwDTFormat&)
  
  Public Declare Function SetForegroundWindow& Lib "user32" (ByVal hWnd&)
  
  Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd&, _
                                                        ByVal lpClassName$, ByVal nMaxCount&)

  Public Declare Function GetWindowRect& Lib "user32" (ByVal hWnd&, lpRect As RECT)
  
  Public Declare Function SetWindowPos& Lib "user32" (ByVal hWnd&, ByVal hWndInsertAfter&, _
                                      ByVal x&, ByVal y&, ByVal cx&, ByVal cy&, ByVal wFlags&)
                                      
  Public Declare Function MessageBox& Lib "user32" Alias "MessageBoxA" (ByVal hWnd&, _
                                                ByVal lpText$, ByVal lpCaption$, ByVal wType&)

  Public Declare Function SetTimer& Lib "user32" (ByVal hWnd&, ByVal nIDEvent&, ByVal uElapse&, _
                                                                            ByVal lpTimerFunc&)
  
  Public Declare Function KillTimer& Lib "user32" (ByVal hWnd&, ByVal nIDEvent&)

Public Sub TimerProc(ByVal hWnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&)
  ' this is a callback function.  This means that windows "calls back" to this function
  ' when it's time for the timer event to fire
  
  ' first thing we do is kill the timer so that no other timer events will fire
  KillTimer hWnd, idEvent
  
  ' select the type of manipulation that we want to perform
  Select Case idEvent
    Case NV_CLOSEMSGBOX '<-- we want to close this messagebox after 4 seconds
      Dim hMessageBox&
      
      ' find the messagebox window
      hMessageBox = FindWindow("#32770", MSG_TITLE)
      
      ' if we found it make sure it has the keyboard focus and then send it an enter to dismiss it
      If hMessageBox Then
        Call SetForegroundWindow(hMessageBox)
        SendKeys "{enter}"
      End If
      
    Case NV_MOVEMSGBOX '<-- we want to move this messagebox
      Dim hMsgBox&, xPoint&, yPoint&
      Dim stMsgBoxRect As RECT, stParentRect As RECT
      
      ' find the messagebox window
      hMsgBox = FindWindow("#32770", "Position A Message Box")
    
      ' if we found it then move it
      If hMsgBox Then
        ' get the rect for the parent window and the messagebox
        Call GetWindowRect(hMsgBox, stMsgBoxRect)
        Call GetWindowRect(hWnd, stParentRect)
        
        ' calculate the position for putting the messagebox in the middle of the form
        xPoint = stParentRect.Left + (((stParentRect.Right - stParentRect.Left) \ 2) - _
                                              ((stMsgBoxRect.Right - stMsgBoxRect.Left) \ 2))
        yPoint = stParentRect.Top + (((stParentRect.Bottom - stParentRect.Top) \ 2) - _
                                              ((stMsgBoxRect.Bottom - stMsgBoxRect.Top) \ 2))
        
        ' make sure the messagebox will not be off the screen.
        If xPoint < 0 Then xPoint = 0
        If yPoint < 0 Then yPoint = 0
        If (xPoint + (stMsgBoxRect.Right - stMsgBoxRect.Left)) > _
                                          (Screen.Width \ Screen.TwipsPerPixelX) Then
          xPoint = (Screen.Width \ Screen.TwipsPerPixelX) - (stMsgBoxRect.Right - stMsgBoxRect.Left)
        End If
        If (yPoint + (stMsgBoxRect.Bottom - stMsgBoxRect.Top)) > _
                                          (Screen.Height \ Screen.TwipsPerPixelY) Then
          yPoint = (Screen.Height \ Screen.TwipsPerPixelY) - (stMsgBoxRect.Bottom - stMsgBoxRect.Top)
        End If
        
        
        ' move the messagebox
        Call SetWindowPos(hMsgBox, HWND_TOP, xPoint, yPoint, _
                                        API_FALSE, API_FALSE, SWP_NOZORDER Or SWP_NOSIZE)
      End If
      
      ' unlock the desktop
      Call LockWindowUpdate(API_FALSE)
      
      
    Case NV_MSGBOXCHNGFONT '<-- we want to change the font for this messagebox
      '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      ' NOTE: Changing the font of a message box is not recomemded!!
      '       This portion of the demo is just provided to show some of the possibilities
      '       for manipulating other windows using the Windows API.
      '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
      
      ' find the messagebox window
      hMsgBox = FindWindow("#32770", "Change The Message Box Font")
    
      ' if we found it then find the static control that holds the text...
      If hMsgBox Then
        Dim hStatic&, hButton&, stMsgBoxRect2 As RECT
        Dim stStaticRect As RECT, stButtonRect As RECT
        
        ' find the static control that holds the message text
        hStatic = FindWindowEx(hMsgBox, API_FALSE, "Static", MSGBOXTEXT)
        hButton = FindWindowEx(hMsgBox, API_FALSE, "Button", "OK")
        
        ' if we found it, change the text and resize the static control so it will be displayed
        If hStatic Then
          ' get the rects of the message box and the static control before we change the font
          Call GetWindowRect(hMsgBox, stMsgBoxRect2)
          Call GetWindowRect(hStatic, stStaticRect)
          Call GetWindowRect(hButton, stButtonRect)
          
          ' set the font we borrowed from the form into the static control
          Call SendMessage(hStatic, WM_SETFONT, g_hBoldFont, ByVal API_TRUE)
          
          With stStaticRect
            ' convert the rect from screen coordinates to client coordinates
            Call ScreenToClientLong(hMsgBox, .Left)
            Call ScreenToClientLong(hMsgBox, .Right)
            
            Dim nRectHeight&, nHeightDifference&, hStaticDC&
            
            ' get the current height of the static control
            nHeightDifference = .Bottom - .Top
            
            ' get the device context of the static control to pass to DrawText
            hStaticDC = GetDC(hStatic)
            
            ' use DrawText to calculate the new height of the static control
            nRectHeight = DrawText(hStaticDC, MSGBOXTEXT, (-1&), stStaticRect, _
                                              DT_CALCRECT Or DT_EDITCONTROL Or DT_WORDBREAK)
            
            ' release the DC
            Call ReleaseDC(hStatic, hStaticDC)
            
            ' calculate the difference in height
            nHeightDifference = nRectHeight - nHeightDifference
            
            ' resize the static control so that the new larger bold text will fit in the messagebox
            Call MoveWindow(hStatic, .Left, .Top, .Right - .Left, nRectHeight, API_TRUE)
          End With
            
          ' move the button to the new position
          With stButtonRect
            ' convert the rect from screen coordinates to client coordinates
            Call ScreenToClientLong(hMsgBox, .Left)
            Call ScreenToClientLong(hMsgBox, .Right)
            
             ' move the button
            Call MoveWindow(hButton, .Left, .Top + nHeightDifference, .Right - .Left, .Bottom - .Top, API_TRUE)
          End With
          
          With stMsgBoxRect2
            ' resize and reposition the messagebox
            Call MoveWindow(hMsgBox, .Left, .Top - (nHeightDifference \ 2), .Right - .Left, (.Bottom - .Top) + nHeightDifference, API_TRUE)
          
            ' NOTE: if your message is very long, you may need to add code to make sure the messagebox
            ' will not run off the screen....
          End With
        End If
      End If
      
      ' unlock the desktop
      Call LockWindowUpdate(API_FALSE)
  
  End Select
  
End Sub


