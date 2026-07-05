Option Explicit

Dim x, y
Dim Excel
Dim wsh

'Shell Object Creation
Set Excel = WScript.CreateObject("Excel.Application")
Set wsh = WScript.CreateObject("WScript.Shell")


'Key Codes
'Const VK_SHIFT = &H10

'Constants
Const MOUSEEVENTF_ABSOLUTE = 32768
Const MOUSE_MOVE = &H1
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4


Const SCREEN_X = 1920
Const SCREEN_Y = 1080

'API
Sub API_mouse_event(dwFlags, dx, dy, dwData, dwExtraInfo) 
    Dim strFunction 
    Const API_STRING = "CALL(""user32"",""mouse_event"",""JJJJJJ"", $1, $2, $3, $4, $5)" 
    strFunction = Replace(Replace(Replace(Replace(Replace(API_STRING, "$1", dwFlags), "$2", dx), "$3", dy), "$4", dwData), "$5", dwExtraInfo) 
    Call Excel.ExecuteExcel4Macro(strFunction) 
End Sub

Function API_GetMessagePos() 
    Dim ret, strHex, x, y 
    Dim strFunction 
    Const API_STRING = "CALL(""user32"",""GetMessagePos"",""J"")" 
    strFunction = API_STRING 
    ret = Excel.ExecuteExcel4Macro(strFunction) 
    strHex = Right("00000000" & Hex(ret), 8) 
    x = CLng("&H" & Right(strHex, 4)) 
    y = CLng("&H" &  Left(strHex, 4)) 
    API_GetMessagePos = Array(x, y) 
End Function 

'Mouse Pointer Move
Sub MouseMove(x, y)
    Dim pos_x, pos_y, dwFlags

    dwFlags = MOUSEEVENTF_ABSOLUTE + MOUSE_MOVE
    pos_x = Int(x * 1 / SCREEN_X)
    pos_y = Int(y * 1 / SCREEN_Y)
    Call API_mouse_event(dwFlags, pos_x, pos_y, 0, 0)
    WScript.Sleep 100
End Sub

Function get_point(ret_x,ret_y)
    Dim myPoint
    myPoint = API_GetMessagePos 
    ret_x = myPoint(0) 
    ret_y = myPoint(1) 
End Function

Dim x_pointer,y_pointer

Do While True 
    Call get_point(x_pointer,y_pointer)
    Dim angle,radian
    angle = Int((359 - 0 + 1) * Rnd + 0)
    radian = angle*3.14/180
    x_pointer=Int(SCREEN_X/2)
    y_pointer=Int(SCREEN_Y/2)
    x_pointer = x_pointer+Int(1.3+10*Cos(radian))
    y_pointer = y_pointer+Int(1.3+10*Sin(radian))
    MouseMove x_pointer , y_pointer  
    WScript.Sleep 290000
    
Loop

