
#INCLUDE ONCE "WIN32API.INC"

GLOBAL hHookMsgBox AS LONG
GLOBAL hDlgMsgBox AS LONG

FUNCTION CenterMessageBox_CB (BYVAL lMsg AS LONG, BYVAL wParam AS LONG, BYVAL lParam AS LONG) AS LONG
    LOCAL rc1  AS RECT
    LOCAL rc2  AS RECT
    LOCAL x    AS LONG
    LOCAL y    AS LONG

    IF lMsg = %HCBT_ACTIVATE THEN
        IF hDlgMsgBox <> %HWND_DESKTOP THEN
            GetWindowRect hDlgMsgBox, rc1
            GetWindowRect wParam, rc2
            x    = (rc1.nLeft + rc1.nRight - rc2.nRight + rc2.nLeft) / 2
            y    = (rc1.nTop + rc1.nBottom - rc2.nBottom + rc2.nTop) / 2
            SetWindowPos wParam, %NULL, x, y, 0, 0, %SWP_NOSIZE   ' Or %SWP_NOACTIVATE
            'SetWindowPos wParam, %HWND_TOPMOST, x, y, 0, 0, %SWP_NOSIZE   ' Or %SWP_NOACTIVATE
        END IF
        UnhookWindowsHookEx hHookMsgBox
    END IF
END FUNCTION

FUNCTION CenterMessageBox (BYVAL hDlg AS DWORD, pMsg AS STRING, pCaption AS STRING, BYVAL pStyle AS LONG) AS LONG
    hDlgMsgBox  = hDlg
    hHookMsgBox = SetWindowsHookEx(%WH_CBT, CODEPTR(CenterMessageBox_CB), GetModuleHandle(BYVAL 0&), GetCurrentThreadId)
    FUNCTION    = MessageBox (hDlg, (pMsg), (pCaption), pStyle)
END FUNCTION

FUNCTION CMSGBOX (BYVAL hDlg AS DWORD, pMsg AS STRING, pCaption AS STRING, BYVAL pStyle AS LONG) AS LONG
    hDlgMsgBox  = hDlg
    hHookMsgBox = SetWindowsHookEx(%WH_CBT, CODEPTR(CenterMessageBox_CB), GetModuleHandle(BYVAL 0&), GetCurrentThreadId)
    FUNCTION    = MessageBox (hDlg, (pMsg), (pCaption), pStyle)
END FUNCTION
