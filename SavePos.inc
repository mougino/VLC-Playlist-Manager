
' Just call CB_SAVEPOS(id_string) at the very beginning of your Dialog CALLBACK
' to save and restore position of your Dialog upon creation / closing

'------------------------------------------------------------------------------
MACRO CB_SAVEPOS
    SELECT CASE AS LONG CB.MSG
        CASE %WM_INITDIALOG
            LoadDialogPos CB.HNDL, EXE.NAME$
        CASE %WM_SYSCOMMAND
            SaveDialogPos CB.HNDL, EXE.NAME$
    END SELECT
END MACRO
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION Exist(BYVAL fileOrFolder AS STRING) AS LONG
    LOCAL Dummy&
    Dummy& = GETATTR(fileOrFolder)
    FUNCTION = (ERRCLEAR = 0)
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION LocalAppData AS STRING
    LOCAL szBaseFolder AS ASCIIZ * %MAX_PATH
    ShGetFolderPath (BYVAL 0, %CSIDL_LOCAL_APPDATA, BYVAL 0, BYVAL 0, szBaseFolder)
    FUNCTION = TRIM$(szBaseFolder) + "\"
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION RoamingAppData AS STRING
    LOCAL e AS STRING
    e = LCASE$(LocalAppData())
    REPLACE "\local\" WITH "\roaming\" IN e
    FUNCTION = e
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
SUB SaveDialogPos(BYVAL hDlg AS DWORD, BYVAL POS_FILE AS STRING)
    LOCAL ff, x, y AS LONG
    IF IsIconic(hDlg) THEN EXIT SUB
    DIALOG GET LOC hDlg TO x, y
    ff = FREEFILE
    OPEN LocalAppData + POS_FILE + ".pos" FOR OUTPUT AS #ff
        PRINT #ff, x
        PRINT #ff, y
    CLOSE #ff
END SUB
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
SUB LoadDialogPos(BYVAL hDlg AS DWORD, BYVAL POS_FILE AS STRING)
    LOCAL ff, x, y AS LONG
    LOCAL e AS STRING
    IF NOT EXIST(LocalAppData + POS_FILE + ".pos") THEN EXIT SUB
    ff = FREEFILE
    OPEN LocalAppData + POS_FILE + ".pos" FOR INPUT AS #ff
        LINE INPUT #ff, e : x = VAL(e)
        LINE INPUT #ff, e : y = VAL(e)
    CLOSE #ff
    IF x <> 0 OR y <> 0 THEN
        DIALOG SET LOC hDlg, x, y
    END IF
END SUB
'------------------------------------------------------------------------------
