#COMPILE EXE
#DIM ALL

#INCLUDE ONCE "win32api.inc"
#INCLUDE ONCE "shfolder.inc"
#INCLUDE ONCE "commctrl.inc"
#INCLUDE ONCE "registry.inc"
#INCLUDE ONCE "savepos.inc"
#INCLUDE ONCE "centermsgbox.inc"

#RESOURCE ".\res\vpm.pbr"

%WM_LOADPL      = %WM_USER + 1000
%WM_STARTPL     = %WM_USER + 1001
%WM_RESUMEPL    = %WM_USER + 1002
%WM_UPDATEMF    = %WM_USER + 1003

' TODO:
' [ ] Fix long paths in recent playlist combobox
' [X] Add eye icon to show activity of VLC Listener
' [X] Add ! icon to alert last VLC media couldn't be found
' [X] CenterMsgBox()
' [X] Fix relative paths broken when opened from local playlist
' [X] Fix xspf support
' [X] Fix recent playlist combobox duplicates
' [X] Fix empty recent playlists combobox when opening an invalid playlist
' [X] Double-click on a media file resumes the playlist at this point
' [X] Button "Start playlist" selects first media then start playlist
' [X] Button "Resume playlist" selects last media and resumes playlist
' [X] Make dialog able to receive files by drag & drop
'   [X] Start/resume drag and dropped playlist
' [X] Save last played file from playlist
' [X] Buttons:
'   [X] Start/Resume playlist
'   [X] Help
' [X] Create a temp playlist starting at current media till last media of playlist
' [X] Make VLC run this temp playlist
' [X] VLC listener (every 2s)
'   [X] check last played media in the .ini, update LastMf + listview accordingly
'   [X] check if VLC has been killed -> if yes stop VLC listener

'------------------------------------------------------------------------------
GLOBAL VlcPath  AS STRING
GLOBAL LastPl   AS STRING ' the last playlist opened
GLOBAL LastMf   AS STRING ' the last media file played from the current playlist
GLOBAL Mf()     AS STRING ' the list of media files for the current playlist
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION GetVLCLastPlayedMedia() AS STRING
    LOCAL zstr AS ASCIIZ * %MAX_PATH
    LOCAL i AS LONG
    GetPrivateProfileString "RecentsMRL", "list", "", zstr, %MAX_PATH, _
        RoamingAppData + "vlc\vlc-qt-interface.ini"
    i = INSTR(zstr, ", ")
    IF i > 0 THEN zstr = LEFT$(zstr, i-1)
    FUNCTION = TRIM$(zstr)
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION PBMAIN () AS LONG
    LOCAL e  AS STRING
    LOCAL ff AS LONG

    ' Get VLC path from registry
    LET VlcPath = GETREGVALUE(%HKEY_LOCAL_MACHINE, "SOFTWARE\VideoLAN\VLC", "")
    IF VlcPath = "" THEN
        CMSGBOX 0, "VLC not detected on your system",EXE.NAME$,%MB_ICONERROR
        EXIT FUNCTION
    END IF

    ' Get last playlist, if any
    e = LocalAppData + EXE.NAME$ + ".last.pl"
    IF EXIST(e) THEN
        ff = FREEFILE
        OPEN e FOR INPUT AS #ff
        LINE INPUT #ff, LastPl
        CLOSE #ff
    END IF

    ' Show main dialog
    ShowDialog 0

END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION IsVlcAlive() AS LONG
' Search for a VLC instance currently running
    LOCAL hWnd AS DWORD
    LOCAL length AS LONG
    LOCAL e AS STRING
    FUNCTION = 0
    hWnd = FindWindow("Qt5QWindowIcon", "") ' VLC dialog class
    IF hWnd = 0 THEN EXIT FUNCTION
    length = SendMessage(hWnd, %WM_GETTEXTLENGTH, 0, 0) + 1    ' Get VLC caption length
    e = SPACE$(length)
    length = SendMessage(hWnd, %WM_GETTEXT, length, STRPTR(e)) ' Get VLC caption
    IF INSTR(e, "VLC media player") = 0 THEN EXIT FUNCTION
    FUNCTION = 1
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
THREAD FUNCTION VlcListener(BYVAL hDlg AS DWORD) AS DWORD
    LOCAL curr_vid  AS STRING
    LOCAL last_vid  AS STRING
    LOCAL found     AS LONG

    curr_vid = LastMf

    DO
        SLEEP 2000
        IF ISFALSE IsVlcAlive() THEN EXIT LOOP
        last_vid = GetVLCLastPlayedMedia()
        IF last_vid = "" THEN
            CONTROL SHOW STATE hDlg, 4002, %SW_RESTORE ' VLC last media file not accessible
        ELSE
            CONTROL SHOW STATE hDlg, 4002, %SW_HIDE
            IF last_vid <> curr_vid THEN                ' VLC changed the media played
                LISTVIEW FIND EXACT hDlg, 2001, 1, UrlDecode(last_vid) TO found
                IF found THEN
                    lastMf = last_vid
                    DIALOG POST hDlg, %WM_UPDATEMF, 0, 0
                    curr_vid = last_vid
                END IF
            END IF
        END IF
    LOOP

    FUNCTION = 1 ' done
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION DqProtect(s AS STRING) AS STRING
    IF INSTR(s, " ") > 0 THEN FUNCTION = $DQ + s + $DQ ELSE FUNCTION = s
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
SUB UpdateLV(BYVAL hDlg AS DWORD)
    LOCAL i, found AS LONG
    LISTVIEW FIND EXACT hDlg, 2001, 1, UrlDecode(LastMf) TO found
    IF found = 0 THEN EXIT SUB
    FOR i = LBOUND(Mf) TO UBOUND(Mf)
        IF i < found THEN
            LISTVIEW SET IMAGE hDlg, 2001, i, 2 ' played
        ELSEIF i = found THEN
            LISTVIEW SET IMAGE hDlg, 2001, i, 1 ' play
            LISTVIEW SELECT hDlg, 2001, i
            LISTVIEW VISIBLE hDlg, 2001, i
        ELSE
            LISTVIEW SET IMAGE hDlg, 2001, i, 0 ' no icon (media yet to be played)
        END IF
    NEXT
END SUB
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
SUB CreateLocalM3uStartingFromSel(BYVAL hDlg AS DWORD)
    LOCAL i, k, ff AS LONG
    LOCAL e AS STRING
    LOCAL p AS STRING

    ' Get selected media
    LISTVIEW GET SELECT hDlg, 2001 TO k
    LISTVIEW GET USER hDlg, 2001, k TO k

    ' Create absolute path if needed
    p = FPath(lastPl)

    ' Create local m3u playlist
    ff = FREEFILE
    OPEN LocalAppData + FName(LastPl) + ".m3u" FOR OUTPUT AS #ff
    PRINT #ff, "#EXTM3U"
    FOR i = LBOUND(Mf) TO UBOUND(Mf)
        IF i >= k THEN
            PRINT #ff, "#EXTINF:," + FName(UrlDecode(Mf(i)))
            e = Mf(i)
            IF MID$(LTRIM$(e, "file:///"), 2, 2) <> ":/" THEN
                e = p + LTRIM$(e, "file:///")
            END IF
            PRINT #ff, e
        END IF
    NEXT
    CLOSE #ff
END SUB
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
MACRO LaunchLocalM3u
    MACROTEMP lRes
    DIM lRes AS LONG
    lRes = SHELL(VlcPath + " " + DqProtect(LocalAppData + FName(LastPl) + ".m3u"))
    IF hThread = 0 THEN
        THREAD CREATE VlcListener(CB.HNDL) TO hThread
        idTimer = SetTimer(CB.HNDL, 999, 2000, BYVAL 0)
        CONTROL SHOW STATE CB.HNDL, 4001, %SW_RESTORE ' VLC Listener thread active
    END IF
END MACRO
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
MACRO KillVlcListener
    THREAD CLOSE hThread TO i
    hThread = 0
    KillTimer CB.HNDL, idTimer
    idTimer = 0
    CONTROL SHOW STATE CB.HNDL, 4001, %SW_HIDE
END MACRO
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
CALLBACK FUNCTION ProcDialog
    STATIC hThread AS DWORD
    STATIC idTimer AS LONG
    LOCAL e  AS STRING
    LOCAL i  AS LONG

    CB_SAVEPOS ' Save & restore dialog position

    SELECT CASE AS LONG CB.MSG

        ' Initialization handler
        ' ---------------------------------------------
        CASE %WM_INITDIALOG
            LoadRecentList CB.HNDL
            DIALOG POST CB.HNDL, %WM_LOADPL, 0, 0

        ' User message: load last playlist, if any
        ' ---------------------------------------------
        CASE %WM_LOADPL
            e = TRIM$(COMMAND$)
            IF e <> "" AND EXIST(e) THEN        ' ...passed as argument
                LoadPlaylist CB.HNDL, e
            ELSEIF LastPl <> "" THEN
                LoadPlaylist CB.HNDL, LastPl    ' ...or from last session
            END IF

        ' User message: start playlist from the top
        ' ---------------------------------------------
        CASE %WM_STARTPL
            CreateLocalM3uStartingFromSel CB.HNDL
            CONTROL SET TEXT CB.HNDL, 3001, "Resume playlist"
            LaunchLocalM3u CB.HNDL

        ' User message: resume playlist
        ' ---------------------------------------------
        CASE %WM_RESUMEPL
            CreateLocalM3uStartingFromSel CB.HNDL
            LaunchLocalM3u

        ' User message: last played media file changed
        ' ---------------------------------------------
        CASE %WM_UPDATEMF
            UpdateLV CB.HNDL
            CreateLocalM3uStartingFromSel CB.HNDL

        ' Drag and drop playlist on program to open it
        ' ---------------------------------------------
        CASE %WM_DROPFILES
            IF DragQueryFile(CB.WPARAM, -1, BYVAL 0, 0) > 1 THEN EXIT FUNCTION ' ignore multiple
            LOCAL zStr AS ASCIIZ * %MAX_PATH
            DragQueryFile CB.WPARAM, i, zStr, SIZEOF(zStr)
            e = TRIM$(zStr)
            DragFinish CB.WPARAM
            LoadPlaylist CB.HNDL, e

        ' Timer (2s) to check on VLC Listener thread
        ' ---------------------------------------------
        CASE %WM_TIMER
            IF hThread <> 0 THEN
                THREAD STATUS hThread TO i
                IF i <> &H103 THEN
                    KillVlcListener
                END IF
            END IF

        ' Double-click on listview item to lauch media file
        ' --------------------------------------------------
        CASE %WM_NOTIFY
            LOCAL pLV AS LVUNION PTR
            pLV = CB.LPARAM
            IF 2001 = @pLV.NMHDR.IDFROM _
            AND %NM_DBLCLK = @pLV.NMHDR.CODE THEN
                ' Force-stop VLC Listener to avoid resuming an old "last played media"
                KillVlcListener
                ' Set double-clicked media as last media file
                LISTVIEW GET SELECT CB.HNDL, 2001 TO i
                LISTVIEW GET USER CB.HNDL, 2001, i TO i
                LastMf = Mf(i)
                ' Update the listview, create and start local playlist
                UpdateLV CB.HNDL
                CreateLocalM3uStartingFromSel CB.HNDL
                LaunchLocalM3u
'                CMSGBOX CB.HNDL,LastMf,EXE.NAME$,0
            END IF

        CASE %WM_COMMAND                ' Process control notifications
            SELECT CASE AS LONG CB.CTL

                ' Double-click on "Playlists" label: secret feature!
                ' -----------------------------------------------------------------
                CASE 1001
                    IF CB.CTLMSG = %STN_DBLCLK THEN
                        i = SHELL("notepad " + LocalAppData + EXE.NAME$ + ".lst")
                    END IF

                ' New selection in "Recent playlists" combobox: change playlist
                ' -----------------------------------------------------------------
                CASE 1002
                    IF CB.CTLMSG = %CBN_SELENDOK THEN
                        COMBOBOX GET SELECT CB.HNDL, CB.CTL TO i
                        COMBOBOX GET TEXT CB.HNDL, CB.CTL, i TO e
                        IF LEFT$(e, 1) = "<" THEN EXIT FUNCTION ' <none loaded>
                        LoadPlaylist CB.HNDL, e
                    END IF

                ' Button "Load/Change playlist"
                ' -----------------------------------------------------------------
                CASE 1003
                    IF CB.CTLMSG = %BN_CLICKED OR CB.CTLMSG = 1 THEN
                        e = CHR$("VLC playlist (*.m3u *.m3u8 *.xspf)", 0)
                        e += CHR$("*.m3u;*.m3u8;*.xspf", 0)
                        DISPLAY OPENFILE CB.HNDL, -120, 0, "Load playlist", "", e, _
                            "", "", %OFN_FILEMUSTEXIST TO e
                        IF e = "" THEN EXIT FUNCTION
                        LoadPlaylist CB.HNDL, e
                    END IF

                ' Button "Start/Resume playlist"
                ' -----------------------------------------------------------------
                CASE 3001
                    CreateLocalM3uStartingFromSel CB.HNDL
                    CONTROL SET TEXT CB.HNDL, 3001, "Resume playlist"
                    LaunchLocalM3u CB.HNDL

                ' Button "Help"
                ' -----------------------------------------------------------------
                CASE 3002
                    ShellExecute %NULL, "open", "http://mougino.free.fr/vpm/", "", "", %SW_SHOW

            END SELECT
    END SELECT
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
SUB ParsePlaylist(BYVAL file AS STRING, BYREF arr() AS STRING)
' Parse a playlist and return an array of its media files
    LOCAL ff AS LONG
    LOCAL n  AS LONG
    LOCAL e  AS STRING

    IF NOT EXIST(file) THEN EXIT SUB
    ff = FREEFILE
    OPEN file FOR INPUT ACCESS READ LOCK SHARED AS #ff

    IF RIGHT$(LCASE$(file),5) = ".xspf" THEN
        DO
            LINE INPUT #ff, e : e = TRIM$(e, ANY $SPC+$TAB)
            IF INSTR(e, "<location>") <> 1 THEN ITERATE DO
            e = LTRIM$(e, "<location>")
            e = RTRIM$(e, "</location>")
            IF n = 0 THEN DIM arr(0) ELSE REDIM PRESERVE arr(n)
            arr(n) = e
            INCR n
        LOOP UNTIL EOF(#ff)

    ELSEIF RIGHT$(LCASE$(file),4) = ".m3u" OR RIGHT$(LCASE$(file),5) = ".m3u8" THEN
        DO
            LINE INPUT #ff, e
            IF LEFT$(e,1) = "#" THEN ITERATE DO
            IF n = 0 THEN DIM arr(0) ELSE REDIM PRESERVE arr(n)
            arr(n) = e
            INCR n
        LOOP UNTIL EOF(#ff)
    END IF
    CLOSE #ff
END SUB
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION FPath(filePath AS STRING) AS STRING
' Return url-encoded path for a file, including last slash
    LOCAL r AS STRING
    LOCAL i AS LONG
    r = filePath
    i = INSTR(-1, r, "\")
    IF i > 0 THEN r = LEFT$(r, i)
    REPLACE "\" WITH "/" IN r
    REPLACE "%" WITH "%25" IN r
    REPLACE " " WITH "%20" IN r
    FUNCTION = "file:///" + r
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION FName(filePath AS STRING) AS STRING
' Return file name without extension from a full file uri
    LOCAL r AS STRING
    LOCAL i AS LONG
    r = filePath
    REPLACE "/" WITH "\" IN r
    i = INSTR(-1, r, "\")
    IF i > 0 THEN r = MID$(r, i+1)
    i = INSTR(-1, r, ".")
    IF i > 0 THEN r = LEFT$(r, i-1)
    FUNCTION = r
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION UrlDecode(s AS STRING) AS STRING
' Decode an url encoded URI
' e.g. "file:///C:/file%20name.ext" returns "C:\file name.ext"
    LOCAL r AS STRING
    LOCAL i AS LONG
    r = LTRIM$(s, "file:///")
    i = INSTR(r, "%")
    WHILE i > 0
        IF i < LEN(r)-2 THEN r = LEFT$(r,i-1) + CHR$(VAL("&H0"+MID$(r,i+1,2))) + MID$(r,i+3)
        i = INSTR(i+1, r, "%")
    WEND
    REPLACE "/" WITH "\" IN r
    FUNCTION = r
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
SUB LoadPlaylist(hDlg AS DWORD, playlistPath AS STRING)
    LOCAL e     AS STRING
    LOCAL i, ff AS LONG
    LOCAL found AS LONG

    ' Reset last played media file
    LastMf = ""

    ' Load playlist content into listview
    LISTVIEW RESET hDlg, 2001
    REDIM Mf(-1)
    IF NOT EXIST(playlistPath) THEN
        CONTROL DISABLE hDlg, 3001 ' Start/Resume playlist
        EXIT SUB
    END IF
    ParsePlaylist playlistPath, Mf()
    IF UBOUND(Mf) < 0 THEN
        CONTROL SET TEXT hDlg, 3001, "Start playlist"
        CONTROL DISABLE hDlg, 3001 ' Start/Resume playlist
        CMSGBOX hDlg, "Invalid playlist", EXE.NAME$, %MB_ICONWARNING
        EXIT SUB
    END IF
    CONTROL ENABLE hDlg, 3001 ' Start/Resume playlist
    FOR i = UBOUND(Mf) TO LBOUND(Mf) STEP -1
        LISTVIEW INSERT ITEM hDlg, 2001, 1, 0, UrlDecode(Mf(i))
        LISTVIEW SET USER hDlg, 2001, 1, i ' pointer to media file path: Mf(i)
    NEXT

    ' Update header
    CONTROL SET TEXT  hDlg, 1002, playlistPath
    CONTROL SET TEXT  hDlg, 1003, "Change" ' Load button >> Change button
    AddToRecentList hDlg, playlistPath

    ' Save last playlist opened
    LastPl = playlistPath
    e = LocalAppData + EXE.NAME$ + ".last.pl"
    ff = FREEFILE
    OPEN e FOR OUTPUT AS #ff
    PRINT #ff, playlistPath
    CLOSE #ff

    ' Get last played media file for this playlist, if any
    e = LocalAppData + FName(playlistPath) + ".m3u"
    IF EXIST(e) THEN
        ff = FREEFILE
        OPEN e FOR INPUT ACCESS READ LOCK SHARED AS #ff
        LINE INPUT #ff, e ' #EXTM3U
        LINE INPUT #ff, e ' #EXTINF
        LINE INPUT #ff, LastMf
        CLOSE #ff
    END IF

    ' New playlist >> propose to start it
    LISTVIEW FIND EXACT hDlg, 2001, 1, UrlDecode(LastMf) TO found
    IF LastMf = "" OR found = 0 THEN
        CONTROL SET TEXT hDlg, 3001, "Start playlist"
        LISTVIEW SET IMAGE hDlg, 2001, 1, 1 ' play
        LISTVIEW SELECT hDlg, 2001, 1
        i = CMSGBOX(hDlg, "Start playlist ?", EXE.NAME$, %MB_ICONQUESTION OR %MB_YESNO)
        IF i = %IDYES THEN DIALOG POST hDlg, %WM_STARTPL, 0, 0
        EXIT SUB
    END IF

    ' Ongoing playlist >> mark read files and current file (where to resume)
    CONTROL SET TEXT hDlg, 3001, "Resume playlist"
    FOR i = LBOUND(Mf) TO UBOUND(Mf)
        IF i < found THEN
            LISTVIEW SET IMAGE hDlg, 2001, i, 2 ' played
        ELSEIF i = found THEN
            LISTVIEW SET IMAGE hDlg, 2001, i, 1 ' play
            LISTVIEW SELECT hDlg, 2001, i
            LISTVIEW VISIBLE hDlg, 2001, i
            EXIT FOR
        END IF
    NEXT
    i = CMSGBOX(hDlg, "Resume playlist ?", EXE.NAME$, %MB_ICONQUESTION OR %MB_YESNO)
    IF i = %IDYES THEN DIALOG POST hDlg, %WM_RESUMEPL, 0, 0

END SUB
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
SUB AddToRecentList(hDlg AS DWORD, playlistPath AS STRING)
    LOCAL found AS LONG
    LOCAL ff AS LONG

    COMBOBOX FIND EXACT hDlg, 1002, 1, playlistPath TO found
    IF found = 0 THEN
        ff = FREEFILE
        OPEN LocalAppData + EXE.NAME$ + ".lst" FOR APPEND AS  #ff
            PRINT #ff, playlistPath
        CLOSE #ff
    END IF
    LoadRecentList hDlg
END SUB
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
SUB LoadRecentList(hDlg AS DWORD)
    LOCAL noproj, found AS LONG
    LOCAL ff AS LONG
    LOCAL e AS STRING

    COMBOBOX RESET hDlg, 1002
    e = LocalAppData + EXE.NAME$ + ".lst"

    IF EXIST(e) THEN
        ff = FREEFILE
        OPEN e FOR INPUT AS  #ff
        DO
            LINE INPUT #ff, e
            IF LEN(e)>0 AND EXIST(e) THEN _
              COMBOBOX INSERT hDlg, 1002, 0, e
        LOOP UNTIL EOF(#ff)
        CLOSE #ff
        COMBOBOX FIND EXACT hDlg, 1002, 1, LastPl TO found
    ELSE
        COMBOBOX ADD hDlg, 1002, "<none loaded>"
        found = 1
    END IF

    COMBOBOX SELECT hDlg, 1002, found
END SUB
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION ShowDialog(BYVAL hParent AS LONG) AS LONG
    LOCAL lRes AS LONG
    LOCAL lSty AS LONG
    LOCAL hDlg AS DWORD
    LOCAL hIml AS DWORD

    DIALOG NEW PIXELS, hParent, EXE.NAME$, , , 340, 380, %WS_POPUP _
        OR %WS_BORDER OR %WS_DLGFRAME OR %WS_SYSMENU OR %WS_CLIPSIBLINGS OR _
        %WS_VISIBLE OR %DS_MODALFRAME OR %DS_3DLOOK OR %DS_NOFAILCREATE OR _
        %DS_SETFONT OR %WS_CAPTION OR %WS_MINIMIZEBOX, %WS_EX_WINDOWEDGE OR _
        %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
        %WS_EX_RIGHTSCROLLBAR, TO hDlg

    CONTROL ADD LABEL,      hDlg, 1001, "Playlist :", 8, 8, 48, 16, %SS_NOTIFY
    CONTROL ADD COMBOBOX,   hDlg, 1002, , 56, 5, 230, 18*10, _
        %CBS_DROPDOWNLIST, %WS_EX_CLIENTEDGE OR %WS_EX_LEFT
    CONTROL ADD BUTTON,     hDlg, 1003, "Load", 286, 5, 48, 22, _
        %BS_CENTER OR %BS_VCENTER OR %WS_GROUP OR %WS_TABSTOP

    CONTROL ADD LISTVIEW,   hDlg, 2001, "", 8, 38, 324, 280, %LVS_NOCOLUMNHEADER OR _
        %LVS_REPORT OR %LVS_SHOWSELALWAYS OR %LVS_SINGLESEL, %WS_EX_CLIENTEDGE
    LISTVIEW INSERT COLUMN  hDlg, 2001, 1, "Media file", 320, %SS_LEFT
    LISTVIEW GET STYLEXX    hDlg, 2001 TO lSty
    LISTVIEW SET STYLEXX    hDlg, 2001, lSty OR %LVS_EX_FULLROWSELECT OR _
        %LVS_EX_GRIDLINES OR %LVS_EX_INFOTIP

    CONTROL ADD BUTTON,     hDlg, 3001, "Start playlist", 8, 330, 150, 40
    CONTROL ADD BUTTON,     hDlg, 3002, "Help", 180, 330, 80, 40

    ' Add status icons
    CONTROL ADD IMAGE,      hDlg, 4001, "ICO4", 312, 354, 16, 16, %SS_NOTIFY ' eye  > VLC listener thread active
    CONTROL SHOW STATE      hDlg, 4001, %SW_HIDE
    CONTROL ADD IMAGE,      hDlg, 4002, "ICO5", 312, 336, 16, 16, %SS_NOTIFY '  !   > VLC last media not found
    CONTROL SHOW STATE      hDlg, 4002, %SW_HIDE

    ' Attach play/played imagelist to the listview
    IMAGELIST NEW ICON 16, 16, 24, 6 TO hIml
    IMAGELIST ADD ICON hIml, "ICO2" ' play
    IMAGELIST ADD ICON hIml, "ICO3" ' played
    LISTVIEW SET IMAGELIST  hDlg, 2001, hIml, 1

    DIALOG  SET ICON        hDlg, "ICO1"
    DragAcceptFiles         hDlg, %True
    DIALOG SHOW MODAL       hDlg, CALL ProcDialog TO lRes
    FUNCTION = lRes
END FUNCTION
'------------------------------------------------------------------------------
