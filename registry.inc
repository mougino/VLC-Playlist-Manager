'REGISTRY.INC
'General Purpose Registry Access
'Uses: win32api.inc

'Functions:
' DelRegValue(lpKey As Long,ByVal cMainkey As String, ByVal Key As String) As Long
' GetRegValue(lpKey As Long,ByVal cMainkey As String, ByVal Key As String) As String
' SetRegValue(lpKey As Long,ByVal cMainkey As String, ByVal Key As String, ByVal Setting As String) As Long
' DelRegKey(lpKey As Long, Key As String) Export As Long
' IsExtensionRegistered(FileExt AS STRING) AS LONG
' RegisterExtension(FileExt AS STRING, FileDesc AS STRING) AS LONG

'------------------------------------------------------------------------------
FUNCTION DelRegValue(lpKey AS LONG,BYVAL cMainkey AS STRING, BYVAL Key AS STRING) AS LONG
    ON ERROR RESUME NEXT
    LOCAL RetCode AS LONG
    LOCAL hKey AS LONG
    DIM acMainkey AS ASCIIZ * 300
     acMainkey = cMainkey
     RetCode = RegOpenKeyEx(lpKey, acMainkey, 0&, %KEY_ALL_ACCESS, hKey)
     IF RetCode = %ERROR_SUCCESS THEN
       IF Key$ = "*" THEN Key$ = CHR$(0,0)
       RetCode = RegDeleteValue(hKey, BYVAL STRPTR(Key$))
     END IF
     RegCloseKey hKey
     FUNCTION = RetCode
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION GetRegValue(lpKey AS LONG,BYVAL cMainkey AS STRING, BYVAL Key AS STRING) AS STRING
    ON ERROR RESUME NEXT
    DIM RetCode AS LONG
    DIM hKey AS LONG
    DIM KeyNameA AS ASCIIZ * 256
    DIM zTmp AS ASCIIZ * 256
    DIM acMainKey AS ASCIIZ * 300
    DIM ZZZ AS STRING
    DIM szdat&, cbData&, KeyType&
     acMainKey = cMainKey
     RetCode = RegOpenKeyEx(lpKey, acMainkey, 0&, %KEY_ALL_ACCESS, hKey)
     IF RetCode = %ERROR_SUCCESS THEN
       IF Key$ = "*" THEN Key$ = CHR$(0,0)
       szdat&=256
       DIM zbuffer AS ASCIIZ*256
       KeyNameA = Key
       cbData& = SIZEOF(zTmp)
       RetCode = RegQueryValueEx(BYVAL hKey, KeyNameA, BYVAL 0, KeyType&, zTmp, cbData&)
       ZZZ = zTmp
       FUNCTION = ZZZ
       EXIT FUNCTION
     END IF
     FUNCTION = ""
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION SetRegValue(lpKey AS LONG,BYVAL cMainkey AS STRING, BYVAL Key AS STRING, BYVAL Setting AS STRING) AS LONG
    ON ERROR RESUME NEXT
    LOCAL hKey   AS LONG
    LOCAL Result AS LONG
    LOCAL zText  AS ASCIIZ * 2048
    IF Key$ = "*" THEN Key$ = CHR$(0,0)
    IF RegCreateKeyEx(lpKey, cMainKey + CHR$(0),0, "", %REG_OPTION_NON_VOLATILE, _
                       %KEY_ALL_ACCESS, BYVAL %NULL, hKey, Result) <> %ERROR_SUCCESS THEN
        FUNCTION = 0
        EXIT FUNCTION
    END IF
    zText = Setting
    IF LEN(Setting) THEN
        RegSetValueEx hKey, Key+CHR$(0), 0, %REG_SZ, zText, LEN(Setting)+1
    ELSE
        RegSetValueEx hKey, Key+CHR$(0), 0, %REG_SZ, zText, 1
    END IF
    RegCloseKey hKey
    FUNCTION = 0
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
FUNCTION DelRegKey(lpKey AS LONG, Key AS STRING) EXPORT AS LONG
    ON ERROR RESUME NEXT
    LOCAL zStrKey AS ASCIIZ * 255, hKey AS LONG, Result AS LONG
    zStrKey = Key
    IF RegOpenKeyEx(BYVAL lpKey,zStrKey,0,  %KEY_ALL_ACCESS,  hKey) <> %ERROR_SUCCESS THEN  EXIT FUNCTION
    Result = RegDeleteKey(lpKey, zStrKey)
    RegCloseKey hKey
    FUNCTION = Result
END FUNCTION
'------------------------------------------------------------------------------

'
'  RegOpenSection
'------------------------------------------------------------------------------
FUNCTION RegOpenSection (BYVAL Key AS LONG, Section AS ASCIIZ) AS LONG

   DIM hKey AS LONG
   DIM Result AS LONG

   ' Create the section
   IF RegCreateKeyEx(Key, Section, 0, "", %REG_OPTION_NON_VOLATILE, _
                     %KEY_ALL_ACCESS, BYVAL %NULL, hKey, Result) <> %ERROR_SUCCESS THEN
      EXIT FUNCTION
   END IF

   '- Return the registry key handle
   FUNCTION = hKey

END FUNCTION

'
'  RegClose
'------------------------------------------------------------------------------
FUNCTION RegClose (BYVAL hKey AS LONG) AS LONG

      RegCloseKey hKey

END FUNCTION

'
'  RegSetString
'------------------------------------------------------------------------------
FUNCTION RegSetString (BYVAL hKey AS LONG, Entry AS ASCIIZ, Value AS ASCIIZ) AS LONG

   '- Save the value for the entry
      IF LEN(Value) THEN
      FUNCTION = RegSetValueEx(hKey, Entry, 0, %REG_SZ, Value, LEN(Value) + 1)
      ELSE
      FUNCTION = RegSetValueEx(hKey, Entry, 0, %REG_SZ, BYVAL %NULL, 0)
      END IF

END FUNCTION

'
'  IsExtensionRegistered
'------------------------------------------------------------------------------
FUNCTION IsExtensionRegistered(FileExt AS STRING) AS LONG

   DIM hKey AS LONG
   DIM zExtension AS ASCIIZ * 10
   DIM zBuffer AS ASCIIZ * 300

   zExtension = FileExt ' ".ext"
   FUNCTION = %False

   '- Open the key
   hKey = RegOpenSection(%HKEY_CLASSES_ROOT, zExtension)
   IF ISFALSE(hKey) THEN EXIT FUNCTION

   '- Check it's value
   IF RegQueryValueEx(hKey, "", 0, %REG_SZ, zBuffer, SIZEOF(zBuffer) - 1) = %ERROR_SUCCESS THEN
      'ShowMessage "zBuffer = " + zBuffer
      IF TRIM$(zBuffer) <> "" THEN FUNCTION = %True
   END IF
   RegClose hKey

END FUNCTION

'
'     RegisterExtension
'
'  This function associates the extension in the registry
'
'  It is called like this:
'
'  RegisterProgram ".ext", "Program Description", "c:\windows\system\myprog.exe %1"
'------------------------------------------------------------------------------
FUNCTION RegisterExtension(FileExt AS STRING, FileDesc AS STRING) AS LONG
   DIM hKey AS LONG
   DIM zExtension AS ASCIIZ * 10
   DIM zDescription AS ASCIIZ * 50
   DIM zOpen AS ASCIIZ * 300

   '- Set the variables
   zExtension = FileExt
   zDescription = FileDesc
   zOpen = EXE.FULL$ + " %1"

   '- Write out extension
      hKey = RegOpenSection(%HKEY_CLASSES_ROOT, zExtension)

      IF ISFALSE(hKey) THEN EXIT FUNCTION

      RegSetString hKey, "", zDescription
      RegClose hKey

   '- Write out description
      hKey = RegOpenSection(%HKEY_CLASSES_ROOT, zDescription)

      IF ISFALSE(hKey) THEN EXIT FUNCTION

      RegSetString hKey, "", zDescription
      RegClose hKey

   '- The open instruction
      hKey = RegOpenSection(%HKEY_CLASSES_ROOT, TRIM$(zDescription) + "\shell\open\command")

      IF ISFALSE(hKey) THEN EXIT FUNCTION

      RegSetString hKey, "", zOpen
      RegClose hKey

      FUNCTION = %True

END FUNCTION
