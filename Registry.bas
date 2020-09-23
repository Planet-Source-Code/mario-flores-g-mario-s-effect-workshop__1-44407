Attribute VB_Name = "Registry"
 '------------------------------------------------------------------------------
 '
 'THIS MODULE IS USED TO CHECK THE INSATALLED EFFECTS THAT ARE MANIPULATED IN
 '                           MARIO'S EFFECT WORKSHOP.
 '
 '                          sistec_de_juarez@hotmail.com
 '------------------------------------------------------------------------------
 
Option Explicit
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long   'Abre un registro
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long                                                                                                                          'Cierra un registro
Public Const KEY_READ = &H20019

'Variables para hkey
Public Const HKEY_CLASSES_ROOT = &H80000000

'CHEcker Utility for TRANSITIONS MAX by MArio FLores

Public Function ViewStringRegistryValue(Llave As Long, SubKey As String) As String
    Dim retval As Long
    Dim Key As Long
    
    retval = RegOpenKeyEx(Llave, SubKey, 0, KEY_READ, Key)
    
    If retval <> 0 Then ViewStringRegistryValue = "NO"
    If retval = 0 Then ViewStringRegistryValue = "YES"
     
    retval = RegCloseKey(Key)  'Cerramos el registro
End Function
