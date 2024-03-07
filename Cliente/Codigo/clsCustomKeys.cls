VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsCustomKeys.cls - Allows the User to Customize Keys.
'
' Creation Date: 08-03-07
' Created and Implemented by Rapsodius
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
'Customizes User Keys.
'
'@author Rapsodius
'@version 1.2.0
'@remarks None
'@date 20070803

Option Explicit

''
'Number of Customizable Keys
Private Const NUM_CUSTOM_KEYS As Byte = 27

''
'Error Raised in case the key file doesn't exist
Private Const ERROR_NO_KEY_FILE As Long = vbObjectError + 16

''
'Number of Independent Keyboard Configurations
Private Const NUM_KEY_CONFIG As Byte = 6

''
'The relative path from the application's location where the key config file is.
Private Const KEY_CONFIG_FILE As String = "\init\BindKeys.bin"

''
'Array of mapped keys
Private mappedKeys(1 To NUM_CUSTOM_KEYS, 0 To NUM_KEY_CONFIG) As Integer


''
'Index of keys on array
'
'@param mKeyUp Key for moving character up
'@param mKeyDown Key for moving character down
'@param mKeyLeft Key for moving character left
'@param mKeyRight Key for moving character right
'@param mKeyToggleMusic Key for toggling background music
'@param mKeyRequestRefresh Key for sending a request refreshing packet
'@param mKeyToggleNames Key for toggling characters' names
'@param mKeyGetObject Key for retrieving an object from the ground
'@param mKeyToggleCombatMode Key for toggling combat mode
'@param mKeyEquipObject Key for equiping an inventory object
'@param mKeyTamAnimal Key for taming animals
'@param mKeySteal Key for Stealing other users' gold or objects
'@param mKeyToggleSafeMode Key for toggling Safe Mode
'@param mKeyToggleResuscitationSafe Key for toggling Resuscitation Safe
'@param mKeyHide Key for hiding the character
'@param mKeyDropObject Key for dropping inventory objects
'@param mKeyUseObject Key for using inventory objects
'@param mKeyAttack Key for attacking other Users/NPCs
'@param mKeyTalk Key for talking (to all)
'@param mKeyTalkWithGuild Key for talking (only to guild members)
'@param mKeyToggleFPS Key for toggling the FPS counter
'@param mKeyTakeScreenShot Key for taking a snapshot of the screen
'@param mKeyShowOptions Key for showing the "Options" form
'@param mKeyMeditate Key for start/stop meditation
'@param mKeyCastSpellMacro Key for start/stop the Cast Spell Macro
'@param mKeyWorkMacro Key for start/stop the Work Macro
'@param mKeyExitGame Key for close the game
Public Enum eKeyType
    mKeyUp = 1
    mKeyDown
    mKeyLeft
    mKeyRight
    mKeyToggleMusic
    mKeyRequestRefresh
    mKeyToggleNames
    mKeyGetObject
    mKeyToggleCombatMode
    mKeyEquipObject
    mKeyTamAnimal
    mKeySteal
    mKeyToggleSafeMode
    mKeyToggleResuscitationSafe
    mKeyHide
    mKeyDropObject
    mKeyUseObject
    mKeyAttack
    mKeyTalk
    mKeyTalkWithGuild
    mKeyToggleFPS
    'mKeyTakeScreenShot
  '  mKeyShowOptions
    mKeyMeditate
    mKeyCastSpellMacro
    mKeyWorkMacro
    mKeyExitGame
End Enum

''
'Which keyboard is currently in use.
Private numKeyboard As Byte

''
'Shows a readable name for a specific Key code
'
'@param KeyCode Key code to process
'@return An User readable name for the key code or a null string if the key is not valid
'@remarks All keys not listed here are considered not valid

Public Function ReadableName(ByVal KeyCode As Integer) As String
'***************************************************
' Author: Rapsodius
' Last Modification: 08-04-07
'
'***************************************************
    Select Case KeyCode
        Case KeyCodeConstants.vbKeyA To KeyCodeConstants.vbKeyZ
            ReadableName = Chr$(KeyCode)
        Case KeyCodeConstants.vbKeyNumpad0 To KeyCodeConstants.vbKeyNumpad9
            ReadableName = KeyCode - KeyCodeConstants.vbKeyNumpad0 & " (teclado num�rico)"
        Case KeyCodeConstants.vbKeyF1 To KeyCodeConstants.vbKeyF16
            ReadableName = "F" & KeyCode - KeyCodeConstants.vbKeyF1 + 1
        Case KeyCodeConstants.vbKeyMultiply
            ReadableName = "* (teclado numerico)"
        Case KeyCodeConstants.vbKeyAdd
            ReadableName = "+ (teclado numerico)"
        Case KeyCodeConstants.vbKeySeparator
            'Note: Separator appears in MSDN as the Enter key for
            'the numpad, but MS recognizes it as the normal
            'Enter key
            ReadableName = "Enter (teclado num�rico)"
        Case KeyCodeConstants.vbKeySubtract
            ReadableName = "- (teclado numerico)"
        Case KeyCodeConstants.vbKeyDecimal
            ReadableName = ". (teclado numerico)"
        Case KeyCodeConstants.vbKeyDivide
            ReadableName = "/ (teclado numerico)"
        Case KeyCodeConstants.vbKeyShift
            ReadableName = "Shift"
        Case KeyCodeConstants.vbKeyControl
            ReadableName = "Control"
        Case KeyCodeConstants.vbKeyMenu 'Alt
            ReadableName = "Alt"
        Case KeyCodeConstants.vbKeyPause
            ReadableName = "Pausa"
        Case KeyCodeConstants.vbKeyCapital
            ReadableName = "Blq Mayus"
        Case KeyCodeConstants.vbKeyEscape
            ReadableName = "Esc"
        Case KeyCodeConstants.vbKeyPageUp
            ReadableName = "Ret Pag"
        Case KeyCodeConstants.vbKeyPageDown
            ReadableName = "Av Pag"
        Case KeyCodeConstants.vbKeyEnd
            ReadableName = "Fin"
        Case KeyCodeConstants.vbKeyHome
            ReadableName = "Inicio"
        Case KeyCodeConstants.vbKeyLeft
            ReadableName = "Izquierda"
        Case KeyCodeConstants.vbKeyUp
            ReadableName = "Arriba"
        Case KeyCodeConstants.vbKeyRight
            ReadableName = "Derecha"
        Case KeyCodeConstants.vbKeyDown
            ReadableName = "Abajo"
        Case KeyCodeConstants.vbKeyInsert
            ReadableName = "Insertar"
        Case KeyCodeConstants.vbKeyNumlock
            ReadableName = "Blq Num"
        Case KeyCodeConstants.vbKeyScrollLock
            ReadableName = "Blq Despl"
        Case KeyCodeConstants.vbKeySpace
            ReadableName = "Barra Espaciadora"
        Case KeyCodeConstants.vbKeyBack
            ReadableName = "Borrar"
     '   Case KeyCodeConstants.vbKeyReturn
         '   ReadableName = "Enter"
        Case KeyCodeConstants.vbKeyDelete
            ReadableName = "Suprimir"
        Case Else
            'In all other cases, key is considered not valid.
            ReadableName = Constants.vbNullString
    End Select
End Function

''
'Reads custom keys from a file

Public Sub LoadCustomKeys()
'**************************************************************
' LoadCustomKeys - Reads binded keys from file, raises an
' error if file doesn't exist or if there is an invalid key
' Creation Date: 08-21-07
'
' Last Modification: 09-20-07 by Rapsodius
' Added Independent keyboard configuration
'**************************************************************
    Dim i As Long
    Dim k As Long
    Dim FileNum As Integer
    FileNum = FreeFile()
    
    'Inits keyboard config #0 with default configurations
    Call InitDefaults
        
    Open App.Path & KEY_CONFIG_FILE For Binary Access Read Lock Read Write As FileNum
        Get #1, , numKeyboard
        For k = 1 To NUM_KEY_CONFIG
            For i = 1 To NUM_CUSTOM_KEYS
                Get FileNum, , mappedKeys(i, k)
                If LenB(ReadableName(mappedKeys(i, k))) = 0 Then
                    Close FileNum
                    Call Err.Raise(ERROR_NO_KEY_FILE)
                    Exit Sub
                End If
            Next i
        Next k
    Close FileNum
End Sub

''
' Saves custom keys to a file

Public Sub SaveCustomKeys()
'**************************************************************
' SaveCustomKeys - Saves binded keys to the file
'
' Creation Date: 08-04-07
' Created and Implemented by Rapsodius
'**************************************************************
    Dim FileNum As Integer
    Dim i As Long
    Dim k As Long
    FileNum = FreeFile()
    
    Open App.Path & KEY_CONFIG_FILE For Binary Access Write As FileNum
        Put FileNum, , numKeyboard
        For k = 1 To NUM_KEY_CONFIG
            For i = 1 To NUM_CUSTOM_KEYS
                Put FileNum, , mappedKeys(i, k)
            Next i
        Next k
    Close #FileNum
End Sub

''
'Gets the currently active key configuration
'
'@return The keyboard index in use

Public Property Get KeybConfig() As Byte
'**************************************************************
' KeybConfig - Retrieves the active key configuration.
'
' Creation Date: 09-20-07
' Created and Implemented by Rapsodius
'**************************************************************

    KeybConfig = numKeyboard
End Property

''
'Sets the active key configuration
'
'@param newValue Key configuration to set as active.

Public Property Let KeybConfig(ByVal NewValue As Byte)
'**************************************************************
' KeybConfig - Sets the active key configuration.
'
' Creation Date: 09-20-07
' Created and Implemented by Rapsodius
'**************************************************************
    If NewValue < 0 Or NewValue > NUM_KEY_CONFIG Then Exit Property
    
    If NewValue <> numKeyboard Then
        numKeyboard = NewValue
    End If
End Property

''
'Gets a key code
'
'@param index Index of the key code to retrieve
'@return The Key code that belongs to index

Public Property Get BindedKey(ByVal Index As eKeyType) As Integer
'***************************************************
' Author: Rapsodius
' Last Modification: 08-04-07
'
' Creation Date: 08-04-07
' Created and Implemented by Rapsodius
'**************************************************************
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
    BindedKey = mappedKeys(Index, numKeyboard)
End Property

''
'Sets a key code
'
'@param index Index of the key to change
'@param NewVal New key code

Public Property Let BindedKey(ByVal Index As eKeyType, ByVal NewVal As Integer)
'***************************************************
' Author: Rapsodius
' Creation Date: 08-05-07
'
' Last Modification: 08-14-07 by Rapsodius
' Created and Implemented by Rapsodius
'**************************************************************
    If Index < 1 Or Index > NUM_CUSTOM_KEYS Then Exit Property
    
    If LenB(ReadableName(NewVal)) = 0 Then 'If key is not valid...
        Exit Property                      'Rejects the new assignment
    End If
    
    If KeyAssigned(NewVal) Then 'Also reject it in case key is already assigned
        Exit Property
    End If
    
    mappedKeys(Index, numKeyboard) = NewVal
End Property

''
'Inits Default keys on config number 0.

Public Sub InitDefaults()
'**************************************************************
' InitDefaults - Loads default keys on config number 0.
'
' Creation Date: 09-20-07
' Created and Implemented by Rapsodius
'**************************************************************
    mappedKeys(eKeyType.mKeyUp, 0) = KeyCodeConstants.vbKeyUp
    mappedKeys(eKeyType.mKeyDown, 0) = KeyCodeConstants.vbKeyDown
    mappedKeys(eKeyType.mKeyLeft, 0) = KeyCodeConstants.vbKeyLeft
    mappedKeys(eKeyType.mKeyRight, 0) = KeyCodeConstants.vbKeyRight
    mappedKeys(eKeyType.mKeyToggleMusic, 0) = KeyCodeConstants.vbKeyM
    mappedKeys(eKeyType.mKeyRequestRefresh, 0) = KeyCodeConstants.vbKeyL
    mappedKeys(eKeyType.mKeyToggleNames, 0) = KeyCodeConstants.vbKeyN
    mappedKeys(eKeyType.mKeyGetObject, 0) = KeyCodeConstants.vbKeyA
    mappedKeys(eKeyType.mKeyToggleCombatMode, 0) = KeyCodeConstants.vbKeyC
    mappedKeys(eKeyType.mKeyEquipObject, 0) = KeyCodeConstants.vbKeyE
    mappedKeys(eKeyType.mKeyTamAnimal, 0) = KeyCodeConstants.vbKeyD
    mappedKeys(eKeyType.mKeySteal, 0) = KeyCodeConstants.vbKeyR
    mappedKeys(eKeyType.mKeyToggleSafeMode, 0) = KeyCodeConstants.vbKeyMultiply
    mappedKeys(eKeyType.mKeyToggleResuscitationSafe, 0) = KeyCodeConstants.vbKeyEnd
    mappedKeys(eKeyType.mKeyHide, 0) = KeyCodeConstants.vbKeyO
    mappedKeys(eKeyType.mKeyDropObject, 0) = KeyCodeConstants.vbKeyT
    mappedKeys(eKeyType.mKeyUseObject, 0) = KeyCodeConstants.vbKeyU
    mappedKeys(eKeyType.mKeyAttack, 0) = KeyCodeConstants.vbKeyControl
    mappedKeys(eKeyType.mKeyTalk, 0) = KeyCodeConstants.vbKeyS
    mappedKeys(eKeyType.mKeyTalkWithGuild, 0) = KeyCodeConstants.vbKeyDelete
    mappedKeys(eKeyType.mKeyToggleFPS, 0) = KeyCodeConstants.vbKeyF4
  '  mappedKeys(eKeyType.mKeyTakeScreenShot, 0) = KeyCodeConstants.vbKeyF5
   ' mappedKeys(eKeyType.mKeyShowOptions, 0) = KeyCodeConstants.vbKeyF5
    mappedKeys(eKeyType.mKeyMeditate, 0) = KeyCodeConstants.vbKeyF6
    mappedKeys(eKeyType.mKeyCastSpellMacro, 0) = KeyCodeConstants.vbKeyF7
    mappedKeys(eKeyType.mKeyWorkMacro, 0) = KeyCodeConstants.vbKeyF8
    mappedKeys(eKeyType.mKeyExitGame, 0) = KeyCodeConstants.vbKeyF12
End Sub

''
'Loads default keys

Public Sub LoadDefaults()
'**************************************************************
' LoadDefaults - Loads default key codes in case of error.
' Creation Date: 08-05-07
'
' Last Modification: 10-05-07 by Rapsodius
' Now default key loading is a direct copy from key config #0,
' which is previously initialized.
'**************************************************************
Dim k As Long
Dim i As Long

    For k = 1 To NUM_KEY_CONFIG
        For i = 1 To NUM_CUSTOM_KEYS
            mappedKeys(i, k) = mappedKeys(i, 0)
        Next i
    Next k
End Sub

''
'Returns whether a key is binded to an action
'
'@param KeyCode Key code to check
'@return True if key is mapped, false otherwise

Public Function KeyAssigned(ByVal KeyCode As Integer) As Boolean
'***************************************************
' Author: Rapsodius
' Creation Date: 08-07-07
'
' Last Modification: 09-20-07 by Rapsodius
' Created and Implemented by Rapsodius
'**************************************************************
    Dim Counter As Long
    
    KeyAssigned = False
    
    For Counter = 1 To NUM_CUSTOM_KEYS
        If mappedKeys(Counter, numKeyboard) = KeyCode Then
            KeyAssigned = True
            Exit For
        End If
    Next Counter
End Function

''
'Retrieves the number of customizable keys
'
'@return The number of customizable keys

Public Property Get count() As Byte
'***************************************************
'Author: Rapsodius
'Last Modification: 08-04-07
'
'***************************************************
    count = NUM_CUSTOM_KEYS
End Property

''
'Retrieves the number of keyboard configurations
'
'@return The number of independent keyboard configurations

Public Property Get KeyCount() As Byte
'***************************************************
'Author: Rapsodius
'Last Modification: 08-04-07
'
'***************************************************
    KeyCount = NUM_KEY_CONFIG
End Property

''
'Constructor. Loads Keys from file

Private Sub Class_Initialize()
'***************************************************
'Author: Rapsodius
'Last Modification: 08-04-07
'
'***************************************************
On Error GoTo ErrHand:
    Call LoadCustomKeys
    Exit Sub
    
ErrHand:
    If Err.Number = ERROR_NO_KEY_FILE Then
        Call LoadDefaults
        Call SaveCustomKeys
        Resume Next
    End If
End Sub

''
'Destructor. Saves Keys to file

Private Sub Class_Terminate()
'***************************************************
'Author: Rapsodius
'Last Modification: 08-04-07
'
'***************************************************
    Call SaveCustomKeys
End Sub