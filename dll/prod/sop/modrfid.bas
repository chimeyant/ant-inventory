Attribute VB_Name = "modrfid"
'===========================================================
'   Disposition
'===========================================================
Global Const SCARD_LEAVE_CARD = 0 ' Don't do anything special on close
Global Const SCARD_RESET_CARD = 1 ' Reset the card on close
Global Const SCARD_UNPOWER_CARD = 2 ' Power down the card on close
Global Const SCARD_EJECT_CARD = 3 ' Eject the card on close
'===========================================================
Global Const SCARD_S_SUCCESS = 0
Global Const SCARD_SHARE_SHARED = 2 ' This application is willing to share this
                                ' card with other applications.
Global Const SCARD_PROTOCOL_T0 = &H1                  ' T=0 is the active protocol.
Global Const SCARD_PROTOCOL_T1 = &H2                  ' T=1 is the active protocol.
Global Const SCARD_SCOPE_USER = 0 ' The context is a user context, and any
                                  ' database operations are performed within the
                                  ' domain of the user.

Public Declare Function SCardEstablishContext Lib "winscard.dll" (ByVal dwScope As Long, _
                                                                  ByVal pvReserved1 As Long, _
                                                                  ByVal pvReserved2 As Long, _
                                                                  ByRef phContext As Long) As Long
                                                                  
Public Declare Function SCardListReaders Lib "winscard.dll" Alias "SCardListReadersA" (ByVal hContext As Long, _
                                                            ByVal mzGroup As String, _
                                                            ByVal ReaderList As String, _
                                                            ByRef pcchReaders As Long) As Long
                                                                  
Public Declare Function SCardConnect Lib "winscard.dll" Alias "SCardConnectA" (ByVal hContext As Long, _
                                                                               ByVal szReaderName As String, _
                                                                               ByVal dwShareMode As Long, _
                                                                               ByVal dwPrefProtocol As Long, _
                                                                               ByRef hCard As Long, _
                                                                               ByRef ActiveProtocol As Long) As Long
                                                                  
Public Declare Function SCardTransmit Lib "winscard.dll" (ByVal hCard As Long, _
                                                          pioSendRequest As SCARD_IO_REQUEST, _
                                                          ByRef SendBuff As Byte, _
                                                          ByVal SendBuffLen As Long, _
                                                          ByRef pioRecvRequest As SCARD_IO_REQUEST, _
                                                          ByRef RecvBuff As Byte, _
                                                          ByRef RecvBuffLen As Long) As Long
                                                                  
Public Declare Function SCardReleaseContext Lib "winscard.dll" (ByVal hContext As Long) As Long
                                                                  
Public Declare Function SCardDisconnect Lib "winscard.dll" (ByVal hCard As Long, _
                                                            ByVal Disposistion As Long) As Long
                                                                  

Public Type SCARD_IO_REQUEST
    dwProtocol As Long
    cbPciLength As Long
End Type

Public Sub LoadListToControl(ByVal Ctrl As ComboBox, ByVal ReaderList As String)
    Dim sTemp As String
    Dim indx As Integer

    indx = 1
    sTemp = ""
    Ctrl.Clear
    
    While (Mid(ReaderList, indx, 1) <> vbNullChar)
        While (Mid(ReaderList, indx, 1) <> vbNullChar)
           sTemp = sTemp + Mid(ReaderList, indx, 1)
           indx = indx + 1
        Wend
        
        indx = indx + 1
        Ctrl.AddItem sTemp
        sTemp = ""
    Wend
End Sub
                                                                  
Public Function GetScardErrMsg(ByVal ReturnCode As Long) As String
  Select Case ReturnCode
    Case SCARD_E_CANCELLED
    GetScardErrMsg = "The action was canceled by an SCardCancel request."
    Case SCARD_E_CANT_DISPOSE
    GetScardErrMsg = "The system could not dispose of the media in the requested manner."
    Case SCARD_E_CARD_UNSUPPORTED
    GetScardErrMsg = "The smart card does not meet minimal requirements for support."
    Case SCARD_E_DUPLICATE_READER
    GetScardErrMsg = "The reader driver didn't produce a unique reader name."
    Case SCARD_E_INSUFFICIENT_BUFFER
    GetScardErrMsg = "The data buffer for returned data is too small for the returned data."
    Case SCARD_E_INVALID_ATR
    GetScardErrMsg = "An ATR string obtained from the registry is not a valid ATR string."
    Case SCARD_E_INVALID_HANDLE
    GetScardErrMsg = "The supplied handle was invalid."
    Case SCARD_E_INVALID_PARAMETER
    GetScardErrMsg = "One or more of the supplied parameters could not be properly interpreted."
    Case SCARD_E_INVALID_TARGET
    GetScardErrMsg = "Registry startup information is missing or invalid."
    Case SCARD_E_INVALID_VALUE
    GetScardErrMsg = "One or more of the supplied parameter values could not be properly interpreted."
    Case SCARD_E_NOT_READY
    GetScardErrMsg = "The reader or card is not ready to accept commands."
    Case SCARD_E_NOT_TRANSACTED
    GetScardErrMsg = "An attempt was made to end a non-existent transaction."
    Case SCARD_E_NO_MEMORY
    GetScardErrMsg = "Not enough memory available to complete this command."
    Case SCARD_E_NO_SERVICE
    GetScardErrMsg = "The smart card resource manager is not running."
    Case SCARD_E_NO_SMARTCARD
    GetScardErrMsg = "The operation requires a smart card, but no smart card is currently in the device."
    Case SCARD_E_PCI_TOO_SMALL
    GetScardErrMsg = "The PCI receive buffer was too small."
    Case SCARD_E_PROTO_MISMATCH
    GetScardErrMsg = "The requested protocols are incompatible with the protocol currently in use with the card."
    Case SCARD_E_READER_UNAVAILABLE
    GetScardErrMsg = "The specified reader is not currently available for use."
    Case SCARD_E_READER_UNSUPPORTED
    GetScardErrMsg = "The reader driver does not meet minimal requirements for support."
    Case SCARD_E_SERVICE_STOPPED
    GetScardErrMsg = "The smart card resource manager has shut down."
    Case SCARD_E_SHARING_VIOLATION
    GetScardErrMsg = "The smart card cannot be accessed because of other outstanding connections."
    Case SCARD_E_SYSTEM_CANCELLED
    GetScardErrMsg = "The action was canceled by the system, presumably to log off or shut down."
    Case SCARD_E_TIMEOUT
    GetScardErrMsg = "The user-specified timeout value has expired."
    Case SCARD_E_UNKNOWN_CARD
    GetScardErrMsg = "The specified smart card name is not recognized."
    Case SCARD_E_UNKNOWN_READER
    GetScardErrMsg = "The specified reader name is not recognized."
    Case SCARD_F_COMM_ERROR
    GetScardErrMsg = "An internal communications error has been detected."
    Case SCARD_F_INTERNAL_ERROR
    GetScardErrMsg = "An internal consistency check failed."
    Case SCARD_F_UNKNOWN_ERROR
    GetScardErrMsg = "An internal error has been detected, but the source is unknown."
    Case SCARD_F_WAITED_TOO_LONG
    GetScardErrMsg = "An internal consistency timer has expired."
    Case SCARD_S_SUCCESS
    GetScardErrMsg = "No error was encountered."
    Case SCARD_W_REMOVED_CARD
    GetScardErrMsg = "The smart card has been removed, so that further communication is not possible."
    Case SCARD_W_RESET_CARD
    GetScardErrMsg = "The smart card has been reset, so any shared state information is invalid."
    Case SCARD_W_UNPOWERED_CARD
    GetScardErrMsg = "Power has been removed from the smart card, so that further communication is not possible."
    Case SCARD_W_UNRESPONSIVE_CARD
    GetScardErrMsg = "The smart card is not responding to a reset."
    Case SCARD_W_UNSUPPORTED_CARD
    GetScardErrMsg = "The reader cannot communicate with the card, due to ATR string configuration conflicts."
    Case Else
    GetScardErrMsg = "Device is not connected or The card is not on the device!" '"?"
    End Select
End Function



