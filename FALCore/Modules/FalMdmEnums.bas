Attribute VB_Name = "FalMdmEnums"

Option Explicit

'@brief Defines the valid modes for an ICCAP Input parameter.
Public Enum MdmInputMode
    mdmModeUnknown = 0
    mdmModeV
    mdmModeI
    mdmModeF
    mdmModeT
    mdmModeP
    mdmModeU
    mdmModeW
End Enum

'@brief Defines the valid types for an ICCAP Output parameter.
Public Enum MdmOutputType
    mdmTypeUnknown = 0
    mdmTypeV
    mdmTypeI
    mdmTypeC
    mdmTypeG
    mdmTypeT
    mdmTypeS
    mdmTypeH
    mdmTypeZ
    mdmTypeY
    mdmTypeK
    mdmTypeA
    mdmTypeN
    mdmTypeU
End Enum

'@brief Defines the valid sweep types for an ICCAP Input parameter.
Public Enum MdmSweepType
    mdmSweepUnknown = 0
    mdmSweepLIN
    mdmSweepLOG
    mdmSweepSYNC
    mdmSweepLIST
    mdmSweepCON
    mdmSweepAC
    mdmSweepHB
    mdmSweepEXP
    mdmSweepPULSE
    mdmSweepPWL
    mdmSweepSFFM
    mdmSweepSIN
    mdmSweepTDR
    mdmSweepSEG
End Enum

'@brief Defines the types of validation errors that can occur.
Public Enum MdmValidationErrorType
    errTypeGeneric = 0
    errTypeRowCountMismatch
    errTypeUndefinedHeader
    errTypeSweepNumPointsMismatch
    errTypeSweepStartMismatch
    errTypeSweepStopMismatch
    errTypeMissingSweepData
End Enum