VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMdmInputParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' --- Base Properties ---
Private pName As String
Private pMode As MdmInputMode
Private pSweepType As MdmSweepType

' --- Mode Options (Table 55) ---
' For Mode = V, U
Private pPlusNode As String
Private pMinusNode As String
' For Mode = I
Private pToNode As String
Private pFromNode As String
' For Mode = P
Private pParamName As String
' For Mode = W
' Uses pPlusNode, pMinusNode
Private pPowerUnit As String ' dBm(d) or Watts(W)
Private pResistance As Double
Private pFund As Variant
' Common for V, U, I, W
Private pUnit As String
Private pCompliance As Variant

' --- Sweep Options (Table 56) ---
' For LIN, LOG, LIST, HB, SEG
Private pSweepOrder As Long
' For LIN, LOG
Private pStart As Double
Private pStop As Double
Private pNumPoints As Long
' For LIN
Private pStepSize As Double
' For LOG
Private pSweepScale As String ' D or O
Private pTotalNumPoints As Long
' For SYNC
Private pRatio As Double
Private pOffset As Double
Private pMasterSweep As String
' For LIST
Private pListValues As Variant ' Array
' For CON, HB
Private pValue As Variant
' For AC
Private pMagnitude As Double
Private pPhase As Double
' For HB
Private pOrder As Long
' For EXP
Private pInitialValue As Double
Private pPulsedValue As Double
Private pRiseDelay As Double
Private pRiseConst As Double
Private pFallDelay As Double
Private pFallConst As Double
' For PULSE, TDR
Private pDelayTime As Double
Private pRiseTime As Double
Private pFallTime As Double
Private pPulseWidth As Double
Private pPeriod As Double
' For PWL
Private pTimeValuePairs As Dictionary
Private pStartTime As Double
Private pRepeatTimes As Long
' For SFFM
Private pAmplitude As Double
Private pCarrierFrequency As Double
Private pModulIndex As Double
Private pSignalFrequency As Double
' For SIN
' Uses pOffset, pAmplitude, pDelayTime, pPhase
Private pFrequency As Double
Private pDampFactor As Double
' For TDR
' Uses pInitialValue, pPulsedValue, pDelayTime, pRiseTime, pFallTime, pPulseWidth, pPeriod, pResistance
' For SEG
Private pNumSegments As Long
' Uses pNumPoints
Private pSegmentValues As Variant ' Array

' --- Property Get/Let for Base Properties ---
Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(ByVal Value As String)
    pName = Value
End Property

Public Property Get Mode() As MdmInputMode
    Mode = pMode
End Property
Public Property Let Mode(ByVal Value As MdmInputMode)
    pMode = Value
End Property

Public Property Get SweepType() As MdmSweepType
    SweepType = pSweepType
End Property
Public Property Let SweepType(ByVal Value As MdmSweepType)
    pSweepType = Value
End Property

' --- Property Get/Let for Mode Options ---
Public Property Get PlusNode() As String
    PlusNode = pPlusNode
End Property
Public Property Let PlusNode(ByVal Value As String)
    pPlusNode = Value
End Property

Public Property Get MinusNode() As String
    MinusNode = pMinusNode
End Property
Public Property Let MinusNode(ByVal Value As String)
    pMinusNode = Value
End Property

Public Property Get ToNode() As String
    ToNode = pToNode
End Property
Public Property Let ToNode(ByVal Value As String)
    pToNode = Value
End Property

Public Property Get FromNode() As String
    FromNode = pFromNode
End Property
Public Property Let FromNode(ByVal Value As String)
    pFromNode = Value
End Property

Public Property Get ParamName() As String
    ParamName = pParamName
End Property
Public Property Let ParamName(ByVal Value As String)
    pParamName = Value
End Property

Public Property Get PowerUnit() As String
    PowerUnit = pPowerUnit
End Property
Public Property Let PowerUnit(ByVal Value As String)
    pPowerUnit = Value
End Property

Public Property Get Resistance() As Double
    Resistance = pResistance
End Property
Public Property Let Resistance(ByVal Value As Double)
    pResistance = Value
End Property

Public Property Get Fund() As Variant
    Fund = pFund
End Property
Public Property Let Fund(ByVal Value As Variant)
    pFund = Value
End Property

Public Property Get Unit() As String
    Unit = pUnit
End Property
Public Property Let Unit(ByVal Value As String)
    pUnit = Value
End Property

Public Property Get Compliance() As Variant
    Compliance = pCompliance
End Property
Public Property Let Compliance(ByVal Value As Variant)
    pCompliance = Value
End Property

' --- Property Get/Let for Sweep Options ---
Public Property Get SweepOrder() As Long
    SweepOrder = pSweepOrder
End Property
Public Property Let SweepOrder(ByVal Value As Long)
    pSweepOrder = Value
End Property

Public Property Get Start() As Double
    Start = pStart
End Property
Public Property Let Start(ByVal Value As Double)
    pStart = Value
End Property

Public Property Get Stop() As Double
    Stop = pStop
End Property
Public Property Let Stop(ByVal Value As Double)
    pStop = Value
End Property

Public Property Get NumPoints() As Long
    NumPoints = pNumPoints
End Property
Public Property Let NumPoints(ByVal Value As Long)
    pNumPoints = Value
End Property

Public Property Get StepSize() As Double
    StepSize = pStepSize
End Property
Public Property Let StepSize(ByVal Value As Double)
    pStepSize = Value
End Property

Public Property Get SweepScale() As String
    SweepScale = pSweepScale
End Property
Public Property Let SweepScale(ByVal Value As String)
    pSweepScale = Value
End Property

Public Property Get TotalNumPoints() As Long
    TotalNumPoints = pTotalNumPoints
End Property
Public Property Let TotalNumPoints(ByVal Value As Long)
    pTotalNumPoints = Value
End Property

Public Property Get Ratio() As Double
    Ratio = pRatio
End Property
Public Property Let Ratio(ByVal Value As Double)
    pRatio = Value
End Property

Public Property Get Offset() As Double
    Offset = pOffset
End Property
Public Property Let Offset(ByVal Value As Double)
    pOffset = Value
End Property

Public Property Get MasterSweep() As String
    MasterSweep = pMasterSweep
End Property
Public Property Let MasterSweep(ByVal Value As String)
    pMasterSweep = Value
End Property

Public Property Get ListValues() As Variant
    ListValues = pListValues
End Property
Public Property Let ListValues(ByVal Value As Variant)
    pListValues = Value
End Property

Public Property Get Value() As Variant
    Value = pValue
End Property
Public Property Let Value(ByVal Value As Variant)
    pValue = Value
End Property

Public Property Get Magnitude() As Double
    Magnitude = pMagnitude
End Property
Public Property Let Magnitude(ByVal Value As Double)
    pMagnitude = Value
End Property

Public Property Get Phase() As Double
    Phase = pPhase
End Property
Public Property Let Phase(ByVal Value As Double)
    pPhase = Value
End Property

Public Property Get Order() As Long
    Order = pOrder
End Property
Public Property Let Order(ByVal Value As Long)
    pOrder = Value
End Property

Public Property Get InitialValue() As Double
    InitialValue = pInitialValue
End Property
Public Property Let InitialValue(ByVal Value As Double)
    pInitialValue = Value
End Property

Public Property Get PulsedValue() As Double
    PulsedValue = pPulsedValue
End Property
Public Property Let PulsedValue(ByVal Value As Double)
    pPulsedValue = Value
End Property

Public Property Get RiseDelay() As Double
    RiseDelay = pRiseDelay
End Property
Public Property Let RiseDelay(ByVal Value As Double)
    pRiseDelay = Value
End Property

Public Property Get RiseConst() As Double
    RiseConst = pRiseConst
End Property
Public Property Let RiseConst(ByVal Value As Double)
    pRiseConst = Value
End Property

Public Property Get FallDelay() As Double
    FallDelay = pFallDelay
End Property
Public Property Let FallDelay(ByVal Value As Double)
    pFallDelay = Value
End Property

Public Property Get FallConst() As Double
    FallConst = pFallConst
End Property
Public Property Let FallConst(ByVal Value As Double)
    pFallConst = Value
End Property

Public Property Get DelayTime() As Double
    DelayTime = pDelayTime
End Property
Public Property Let DelayTime(ByVal Value As Double)
    pDelayTime = Value
End Property

Public Property Get RiseTime() As Double
    RiseTime = pRiseTime
End Property
Public Property Let RiseTime(ByVal Value As Double)
    pRiseTime = Value
End Property

Public Property Get FallTime() As Double
    FallTime = pFallTime
End Property
Public Property Let FallTime(ByVal Value As Double)
    pFallTime = Value
End Property

Public Property Get PulseWidth() As Double
    PulseWidth = pPulseWidth
End Property
Public Property Let PulseWidth(ByVal Value As Double)
    pPulseWidth = Value
End Property

Public Property Get Period() As Double
    Period = pPeriod
End Property
Public Property Let Period(ByVal Value As Double)
    pPeriod = Value
End Property

Public Property Get TimeValuePairs() As Dictionary
    Set TimeValuePairs = pTimeValuePairs
End Property
Public Property Let TimeValuePairs(ByVal Value As Dictionary)
    Set pTimeValuePairs = Value
End Property

Public Property Get StartTime() As Double
    StartTime = pStartTime
End Property
Public Property Let StartTime(ByVal Value As Double)
    pStartTime = Value
End Property

Public Property Get RepeatTimes() As Long
    RepeatTimes = pRepeatTimes
End Property
Public Property Let RepeatTimes(ByVal Value As Long)
    pRepeatTimes = Value
End Property

Public Property Get Amplitude() As Double
    Amplitude = pAmplitude
End Property
Public Property Let Amplitude(ByVal Value As Double)
    pAmplitude = Value
End Property

Public Property Get CarrierFrequency() As Double
    CarrierFrequency = pCarrierFrequency
End Property
Public Property Let CarrierFrequency(ByVal Value As Double)
    pCarrierFrequency = Value
End Property

Public Property Get ModulIndex() As Double
    ModulIndex = pModulIndex
End Property
Public Property Let ModulIndex(ByVal Value As Double)
    pModulIndex = Value
End Property

Public Property Get SignalFrequency() As Double
    SignalFrequency = pSignalFrequency
End Property
Public Property Let SignalFrequency(ByVal Value As Double)
    pSignalFrequency = Value
End Property

Public Property Get Frequency() As Double
    Frequency = pFrequency
End Property
Public Property Let Frequency(ByVal Value As Double)
    pFrequency = Value
End Property

Public Property Get DampFactor() As Double
    DampFactor = pDampFactor
End Property
Public Property Let DampFactor(ByVal Value As Double)
    pDampFactor = Value
End Property

Public Property Get NumSegments() As Long
    NumSegments = pNumSegments
End Property
Public Property Let NumSegments(ByVal Value As Long)
    pNumSegments = Value
End Property

Public Property Get SegmentValues() As Variant
    SegmentValues = pSegmentValues
End Property
Public Property Let SegmentValues(ByVal Value As Variant)
    pSegmentValues = Value
End Property

Private Sub Class_Initialize()
    ' Initialize dictionary for PWL sweep type
    Set pTimeValuePairs = New Dictionary
End Sub