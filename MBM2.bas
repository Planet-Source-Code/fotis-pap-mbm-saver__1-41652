Attribute VB_Name = "MBMnew"
Option Explicit

'*********************************************************************************************
'*  API Declarations to open the shared memory and to copy the contained information ...
'*********************************************************************************************

Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
Private Const FILE_MAP_READ = &H4
Private Const FILE_MAP_WRITE = &H2
Private Declare Sub CopyMemoryX Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub CopyMemoryBack Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As Any, ByVal Length As Long)

'*********************************************************************************************
'* Shared Data Type Definitions for VB
'*********************************************************************************************

Public Enum TBusType
    btISA = 0
    btSMBus = 1
    btVIA686ABus = 2
    btDirectIO = 3
End Enum

Public Enum TSMBType
    smtSMBIntel = 0
    smtSMBAMD = 1
    smtSMBALi = 2
    smtSMBNForce = 3
    smtSMBSIS = 4
End Enum

Public Enum TSensorType
    stUnknown = 0
    stTemperature = 1
    stVoltage = 2
    stFan = 3
    stMhz = 4
    stPercentage = 5
End Enum


Public Type TSharedIndex
    iType   As TSensorType                  ' type of sensor
    count   As Integer                      ' number of sensor for that type
End Type


Public Type TSharedSensor
    ssType      As Byte                     ' type of sensor
    ssName      As String * 12              ' name of sensor. array [0..11] of AnsiChar
    sspadding1  As String * 3               ' padding of 3 byte
    ssCurrent   As Double                   ' current value
    ssLow       As Double                   ' lowest readout
    ssHigh      As Double                   ' highest readout
    ssCount     As Long                     ' total number of readout
    sspadding2  As String * 4               ' padding of 4 byte
    ssTotal     As Double                   ' total amout of all readouts  -> Not Working. Don't know how to convert double to extended
    sspadding3  As String * 6               ' padding of 6 byte
    ssAlarm1    As Double                   ' temp & fan: low alarm; voltage: % off;
    ssAlarm2    As Double                   ' temp: high alarm
End Type

' This part is probably not working correctly
' I don't know what values I should read
Public Type TSharedInfo
    siSMB_Base       As Integer             ' SMBus base address
    sismb_type       As Byte                ' SMBus/Isa bus used to access chip
    siSMB_code       As Byte                ' SMBus sub type, Intel, AMD or ALi
    siSMB_Addr       As Byte                ' Address of sensor chip on SMBus
    siSMB_Name       As String * 41         ' Nice name for SMBus, array [0..40] of AnsiChar
    siISA_Base       As Integer             ' ISA base address of sensor chip on ISA
    siChipType       As Integer             ' Chip nr, connects with Chipinfo.ini
    siVoltageSubType As Byte                ' Subvoltage option selected
    sPad             As String * 4
End Type


Public Type TSharedData
    sdVersion As Double                     ' version number (example: 51090)
    sdIndex(0 To 9) As TSharedIndex         ' Sensor index
    sdSensor(0 To 99) As TSharedSensor      ' sensor info
    sdInfo    As TSharedInfo                ' misc. info
    sdStart   As String * 41                ' start time
    sdCurrent As String * 41                ' current time
    sdPath    As String * 256               ' MBM path
End Type


'*********************************************************************************************
'* Shared Data Type Definitions for VB
'*********************************************************************************************

Public Function MBM_GetData() As TSharedData

    Dim myDataStruct As TSharedData

    Dim myMBMFile As Long
    Dim myMBMMem As Long
    
    myMBMFile = OpenFileMapping(FILE_MAP_READ, False, "$M$B$M$5$S$D$")
    If myMBMFile = 0 Then
        Exit Function
    End If
    
    myMBMMem = MapViewOfFile(myMBMFile, FILE_MAP_READ, 0, 0, 0)
    CopyMemoryX myDataStruct, myMBMMem, Len(myDataStruct)
    UnmapViewOfFile myMBMMem
    
    CloseHandle myMBMFile

    MBM_GetData = myDataStruct

End Function


Public Sub MBM_SetTemp(temp As Integer, value As Integer)

    Dim myDataStruct As TSharedData

    Dim myMBMFile As Long
    Dim myMBMMem As Long
    
    myMBMFile = OpenFileMapping(FILE_MAP_WRITE, False, "$M$B$M$5$S$D$")
    If myMBMFile = 0 Then
        Exit Sub
    End If
    
    myMBMMem = MapViewOfFile(myMBMFile, FILE_MAP_WRITE, 0, 0, 0)
    CopyMemoryX myDataStruct, myMBMMem, Len(myDataStruct)
        
    'write the value
    myDataStruct.sdSensor(2).ssCurrent = value

    'copy back
    CopyMemoryBack myMBMMem, myDataStruct, Len(myDataStruct)

    UnmapViewOfFile myMBMMem
    
    CloseHandle myMBMFile

End Sub
