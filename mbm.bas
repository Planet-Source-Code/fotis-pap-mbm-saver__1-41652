Attribute VB_Name = "modMBMAccess"
Option Explicit
Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
Private Const FILE_MAP_READ = &H4
Private Declare Sub CopyMemoryX Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Const numMBMTemperatures = 10
Private Const numMBMVoltages = 10
Private Const numMBMFans = 10
Private Const numMBMCPUs = 4
Public Type MBMSharedData1
      STemperature(1 To numMBMTemperatures) As Long     ': array [1..10]  of Integer;   // Holding the 10 possible temps
      SVoltage(1 To numMBMVoltages) As Double           ': array [1..7]   of Real;      // Holding the 7 possible voltages
      SFan(1 To numMBMFans) As Long                     ': array [1..4]   of Integer;   // Holding the 4 possible fans
      SMHZ As Long                                      ': Integer;                     // CPU freq
      SNrCPU As Byte                                    ': Byte;                        // Number of CPU's
      SCPUUsage(1 To numMBMCPUs) As Double              ': array [1..4]   of Real;
End Type
Public Type MBMSharedName1
      STempName(1 To numMBMTemperatures) As String * 11 ': array [1..10] of array [0..10] of Char;  // array 10 deep for a name 11 char long
      SVoltName(1 To numMBMVoltages) As String * 11     ': array [1..7]  of array [0..10] of Char;  // array 7  deep for a name 11 char long
      SFanName(1 To numMBMFans) As String * 11          ': array [1..4]  of array [0..10] of Char;  // array 4  deep for a name 11 char long
      SCPUName As String * 11                           ': array [0..10] of Char;                   // for name 11 char long
      SCPUUsageName As String * 11                      ': array [0..10] of Char;                   // for name 11 char long
End Type
Public Type MBMSharedInfo1
      SMBM_Version As String * 11                       ': array [0..10]  of Char;      // the version number (example: 5.1 or 5.09)
      SSMB_Base As Long                                 ': Word;                        // SMBus base address
      SSMB_Type As Integer                              ': TBus;                        // SMBus/Isa bus used to access chip
      SSMB_Code As Integer                              ': TSMB;                        // SMBus sub type, Intel based or AMD
      SSMB_Addr As Byte                                 ': Byte;                        // Address of sensor chip on SMBus
      SSMB_Name As String * 41                          ': array [0..40]  of Char;      // Nice name for SMBus
      SISA_Base As Long                                 ': Word;                        // ISA base address of sensor chip on ISA
      SChipType As Long                                 ': Integer;                     // Chip nr, connects with Chipinfo.ini
      SVoltageSubType As Byte                           ': Byte;                        // Subvoltage option selected
End Type
Public Type MBMSharedHL1                                 '                               // Avarage = A / C
    STempC(1 To numMBMTemperatures) As Long             ': array [1..10]  of LongInt;   // total number of readouds of temp
    STempA(1 To numMBMTemperatures) As String * 10      ': array [1..10]  of Extended;  // total amount of all readouts of temp
    STempL(1 To numMBMTemperatures) As Double           ': array [1..10]  of Real;      // lowest value so far of temp
    STempH(1 To numMBMTemperatures) As Double           ': array [1..10]  of Real;      // highest value so far of temp
    SVoltC(1 To numMBMVoltages) As Long                 ': array [1..7]   of LongInt;   // total number of readouts of voltage
    SVoltA(1 To numMBMVoltages) As String * 10          ': array [1..7]   of Extended;  // total amount of all readouts of voltage
    SVoltL(1 To numMBMVoltages) As Double               ': array [1..7]   of Real;      // lowest value so far of voltage
    SVoltH(1 To numMBMVoltages) As Double               ': array [1..7]   of Real;      // highest value so far of voltage
    SFanC(1 To numMBMFans) As Long                      ': array [1..4]   of LongInt;   // total number of readouds of fan
    SFanA(1 To numMBMFans) As String * 10               ': array [1..4]   of Extended;  // total amount of all readouts of fan
    SFanL(1 To numMBMFans) As Double                    ': array [1..4]   of Real;      // lowest value so far of fan
    SFanH(1 To numMBMFans) As Double                    ': array [1..4]   of Real;      // highest value so far of fan
    SStart As String * 41                               ': array [0..40]  of Char;      // starting time
    SCurrent As String * 41                             ': array [0..40]  of Char;      // current time
    SCPUC(1 To numMBMCPUs) As Long                      ': array [1..4]   of LongInt;
    SCPUA(1 To numMBMCPUs) As String * 10               ': array [1..4]   of Extended;
    SCPUL(1 To numMBMCPUs) As Double                    ': array [1..4]   of Real;
    SCPUH(1 To numMBMCPUs) As Double                    ': array [1..4]   of Real;
End Type
Public Function MBM_CheckVersion() As Boolean

    Dim myInfo As MBMSharedInfo1
    Dim retVal As Boolean
    
    myInfo = MBM_GetInfo
    If (Left(myInfo.SMBM_Version, 1)) = "V" Or (Left(myInfo.SMBM_Version, 1)) = "v" Then
            retVal = True
        ElseIf Mid(myInfo.SMBM_Version, 3, 1) > 0 Then
            retVal = True
            Else
                retVal = False
    End If

    MBM_CheckVersion = retVal
End Function
Public Function MBM_GetData1() As MBMSharedData1

    Dim myDataStruct As MBMSharedData1

    Dim myMBMFile As Long
    Dim myMBMMem As Long
    
    myMBMFile = OpenFileMapping(FILE_MAP_READ, False, "$M$B$M$5$D$")
    If myMBMFile = 0 Then
        Exit Function
    End If
    
    myMBMMem = MapViewOfFile(myMBMFile, FILE_MAP_READ, 0, 0, 0)
    CopyMemoryX myDataStruct, myMBMMem, Len(myDataStruct)
    UnmapViewOfFile myMBMMem
    
    CloseHandle myMBMFile

    MBM_GetData1 = myDataStruct

End Function

Public Function MBM_GetName() As MBMSharedName1

    Dim myDataStruct As MBMSharedName1

    Dim myMBMFile As Long
    Dim myMBMMem As Long
    
    myMBMFile = OpenFileMapping(FILE_MAP_READ, False, "$M$B$M$5$N$")
    If myMBMFile = 0 Then
         Exit Function
    End If
    
    myMBMMem = MapViewOfFile(myMBMFile, FILE_MAP_READ, 0, 0, 0)
    CopyMemoryX myDataStruct, myMBMMem, Len(myDataStruct)
    UnmapViewOfFile myMBMMem
    
    CloseHandle myMBMFile

    MBM_GetName = myDataStruct

End Function

Public Function MBM_GetHighLow() As MBMSharedHL1

    Dim myDataStruct As MBMSharedHL1

    Dim myMBMFile As Long
    Dim myMBMMem As Long
    
    myMBMFile = OpenFileMapping(FILE_MAP_READ, False, "$M$B$M$5$H$")
    If myMBMFile = 0 Then
        Exit Function
    End If
    
    myMBMMem = MapViewOfFile(myMBMFile, FILE_MAP_READ, 0, 0, 0)
    CopyMemoryX myDataStruct, myMBMMem, Len(myDataStruct)
    UnmapViewOfFile myMBMMem
    
    CloseHandle myMBMFile

    MBM_GetHighLow = myDataStruct

End Function

Public Function MBM_GetInfo() As MBMSharedInfo1

    Dim myDataStruct As MBMSharedInfo1

    Dim myMBMFile As Long
    Dim myMBMMem As Long
    
    myMBMFile = OpenFileMapping(FILE_MAP_READ, False, "$M$B$M$5$I$")
    If myMBMFile = 0 Then
         Exit Function
    End If
    
    myMBMMem = MapViewOfFile(myMBMFile, FILE_MAP_READ, 0, 0, 0)
    CopyMemoryX myDataStruct, myMBMMem, Len(myDataStruct)
    UnmapViewOfFile myMBMMem
    
    CloseHandle myMBMFile

    MBM_GetInfo = myDataStruct

End Function
