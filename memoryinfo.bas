Attribute VB_Name = "memoryinfo"
'From planet-Source-code
'with some changes
Option Explicit
Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type
Private memoryinfo As MEMORYSTATUS
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Function GetMemoryInfo()
  Dim totalmem, availmem As Long
  DoEvents
  GlobalMemoryStatus memoryinfo
  totalmem = Int(memoryinfo.dwTotalPhys / 1044032 * 10 + 0.5) / 10
  availmem = Int(memoryinfo.dwAvailPhys / 1044032 * 10 + 0.5) / 10
  GetMemoryInfo = availmem & "/" & totalmem & "MB"
End Function
