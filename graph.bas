Attribute VB_Name = "graph"
Option Explicit
Private lw() As Integer
Private hgh() As Integer
Private sizech(1 To 50) As Long
Private moi() As Integer 'Divider of average
Private mos() As Integer 'The average
Private mo() As Long  'temp for average
Private initialized(1 To 50) As Boolean
Private history() As Integer


Public Sub AddValue(mValue, NumOfGraphs As Integer, CurGraph As Integer, PictObj As PictureBox, UpperLimit As Long, LowerLimit As Long, MaxPerPage As Integer, ShowAverage As Boolean, ShowLow As Boolean, showHigh As Boolean, linecolor As ColorConstants, bbackcolor As ColorConstants, fontscolor As ColorConstants)
 Dim i As Integer
 Dim LnI As Integer
 If sizech(CurGraph) <> PictObj.Height + PictObj.Width + UpperLimit * LowerLimit Then initialized(CurGraph) = False
 
 PictObj.Cls

'initialization of picturebox
If initialized(CurGraph) = False Then
 sizech(CurGraph) = PictObj.Height + PictObj.Width + UpperLimit * LowerLimit
 ReDim Preserve history(0 To MaxPerPage, 0 To NumOfGraphs)
 ReDim Preserve lw(1 To NumOfGraphs)
 ReDim Preserve hgh(1 To NumOfGraphs)
 ReDim Preserve mo(1 To NumOfGraphs)
 ReDim Preserve moi(1 To NumOfGraphs)
 ReDim Preserve mos(1 To NumOfGraphs)
  PictObj.ForeColor = fontscolor
  PictObj.BackColor = bbackcolor
  PictObj.ScaleHeight = UpperLimit - LowerLimit
  PictObj.ScaleWidth = MaxPerPage
  PictObj.AutoRedraw = True
  If lw(CurGraph) = 0 Then lw(CurGraph) = UpperLimit
  If hgh(CurGraph) = 0 Then hgh(CurGraph) = LowerLimit
  initialized(CurGraph) = True
End If

 'Average calculation
If ShowAverage = True Then
  moi(CurGraph) = moi(CurGraph) + 1
  mo(CurGraph) = mo(CurGraph) + mValue
  mos(CurGraph) = Int(mo(CurGraph) / moi(CurGraph))
End If
 'Move left all the lines
  For LnI = LBound(history) To UBound(history) - 1
   history(LnI, CurGraph) = history(LnI + 1, CurGraph)
  Next LnI
 'Add the value
  history(UBound(history), CurGraph) = mValue
    
    
 'Prints the lines and the numbers
  PictObj.ForeColor = fontscolor
  PictObj.Print UpperLimit
  For i = 1 To Int(PictObj.Height / 280)
   PictObj.CurrentX = 0
   PictObj.CurrentY = i * PictObj.ScaleHeight / Int(PictObj.Height / 280)
   PictObj.Print Int(UpperLimit - i * PictObj.ScaleHeight / Int(PictObj.Height / 280))
   PictObj.Line (0, i * PictObj.ScaleHeight / Int(PictObj.Height / 280))-(PictObj.Width, i * PictObj.ScaleHeight / Int(PictObj.Height / 280)), &H8000&
  Next i
  
  'The average line
If ShowAverage = True Then
 PictObj.DrawStyle = 1
 PictObj.Line (0, UpperLimit - mos(CurGraph))-(PictObj.Width, UpperLimit - mos(CurGraph)), &HFF00FF
 PictObj.CurrentX = 20
 PictObj.CurrentY = UpperLimit - mos(CurGraph)
 PictObj.ForeColor = &HFF00FF
 PictObj.Print mos(CurGraph)
End If
'the lower and higher limit
If showHigh = True Then
 If mValue > hgh(CurGraph) Then hgh(CurGraph) = mValue
 PictObj.CurrentX = 10
 PictObj.CurrentY = UpperLimit - hgh(CurGraph)
 PictObj.ForeColor = vbRed
 PictObj.Print hgh(CurGraph)
 PictObj.DrawStyle = 2
 PictObj.Line (0, UpperLimit - hgh(CurGraph))-(PictObj.Width, UpperLimit - hgh(CurGraph)), vbRed
End If
If ShowLow = True Then
 If mValue < lw(CurGraph) Then lw(CurGraph) = mValue
 PictObj.CurrentY = UpperLimit - lw(CurGraph) - PictObj.ScaleHeight / Int(PictObj.Height / 180)
 PictObj.CurrentX = 30
 PictObj.ForeColor = &HFF0000
 PictObj.Print lw(CurGraph)
 PictObj.DrawStyle = 2
 PictObj.Line (0, UpperLimit - lw(CurGraph))-(PictObj.Width, UpperLimit - lw(CurGraph)), &HFF0000
End If
  'The main lines
    For LnI = LBound(history) To UBound(history) - 1
     PictObj.DrawStyle = 0
     PictObj.Line (LnI, UpperLimit - history(LnI, CurGraph))-(LnI + 1, UpperLimit - history(LnI + 1, CurGraph)), linecolor
    Next LnI
    
End Sub
