Attribute VB_Name = "modNipaApp"
Option Explicit
Public Sub GoHome(): Worksheets("대시보드").Activate: End Sub
Public Sub GoExplorer(): Worksheets("사업탐색").Activate: RefreshExplorer: End Sub
Public Sub GoLayer(): Worksheets("레이어 현황").Activate: End Sub
Public Sub GoHeadquarters(): Worksheets("본부별 현황").Activate: End Sub
Public Sub GoDatabase(): Worksheets("사업 DB").Activate: End Sub
Public Sub NewProject(): frmProjectEdit.LoadProject 0: frmProjectEdit.Show: End Sub
Public Sub EditProject(ByVal sourceRow As Long): frmProjectEdit.LoadProject sourceRow: frmProjectEdit.Show: End Sub
Public Sub ShowProjectCard(ByVal sourceRow As Long): If sourceRow>=2 Then frmProjectCard.LoadProject sourceRow: frmProjectCard.Show
End Sub
Public Sub PrepareProjectCard(ByVal sourceRow As Long): If sourceRow>=2 Then frmProjectCard.LoadProject sourceRow
End Sub
Public Sub ResetExplorer(): With Worksheets("사업탐색"): .Range("B8").Value="": .Range("C8").Value="전체": .Range("D8").Value="전체": End With: RefreshExplorer: End Sub
Public Sub RefreshExplorer()
 Dim src As Worksheet, dst As Worksheet, lastRow As Long, outRow As Long, r As Long, q As String, layer As String, hq As String, hay As String
 Set src=Worksheets("사업 DB"): Set dst=Worksheets("사업탐색"): q=LCase(Trim(dst.Range("B8").Value)): layer=dst.Range("C8").Value: hq=dst.Range("D8").Value
 dst.Range("A11:F1000").ClearContents: dst.Range("A11:F1000").Borders.LineStyle=xlNone: dst.Range("A11:F11").Value=Array("원본행","사업명","사업코드","레이어","NIPA 본부","예산(백만원)")
 With dst.Range("A11:F11"): .Font.Bold=True: .Font.Color=vbWhite: .Interior.Color=RGB(36,95,158): .HorizontalAlignment=xlCenter: End With
 lastRow=src.Cells(src.Rows.Count,"B").End(xlUp).Row: outRow=12
 For r=2 To lastRow
  If Trim(src.Cells(r,2).Value)<>"" Then
   hay=LCase(src.Cells(r,2).Value & " " & src.Cells(r,3).Value & " " & src.Cells(r,4).Value)
   If (q="" Or InStr(hay,q)>0) And (layer="전체" Or src.Cells(r,6).Value=layer) And (hq="전체" Or InStr(src.Cells(r,8).Value,hq)>0) Then
    dst.Cells(outRow,1).Value=r: dst.Cells(outRow,2).Value=src.Cells(r,2).Value: dst.Cells(outRow,3).Value=src.Cells(r,4).Value: dst.Cells(outRow,4).Value=src.Cells(r,6).Value: dst.Cells(outRow,5).Value=src.Cells(r,8).Value: dst.Cells(outRow,6).Value=src.Cells(r,9).Value: outRow=outRow+1
   End If
  End If
 Next r
 If outRow>12 Then With dst.Range("A12:F" & outRow-1): .Borders.LineStyle=xlContinuous: .RowHeight=28: .VerticalAlignment=xlCenter: End With
 dst.Columns("A").Hidden=True: dst.Range("F12:F" & Application.Max(12,outRow-1)).NumberFormat="#,##0"
End Sub
Public Sub FilterFromLayer(): With Worksheets("사업탐색"): .Range("C8").Value=ActiveCell.Value: .Range("D8").Value="전체": End With: GoExplorer: End Sub
Public Sub FilterFromHeadquarters(): With Worksheets("사업탐색"): .Range("D8").Value=ActiveCell.Value: .Range("C8").Value="전체": End With: GoExplorer: End Sub
Public Sub ShowSelectedProject()
 Dim r As Long: If ActiveSheet.Name<>"사업탐색" Then Exit Sub: r=ActiveCell.Row
 If r<12 Or Cells(r,1).Value="" Then MsgBox "목록에서 확인할 사업 행을 먼저 선택하세요.",vbInformation: Exit Sub
 ShowProjectCard CLng(Cells(r,1).Value)
End Sub
