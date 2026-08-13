$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$output = Join-Path $root 'NIPA_2026_예산_업무앱.xlsm'
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$book = $excel.Workbooks.Add()

$db=$book.Worksheets.Item(1); $db.Name='사업 DB'
$dashboard=$book.Worksheets.Add($db); $dashboard.Name='대시보드'
$layerSheet=$book.Worksheets.Add($db); $layerSheet.Name='레이어 현황'
$hqSheet=$book.Worksheets.Add($db); $hqSheet.Name='본부별 현황'
$guideSheet=$book.Worksheets.Add($db); $guideSheet.Name='신규입력 안내'
$data=(Get-Content -Raw -Encoding UTF8 (Join-Path $root 'nipa-ai-budget-projects.json') | ConvertFrom-Json).rows
$headers=@('관리상태','사업명','세부사업명','사업코드','회계','레이어','과기정통부 소관','NIPA 본부','2026년 예산(백만원)','개요','주요내용','추진방향')
$layers=@('AI인프라','AI반도체','AI모델','지역AX','피지컬AI','글로벌 이니셔티브','AI생태계','산업전략','기관운영비')
$headquarters=@('정책기획단','SW융합본부','AI인프라본부','AI반도체본부','AI활용본부','지역AX본부','글로벌본부','경영기획본부')
$html=Get-Content -Raw -Encoding UTF8 (Join-Path $root 'index.html')
$hqJson=[regex]::Match($html,'const NIPA_HQ_PROJECTS = (\[[\s\S]*?\]);\s*const NIPA_HEADQUARTERS').Groups[1].Value
$hqRows=if($hqJson){$hqJson|ConvertFrom-Json}else{@()}; $hqMap=@{}
foreach($item in $hqRows){$name=if($item.detailTitle){$item.detailTitle}else{([string]$item.title -replace '^\(내역\)\s*','')};$parent=if($item.parentTitle){$item.parentTitle}else{$item.title};$key="$($item.code)|$parent|$name";if(!$hqMap.ContainsKey($key)){$hqMap[$key]=New-Object System.Collections.ArrayList};if($item.nipaHeadquarters -and !$hqMap[$key].Contains($item.nipaHeadquarters)){[void]$hqMap[$key].Add($item.nipaHeadquarters)}}

$db.Cells.Font.Name='맑은 고딕'; for($c=0;$c -lt $headers.Count;$c++){$db.Cells.Item(1,$c+1).Value2=$headers[$c]}
$db.Range('A1:L1').Font.Bold=$true;$db.Range('A1:L1').Font.Color=0xFFFFFF;$db.Range('A1:L1').Interior.Color=0x9E5F24;$db.Range('A1:L1').HorizontalAlignment=-4108
$r=2
foreach($item in $data){$name=if($item.detailTitle){$item.detailTitle}else{([string]$item.title -replace '^\(내역\)\s*','')};$parent=if($item.parentTitle){$item.parentTitle}else{$item.title};$key="$($item.code)|$parent|$name";$hq=if($hqMap.ContainsKey($key)){[string]::Join(' · ',[string[]]$hqMap[$key])}else{''};$main=if($item.sourceMainItems){[string]::Join("`n",[string[]]$item.sourceMainItems)}else{$item.mainPoints};$values=@('기존',$name,$parent,$item.code,$item.account,$item.layer,$item.mappedOrgName,$hq,'',$item.overview,$main,$item.direction);for($c=0;$c -lt $values.Count;$c++){$db.Cells.Item($r,$c+1).Value2=[string]$values[$c]};$db.Cells.Item($r,9).Value2=[double]$item.budget;$r++}
$lastData=$r-1;$table=$db.ListObjects.Add(1,$db.Range("A1:L$lastData"),$null,1);$table.Name='사업DB';$table.TableStyle='TableStyleMedium2';$db.Columns('A').ColumnWidth=9;$db.Columns('B').ColumnWidth=34;$db.Columns('C').ColumnWidth=38;$db.Columns('D:F').ColumnWidth=16;$db.Columns('G').ColumnWidth=48;$db.Columns('H').ColumnWidth=24;$db.Columns('I').ColumnWidth=20;$db.Columns('J:L').ColumnWidth=55;$db.Range("I2:I$lastData").NumberFormat='#,##0';$db.Range("J2:L$lastData").WrapText=$true;$db.Rows("2:$lastData").RowHeight=38;$db.Application.ActiveWindow.SplitRow=1;$db.Application.ActiveWindow.FreezePanes=$true

foreach($ws in @($dashboard,$layerSheet,$hqSheet,$guideSheet)){$ws.Cells.Font.Name='맑은 고딕';$ws.Application.ActiveWindow.DisplayGridlines=$false;$ws.Range('B2:J3').Merge();$ws.Range('B2').Font.Size=22;$ws.Range('B2').Font.Bold=$true;$ws.Range('B2').Font.Color=0x5B3712}
$dashboard.Range('B2').Value2='정보통신산업진흥원 2026 예산 업무 앱';$dashboard.Range('B5').Value2='사업 수';$dashboard.Range('B6').Formula='=COUNTA(사업DB[사업코드])';$dashboard.Range('E5').Value2='총예산(백만원)';$dashboard.Range('E6').Formula='=SUM(사업DB[2026년 예산(백만원)])';$dashboard.Range('H5').Value2='레이어 수';$dashboard.Range('H6').Value2=9
foreach($range in @('B5:C6','E5:F6','H5:I6')){$dashboard.Range($range).Interior.Color=0xF9F2EA;$dashboard.Range($range).Borders.LineStyle=1;$dashboard.Range($range).HorizontalAlignment=-4108;$dashboard.Range($range).VerticalAlignment=-4108};$dashboard.Range('B6,E6,H6').Font.Size=18;$dashboard.Range('B6,E6,H6').Font.Bold=$true
$layerSheet.Range('B2').Value2='NIPA 2026 레이어 현황';$layerSheet.Range('B6:D6').Value2=@('레이어','사업 수','예산(백만원)');$hqSheet.Range('B2').Value2='NIPA 2026 본부별 현황';$hqSheet.Range('B6:D6').Value2=@('본부','사업 수','예산(백만원)')
for($i=0;$i -lt $layers.Count;$i++){$rr=7+$i;$layerSheet.Cells.Item($rr,2).Value2=$layers[$i];$layerSheet.Cells.Item($rr,3).Formula="=COUNTIF(사업DB[레이어],B$rr)";$layerSheet.Cells.Item($rr,4).Formula="=SUMIF(사업DB[레이어],B$rr,사업DB[2026년 예산(백만원)])"}
for($i=0;$i -lt $headquarters.Count;$i++){$rr=7+$i;$hqSheet.Cells.Item($rr,2).Value2=$headquarters[$i];$hqSheet.Cells.Item($rr,3).Formula="=COUNTIF(사업DB[NIPA 본부],""*""&B$rr&""*"")";$hqSheet.Cells.Item($rr,4).Formula="=SUMIF(사업DB[NIPA 본부],""*""&B$rr&""*"",사업DB[2026년 예산(백만원)])"}
foreach($ws in @($layerSheet,$hqSheet)){$ws.Range('B6:D6').Font.Bold=$true;$ws.Range('B6:D6').Font.Color=0xFFFFFF;$ws.Range('B6:D6').Interior.Color=0x9E5F24;$ws.Columns('B').ColumnWidth=28;$ws.Columns('C:D').ColumnWidth=18;$ws.Range('D7:D20').NumberFormat='#,##0';$ws.Range('B6:D20').Borders.LineStyle=1}
$guideSheet.Range('B2').Value2='사용 안내';$guideSheet.Range('B6').Value2='사업 탐색';$guideSheet.Range('C6').Value2='검색 조건을 설정하고 행을 더블클릭하면 상세 카드가 열립니다.';$guideSheet.Range('B7').Value2='신규 등록';$guideSheet.Range('C7').Value2='홈 또는 사업탐색 화면의 신규 등록 버튼을 사용합니다.';$guideSheet.Range('B8').Value2='기존 수정';$guideSheet.Range('C8').Value2='상세 카드의 수정 버튼을 사용하거나 사업 DB 셀을 직접 편집합니다.';$guideSheet.Columns('B').ColumnWidth=22;$guideSheet.Columns('C').ColumnWidth=75;$guideSheet.Range('B6:C8').Borders.LineStyle=1;$guideSheet.Range('B6:B8').Font.Bold=$true

function Add-Control($designer, $type, $name, $caption, $left, $top, $width, $height) {
  $c = $designer.Controls.Add($type, $name, $true)
  $c.Left=$left; $c.Top=$top; $c.Width=$width; $c.Height=$height
  if ($caption -ne $null) { $c.Caption=$caption }
  try { $c.Font.Name='맑은 고딕'; $c.Font.Size=10 } catch {}
  return $c
}
function Add-Button($sheet,$name,$caption,$left,$top,$width,$macro,$color=0x245F9E) {
  $shape=$sheet.Shapes.AddShape(5,$left,$top,$width,30)
  $shape.Name=$name; $shape.TextFrame2.TextRange.Text=$caption
  $shape.TextFrame2.TextRange.Font.Name='맑은 고딕'; $shape.TextFrame2.TextRange.Font.Size=11; $shape.TextFrame2.TextRange.Font.Bold=$true
  $shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB=0xFFFFFF
  $shape.Fill.ForeColor.RGB=$color; $shape.Line.ForeColor.RGB=$color; $shape.OnAction=$macro
}

try {
  Write-Output 'STEP=UI_SHEETS'
  $detail=$book.Worksheets.Add();$detail.Name='상세카드';$detail.Cells.Font.Name='맑은 고딕';$detail.Columns('A').ColumnWidth=3;$detail.Columns('B:J').ColumnWidth=13;$detail.Range('B2:J3').Merge();$detail.Range('B2').Font.Size=20;$detail.Range('B2').Font.Bold=$true;$detail.Range('B2').Font.Color=0x5B3712;$detail.Range('B5:J5').Merge();$detail.Range('B7:J9').Merge();$detail.Range('B11:J16').Merge();$detail.Range('B18:J22').Merge();foreach($rg in @('B7:J9','B11:J16','B18:J22')){$detail.Range($rg).Interior.Color=0xF9F2EA;$detail.Range($rg).Borders.LineStyle=1;$detail.Range($rg).WrapText=$true;$detail.Range($rg).VerticalAlignment=-4160};$detail.Range('B6').Value2='개요';$detail.Range('B10').Value2='주요내용';$detail.Range('B17').Value2='추진방향';$detail.Range('B6,B10,B17').Font.Bold=$true;$detail.Application.ActiveWindow.DisplayGridlines=$false
  $editor=$book.Worksheets.Add();$editor.Name='등록수정';$editor.Cells.Font.Name='맑은 고딕';$editor.Columns('A').ColumnWidth=3;$editor.Columns('B:C').ColumnWidth=14;$editor.Columns('D:J').ColumnWidth=12;$editor.Range('B2:J3').Merge();$editor.Range('B2').Value2='사업 등록·수정';$editor.Range('B2').Font.Size=20;$editor.Range('B2').Font.Bold=$true;$editor.Range('B2').Font.Color=0x5B3712;$editLabels=@('사업명','세부사업명','사업코드','회계','레이어','과기정통부 소관','NIPA 본부','2026년 예산(백만원)','개요','주요내용','추진방향');for($i=0;$i -lt $editLabels.Count;$i++){$rr=5+$i;$editor.Range("B$rr:C$rr").Merge();$editor.Range("D$rr:J$rr").Merge();$editor.Cells.Item($rr,2).Value2=$editLabels[$i];$editor.Cells.Item($rr,2).Font.Bold=$true;$editor.Range("B$rr:J$rr").Borders.LineStyle=1;if($i -ge 8){$editor.Rows($rr).RowHeight=55;$editor.Cells.Item($rr,4).WrapText=$true}};$editor.Application.ActiveWindow.DisplayGridlines=$false
  $explorer = $book.Worksheets.Add(); $explorer.Name='사업탐색'
  $explorer.Cells.Font.Name='맑은 고딕'; $explorer.Columns('A').ColumnWidth=7; $explorer.Columns('B').ColumnWidth=38; $explorer.Columns('C').ColumnWidth=15; $explorer.Columns('D').ColumnWidth=20; $explorer.Columns('E').ColumnWidth=25; $explorer.Columns('F').ColumnWidth=18
  $explorer.Range('B2:F3').Merge(); $explorer.Range('B2').Value2='NIPA 사업 탐색'; $explorer.Range('B2').Font.Size=22; $explorer.Range('B2').Font.Bold=$true; $explorer.Range('B2').Font.Color=0x5B3712
  $explorer.Range('B4:F4').Merge(); $explorer.Range('B4').Value2='검색·레이어·본부 조건으로 사업을 찾고 행을 더블클릭하면 상세 카드가 열립니다.'; $explorer.Range('B4').Font.Color=0x8A7560
  $explorer.Range('B7').Value2='검색어'; $explorer.Range('B8').Value2=''; $explorer.Range('C7').Value2='레이어'; $explorer.Range('C8').Value2='전체'; $explorer.Range('D7').Value2='NIPA 본부'; $explorer.Range('D8').Value2='전체'
  $explorer.Range('B7:D8').Borders.LineStyle=1; $explorer.Range('B7:D7').Font.Bold=$true; $explorer.Range('B7:D7').Interior.Color=0xF9F2EA
  Add-Button $explorer 'btnSearch' '검색 적용' 600 42 100 'RefreshExplorer'
  Add-Button $explorer 'btnReset' '조건 초기화' 710 42 100 'ResetExplorer' 0x9B8060
  Add-Button $explorer 'btnNew' '신규 사업 등록' 820 42 120 'NewProject' 0x588B45
  Add-Button $explorer 'btnHome' '홈으로' 600 78 100 'GoHome' 0x9B8060
  $explorer.Application.ActiveWindow.DisplayGridlines=$false

  Write-Output 'STEP=VBA_MODULE'
  $excel.Visible=$true
  $vb = $book.VBProject
  $modulePath=Join-Path $root 'work\modNipaApp-import.bas';$moduleText=[IO.File]::ReadAllText((Join-Path $root 'modNipaApp.bas'),[Text.Encoding]::UTF8);[IO.File]::WriteAllText($modulePath,$moduleText,[Text.Encoding]::Default);$module=$vb.VBComponents.Import($modulePath)
  if($false){
  $module = $vb.VBComponents.Add(1); $module.Name='modNipaApp'
  $module.CodeModule.AddFromString(@'
Option Explicit
Public Sub GoHome(): Worksheets("대시보드").Activate: End Sub
Public Sub GoExplorer(): Worksheets("사업탐색").Activate: RefreshExplorer: End Sub
Public Sub GoLayer(): Worksheets("레이어 현황").Activate: End Sub
Public Sub GoHeadquarters(): Worksheets("본부별 현황").Activate: End Sub
Public Sub GoDatabase(): Worksheets("사업 DB").Activate: End Sub
Public Sub NewProject(): Dim ws As Worksheet: Set ws=Worksheets("등록수정"): ws.Range("A1").Value=0: ws.Range("D5:J15").ClearContents: ws.Activate: End Sub
Public Sub EditProject(ByVal sourceRow As Long)
 Dim src As Worksheet, ws As Worksheet, i As Long: Set src=Worksheets("사업 DB"): Set ws=Worksheets("등록수정"): ws.Range("A1").Value=sourceRow
 For i=0 To 10: ws.Cells(5+i,4).Value=src.Cells(sourceRow,2+i).Value: Next i: ws.Activate
End Sub
Public Sub ShowProjectCard(ByVal sourceRow As Long)
 Dim src As Worksheet, ws As Worksheet: If sourceRow<2 Then Exit Sub
 Set src=Worksheets("사업 DB"): Set ws=Worksheets("상세카드"): ws.Range("A1").Value=sourceRow: ws.Range("B2").Value=src.Cells(sourceRow,2).Value
 ws.Range("B5").Value=src.Cells(sourceRow,4).Value & "  |  " & src.Cells(sourceRow,5).Value & "  |  " & src.Cells(sourceRow,6).Value & "  |  " & src.Cells(sourceRow,8).Value & "  |  " & Format(src.Cells(sourceRow,9).Value,"#,##0") & " 백만원"
 ws.Range("B7").Value=src.Cells(sourceRow,10).Value: ws.Range("B11").Value=src.Cells(sourceRow,11).Value: ws.Range("B18").Value=src.Cells(sourceRow,12).Value: ws.Activate
End Sub
Public Sub EditCurrentProject(): EditProject CLng(Worksheets("상세카드").Range("A1").Value): End Sub
Public Sub SaveProject()
 Dim src As Worksheet, ws As Worksheet, r As Long, i As Long: Set src=Worksheets("사업 DB"): Set ws=Worksheets("등록수정"): r=CLng(ws.Range("A1").Value)
 If Trim(ws.Range("D5").Value)="" Or Trim(ws.Range("D7").Value)="" Then MsgBox "사업명과 사업코드는 필수입니다.",vbExclamation: Exit Sub
 If r=0 Then r=src.Cells(src.Rows.Count,"B").End(xlUp).Row+1: src.Cells(r,1).Value="신규"
 For i=0 To 10: src.Cells(r,2+i).Value=ws.Cells(5+i,4).Value: Next i
 MsgBox "저장했습니다.",vbInformation: RefreshExplorer: ShowProjectCard r
End Sub
Public Sub ResetExplorer()
 With Worksheets("사업탐색"): .Range("B8").Value="": .Range("C8").Value="전체": .Range("D8").Value="전체": End With
 RefreshExplorer
End Sub
Public Sub RefreshExplorer()
 Dim src As Worksheet, dst As Worksheet, lastRow As Long, outRow As Long, r As Long, q As String, layer As String, hq As String, hay As String
 Set src=Worksheets("사업 DB"): Set dst=Worksheets("사업탐색")
 q=LCase(Trim(dst.Range("B8").Value)): layer=dst.Range("C8").Value: hq=dst.Range("D8").Value
 dst.Range("A11:F1000").ClearContents: dst.Range("A11:F1000").Borders.LineStyle=xlNone
 dst.Range("A11:F11").Value=Array("원본행","사업명","사업코드","레이어","NIPA 본부","예산(백만원)")
 With dst.Range("A11:F11"): .Font.Bold=True: .Font.Color=vbWhite: .Interior.Color=RGB(36,95,158): .HorizontalAlignment=xlCenter: End With
 lastRow=src.Cells(src.Rows.Count,"B").End(xlUp).Row: outRow=12
 For r=2 To lastRow
  If Trim(src.Cells(r,2).Value)<>"" Then
   hay=LCase(src.Cells(r,2).Value & " " & src.Cells(r,3).Value & " " & src.Cells(r,4).Value)
   If (q="" Or InStr(hay,q)>0) And (layer="전체" Or src.Cells(r,6).Value=layer) And (hq="전체" Or InStr(src.Cells(r,8).Value,hq)>0) Then
    dst.Cells(outRow,1).Value=r: dst.Cells(outRow,2).Value=src.Cells(r,2).Value: dst.Cells(outRow,3).Value=src.Cells(r,4).Value
    dst.Cells(outRow,4).Value=src.Cells(r,6).Value: dst.Cells(outRow,5).Value=src.Cells(r,8).Value: dst.Cells(outRow,6).Value=src.Cells(r,9).Value: outRow=outRow+1
   End If
  End If
 Next r
 If outRow>12 Then With dst.Range("A12:F" & outRow-1): .Borders.LineStyle=xlContinuous: .RowHeight=28: .VerticalAlignment=xlCenter: End With
 dst.Columns("A").Hidden=True: dst.Range("F12:F" & Application.Max(12,outRow-1)).NumberFormat="#,##0"
End Sub
Public Sub FilterFromLayer(): With Worksheets("사업탐색"): .Range("C8").Value=ActiveCell.Value: .Range("D8").Value="전체": End With: GoExplorer: End Sub
Public Sub FilterFromHeadquarters(): With Worksheets("사업탐색"): .Range("D8").Value=ActiveCell.Value: .Range("C8").Value="전체": End With: GoExplorer: End Sub
Public Sub ShowSelectedProject(): Dim r As Long: r=ActiveCell.Row: If ActiveSheet.Name="사업탐색" And r>=12 Then ShowProjectCard CLng(Cells(r,1).Value): End Sub
'@)
  }

  if($true){
  $card = $vb.VBComponents.Add(3); $card.Name='frmProjectCard'; $card.Properties.Item('Caption').Value='NIPA 사업 상세'; $card.Properties.Item('Width').Value=620; $card.Properties.Item('Height').Value=600
  $d=$card.Designer
  $head=Add-Control $d 'Forms.Label.1' 'lblTitle' '사업명' 18 14 560 42; try{$head.Font.Size=16;$head.Font.Bold=$true}catch{}; $head.ForeColor=0x5B3712
  Add-Control $d 'Forms.Label.1' 'lblMeta' '' 18 58 560 32 | Out-Null
  $labels=@(@('개요',98),@('주요내용',198),@('추진방향',355))
  foreach($item in $labels){$l=Add-Control $d 'Forms.Label.1' ('lbl'+$item[0]) $item[0] 18 $item[1] 90 20;try{$l.Font.Bold=$true}catch{};$l.ForeColor=0x5B3712}
  $txtOverview=Add-Control $d 'Forms.TextBox.1' 'txtOverview' $null 18 120 560 68; $txtOverview.MultiLine=$true; $txtOverview.WordWrap=$true; $txtOverview.Locked=$true; $txtOverview.ScrollBars=2
  $txtMain=Add-Control $d 'Forms.TextBox.1' 'txtMain' $null 18 220 560 125; $txtMain.MultiLine=$true; $txtMain.WordWrap=$true; $txtMain.Locked=$true; $txtMain.ScrollBars=2
  $txtDirection=Add-Control $d 'Forms.TextBox.1' 'txtDirection' $null 18 377 560 100; $txtDirection.MultiLine=$true; $txtDirection.WordWrap=$true; $txtDirection.Locked=$true; $txtDirection.ScrollBars=2
  Add-Control $d 'Forms.CommandButton.1' 'btnEdit' '이 사업 수정' 350 500 110 32 | Out-Null
  Add-Control $d 'Forms.CommandButton.1' 'btnClose' '닫기' 470 500 90 32 | Out-Null
  $card.CodeModule.AddFromString(@'
Option Explicit
Public Sub LoadProject(ByVal r As Long)
 Dim ws As Worksheet: Set ws=Worksheets("사업 DB"): Me.Tag=CStr(r)
 lblTitle.Caption=ws.Cells(r,2).Value
 lblMeta.Caption=ws.Cells(r,4).Value & "  |  " & ws.Cells(r,5).Value & "  |  " & ws.Cells(r,6).Value & vbCrLf & ws.Cells(r,8).Value & "  |  " & Format(ws.Cells(r,9).Value,"#,##0") & " 백만원"
 txtOverview.Value=ws.Cells(r,10).Value: txtMain.Value=ws.Cells(r,11).Value: txtDirection.Value=ws.Cells(r,12).Value
End Sub
Private Sub btnClose_Click(): Unload Me: End Sub
Private Sub btnEdit_Click(): Dim r As Long: r=CLng(Me.Tag): Unload Me: frmProjectEdit.LoadProject r: frmProjectEdit.Show: End Sub
'@)

  $edit = $vb.VBComponents.Add(3); $edit.Name='frmProjectEdit'; $edit.Properties.Item('Caption').Value='NIPA 사업 등록·수정'; $edit.Properties.Item('Width').Value=650; $edit.Properties.Item('Height').Value=650
  $ed=$edit.Designer; $fieldNames=@('사업명','세부사업명','사업코드','회계','레이어','과기정통부 소관','NIPA 본부','2026년 예산(백만원)','개요','주요내용','추진방향'); $controlNames=@('txtName','txtParent','txtCode','cmbAccount','cmbLayer','txtOrg','txtHq','txtBudget','txtOverview','txtMain','txtDirection')
  $top=18
  for($i=0;$i -lt $fieldNames.Count;$i++){ $lab=Add-Control $ed 'Forms.Label.1' ('lab'+$i) $fieldNames[$i] 16 $top 145 20; try{$lab.Font.Bold=$true}catch{}; $type=if($controlNames[$i] -like 'cmb*'){'Forms.ComboBox.1'}else{'Forms.TextBox.1'}; $height=if($i -ge 8){55}else{23}; $ctl=Add-Control $ed $type $controlNames[$i] $null 165 $top 440 $height; if($i -ge 8){$ctl.MultiLine=$true;$ctl.WordWrap=$true;$ctl.ScrollBars=2}; $top += $height+9 }
  Add-Control $ed 'Forms.CommandButton.1' 'btnSave' '저장' 405 565 95 32 | Out-Null; Add-Control $ed 'Forms.CommandButton.1' 'btnCancel' '취소' 510 565 95 32 | Out-Null
  $edit.CodeModule.AddFromString(@'
Option Explicit
Public Sub LoadProject(ByVal r As Long)
 Dim ws As Worksheet: Set ws=Worksheets("사업 DB"): Me.Tag=CStr(r)
 cmbAccount.Clear: cmbAccount.List=Array("일반회계","특별회계","기금")
 cmbLayer.Clear: cmbLayer.List=Array("AI인프라","AI반도체","AI모델","지역AX","피지컬AI","글로벌 이니셔티브","AI생태계","산업전략","기관운영비")
 If r=0 Then Me.Caption="신규 사업 등록": Exit Sub
 Me.Caption="기존 사업 수정": txtName=ws.Cells(r,2): txtParent=ws.Cells(r,3): txtCode=ws.Cells(r,4): cmbAccount=ws.Cells(r,5): cmbLayer=ws.Cells(r,6): txtOrg=ws.Cells(r,7): txtHq=ws.Cells(r,8): txtBudget=ws.Cells(r,9): txtOverview=ws.Cells(r,10): txtMain=ws.Cells(r,11): txtDirection=ws.Cells(r,12)
End Sub
Private Sub btnSave_Click()
 Dim ws As Worksheet, r As Long: Set ws=Worksheets("사업 DB"): r=CLng(Me.Tag)
 If Trim(txtName.Value)="" Or Trim(txtCode.Value)="" Then MsgBox "사업명과 사업코드는 필수입니다.",vbExclamation: Exit Sub
 If r=0 Then r=ws.Cells(ws.Rows.Count,"B").End(xlUp).Row+1: ws.Cells(r,1).Value="신규"
 ws.Cells(r,2)=txtName: ws.Cells(r,3)=txtParent: ws.Cells(r,4)=txtCode: ws.Cells(r,5)=cmbAccount: ws.Cells(r,6)=cmbLayer: ws.Cells(r,7)=txtOrg: ws.Cells(r,8)=txtHq: ws.Cells(r,9)=Val(Replace(txtBudget,",","")): ws.Cells(r,10)=txtOverview: ws.Cells(r,11)=txtMain: ws.Cells(r,12)=txtDirection
 MsgBox "저장했습니다.",vbInformation: Unload Me: RefreshExplorer
End Sub
Private Sub btnCancel_Click(): Unload Me: End Sub
'@)
  }

  Write-Output 'STEP=SHEET_EVENTS'
  if($true){
  $sheetModule=$vb.VBComponents.Item($explorer.CodeName)
  $sheetModule.CodeModule.AddFromString(@'
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
 If Target.Row>=12 And Cells(Target.Row,1).Value<>"" Then Cancel=True: ShowProjectCard CLng(Cells(Target.Row,1).Value)
End Sub
'@)
  foreach($pair in @(@('레이어 현황','FilterFromLayer'),@('본부별 현황','FilterFromHeadquarters'))){$ws=$book.Worksheets.Item($pair[0]);$comp=$vb.VBComponents.Item($ws.CodeName);$macro=$pair[1];$comp.CodeModule.AddFromString("Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)`r`n If Target.Column=2 And Target.Row>=7 Then Cancel=True: $macro`r`nEnd Sub")}
  $thisBook=$vb.VBComponents|Where-Object{$_.Type -eq 100 -and $_.Name -like 'ThisWorkbook*'}|Select-Object -First 1;if($thisBook){$thisBook.CodeModule.AddFromString("Private Sub Workbook_Open()`r`n Worksheets(""대시보드"").Activate`r`n RefreshExplorer`r`nEnd Sub")}
  }

  Write-Output 'STEP=BUTTONS'
  $dashboardSheet=$book.Worksheets.Item('대시보드'); Add-Button $dashboardSheet 'goExplorer' '사업 탐색' 620 42 110 'GoExplorer'; Add-Button $dashboardSheet 'newProject' '신규 등록' 740 42 110 'NewProject' 0x588B45; Add-Button $dashboardSheet 'goLayer' '레이어 현황' 620 78 110 'GoLayer' 0x9B8060; Add-Button $dashboardSheet 'goHq' '본부별 현황' 740 78 110 'GoHeadquarters' 0x9B8060; Add-Button $dashboardSheet 'goDb' '사업 DB' 860 78 90 'GoDatabase' 0x9B8060
  Add-Button $explorer 'showSelected' '선택 사업 상세' 710 78 120 'ShowSelectedProject' 0x588B45
  Add-Button $layerSheet 'filterLayer' '선택 레이어 사업보기' 620 42 150 'FilterFromLayer' 0x588B45;Add-Button $hqSheet 'filterHq' '선택 본부 사업보기' 620 42 150 'FilterFromHeadquarters' 0x588B45
  $detail.Visible=2;$editor.Visible=2
  $book.Worksheets.Item('사업 DB').Visible=1
  Write-Output 'STEP=SAVE'
  $dashboard.Activate()
  $book.SaveAs($output,52)
  Write-Output 'STEP=DONE'
} finally {
  if($book){$book.Close($true)|Out-Null}; if($excel){$excel.Quit()};
  if($book){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)}; if($excel){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)}
}
Write-Output $output
