<#
.SYNOPSIS
    현재 폴더의 모든 xlsx 파일을 하나의 xlsx 파일로 병합합니다.
    이미 존재하는 merged.xlsx 또는 merged-N.xlsx 파일은 입력에서 제외하며,
    출력 파일명은 중복을 피해 자동으로 생성합니다.

.DESCRIPTION
    각 입력 파일의 데이터는 결과 파일의 별도 시트로 추가됩니다.
    시트 이름은 원본 파일명으로 변경됩니다.
#>

param (
    # 우클릭으로 넘어오는 경로($args[0])가 있으면 그걸 쓰고, 없으면 현재 위치 사용
    [string]$TargetDirectory = (Get-Location).Path
)

# 1. 작업 경로 설정
$Path = Resolve-Path $TargetDirectory
Write-Host "작업 위치: $Path" -ForegroundColor Cyan

# 2. 제외할 파일명 패턴 (merged.xlsx, merged-1.xlsx ...)
$ExcludePattern = "^merged(-\d+)?\.xlsx$"

# 3. 출력 파일명 결정 로직
$BaseName = "merged"
$Extension = ".xlsx"
$Counter = 0
$OutputFileName = "$BaseName$Extension"

while (Test-Path (Join-Path $Path $OutputFileName)) {
    $Counter++
    $OutputFileName = "$BaseName-$Counter$Extension"
}
$OutputFilePath = Join-Path $Path $OutputFileName

Write-Host "출력 예정 파일: $OutputFileName" -ForegroundColor Yellow

# 4. 병합 대상 파일 리스트 확보
# 제외 패턴에 맞는 파일은 리스트에서 제거
$SourceFiles = Get-ChildItem -Path $Path -Filter *.xlsx | 
               Where-Object { $_.Name -notmatch $ExcludePattern }

if ($SourceFiles.Count -eq 0) {
    Write-Warning "병합할 .xlsx 파일을 찾을 수 없습니다."
    exit
}

Write-Host "병합 대상: $($SourceFiles.Count)개 파일"

# 5. Excel COM Object 시작
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false # 덮어쓰기 경고 등 팝업 억제

try {
    # 결과물이 될 새 워크북 생성
    $DestWb = $Excel.Workbooks.Add()
    
    # 기본 생성된 시트(Sheet1) 갯수 파악 (나중에 삭제하기 위해)
    $InitialSheetCount = $DestWb.Worksheets.Count

    foreach ($File in $SourceFiles) {
        Write-Host "처리 중... $($File.Name)" -NoNewline

        try {
            $SourceWb = $Excel.Workbooks.Open($File.FullName)
            
            # 소스 파일의 모든 시트를 결과 파일의 맨 뒤로 복사
            # (각각 sheet으로 하라고 하셨으므로, 파일명을 시트명으로 씁니다)
            $SourceWb.Worksheets.Copy([Type]::Missing, $DestWb.Worksheets.Item($DestWb.Worksheets.Count))
            
            # 방금 복사된 시트(맨 마지막 시트)의 이름을 파일명으로 변경 시도
            # (주의: 엑셀 시트 이름은 31자 제한 및 특수문자 제한이 있어 실패할 수 있음)
            $NewSheet = $DestWb.Worksheets.Item($DestWb.Worksheets.Count)
            try {
                # 확장자 뺀 파일명을 시트 이름으로 사용
                $NewSheet.Name = $File.BaseName 
            }
            catch {
                # 이름 충돌이나 길이 제한 등으로 실패 시 경고만 출력하고 넘어감
                Write-Host " (시트명 변경 건너뜀)" -ForegroundColor DarkGray -NoNewline
            }

            $SourceWb.Close($false) # 저장하지 않고 닫기
            Write-Host " [완료]" -ForegroundColor Green
        }
        catch {
            Write-Host " [실패: $($_.Exception.Message)]" -ForegroundColor Red
            if ($SourceWb) { $SourceWb.Close($false) }
        }
    }

    # 6. 초기 빈 시트(Sheet1) 정리
    # 데이터가 있는 시트만 남기기 위해 처음에 있던 빈 시트는 삭제
    # (단, 병합된 시트가 하나도 없으면 에러나므로 체크)
    if ($DestWb.Worksheets.Count > $InitialSheetCount) {
        for ($i = 1; $i -le $InitialSheetCount; $i++) {
            $DestWb.Worksheets.Item(1).Delete()
        }
    }

    # 7. 저장
    $DestWb.SaveAs($OutputFilePath)
    $DestWb.Close($true)
    
    Write-Host "`n모든 작업이 완료되었습니다: $OutputFileName" -ForegroundColor Cyan
}
catch {
    Write-Error "스크립트 실행 중 치명적 오류 발생: $($_.Exception.Message)"
}
finally {
    # 8. 메모리 해제 (Excel 프로세스가 남지 않도록 중요)
    $Excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
