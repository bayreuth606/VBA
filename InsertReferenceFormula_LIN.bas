Attribute VB_Name = "Module1"
Sub InsertReferenceFormula_LIN()
    Dim compareSheet As Worksheet
    Dim testSheet As Worksheet
    Dim groupRange As Range
    Dim widthRange As Range
    Dim cdRange As Range
    Dim gapRange As Range
    Dim widthValue As Double
    Dim groupValue As String
    Dim matchIndex As Long
    Dim k As Long
    
    On Error GoTo ErrorHandler ' 오류 핸들링 추가
    
    ' 시트 정의
    Set compareSheet = ThisWorkbook.Sheets("Compare_CD")
    Set testSheet = ThisWorkbook.Sheets("test")
    
    ' Compare_CD의 데이터 범위를 열별로 설정
    Set groupRange = compareSheet.Range("A2:A131") ' 그룹(A열)
    Set widthRange = compareSheet.Range("B2:B131") ' WIDTH(B열)
    Set cdRange = compareSheet.Range("D2:D131")    ' CD(C열)
    Set gapRange = compareSheet.Range("C2:C131")   ' GAP(H열)

    ' Test 시트 데이터 채우기
    Dim i As Long, j As Long
    For i = 2 To 5 ' A, B, C 그룹 반복 (Test 시트의 X)
        For j = 2 To 11 ' WIDTH 값 반복 (Test 시트의 Y)
            ' Test 시트에서 기준 값 가져오기
            widthValue = testSheet.Cells(j, 1).Value ' WIDTH 값 (A열)
            groupValue = testSheet.Cells(1, i).Value ' 그룹 값 (첫 번째 행)
            
            ' Compare_CD 시트에서 조건에 맞는 값 찾기
            matchIndex = 0
            For k = 1 To groupRange.Rows.Count
                ' 그룹이 'LS'일 경우 추가 조건 확인
                If groupValue = "LS" Then
                    If groupRange.Cells(k, 1).Value = groupValue And _
                       gapRange.Cells(k, 1).Value = widthRange.Cells(k, 1).Value And _
                       widthRange.Cells(k, 1).Value = widthValue Then
                        matchIndex = k ' 인덱스 저장
                        Exit For
                    End If
                Else
                    ' 일반 조건
                    If groupRange.Cells(k, 1).Value = groupValue And _
                       widthRange.Cells(k, 1).Value = widthValue Then
                        matchIndex = k ' 인덱스 저장
                        Exit For
                    End If
                End If
            Next k
            
            ' 결과를 Test 시트에 수식으로 삽입
            If matchIndex > 0 Then
                ' cdRange 활용하여 수식을 생성
                testSheet.Cells(j, i).Formula = "=Compare_CD!" & cdRange.Cells(matchIndex, 1).Address(External:=False)
            Else
                testSheet.Cells(j, i).Value = "N/A" ' 값이 없을 경우 N/A로 표시
            End If
        Next j
    Next i
    
    MsgBox "수식이 성공적으로 삽입되었습니다!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "오류 발생: " & Err.Description, vbCritical
End Sub
