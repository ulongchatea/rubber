# rubber

Sub 머릿말_내용_지우기()
    ' 현재 문서의 머릿말(Primary Header) 내용을 지웁니다.
    ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = ""

    ' 현재 문서의 바닥글(Primary Footer) 내용을 지웁니다.
    ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).Range.Text = ""

    ' 현재 문서의 페이지 설정 변경
    With ActiveDocument.PageSetup
        ' 머리글(위쪽) 설정 (3cm)
        .TopMargin = CentimetersToPoints(3)
        ' 바닥글(아래쪽) 설정 (0.5cm)
        .BottomMargin = CentimetersToPoints(0.5)
    End With

    ' 현재 문서의 페이지 레이아웃 설정 변경
    With ActiveDocument.PageSetup
        ' 머리글의 위쪽 여백을 3cm로 설정
        .HeaderDistance = CentimetersToPoints(3)
    End With
    ' 현재 문서의 페이지 레이아웃 설정 변경
    With ActiveDocument.PageSetup
        ' 바닥글의 아래쪽 여백을 0.5cm로 설정
        .FooterDistance = CentimetersToPoints(0.5)
    End With
    
        ' 현재 문서의 페이지 레이아웃 설정 변경
    With ActiveDocument.PageSetup
        ' 왼쪽 여백을 3cm로 설정
        .LeftMargin = CentimetersToPoints(3)
        ' 오른쪽 여백을 3cm로 설정
        .RightMargin = CentimetersToPoints(3)
    End With
    
    
End Sub

Sub 특정_텍스트_변경_및_검색_셀_오른쪽에_A_입력()
    Dim rng As Range
    Dim findRange As Range
    Dim currentCell As Cell
    Dim nextCellRange As Range
    Dim userInput As String

    userInput = InputBox("값을 입력하세요:", "사용자 입력")

    ' 현재 문서의 전체 범위에서 '제 조 번 호'를 'abcde'로 변경
    Set rng = ActiveDocument.Range
    rng.Find.ClearFormatting
    rng.Find.Replacement.ClearFormatting
    rng.Find.Text = "제 조 번 호"
    rng.Find.Replacement.Text = "abcde"
    rng.Find.Execute Replace:=wdReplaceAll

    ' 'abcde'를 검색한 후 해당 위치의 오른쪽 셀로 이동하고 'A' 입력
    Set findRange = ActiveDocument.Range
    findRange.Find.ClearFormatting
    findRange.Find.Text = "abcde"
    Do While findRange.Find.Execute
        If findRange.Find.Found Then
            ' 'abcde'를 찾은 경우
            Set currentCell = findRange.Cells(1)
            
            ' 오른쪽 셀의 Range 찾기
            Set nextCellRange = currentCell.Range.Next(wdCell)
            
            If Not nextCellRange Is Nothing Then
                ' 오른쪽 셀에 'A' 입력
                nextCellRange.Text = userInput
            End If
        End If
    Loop

    ' 'abcde'를 다시 '제 조 번 호'로 변경
    Set rng = ActiveDocument.Range
    rng.Find.ClearFormatting
    rng.Find.Replacement.ClearFormatting
    rng.Find.Text = "abcde"
    rng.Find.Replacement.Text = "제 조 번 호"
    rng.Find.Execute Replace:=wdReplaceAll
End Sub


Sub ReplaceAndPrint()
    Dim rng As Range
    Dim findRange As Range
    Dim currentCell As Cell
    Dim nextCellRange As Range
    Dim userInput As String
    

    '''''''''ㅇ'''''''''


    userInput = InputBox("값을 입력하세요:", "사용자 입력")

    ' 현재 문서의 전체 범위에서 '제 조 번 호'를 'abcde'로 변경
    Set rng = ActiveDocument.Range
    rng.Find.ClearFormatting
    rng.Find.Replacement.ClearFormatting
    rng.Find.Text = "제품명(품목명)"
    rng.Find.Replacement.Text = "abcde"
    rng.Find.Execute Replace:=wdReplaceAll

    ' 'abcde'를 검색한 후 해당 위치의 오른쪽 셀로 이동하고 'A' 입력
    Set findRange = ActiveDocument.Range
    findRange.Find.ClearFormatting
    findRange.Find.Text = "abcde"
    Do While findRange.Find.Execute
        If findRange.Find.Found Then
            ' 'abcde'를 찾은 경우
            Set currentCell = findRange.Cells(1)
            
            ' 오른쪽 셀의 Range 찾기
            Set nextCellRange = currentCell.Range.Next(wdCell)
            
            If Not nextCellRange Is Nothing Then
                ' 오른쪽 셀에 'A' 입력
                nextCellRange.Text = userInput
            End If
        End If
    Loop

    ' 'abcde'를 다시 '제 조 번 호'로 변경
    Set rng = ActiveDocument.Range
    rng.Find.ClearFormatting
    rng.Find.Replacement.ClearFormatting
    rng.Find.Text = "abcde"
    rng.Find.Replacement.Text = "제품명(품목명)"
    rng.Find.Execute Replace:=wdReplaceAll
    
    
    
    '''''''''ㅇ'''''''''
    
    
        ' 창을 통해 바꿀 단어 입력 받기
    Dim targetWord As String
    targetWord = InputBox("바꿀 단어를 입력하세요:", "단어 입력", "바꿀 단어")
    
    ' 창을 통해 변경할 글자 크기 입력 받기
    Dim replacementSize As String
    replacementSize = InputBox("변경할 글자 크기를 입력하세요:", "글자 크기 입력", "12")
    
    ' 문서 전체에서 단어 찾기
    Selection.Find.ClearFormatting
    Selection.Find.Text = targetWord
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = Val(replacementSize)
    With Selection.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    ' 단어 바꾸기
    Selection.Find.Execute Replace:=wdReplaceAll
    
    
    
'''''''''ㅇ'''''''''

    
    ' 문서 인쇄
    ActiveDocument.PrintOut
          
    
End Sub

Sub ReplaceWordInDocument()
    ' 창을 통해 바꿀 단어 입력 받기
    Dim targetWord As String
    targetWord = InputBox("바꿀 단어를 입력하세요:", "단어 입력", "바꿀 단어")
    
    ' 창을 통해 변경할 글자 크기 입력 받기
    Dim replacementSize As String
    replacementSize = InputBox("변경할 글자 크기를 입력하세요:", "글자 크기 입력", "12")
    
    ' 문서 전체에서 단어 찾기
    Selection.Find.ClearFormatting
    Selection.Find.Text = targetWord
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.Replacement.Font.Size = Val(replacementSize)
    With Selection.Find
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    
    ' 단어 바꾸기
    Selection.Find.Execute Replace:=wdReplaceAll
End Sub

