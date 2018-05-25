
'version 1.1 zhenglei 20180524
Sub ranking()
    Dim fullName As String
   fullName = ActiveSheet.Name
   If (Trim(fullName)) = "" Then
   ' msg ("聯賽名称取得ＥＲＲＯＲ")
         ' Exit Sub
   End If
    
    sheetname = Left(fullName, InStr(fullName, "積分榜") - 1)
      rankingNow (sheetname)
End Sub



Function rankingNow(ByVal leagueName As String)

Dim teamNo As Integer
Dim toRoundNo As Integer
Dim allRoundNo As Integer
Dim circleNo As Integer
Dim linePerRound As Integer
Dim stLine As Integer
Dim edLine As Integer
Dim nameArray() As String
Dim winArray() As Integer
Dim drawArray() As Integer
Dim loseArray() As Integer
Dim goalArray() As Integer
Dim beGoaledArray() As Integer
Dim goalwinArray() As Integer
Dim pntArray() As Integer
Dim indexArray() As Integer

Dim curLine As Integer

Dim homeTeam As String
Dim awayTeam As String
Dim homePoint As Integer
Dim awayPoint As Integer
Dim eleCount As Integer


Dim ind1 As Integer
Dim ind2 As Integer

Dim cstTeam As Integer
Dim cstWin As Integer
Dim cstDraw As Integer
Dim cstLose As Integer
Dim cstGoal As Integer
Dim cstGoaled As Integer
Dim cstGoalWin As Integer
Dim cstPoint As Integer

cstTeam = 2

cstCount = 3
cstWin = 4
cstDraw = 5
cstLose = 6
cstGoal = 7
cstGoaled = 8
cstGoalWin = 9
cstPoint = 10


 Set dict = CreateObject("Scripting.Dictionary")

eleCount = 1

teamNo = Cells(21, 2)
circleNo = Cells(22, 2)

allRoundNo = (is_even(teamNo) - 1) * circleNo
'allRoundNo = 4

toRoundNo = Cells(20, 2)
If Trim(toRoundNo) = "" Or Trim(toRoundNo) = 0 Then
toRoundNo = allRoundNo
End If
linePerRound = 1 + is_even(teamNo) / 2
'linePerRound = 5

stLine = 1
edLine = linePerRound * toRoundNo

ActiveWorkbook.Sheets(leagueName).Select

ReDim nameArray(1 To teamNo)
ReDim winArray(1 To teamNo)
ReDim drawArray(1 To teamNo)
ReDim loseArray(1 To teamNo)
ReDim goalArray(1 To teamNo)
ReDim beGoaledArray(1 To teamNo)
ReDim goalwinArray(1 To teamNo)
ReDim pntArray(1 To teamNo)
ReDim indexArray(1 To teamNo)

' dict.Add nameArray(i), i

For i = 1 To teamNo
         winArray(i) = 0
         drawArray(i) = 0
         loseArray(i) = 0
         goalArray(i) = 0
         beGoaledArray(i) = 0
         goalwinArray(i) = 0
         pntArray(i) = 0
        indexArray(i) = i
Next i

For i = 1 To toRoundNo
    For j = 2 To linePerRound
       curLine = linePerRound * (i - 1) + j

       homeTeam = Cells(curLine, 1).Value
       awayTeam = Cells(curLine, 2).Value
       homePoint = Cells(curLine, 3).Value
       awayPoint = Cells(curLine, 4).Value
           
      ' If Trim(homeTeam) = "" Or Trim(awayTeam) = "" Then
      '     Exit For
      ' End If
        
       If Trim(homeTeam) = "" Then
       ElseIf dict.exists(homeTeam) Then
        ind1 = dict.Item(homeTeam)
       Else
        dict.Add homeTeam, eleCount
        ind1 = eleCount
        eleCount = eleCount + 1
        nameArray(ind1) = homeTeam
       
       End If
       
       If Trim(awayTeam) = "" Then
       ElseIf dict.exists(awayTeam) Then
        ind2 = dict.Item(awayTeam)
       Else
        dict.Add awayTeam, eleCount
        ind2 = eleCount
        eleCount = eleCount + 1
        nameArray(ind2) = awayTeam
       End If
       
       If Cells(curLine, 3).Value = "" Or Cells(curLine, 4).Value = "" Then
       
       Else
            If homePoint - awayPoint > 0 Then
            
             winArray(ind1) = winArray(ind1) + 1
             loseArray(ind2) = loseArray(ind2) + 1
             goalArray(ind1) = goalArray(ind1) + homePoint
             goalArray(ind2) = goalArray(ind2) + awayPoint
             beGoaledArray(ind1) = beGoaledArray(ind1) + awayPoint
             beGoaledArray(ind2) = beGoaledArray(ind2) + homePoint
             goalwinArray(ind1) = goalwinArray(ind1) + homePoint - awayPoint
             goalwinArray(ind2) = goalwinArray(ind2) + awayPoint - homePoint
            
             pntArray(ind1) = pntArray(ind1) + 3
            
            ElseIf homePoint - awayPoint = 0 Then
                 drawArray(ind1) = drawArray(ind1) + 1
                 drawArray(ind2) = drawArray(ind2) + 1
                 
                 goalArray(ind1) = goalArray(ind1) + homePoint
                 goalArray(ind2) = goalArray(ind2) + awayPoint
                 beGoaledArray(ind1) = beGoaledArray(ind1) + awayPoint
                 beGoaledArray(ind2) = beGoaledArray(ind2) + homePoint
            
                 pntArray(ind1) = pntArray(ind1) + 1
                 pntArray(ind2) = pntArray(ind2) + 1
            Else
                 winArray(ind2) = winArray(ind2) + 1
                 loseArray(ind1) = loseArray(ind1) + 1
                 goalArray(ind1) = goalArray(ind1) + homePoint
                 goalArray(ind2) = goalArray(ind2) + awayPoint
                 beGoaledArray(ind1) = beGoaledArray(ind1) + awayPoint
                 beGoaledArray(ind2) = beGoaledArray(ind2) + homePoint
                goalwinArray(ind1) = goalwinArray(ind1) + homePoint - awayPoint
                goalwinArray(ind2) = goalwinArray(ind2) + awayPoint - homePoint
                 pntArray(ind2) = pntArray(ind2) + 3
            End If
       End If
    Next j
Next i

    ownName = leagueName + "積分榜"
    ActiveWorkbook.Sheets(ownName).Select


'積分整理完了で、ランキングをする

For j = 0 To teamNo - 2
    For i = 1 To teamNo - 1 - j
    
    If pntArray(indexArray(i)) < pntArray(indexArray(i + 1)) Then
        temp = indexArray(i)
        indexArray(i) = indexArray(i + 1)
        indexArray(i + 1) = temp
    ElseIf pntArray(indexArray(i)) = pntArray(indexArray(i + 1)) Then
        cmp = 3
        cmp = twoTeamCompare(nameArray(indexArray(i)), nameArray(indexArray(i + 1)), 1, toRoundNo * linePerRound, leagueName)
         ActiveWorkbook.Sheets(ownName).Select
        If cmp = 2 Then
             temp = indexArray(i)
            indexArray(i) = indexArray(i + 1)
            indexArray(i + 1) = temp
        ElseIf cmp = 1 Then
        
        ElseIf cmp = 3 Then
            If goalwinArray(indexArray(i)) < goalwinArray(indexArray(i + 1)) Then
                temp = indexArray(i)
                indexArray(i) = indexArray(i + 1)
                indexArray(i + 1) = temp
            ElseIf goalwinArray(indexArray(i)) = goalwinArray(indexArray(i + 1)) Then
                If goalArray(indexArray(i)) < goalArray(indexArray(i + 1)) Then
                    temp = indexArray(i)
                    indexArray(i) = indexArray(i + 1)
                    indexArray(i + 1) = temp
                ElseIf goalArray(indexArray(i)) = goalArray(indexArray(i + 1)) Then
                    If beGoaledArray(indexArray(i)) > beGoaledArray(indexArray(i + 1)) Then
                        temp = indexArray(i)
                        indexArray(i) = indexArray(i + 1)
                        indexArray(i + 1) = temp
                    End If
                End If
            End If
        
        End If
        

    End If
    
    Next i
Next j
    For i = 1 To teamNo
        Cells(i + 1, cstCount).Value = winArray(indexArray(i)) + drawArray(indexArray(i)) + loseArray(indexArray(i))
        Cells(i + 1, cstTeam).Value = nameArray(indexArray(i))
        Cells(i + 1, cstWin).Value = winArray(indexArray(i))
        Cells(i + 1, cstDraw).Value = drawArray(indexArray(i))
        Cells(i + 1, cstLose).Value = loseArray(indexArray(i))
        Cells(i + 1, cstGoal).Value = goalArray(indexArray(i))
        Cells(i + 1, cstGoaled).Value = beGoaledArray(indexArray(i))
        Cells(i + 1, cstGoalWin).Value = goalwinArray(indexArray(i))
        Cells(i + 1, cstPoint).Value = pntArray(indexArray(i))
    Next i
End Function




Function twoTeamCompare(t1 As String, t2 As String, stLine As Integer, edLine As Integer, ln As String) As Integer

    Dim goalInterval As Integer
    Dim point As Integer
    Dim awaygoal As Integer
    
    goalInterval = 0
    point = 0
    awaygoal = 0
 ActiveWorkbook.Sheets(ln).Select
    For i = stLine To edLine
        If Cells(i, 1).Value = t1 And Cells(i, 2).Value = t2 Then
           If Cells(i, 3).Value = "" Or Cells(i, 4).Value = "" Then
           Else
            If Cells(i, 3).Value - Cells(i, 4).Value > 0 Then
                point = point + 1
            ElseIf Cells(i, 3).Value - Cells(i, 4).Value < 0 Then
                point = point - 1
            End If
           
                goalInterval = goalInterval + Cells(i, 3).Value - Cells(i, 4).Value
                awaygoal = awaygoal + -1 * Cells(i, 4).Value
           End If
        ElseIf Cells(i, 1).Value = t2 And Cells(i, 2).Value = t1 Then
           If Cells(i, 3).Value = "" Or Cells(i, 4).Value = "" Then
           Else
            If Cells(i, 4).Value - Cells(i, 3).Value > 0 Then
                point = point + 1
            ElseIf Cells(i, 4).Value - Cells(i, 3).Value < 0 Then
                point = point - 1
            End If
           
                goalInterval = goalInterval + Cells(i, 4).Value - Cells(i, 3).Value
                awaygoal = awaygoal + Cells(i, 4).Value
           End If
        End If
    Next i
    
    If point > 0 Then
       twoTeamCompare = 1
    ElseIf point < 0 Then
        twoTeamCompare = 2
    Else
        If goalInterval > 0 Then
          twoTeamCompare = 1
        ElseIf goalInterval < 0 Then
          twoTeamCompare = 2
        Else
            If awaygoal > 0 Then
              twoTeamCompare = 1
            ElseIf awaygoal < 0 Then
              twoTeamCompare = 2
            Else: twoTeamCompare = 3
            End If
        End If
    End If
    
End Function


Function is_even(x As Integer) As Integer
    If x Mod 2 = 0 Then
       is_even = x
    Else
       is_even = x + 1
    End If
End Function

