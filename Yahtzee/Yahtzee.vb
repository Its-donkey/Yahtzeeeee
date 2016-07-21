Public Class frmYahtzee
    Private diceValue(4) As Integer
    Private diceRoll As Boolean() = {True, True, True, True, True}
    Private roundNumber As Single = 13
    Private subRoundNumber As Single = 3
    Private gameNumber As Integer = 1
    Private scoreCard(14, 5) As Integer
    Private upperSubTotal, upperTotal As Integer
    Private roundScored As Boolean
    Private mouseListener As Boolean = False
    Private bonusYahtzeeSelection As Integer

    Private Sub btnRollDice_Click(sender As Object, e As EventArgs) Handles btnRollDice.Click
        If subRoundNumber > 0 Then
            rollAllDice()
        End If
    End Sub
    Private Sub diceClick(sender As Object, e As EventArgs) Handles lblDiceOne.Click, lblDiceTwo.Click, lblDiceThree.Click,
        lblDiceFour.Click, lblDiceFive.Click

        Select Case DirectCast(sender, Label).Name

            Case "lblDiceOne"
                changeButtonState(sender, diceRoll(0))

            Case "lblDiceTwo"
                changeButtonState(sender, diceRoll(1))

            Case "lblDiceThree"
                changeButtonState(sender, diceRoll(2))

            Case "lblDiceFour"
                changeButtonState(sender, diceRoll(3))

            Case "lblDiceFive"
                changeButtonState(sender, diceRoll(4))

        End Select
    End Sub
    Private Sub changeButtonState(ByVal buttonName As Object, ByRef diceRollNumber As Boolean)
        If subRoundNumber <> 3 Then
            Dim dice As Label = DirectCast(buttonName, Label)
            diceRollNumber = Not (diceRollNumber)
            If diceRollNumber Then
                dice.BorderStyle = BorderStyle.FixedSingle
            Else
                dice.BorderStyle = BorderStyle.Fixed3D
            End If
        End If
    End Sub

    'Upper Section
    Private Sub selectAces_Click(sender As Object, e As EventArgs) Handles picAces.Click
        If mouseListener = True Then
            bonusYahtzeeSelection = 1
            mouseListener = False
        Else
            scoreGame(1)
        End If
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles picTwos.Click
        If mouseListener = True Then
            bonusYahtzeeSelection = 2
            mouseListener = False
        Else
            scoreGame(2)
        End If
    End Sub
    Private Sub PictureBox1_Click_1(sender As Object, e As EventArgs) Handles picThrees.Click
        If mouseListener = True Then
            bonusYahtzeeSelection = 3
            mouseListener = False
        Else
            scoreGame(3)
        End If
    End Sub
    Private Sub picFours_Click(sender As Object, e As EventArgs) Handles picFours.Click
        If mouseListener = True Then
            bonusYahtzeeSelection = 4
            mouseListener = False
        Else
            scoreGame(4)
        End If
    End Sub
    Private Sub picFives_Click(sender As Object, e As EventArgs) Handles picFives.Click
        If mouseListener = True Then
            bonusYahtzeeSelection = 5
            mouseListener = False
        Else
            scoreGame(5)
        End If
    End Sub
    Private Sub picSixes_Click(sender As Object, e As EventArgs) Handles picSixes.Click
        If mouseListener = True Then
            bonusYahtzeeSelection = 6
            mouseListener = False
        Else
            scoreGame(6)
        End If
    End Sub
    Private Sub lblGameOneUpperSubTotal_TextChanged(sender As Object, e As EventArgs) Handles lblGameOneUpperSubTotal.TextChanged,
        lblGameTwoUpperSubTotal.TextChanged, lblGameThreeUpperSubTotal.TextChanged, lblGameFourUpperSubTotal.TextChanged, lblGameFiveUpperSubTotal.TextChanged,
        lblGameSixUpperSubTotal.TextChanged

        If upperSubTotal >= 63 And scoreCard(6, gameNumber - 1) <> 35 Then
            Me.Controls("lblGame" & NumberToText(gameNumber) & "UpperBonus").Text = CStr(35)
            upperTotal += 35
            Me.Controls("lblGame" & NumberToText(gameNumber) & "UpperTotal").Text = CStr(upperSubTotal)
            scoreCard(6, gameNumber - 1) = 35
        End If
    End Sub
    'Lower Section
    Private Sub picThreeOfAKind_Click(sender As Object, e As EventArgs) Handles picThreeOfAKind.Click
        If mouseListener = True Then
            bonusYahtzeeSelection = 7
        Else
            If scoreCard(7, (gameNumber - 1)) <> 1 And roundScored = False Then
                Dim threeOfAKindScore As Integer

                Array.Sort(diceValue)

                If diceValue(0) = diceValue(1) And diceValue(1) = diceValue(2) Then
                    For i As Integer = 0 To diceValue.GetUpperBound(0)
                        threeOfAKindScore += diceValue(i)
                    Next
                ElseIf diceValue(1) = diceValue(2) And diceValue(2) = diceValue(3) Then
                    For i As Integer = 0 To diceValue.GetUpperBound(0)
                        threeOfAKindScore += diceValue(i)
                    Next
                ElseIf diceValue(2) = diceValue(3) And diceValue(3) = diceValue(4) Then
                    For i As Integer = 0 To diceValue.GetUpperBound(0)
                        threeOfAKindScore += diceValue(i)
                    Next
                Else
                    threeOfAKindScore = 0
                End If


                Me.Controls("lblThreeOfAKindGame" & NumberToText(gameNumber)).Text = CStr(threeOfAKindScore)
                roundScored = True
                btnRollDice.Visible = False
                btnNextRound.Visible = True
                scoreCard(7, (gameNumber - 1)) = 1
                roundNumber -= 1
            End If
        End If
    End Sub
    Private Sub picFourOfAKind_Click(sender As Object, e As EventArgs) Handles picFourOfAKind.Click
        If mouseListener = True Then
            bonusYahtzeeSelection = 8
        Else
            If scoreCard(8, (gameNumber - 1)) <> 1 And roundScored = False Then
                Dim fourOfAKindScore As Integer

                Array.Sort(diceValue)

                If diceValue(0) = diceValue(1) And diceValue(0) = diceValue(2) And diceValue(0) = diceValue(3) Then
                    For i As Integer = 0 To diceValue.GetUpperBound(0)
                        fourOfAKindScore += diceValue(i)
                    Next
                ElseIf diceValue(1) = diceValue(2) And diceValue(1) = diceValue(3) And diceValue(1) = diceValue(4) Then
                    For i As Integer = 0 To diceValue.GetUpperBound(0)
                        fourOfAKindScore += diceValue(i)
                    Next
                Else
                    fourOfAKindScore = 0
                End If


                Me.Controls("lblFourOfAKindGame" & NumberToText(gameNumber)).Text = CStr(fourOfAKindScore)
                roundScored = True
                btnRollDice.Visible = False
                btnNextRound.Visible = True
                scoreCard(8, (gameNumber - 1)) = 1
                roundNumber -= 1
            End If
        End If
    End Sub
    Private Sub picFullHouse_Click(sender As Object, e As EventArgs) Handles picFullHouse.Click
        If mouseListener = True Then
            bonusYahtzeeSelection = 9
        Else
            If scoreCard(9, (gameNumber - 1)) <> 1 And roundScored = False Then
                Dim fullHouseScore As Integer

                Array.Sort(diceValue)

                If diceValue(0) = diceValue(1) And diceValue(2) = diceValue(3) And diceValue(2) = diceValue(4) And diceValue(2) <> diceValue(0) Then
                    fullHouseScore = 25
                ElseIf diceValue(0) = diceValue(1) And diceValue(0) = diceValue(2) And diceValue(3) = diceValue(4) And diceValue(3) <> diceValue(0) Then
                    fullHouseScore = 25
                Else
                    fullHouseScore = 0
                End If


                Me.Controls("lblFullHouseGame" & NumberToText(gameNumber)).Text = CStr(fullHouseScore)
                roundScored = True
                btnRollDice.Visible = False
                btnNextRound.Visible = True
                scoreCard(9, (gameNumber - 1)) = 1
                roundNumber -= 1
            End If
        End If
    End Sub
    Private Sub picSmallStraight_Click(sender As Object, e As EventArgs) Handles picSmallStraight.Click
        If mouseListener = True Then
            bonusYahtzeeSelection = 10
        Else
            If scoreCard(10, (gameNumber - 1)) <> 1 And roundScored = False Then
                Dim smallStraightScore As Integer
                Array.Sort(diceValue)

                For i As Integer = 0 To 3
                    If diceValue(i) = (diceValue(i + 1) - 1) Then
                        smallStraightScore = 30
                    Else
                        If diceValue(i) = (diceValue(i + 1)) Then
                            smallStraightScore = 30
                        Else
                            smallStraightScore = 0
                            Exit For
                        End If
                    End If
                Next

                Me.Controls("lblSmallStraightGame" & NumberToText(gameNumber)).Text = CStr(smallStraightScore)

                roundScored = True
                btnRollDice.Visible = False
                btnNextRound.Visible = True
                scoreCard(10, (gameNumber - 1)) = 1
                roundNumber -= 1
            End If
        End If
    End Sub
    Private Sub picLongStraight_Click(sender As Object, e As EventArgs) Handles picLongStraight.Click
        If mouseListener = True Then
            bonusYahtzeeSelection = 11
        Else
            If scoreCard(11, (gameNumber - 1)) <> 1 And roundScored = False Then
                Dim longStraightScore As Integer
                Array.Sort(diceValue)

                For i As Integer = 0 To 3
                    If diceValue(i) = (diceValue(i + 1) - 1) Then
                        longStraightScore = 40
                    Else
                        longStraightScore = 0
                        Exit For
                    End If
                Next

                Me.Controls("lblLongStraightGame" & NumberToText(gameNumber)).Text = CStr(longStraightScore)

                roundScored = True
                btnRollDice.Visible = False
                btnNextRound.Visible = True
                scoreCard(11, (gameNumber - 1)) = 1
                roundNumber -= 1
            End If
        End If
    End Sub
    Public Sub picYahtzee_Click(sender As Object, e As EventArgs) Handles picYahtzee.Click
        If scoreCard(12, (gameNumber - 1)) <> 1 And roundScored = False Then
            Dim yahtzeeScore As Integer
            Array.Sort(diceValue)

            For i As Integer = 0 To 3
                If diceValue(i) = diceValue(i + 1) Then
                    yahtzeeScore = 50
                Else
                    yahtzeeScore = 0
                    Exit For
                End If
            Next

            Me.Controls("lblYahtzeeGame" & NumberToText(gameNumber)).Text = CStr(yahtzeeScore)

            roundScored = True
            btnRollDice.Visible = False
            btnNextRound.Visible = True
            scoreCard(12, (gameNumber - 1)) = 1
            roundNumber -= 1
        End If

    End Sub
    Private Sub picChance_Click(sender As Object, e As EventArgs) Handles picChance.Click
        If mouseListener = True Then
            bonusYahtzeeSelection = 12
        Else
            If scoreCard(13, (gameNumber - 1)) <> 1 And roundScored = False Then
                Dim chanceScore As Integer

                For i As Integer = 0 To diceValue.GetUpperBound(0)
                    chanceScore += diceValue(i)
                Next

                Me.Controls("lblChanceGame" & NumberToText(gameNumber)).Text = CStr(chanceScore)

                roundScored = True
                btnRollDice.Visible = False
                btnNextRound.Visible = True
                scoreCard(13, (gameNumber - 1)) = 1
                roundNumber -= 1
            End If
        End If
    End Sub
    Private Sub picYahtzeeBonus_Click(sender As Object, e As EventArgs) Handles picYahtzeeBonus.Click
        Dim bonusYahtzeeScore, totalBonusYahtzeeScore As Integer

        For i As Integer = 0 To 3
            If diceValue(i) = diceValue(i + 1) Then
                bonusYahtzeeScore = 100
            Else
                bonusYahtzeeScore = 0
                Exit For
            End If
        Next

        If scoreCard(12, (gameNumber - 1)) = 0 And roundScored = False Then
            Dim yahtzeeScore As Integer

            For i As Integer = 0 To 3
                If diceValue(i) = diceValue(i + 1) Then
                    yahtzeeScore = 50
                Else
                    yahtzeeScore = 0
                    Exit For
                End If
            Next

            Me.Controls("lblYahtzeeGame" & NumberToText(gameNumber)).Text = CStr(yahtzeeScore)

            roundScored = True
            btnRollDice.Visible = False
            btnNextRound.Visible = True
            scoreCard(12, (gameNumber - 1)) = 1
            roundNumber -= 1
            Exit Sub
        End If


        If scoreCard(14, (gameNumber - 1)) < 3 And scoreCard(12, (gameNumber - 1)) = 1 And roundScored = False Then

            For i As Integer = 0 To 3
                If diceValue(i) = diceValue(i + 1) Then
                    bonusYahtzeeScore = 100
                Else
                    Exit Sub
                End If
            Next


            Dim scoreValue As Integer = 0
            If scoreCard((diceValue(0) - 1), (gameNumber - 1)) = 0 Then
                For i As Integer = 0 To diceValue.GetUpperBound(0)
                    scoreValue += diceValue(i)
                Next

                Me.Controls("lbl" & NumberToText(diceValue(0)) & "Game" & NumberToText(gameNumber)).Text = CStr(scoreValue)

                scoreCard((diceValue(0) - 1), (gameNumber - 1)) = 1
                roundScored = True
                roundNumber -= 1

                upperSubTotal += scoreValue
                upperTotal += scoreValue
                Me.Controls("lblGame" & NumberToText(gameNumber) & "UpperSubTotal").Text = CStr(upperSubTotal)
                Me.Controls("lblGame" & NumberToText(gameNumber) & "UpperTotal").Text = CStr(upperTotal)

                btnRollDice.Visible = False
                btnNextRound.Visible = True
            ElseIf scoreCard((diceValue(0) - 1), (gameNumber - 1)) = 1 Then

                btnRollDice.Visible = False
                btnNextRound.Visible = False
                MsgBox("You must select another score category on top of the bonus Yahtzee.")

            End If

            scoreCard(14, (gameNumber - 1)) += 1

            Me.Controls("picYahtzeeBonus" & NumberToText(scoreCard(14, (gameNumber - 1))) & "Game" & NumberToText(gameNumber)).BackgroundImage = My.Resources.Extra_Yahtzee_Tick

            If scoreCard(14, (gameNumber - 1)) = 1 Then
                totalBonusYahtzeeScore = bonusYahtzeeScore
            Else
                totalBonusYahtzeeScore = CInt(Me.Controls("lblYahtzeeBonusScoreGame" & NumberToText(gameNumber)).Text)
                totalBonusYahtzeeScore += bonusYahtzeeScore
            End If


            Me.Controls("lblYahtzeeBonusScoreGame" & NumberToText(gameNumber)).Text = CStr(totalBonusYahtzeeScore)



        Else
            Exit Sub
        End If
    End Sub

    Sub rollAllDice()
        Dim random = New Random()
        For i As Integer = 0 To diceValue.GetUpperBound(0)
            If diceRoll(i) Then
                diceValue(i) = random.Next(6, 7)
            End If
        Next

        lblDiceOne.Text = CStr(diceValue(0))
        lblDiceTwo.Text = CStr(diceValue(1))
        lblDiceThree.Text = CStr(diceValue(2))
        lblDiceFour.Text = CStr(diceValue(3))
        lblDiceFive.Text = CStr(diceValue(4))

        subRoundNumber -= 1

        If subRoundNumber = 0 Then
            btnRollDice.Visible = False
        End If
    End Sub
    Private Sub scoreGame(diceSelect As Integer)
        Dim scoreValue As Integer = 0

        If scoreCard((diceSelect - 1), (gameNumber - 1)) = 0 And subRoundNumber <> 3 And roundScored = False Then
            For i As Integer = 0 To diceValue.GetUpperBound(0)
                If diceValue(i) = diceSelect Then
                    scoreValue += diceSelect
                End If
            Next

            Me.Controls("lbl" & NumberToText(diceSelect) & "Game" & NumberToText(gameNumber)).Text = CStr(scoreValue)

            btnRollDice.Visible = False
            btnNextRound.Visible = True
            scoreCard((diceSelect - 1), (gameNumber - 1)) = 1

        End If

        roundNumber -= 1

        upperSubTotal += scoreValue
        upperTotal += scoreValue
        roundScored = True

        Me.Controls("lblGame" & NumberToText(gameNumber) & "UpperSubTotal").Text = CStr(upperSubTotal)
        Me.Controls("lblGame" & NumberToText(gameNumber) & "UpperTotal").Text = CStr(upperTotal)
    End Sub

    Private Sub bonusYahtzeeScoreUpper(diceSelect As Integer)
        Dim scoreValue As Integer = 0

        For i As Integer = 0 To diceValue.GetUpperBound(0)
            If diceValue(i) = diceSelect Then
                scoreValue += diceSelect
            End If
        Next

        Me.Controls("lbl" & NumberToText(diceSelect) & "Game" & NumberToText(gameNumber)).Text = CStr(scoreValue)

        scoreCard((diceSelect - 1), (gameNumber - 1)) = 1
    End Sub
    Private Function NumberToText(ByRef n As Integer) As String
        If n = 1 Then
            Return "One"
        ElseIf n = 2 Then
            Return "Two"
        ElseIf n = 3 Then
            Return "Three"
        ElseIf n = 4 Then
            Return "Four"
        ElseIf n = 5 Then
            Return "Five"
        ElseIf n = 6 Then
            Return "Six"
        Else
            Return ""
        End If
    End Function

    Private Sub btnNextRound_Click(sender As Object, e As EventArgs) Handles btnNextRound.Click
        For i As Integer = 0 To diceRoll.GetUpperBound(0)
            diceRoll(i) = True
        Next

        lblDiceOne.BorderStyle = BorderStyle.FixedSingle
        lblDiceTwo.BorderStyle = BorderStyle.FixedSingle
        lblDiceThree.BorderStyle = BorderStyle.FixedSingle
        lblDiceFour.BorderStyle = BorderStyle.FixedSingle
        lblDiceFive.BorderStyle = BorderStyle.FixedSingle

        subRoundNumber = 3
        If subRoundNumber > 0 Then
            rollAllDice()
        End If
        btnRollDice.Visible = True
        btnNextRound.Visible = False
        roundScored = False
    End Sub

    Private Sub lblGameOneUpperTotal_TextChanged(sender As Object, e As EventArgs) Handles lblGameOneUpperTotal.TextChanged
        lblTotalOfUpperSectionGameOne.Text = lblGameOneUpperTotal.Text

    End Sub
End Class
