Imports System.IO

Public Class frmYahtzee
    Private diceValue(4) As Integer
    Private diceRoll As Boolean() = {True, True, True, True, True}
    Private roundNumber As Single = 13
    Private subRoundNumber As Single = 3
    Private gameNumber As Integer = 1
    Private scoreCard(1, 19, 5) As Integer
    Private roundScored As Boolean
    Private mouseListener As Boolean = False
    Private bonusYahtzeeSelection As Integer
    Private Sub frmYahtzee_Load(sender As Object, e As EventArgs) Handles Me.Load
        frmNameEntry.ShowDialog()
    End Sub
    Private Sub mnuTopBarHelpViewHelp_Click(sender As Object, e As EventArgs) Handles mnuTopBarHelpViewHelp.Click

    End Sub
    Private Sub mnuTopBarHelpRules_Click(sender As Object, e As EventArgs) Handles mnuTopBarHelpRules.Click
        Dim pdf As Byte() = My.Resources.Game_Rules
        Using tmp As New FileStream("Game_Rules.pdf", FileMode.Create)
            tmp.Write(pdf, 0, pdf.Length)
        End Using
        Process.Start("Game_Rules.pdf")
    End Sub
    Private Sub btnRollDice_Click(sender As Object, e As EventArgs) Handles btnRollDice.Click
        If subRoundNumber > 0 Then
            rollAllDice()
        End If
    End Sub
    Private Sub diceClick(sender As Object, e As EventArgs) Handles picDiceOneImage.Click, picDiceTwoImage.Click, picDiceThreeImage.Click, picDiceFourImage.Click, picDiceFiveImage.Click

        Select Case DirectCast(sender, PictureBox).Name

            Case "picDiceOneImage"
                changeButtonState(sender, diceRoll(0))

            Case "picDiceTwoImage"
                changeButtonState(sender, diceRoll(1))

            Case "picDiceThreeImage"
                changeButtonState(sender, diceRoll(2))

            Case "picDiceFourImage"
                changeButtonState(sender, diceRoll(3))

            Case "picDiceFiveImage"
                changeButtonState(sender, diceRoll(4))

        End Select
    End Sub
    Private Sub changeButtonState(ByVal buttonName As Object, ByRef diceRollNumber As Boolean)
        If subRoundNumber <> 3 Then
            Dim dice As PictureBox = DirectCast(buttonName, PictureBox)
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
        scoreGame(1)
    End Sub
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles picTwos.Click
        scoreGame(2)
    End Sub
    Private Sub PictureBox1_Click_1(sender As Object, e As EventArgs) Handles picThrees.Click
        scoreGame(3)
    End Sub
    Private Sub picFours_Click(sender As Object, e As EventArgs) Handles picFours.Click
        scoreGame(4)
    End Sub
    Private Sub picFives_Click(sender As Object, e As EventArgs) Handles picFives.Click
        scoreGame(5)
    End Sub
    Private Sub picSixes_Click(sender As Object, e As EventArgs) Handles picSixes.Click
        scoreGame(6)
    End Sub

    'Lower Section
    Private Sub picThreeOfAKind_Click(sender As Object, e As EventArgs) Handles picThreeOfAKind.Click
        If scoreCard(0, 9, (gameNumber - 1)) = 0 And roundScored = False Then

            Array.Sort(diceValue)

            If diceValue(0) = diceValue(1) And diceValue(1) = diceValue(2) Then
                For i As Integer = 0 To diceValue.GetUpperBound(0)
                    scoreCard(1, 9, (gameNumber - 1)) += diceValue(i)
                Next
            ElseIf diceValue(1) = diceValue(2) And diceValue(2) = diceValue(3) Then
                For i As Integer = 0 To diceValue.GetUpperBound(0)
                    scoreCard(1, 9, (gameNumber - 1)) += diceValue(i)
                Next
            ElseIf diceValue(2) = diceValue(3) And diceValue(3) = diceValue(4) Then
                For i As Integer = 0 To diceValue.GetUpperBound(0)
                    scoreCard(1, 9, (gameNumber - 1)) += diceValue(i)
                Next
            Else
                scoreCard(1, 9, (gameNumber - 1)) = 0
            End If

            Me.Controls("lblThreeOfAKindGame" & NumberToText(gameNumber)).Text = CStr(scoreCard(1, 9, (gameNumber - 1)))

            scoreCard(0, 9, (gameNumber - 1)) = 1
            roundAndGameNumber()
            scoreLowerSection(scoreCard(1, 9, (gameNumber - 1)))
        End If
    End Sub
    Private Sub picFourOfAKind_Click(sender As Object, e As EventArgs) Handles picFourOfAKind.Click
        If scoreCard(0, 10, (gameNumber - 1)) = 0 And roundScored = False Then

            Array.Sort(diceValue)

            If diceValue(0) = diceValue(1) And diceValue(0) = diceValue(2) And diceValue(0) = diceValue(3) Then
                For i As Integer = 0 To diceValue.GetUpperBound(0)
                    scoreCard(1, 10, (gameNumber - 1)) += diceValue(i)
                Next
            ElseIf diceValue(1) = diceValue(2) And diceValue(1) = diceValue(3) And diceValue(1) = diceValue(4) Then
                For i As Integer = 0 To diceValue.GetUpperBound(0)
                    scoreCard(1, 10, (gameNumber - 1)) += diceValue(i)
                Next
            Else
                scoreCard(1, 10, (gameNumber - 1)) = 0
            End If

            Me.Controls("lblFourOfAKindGame" & NumberToText(gameNumber)).Text = CStr(scoreCard(1, 10, (gameNumber - 1)))
            roundAndGameNumber()
            scoreCard(0, 10, (gameNumber - 1)) = 1
            scoreLowerSection(scoreCard(1, 10, (gameNumber - 1)))
        End If
    End Sub
    Private Sub picFullHouse_Click(sender As Object, e As EventArgs) Handles picFullHouse.Click
        If scoreCard(0, 11, (gameNumber - 1)) = 0 And roundScored = False Then

            Array.Sort(diceValue)

            If diceValue(0) = diceValue(1) And diceValue(2) = diceValue(3) And diceValue(2) = diceValue(4) And diceValue(2) <> diceValue(0) Then
                scoreCard(1, 11, (gameNumber - 1)) = 25
            ElseIf diceValue(0) = diceValue(1) And diceValue(0) = diceValue(2) And diceValue(3) = diceValue(4) And diceValue(3) <> diceValue(0) Then
                scoreCard(1, 11, (gameNumber - 1)) = 25
            Else
                scoreCard(1, 11, (gameNumber - 1)) = 0
            End If

            Me.Controls("lblFullHouseGame" & NumberToText(gameNumber)).Text = CStr(scoreCard(1, 11, (gameNumber - 1)))
            roundAndGameNumber()
            scoreLowerSection(scoreCard(1, 11, (gameNumber - 1)))
            scoreCard(0, 11, (gameNumber - 1)) = 1
        End If
    End Sub
    Private Sub picSmallStraight_Click(sender As Object, e As EventArgs) Handles picSmallStraight.Click
        If scoreCard(0, 12, (gameNumber - 1)) = 0 And roundScored = False Then
            Array.Sort(diceValue)

            For i As Integer = 0 To 3
                If diceValue(i) = (diceValue(i + 1) - 1) Then
                    scoreCard(1, 12, (gameNumber - 1)) = 30
                Else
                    If diceValue(i) = (diceValue(i + 1)) Then
                        scoreCard(1, 12, (gameNumber - 1)) = 30
                    Else
                        scoreCard(1, 12, (gameNumber - 1)) = 0
                        Exit For
                    End If
                End If
            Next

            Me.Controls("lblSmallStraightGame" & NumberToText(gameNumber)).Text = CStr(scoreCard(1, 12, (gameNumber - 1)))

            scoreCard(0, 12, (gameNumber - 1)) = 1
            roundAndGameNumber()
            scoreLowerSection(scoreCard(1, 12, (gameNumber - 1)))
        End If
    End Sub
    Private Sub picLongStraight_Click(sender As Object, e As EventArgs) Handles picLongStraight.Click
        If scoreCard(0, 13, (gameNumber - 1)) = 0 And roundScored = False Then

            Array.Sort(diceValue)

            For i As Integer = 0 To 3
                If diceValue(i) = (diceValue(i + 1) - 1) Then
                    scoreCard(1, 13, (gameNumber - 1)) = 40
                Else
                    scoreCard(1, 13, (gameNumber - 1)) = 0
                    Exit For
                End If
            Next

            Me.Controls("lblLongStraightGame" & NumberToText(gameNumber)).Text = CStr(scoreCard(1, 13, (gameNumber - 1)))

            scoreCard(0, 13, (gameNumber - 1)) = 1
            roundAndGameNumber()
            scoreLowerSection(scoreCard(1, 13, (gameNumber - 1)))
        End If
    End Sub
    Public Sub picYahtzee_Click(sender As Object, e As EventArgs) Handles picYahtzee.Click
        If scoreCard(0, 14, (gameNumber - 1)) = 0 And roundScored = False Then
            Array.Sort(diceValue)

            For i As Integer = 0 To 3
                If diceValue(i) = diceValue(i + 1) Then
                    scoreCard(1, 14, (gameNumber - 1)) = 50
                Else
                    scoreCard(1, 14, (gameNumber - 1)) = 0
                    Exit For
                End If
            Next

            Me.Controls("lblYahtzeeGame" & NumberToText(gameNumber)).Text = CStr(scoreCard(1, 14, (gameNumber - 1)))

            scoreCard(0, 14, (gameNumber - 1)) = 1

            roundAndGameNumber()
            scoreLowerSection(scoreCard(1, 14, (gameNumber - 1)))
        End If

    End Sub
    Private Sub picChance_Click(sender As Object, e As EventArgs) Handles picChance.Click
        If scoreCard(0, 15, (gameNumber - 1)) = 0 And roundScored = False Then

            For i As Integer = 0 To diceValue.GetUpperBound(0)
                scoreCard(1, 15, (gameNumber - 1)) += diceValue(i)
            Next

            Me.Controls("lblChanceGame" & NumberToText(gameNumber)).Text = CStr(scoreCard(1, 15, (gameNumber - 1)))

            scoreCard(0, 15, (gameNumber - 1)) = 1

            roundAndGameNumber()
            scoreLowerSection(scoreCard(1, 15, (gameNumber - 1)))
        End If
    End Sub
    Private Sub picYahtzeeBonus_Click(sender As Object, e As EventArgs) Handles picYahtzeeBonus.Click

        For i As Integer = 0 To 3
            If diceValue(i) <> diceValue(i + 1) Then
                Exit Sub
            End If
        Next

        If scoreCard(0, 14, (gameNumber - 1)) = 0 And roundScored = False Then
            Me.Controls("lblYahtzeeGame" & NumberToText(gameNumber)).Text = "50"
            scoreCard(0, 14, (gameNumber - 1)) = 1
            scoreCard(1, 14, (gameNumber - 1)) = 50

            roundAndGameNumber()
            Exit Sub
        ElseIf scoreCard(0, 16, (gameNumber - 1)) < 3 And scoreCard(0, 14, (gameNumber - 1)) = 1 And roundScored = False Then
            Dim scoreValue As Integer = 0

            If scoreCard(0, (diceValue(0) - 1), (gameNumber - 1)) = 0 Then
                scoreGame(diceValue(0))
                roundAndGameNumber()
            ElseIf scoreCard(0, (diceValue(0) - 1), (gameNumber - 1)) = 1 Then
                btnRollDice.Visible = False
                btnNextRound.Visible = False
                MsgBox("You must select another score category on top of the bonus Yahtzee.")
            End If

            scoreCard(0, 16, (gameNumber - 1)) += 1

            Me.Controls("picYahtzeeBonus" & NumberToText(scoreCard(0, 16, (gameNumber - 1))) & "Game" & NumberToText(gameNumber)).BackgroundImage = My.Resources.Extra_Yahtzee_Tick

            Select Case scoreCard(0, 16, (gameNumber - 1))
                Case 0
                    scoreCard(1, 17, 0) = 0
                    Me.Controls("lblYahtzeeBonusScoreGame" & NumberToText(gameNumber)).Text = ""
                Case 1
                    scoreCard(1, 17, 0) = 100
                    Me.Controls("lblYahtzeeBonusScoreGame" & NumberToText(gameNumber)).Text = "100"
                Case 2
                    scoreCard(1, 17, 0) = 200
                    Me.Controls("lblYahtzeeBonusScoreGame" & NumberToText(gameNumber)).Text = "200"
                Case 3
                    scoreCard(1, 17, 0) = 300
                    Me.Controls("lblYahtzeeBonusScoreGame" & NumberToText(gameNumber)).Text = "300"
            End Select
        End If
    End Sub

    Sub rollAllDice()
        Dim random = New Random()

        picDiceOneImage.Image = Nothing
        picDiceTwoImage.Image = Nothing
        picDiceThreeImage.Image = Nothing
        picDiceFourImage.Image = Nothing
        picDiceFiveImage.Image = Nothing

        For i As Integer = 0 To diceValue.GetUpperBound(0)
            If diceRoll(i) Then
                diceValue(i) = random.Next(1, 7)
                Me.Controls("picDice" & NumberToText((i + 1)) & "Image").BackgroundImage = setDiceFaceImage(diceValue(i))
            End If
        Next

        'picDiceOneImage.Image = setDiceFaceImage(diceValue(0))
        'picDiceTwoImage.Image = setDiceFaceImage(diceValue(1))
        'picDiceThreeImage.Image = setDiceFaceImage(diceValue(2))
        'picDiceFourImage.Image = setDiceFaceImage(diceValue(3))
        'picDiceFiveImage.Image = setDiceFaceImage(diceValue(4))


        subRoundNumber -= 1

        If subRoundNumber = 0 Then
            btnRollDice.Visible = False
        End If
    End Sub
    Private Sub scoreGame(diceSelect As Integer)

        If scoreCard(0, (diceSelect - 1), (gameNumber - 1)) = 0 And subRoundNumber <> 3 And roundScored = False Then

            Dim scoreValue As Integer

            For i As Integer = 0 To diceValue.GetUpperBound(0)
                If diceValue(i) = diceSelect Then
                    scoreValue += diceSelect

                End If
            Next

            Me.Controls("lbl" & NumberToText(diceSelect) & "Game" & NumberToText(gameNumber)).Text = CStr(scoreValue)

            scoreCard(0, (diceSelect - 1), (gameNumber - 1)) = 1
            scoreCard(1, (diceSelect - 1), (gameNumber - 1)) = scoreValue
            roundAndGameNumber()

            scoreCard(1, 6, (gameNumber - 1)) += scoreValue

            If scoreCard(1, 6, (gameNumber - 1)) >= 63 And scoreCard(0, 7, gameNumber - 1) = 0 Then
                scoreCard(1, 7, gameNumber - 1) = 35
                scoreCard(0, 7, gameNumber - 1) = 1
                Me.Controls("lblGame" & NumberToText(gameNumber) & "UpperBonus").Text = "35"
                scoreCard(1, 8, (gameNumber - 1)) += 35
            End If

            scoreCard(1, 8, (gameNumber - 1)) += scoreValue

            roundScored = True

            Me.Controls("lblGame" & NumberToText(gameNumber) & "UpperSubTotal").Text = CStr(scoreCard(1, 6, (gameNumber - 1)))
            Me.Controls("lblGame" & NumberToText(gameNumber) & "UpperTotal").Text = CStr(scoreCard(1, 8, (gameNumber - 1)))

            Me.Controls("lblTotalOfUpperSectionGame" & NumberToText(gameNumber)).Text = CStr(scoreCard(1, 8, (gameNumber - 1)))

            scoreCard(1, 19, (gameNumber - 1)) += scoreValue
            Me.Controls("lblGrandTotalScoreGame" & NumberToText(gameNumber)).Text = CStr(scoreCard(1, 19, (gameNumber - 1)))
        End If
    End Sub
    Private Sub scoreLowerSection(scoreValue As Integer)
        scoreCard(1, 18, (gameNumber - 1)) += scoreValue
        Me.Controls("lblTotalOfLowerSectionGame" & NumberToText(gameNumber)).Text = CStr(scoreCard(1, 18, (gameNumber - 1)))
        scoreCard(1, 19, (gameNumber - 1)) += scoreValue
        Me.Controls("lblGrandTotalScoreGame" & NumberToText(gameNumber)).Text = CStr(scoreCard(1, 19, (gameNumber - 1)))
    End Sub
    Private Sub bonusYahtzeeScoreUpper(diceSelect As Integer)
        Dim scoreValue As Integer = 0

        For i As Integer = 0 To diceValue.GetUpperBound(0)
            If diceValue(i) = diceSelect Then
                scoreValue += diceSelect
            End If
        Next

        Me.Controls("lbl" & NumberToText(diceSelect) & "Game" & NumberToText(gameNumber)).Text = CStr(scoreValue)

        scoreCard(0, (diceSelect - 1), (gameNumber - 1)) = 1
    End Sub
    Private Sub btnNextRound_Click(sender As Object, e As EventArgs) Handles btnNextRound.Click, btnStartNewGame.Click
        For i As Integer = 0 To diceRoll.GetUpperBound(0)
            diceRoll(i) = True
        Next

        picDiceOneImage.BorderStyle = BorderStyle.FixedSingle
        picDiceTwoImage.BorderStyle = BorderStyle.FixedSingle
        picDiceThreeImage.BorderStyle = BorderStyle.FixedSingle
        picDiceFourImage.BorderStyle = BorderStyle.FixedSingle
        picDiceFiveImage.BorderStyle = BorderStyle.FixedSingle

        subRoundNumber = 3
        rollAllDice()

        btnRollDice.Visible = True
        btnNextRound.Visible = False
        btnStartNewGame.Visible = False
        roundScored = False
    End Sub
    Private Sub roundAndGameNumber()
        roundNumber -= 1
        roundScored = True

        If roundNumber = 0 And gameNumber < 6 Then
            btnRollDice.Visible = False
            btnNextRound.Visible = False

            Dim newGameOfYahtzee As MsgBoxResult = MsgBox("Your Grand Total score for game " & NumberToText(gameNumber) & " is " & scoreCard(1, 19, (gameNumber - 1)) & Environment.NewLine &
                   Environment.NewLine &
                   "Would you like to start a new Game?", MsgBoxStyle.YesNo, "Yahtzee New Game")

            If newGameOfYahtzee = MsgBoxResult.Yes Then
                btnStartNewGame.Visible = True
                gameNumber += 1
                roundNumber = 13

            End If
        ElseIf roundNumber = 0 And gameNumber = 6 Then
            btnRollDice.Visible = False
            btnNextRound.Visible = False
            btnStartNewGame.Visible = False
        Else
            btnRollDice.Visible = False
            btnNextRound.Visible = True
        End If
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
    Private Function setDiceFaceImage(diceRolledNumber As Integer) As Image
        Select Case diceRolledNumber
            Case 1
                Return My.Resources.Dice_One
            Case 2
                Return My.Resources.Dice_Two
            Case 3
                Return My.Resources.Dice_Three
            Case 4
                Return My.Resources.Dice_Four
            Case 5
                Return My.Resources.Dice_Five
            Case 6
                Return My.Resources.Dice_Six
            Case Else
                Return Nothing
        End Select

    End Function

    Private Sub mnuTopBarFileExit_Click(sender As Object, e As EventArgs) Handles mnuTopBarFileExit.Click
        Close()
    End Sub
End Class
