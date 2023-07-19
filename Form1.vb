Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        driver()
    End Sub
    Private Sub driver()
        ListBox1.Items.Add("Ted")
        ListBox1.Items.Add("Input 80 2 " & PerfBandTed(80, 2))
        ListBox1.Items.Add("Input 90 2 " & PerfBandTed(90, 2))
        ListBox1.Items.Add("Input 80 1 " & PerfBandTed(80, 1))
        ListBox1.Items.Add("Input 80 100 " & PerfBandTed(80, 100))
        ListBox1.Items.Add("Input 110 100 " & PerfBandTed(110, 100))
        'ListBox1.Items.Add("Input 0 0 " & PerfBandTed(0, 0))
        'ListBox1.Items.Add(PerfBandTed("Mr", "Gray"))
        ListBox1.Items.Add("Input -80 -100 " & PerfBandTed(-80, -100))
        ListBox1.Items.Add("Mary")
        ListBox1.Items.Add("Input 80 2 " & PerfBandMary(80, 2))
        ListBox1.Items.Add("Input 90 2 " & PerfBandMary(90, 2))
        ListBox1.Items.Add("Input 80 1 " & PerfBandMary(80, 1))
        ListBox1.Items.Add("Input 80 100 " & PerfBandMary(80, 100))
        ListBox1.Items.Add("Input 110 100 " & PerfBandMary(110, 100))
        ListBox1.Items.Add("Input 0 0 " & PerfBandMary(0, 0))
        ListBox1.Items.Add("Input Mr Gray " & PerfBandMary("Mr", "Gray"))
        ListBox1.Items.Add("Input -80 -100 " & PerfBandMary(-80, -100))
        ListBox1.Items.Add("Alice")
        ListBox1.Items.Add("Input 80 2 " & PerfBandAlice(80, 2))
        ListBox1.Items.Add("Input 90 2 " & PerfBandAlice(90, 2))
        ListBox1.Items.Add("Input 80 1 " & PerfBandAlice(80, 1))
        ListBox1.Items.Add("Input 80 100 " & PerfBandAlice(80, 100))
        ListBox1.Items.Add("Input 110 100 " & PerfBandAlice(110, 100))
        'ListBox1.Items.Add("Input 0 0 " & PerfBandAlice(0, 0))
        'ListBox1.Items.Add(PerfBandAlice("Mr", "Gray"))
        ListBox1.Items.Add("Input -80 -100 " & PerfBandAlice(-80, -100))
    End Sub

    Public Function PerfBandTed(Mark As Integer, Unit As Integer) As Integer
        Dim PMark As Integer, Band As Integer
        PMark = Mark * 2 / Unit
        If PMark < 50 Then
            Band = 1
        ElseIf PMark < 60 Then
            Band = 2
        ElseIf PMark < 70 Then
            Band = 3
        ElseIf PMark < 80 Then
            Band = 4
        ElseIf PMark < 90 Then
            Band = 5
        Else
            Band = 6
        End If
        PerfBandTed = Band
    End Function

    Public Function PerfBandMary(Mark As Object, Unit As Object) As Integer
        On Error GoTo InputError
        'Calculate performance band
        PerfBandMary = Int(((Mark * 2 / Unit) - 30) / 10)
        'Above line can result in -3 to 0. These should all be 1.
        If PerfBandMary < 1 Then PerfBandMary = 1
        'Also full marks results in a band 7. This should be 6.
        If PerfBandMary > 6 Then PerfBandMary = 6
        Exit Function
        'Return a zero for all problems encountered
InputError:
        PerfBandMary = 0
    End Function

    Public Function PerfBandAlice(Mark As Integer, UnitValue As Integer) As Integer
        Dim Percent As Integer, Band As Integer
        'Convert to a mark out of 100
        Percent = Mark / (UnitValue * 50) * 100
        'Convert to a performance band
        Select Case Percent
            Case Is >= 90 : Band = 6
            Case Is >= 80 : Band = 5
            Case Is >= 70 : Band = 4
            Case Is >= 60 : Band = 3
            Case Is >= 50 : Band = 2
            Case Is >= 0 : Band = 1
            Case Else : Band = 0
        End Select
        'Return performance band
        PerfBandAlice = Band
    End Function
End Class
