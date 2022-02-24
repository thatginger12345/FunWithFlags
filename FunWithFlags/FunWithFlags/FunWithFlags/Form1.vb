Public Class Form1
    Dim appPath As String = Application.StartupPath & "\Flags\"
    'create a random number generator
    Private Shared rndGenerator As New System.Random(Now.Millisecond)
    Dim blnFirst As Boolean
    Dim iFirst As Integer
    Dim sFirst As System.Object
    Const cMax As Integer = 8
    Dim iMatches As Integer
    Dim iTries As Integer
    Dim AllFlags(0 To 28) As String
    Dim GameFlags(0 To 15) As String


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'reset all variables
        iMatches = 0
        iTries = 0
        blnFirst = True
        lblMsg.Text = "Pick a Flag"
        lblPick1.Text = ""
        lblPick2.Text = ""
        lblScore.Text = ""
        lblTries.Text = ""
        lblCorrect.Text = ""

        'load all names
        AllFlags(0) = "Albania.png"
        AllFlags(1) = "austria.gif"
        AllFlags(2) = "belgium.gif"
        AllFlags(3) = "bulgaria.gif"
        AllFlags(4) = "cyprus.gif"
        AllFlags(5) = "denmark.gif"
        AllFlags(6) = "england.gif"
        AllFlags(7) = "estonia.gif"
        AllFlags(8) = "france.jpg"
        AllFlags(9) = "germany.gif"
        AllFlags(10) = "greece.gif"
        AllFlags(11) = "hungary.gif"
        AllFlags(12) = "iceland.gif"
        AllFlags(13) = "ireland.gif"
        AllFlags(14) = "italy.gif"
        AllFlags(15) = "malta.png"
        AllFlags(16) = "netherlands.jpg"
        AllFlags(17) = "norway.jpg"
        AllFlags(18) = "poland.gif"
        AllFlags(19) = "portugal.jpg"
        AllFlags(20) = "Romania.png"
        AllFlags(21) = "russia.gif"
        AllFlags(22) = "scotland.gif"
        AllFlags(23) = "spain.gif"
        AllFlags(24) = "sweden.gif"
        AllFlags(25) = "switzerland.gif"
        AllFlags(26) = "turkey.gif"
        AllFlags(27) = "wales.jpg"
        AllFlags(28) = "finland.gif"

        'call procedures to set up game
        Call LoadBacks()
        Call SwapFlags()
        Call CreateGameSet


    End Sub

    Sub LoadBacks()
        'this procedure will reset all the picture boxes
        Dim iControl As Integer
        Dim iCount As Integer
        Dim strControlName As String
        Dim oTemp As System.Object
        'count the object of the form
        iControl = Controls.Count
        'loop through objects
        For iCount = 0 To iControl - 1
            'save the name and the object
            strControlName = Controls(iCount).Name
            oTemp = Controls(iCount)
            If strControlName.Substring(0, 3) = "pic" Then
                oTemp.Enabled = True
                oTemp.Visible = True
                oTemp.Image = Image.FromFile(appPath & "back.jpg")
            End If
        Next
    End Sub

    Sub SwapFlags()
        'this procedure will swap the flags into a new order
        Dim strTemp As String
        Dim iSwap1 As Integer
        Dim iSwap2 As Integer
        Dim iCount As Integer

        'repeat for number of flags
        For iCount = 0 To 30
            iSwap1 = rndGenerator.Next(0, 28)
            iSwap2 = rndGenerator.Next(0, 28)
            'swap the two flags
            strTemp = AllFlags(iSwap1)
            AllFlags(iSwap1) = AllFlags(iSwap2)
            AllFlags(iSwap2) = strTemp
        Next
    End Sub

    Sub CreateGameSet()
        'this will take the first 8 images from the large array
        'and load each one twice to create the game set
        Dim strTemp As String
        Dim iSwap1 As Integer
        Dim iSwap2 As Integer
        Dim iCount As Integer

        'load the game array
        For iCount = 0 To 15 Step 2
            GameFlags(iCount) = AllFlags(iCount)
            GameFlags(iCount + 1) = AllFlags(iCount)
        Next

        'repeat for number of flags
        For iCount = 0 To 15
            iSwap1 = rndGenerator.Next(0, 15)
            iSwap2 = rndGenerator.Next(0, 15)
            'swap the two flags
            strTemp = GameFlags(iSwap1)
            GameFlags(iSwap1) = GameFlags(iSwap2)
            GameFlags(iSwap2) = strTemp
        Next

    End Sub

    Private Sub PicFlag0_Click(sender As Object, e As EventArgs) Handles picFlag0.Click, picFlag1.Click, picFlag2.Click, picFlag3.click, picFlag4.Click, picFlag5.Click, picFlag6.Click, picFlag7.Click, picFlag8.Click, picFlag9.Click, picFlag10.Click, picFlag11.Click, picFlag12.Click, picFlag13.Click, picFlag14.Click, picFlag15.Click
        'this will show the image of the selected flag and check
        'for a match
        Dim bPick As Byte

        'save the tag
        bPick = sender.tag
        'load the image
        sender.image = Image.FromFile(appPath & GameFlags(bPick))
        sender.enabled = False
        'check for first pick
        If blnFirst = True Then
            'store the number and object
            iFirst = bPick
            sFirst = sender
            lblPick1.Text = GetName(bPick)
            blnFirst = False
        Else
            lblPick2.Text = GetName(bPick)
            Me.Refresh()
            Call CheckForMatch(iFirst, bPick, sFirst, sender)
        End If
    End Sub

    Sub CheckForMatch(bNo1 As Byte, bNo2 As Byte, oFlag1 As System.Object, oFlag2 As System.Object)
        'this procedure will check for a match and deal with correct
        'and incorrect answers
        iTries = iTries + 1
        'check for a match
        If GameFlags(bNo1) = GameFlags(bNo2) Then
            iMatches = iMatches + 1
            lblMsg.Text = "Flags ARE Fun!!!"
            System.Threading.Thread.Sleep(500)
            oFlag1.Enabled = False
            oFlag2.Enabled = False
            oFlag1.Visible = False
            oFlag2.Visible = False
        Else
            lblMsg.Text = "Nooooooo! BAZINGA!!"
            System.Threading.Thread.Sleep(500)
            oFlag1.Enabled = True
            oFlag2.Enabled = True
            oFlag1.Image = Image.FromFile(appPath & "back.jpg")
            oFlag2.Image = Image.FromFile(appPath & "back.jpg")
        End If

        lblTries.Text = iTries
        lblCorrect.Text = iMatches
        lblPick1.Text = ""
        lblPick2.Text = ""
        'check for game over
        If iMatches = cMax Then
            lblMsg.Text = "Game Over -  Did you have fun with flags?"
            lblScore.Text = Convert.ToInt32((iMatches / iTries) * 100)
        Else
            lblMsg.Text = "Pick a Flag"
            blnFirst = True
        End If

    End Sub

    Private Function GetName(iFlagNo As Integer) As String
        'this function will return a string of the flag name
        'based on the picture number sent
        Dim strTempName As String
        Dim strCountry As String
        Dim strCap As String
        Dim blen As Byte

        'save length
        blen = Len(GameFlags(iFlagNo))
        'crop first letter
        strCap = GameFlags(iFlagNo).Substring(0, 1).ToUpper
        strTempName = GameFlags(iFlagNo).Substring(1, blen - 5)
        strCountry = strCap & strTempName
        Return strCountry
    End Function

End Class
