Public Class Form1
    Dim word As String
    Dim display As String

    Private Sub readWordFromFile()
        'Σαυτή τη διαδικασία ανοίγουμε το αρχείο c:\words.txt και διαβάζουμε όλες τις γραμμές του
        'Ταυτόχρονα μετράμε πόσες φορές έγινε αυτό και το αποθηκεύουμε στη μεταβλητή NumberOfWords
        'Στη συνέχεια κλείνουμε το αρχείο & παράγουμε έναν τυχαίο αριθμό num στο διάστημα [1, NumberOfWords]
        'Ξανα-ανοίγουμε το αρχείο c:\words.txt & διαβάζουμε τη λέξη που αντιστοιχεί στον τυχαίο αριθμό num
        'Τη λέξη αυτή την εμφανίζουμε στο πλαίσιο LblWord & κλείνουμε το αρχείο

        Dim NumberOfWords As Integer 'Μετρά το πλήθος των γραμμών του αρχείου
        Dim num As Integer 'Κρατά τον τυχαίο αριθμό

        'Ανοιγμα του αρχείου για ανάγνωση με χειριστή αρχείου τον αριθμό 1
        FileOpen(1, "c:\words.txt", OpenMode.Input)

        'Οσο στο αρχείο με χειριστή 1 ΔΕΝ έχει γίνει ανάγνωση του σημείου End Of File
        'δηλαδή όσο στο αρχείο υπάρχουν δεδομένα, διαβάζουμε γραμμές
        Do While Not EOF(1)
            LineInput(1) 'γίνεται ανάγνωση γραμμής από το 1 χωρίς να τοποθετούμε τίποτα σε μεταβλητή
            NumberOfWords = NumberOfWords + 1 'αυξάνουμε τον μετρητή κατά 1 μονάδα
        Loop
        FileClose(1) 'κλείνουμε το αρχείο με χειριστή 1

        'παράγουμε έναν τυχαίο αριθμό στο διάστημα [1, NumberOfWords]
        'όπου NumberOfWords είναι το πλήθος των γραμμών-λέξεων στο αρχείο
        num = Int(Rnd() * NumberOfWords + 1)

        FileOpen(1, "c:\words.txt", OpenMode.Input) 'ανοίγουμε πάλι το αρχείο με χειριστή πάλι 1
        For I = 1 To num - 1
            LineInput(1) 'διαβάζουμε πάλι όλες τις γραμμές μέχρι τη γραμμή (num-1)
        Next
        Input(1, word) 'διαβάζουμε τη λέξη που βρίσκεται στη num γραμμή
        display = MakeDisplay(word)
        LblWord.Text = display 'τοποθετούμε τη λέξη στο πλαίσιο LblWord
        FileClose(1) 'κλείνουμε το αρχείο με χειριστή 1
    End Sub

    Private Sub Bt1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt1.Click, bt2.Click, bt3.Click, bt4.Click, bt5.Click, bt6.Click, bt7.Click, bt8.Click, bt9.Click, bt10.Click, bt11.Click, bt12.Click, bt13.Click, bt14.Click, bt15.Click, bt16.Click, bt17.Click, bt18.Click, bt19.Click, bt20.Click, bt21.Click, bt22.Click, bt23.Click, bt24.Click
        Dim c As Byte, found As Boolean
        found = False 'σημαία που γίνεται true όταν το γράμμα letter βρεθεί στη λέξη w
        Dim ctl As Control
        For c = 2 To Len(word) - 1
            If Mid(word, c, 1) = sender.text Then
                display = LSet(display, c - 1) & sender.text & Microsoft.VisualBasic.Right(display, Len(display) - c)
                LblWord.Text = display
                'το νέο display φτιάχνεται ενώνοντας τα κομμάτια: left (μέχρι τη θέση του γράμματος letter)
                'το γράμμα letter που βρέθηκε ΚΑΙ 
                'το right(παίρνοντας χαρακτήρες από το τέλος της λέξης display μέχρι το γράμμα letter
                found = True
            End If
            sender.enabled = False
        Next
        If display = word Then
            MsgBox("Μπράβο! Βρήκες τη λέξη!")
            For Each ctl In Me.Controls
                If TypeOf ctl Is Button Then
                    ctl.Enabled = False
                    btEnd.Enabled = True
                End If
                btLoadWord.Enabled = True
            Next
        End If
        If found = False Then
            TextBox1.Text -= 1
        End If
        If TextBox1.Text = 0 Then
            MsgBox("Δυστυχώς έχασες :( , η λέξη ήταν: " & word)
            TextBox1.Text = 7
            For Each ctl In Me.Controls
                If TypeOf ctl Is Button Then
                    ctl.Enabled = False
                    btEnd.Enabled = True
                    btLoadWord.Enabled = True
                End If
            Next
        End If
    End Sub

    Public Function MakeDisplay(ByVal w As String) As String 'Φτιάχνει τη λέξη της κρεμάλας με παυλίτσες στα ενδιάμεσα
        Dim display As String, c As Byte
        display = LSet(w, 1)
        For c = 2 To Len(w) - 1
            display = display & "-"
        Next
        display = display & Microsoft.VisualBasic.Right(w, 1)
        Return display
    End Function

    Private Sub btLoadWord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btLoadWord.Click
        TextBox1.Text = 7
        Call readWordFromFile()
        For Each ctl In Me.Controls
            If TypeOf ctl Is Button Then
                ctl.Enabled = True
            End If
        Next
    End Sub

    Private Sub btEnd_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btEnd.Click
        Call writeNameToFile()
        TextBox1.Text = 7
        End
    End Sub

    Private Sub writeNameToFile()
        'Ζητά ένα όνομα και το προσθέτει στα περιεχόμενα του αρχείου c:\players.txt
        'μαζί με την ημερομηνία και ώρα του συστήματος εκείνη τη χρονική στιγμή

        Dim yourName As String 'Το όνομα του παίκτη που θα αποθηκευτεί στο αρχείο

        'Ανοιγμα του αρχείου για εγγραφή με προσθήκη (Append)
        'Αν το ανοίξουμε με OpenMode.Output χάνουμε τα περιεχόμενα που είχε μέχρι τώρα
        FileOpen(5, "c:\players.txt", OpenMode.Append)

        yourName = InputBox("Δώσε το όνομά σου", "Name", "Όνομα..")
        WriteLine(5, yourName, Now)

        FileClose(5) 'κλείνουμε το αρχείο με χειριστή 2
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Call readWordFromFile()
    End Sub
End Class