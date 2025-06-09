'TODO:  Repair, salvage, get new screen (same as create pretty much)
'TODO:  Use real universe unit names
'TODO:  show (%/%) strength for %salvaged/%full units in that unit
'TODO:  Mercenaries Handbook 3055:
'       p110, breach of contract hearing table.  I can make it so contracts cannot be attained for a period of time.
'           2d6, modifier based on how bad they violate the contract, takes 1 week to determine:
'               2-5 = 20K Fine, 1 Month Hiring Ban
'               6-9 = 25K Fine, 3 Month Hiring Ban
'               10-12 = 75K Fine, 6 Month Hiring Ban
'               13-15 = 125K Fine, 12 Month Hiring Ban
'               16-18 = 200K Fine, 18 Month Hiring Ban
'               19-20 = 500K Fine, 24 Month Hiring Ban
'               21-22 = 900K Fine, 36 Month Hiring Ban
'               23 = 1500K Fine, 60 Month Hiring Ban
'               24+ = 4000K Fine, unit disbanded and each member under 60 Month Hiring Ban
'                   Employers breaching contract can pay up to 10x the contract total in fines
'TODO:  Pay 85 C-Bills/month to have name on contract registry (p12, Hot Spots Rules)
'TODO:  Merc3055, p22, I could add an 'R' to add 'Random Lances'
'TODO:  Hiring Halls:  Galatea, Outreach, Solaris VII, Antallos, Astrokaszy, Herotitus, Noisiel, Westerhand
'TODO:  FMMercs, p136 - create a leader, then assemble founding members, then recruit/experience, then get equipment, then get DS/JS, then get payroll/maintenance/support, then create a war chest
'TODO:  FMMercsRev, p137 - create leader, run a path, recruiting, choose hiring hall, combat experience, get DS/JS, then get payroll/maintenance/support, then create a war chest
'TODO:  FMMercsSup, p79 - alternative point/less random based force gen
'TODO:  FMMercsSupII, p89, FMMercsRev additions
'TODO:  Create: When clicking on a unit on the left, select it on the right
'TODO:  Negotiate:  Instead of just 'next contract', list out all the contracts available this month.  Must click advance to see new ones.




Public Class frmMain
    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        End
    End Sub

    Private Sub cmdNew_Click(sender As Object, e As EventArgs) Handles cmdNew.Click
        frmCharacter.Visible = True
        Me.Visible = False
    End Sub

    Private Sub cmdContinue_Click(sender As Object, e As EventArgs) Handles cmdContinue.Click
        Try
            Dim FileList As New SortedDictionary(Of String, Date)
            For Each File As String In System.IO.Directory.GetFiles(Application.StartupPath & "Saves", "*.txt")
                FileList.Add(File, System.IO.File.GetLastWriteTime(File))
            Next
            If FileList.Count = 0 Then
                MsgBox("There were no save files to continue.", vbOKOnly)
                Exit Sub
            Else
                Dim LatestFile As String = (From entry In FileList Order By entry.Value Descending Select entry)(0).Key
                Leader = New cCharacter
                frmCreate.LoadData(Mercenary, LatestFile.ToLower.Replace(".txt", ""))
                frmCreate.ContinueMerc()
            End If
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try

    End Sub

    Private Sub cmdLoad_Click(sender As Object, e As EventArgs) Handles cmdLoad.Click

        Try
            ofd.InitialDirectory = Application.StartupPath & "Saves"
            If ofd.ShowDialog() = DialogResult.OK Then
                frmCreate.LoadData(Mercenary, ofd.FileName)
            End If
            Me.Visible = True
            frmOperations.Show()
        Catch ex As Exception
            Debug.Print(ex.Message.ToString)
        End Try
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'for testing functions
    End Sub
End Class
