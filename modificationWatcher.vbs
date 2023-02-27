Imports System.IO
Imports System.Diagnostics

Module Module1

    Sub Main()
        While True
            ' Get the git log output
            Dim gitLogProcess As New Process()
            gitLogProcess.StartInfo.FileName = "git"
            gitLogProcess.StartInfo.Arguments = "log --pretty=format:%h,%s,%ad"
            gitLogProcess.StartInfo.UseShellExecute = False
            gitLogProcess.StartInfo.RedirectStandardOutput = True
            gitLogProcess.Start()

            Dim gitLogOutput As String = gitLogProcess.StandardOutput.ReadToEnd()

            gitLogProcess.WaitForExit()

            ' Split the output into lines
            Dim gitLogLines() As String = gitLogOutput.Split(Environment.NewLine)

            ' Open the CSV file for writing
            Dim csvFileName As String = "git_log.csv"
            Dim csvFileWriter As New StreamWriter(csvFileName)

            ' Write the header row
            csvFileWriter.WriteLine("Commit ID,Message,Date")

            ' Write each line to the CSV file
            For Each line As String In gitLogLines
                If line.Trim() <> "" Then
                    Dim lineParts() As String = line.Split(",", 3)
                    Dim commitId As String = lineParts(0).Trim()
                    Dim message As String = lineParts(1).Trim()
                    Dim dateStr As String = lineParts(2).Trim()

                    csvFileWriter.WriteLine(commitId & "," & message & "," & dateStr)
                End If
            Next

            csvFileWriter.Close()

            ' Wait for 15 minutes before running the script again
            Threading.Thread.Sleep(15 * 60 * 1000)
        End While
    End Sub

End Module

