Imports System.IO
Imports System.Diagnostics

Module modificationWatcher

Sub Main()
  'call MainLoop()
End Sub

  Sub MainLoop()
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

Function getFilesListInFolder2(sourceFolder)
  'Recursive Function to get all files : https://devblogs.microsoft.com/scripting/how-can-i-get-a-list-of-all-the-files-in-a-folder-and-its-subfolders/
  
  Set objFSO = CreateObject(“Scripting.FileSystemObject”)
  objStartFolder = “sourceFolder”

  Set objFolder = objFSO.GetFolder(objStartFolder)
  Wscript.Echo objFolder.Path
  Set colFiles = objFolder.Files
  For Each objFile in colFiles
      Wscript.Echo objFile.Name
  Next
  Wscript.Echo
  ShowSubfolders objFSO.GetFolder(objStartFolder)
  Sub ShowSubFolders(Folder)
      For Each Subfolder in Folder.SubFolders
          Wscript.Echo Subfolder.Path
          Set objFolder = objFSO.GetFolder(Subfolder.Path)
          Set colFiles = objFolder.Files
          For Each objFile in colFiles
              Wscript.Echo objFile.Name
          Next
          Wscript.Echo
          ShowSubFolders Subfolder
      Next
  End Sub
End Function

Function getFilesListInFolder1(sourceFolder)
  'tentative using : http://www.thescarms.com/dotnet/listfiles.aspx
  'sourceFolder as a string : "\\stccwp0015\Worksresearsh$\130_STELLANTIS\10_CAD\09 - Stellantis 48v - 2x6s1p"
  
    Dim strFileSize As String = ""
    Dim di As New IO.DirectoryInfo(sourceFolder)
    Dim aryFi As IO.FileInfo() = di.GetFiles("*.txt")
    Dim fi As IO.FileInfo

    For Each fi In aryFi
        strFileSize = (Math.Round(fi.Length / 1024)).ToString()
        Console.WriteLine("File Name: {0}", fi.Name)
        Console.WriteLine("File Full Name: {0}", fi.FullName)
        Console.WriteLine("File Size (KB): {0}", strFileSize )
        Console.WriteLine("File Extension: {0}", fi.Extension)
        Console.WriteLine("Last Accessed: {0}", fi.LastAccessTime)
        Console.WriteLine("Read Only: {0}", (fi.Attributes.ReadOnly = True).ToString)
    Next
End Function

End Module

