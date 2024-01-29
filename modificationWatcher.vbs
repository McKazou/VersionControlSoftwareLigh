Option Explicit
'On Error Resume Next

'------------ReUSE OTHER SCRIPT----------------
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
dim currentDir
currentDir = WshShell.CurrentDirectory
WScript.Echo "[modificationWatcher] Started from : "& currentDir
'To USE ME call the add a "include" sub inside the main script like the one at the bottom of this page
Include currentDir&"\Lib\DebugUtility.vbs"
Include currentDir&"\Lib\FileManipulation.vbs"

Dim debugTool
set debugTool = new DebugUtility


Dim CADFileWatcher 
set CADFileWatcher = new modificationWatcher
CADFileWatcher.addFolderRecursive("\\stccwp0015\Worksresearsh$\130_STELLANTIS\10_CAD\09 - Stellantis 48v - 2x6s1p")

'CLASS OBJECT : https://www.tutorialspoint.com/vbscript/vbscript_class_objects.htm
'call test()


Class modificationWatcher

    Private watchFolder()

    Public sub addFolderRecursive(newFolderPath)
        Dim fileManip
        set fileManip = new FileManipulation
        '--------------LISTING FILES----------------------
        Dim filesFound
        set filesFound = fileManip.FileFinder(newFolderPath, TRUE, 15)

        debugTool.startWrittingInConsole()
        debugTool.printTable(filesFound)
        debugTool.stopWrittingInConsole()
    end sub

    Public Sub addFolderToWatch(newFolderPath)
        watchFolder.Add(newFolderPath)
    End Sub

    Public Sub removeFolderToWatch() 
        watchFolder.Remove(newFolderPath)
    end sub

    '----------------EVENTS----------------
    Private Sub Class_Initialize(  )
        'Getting all processRunning at object creation
        Set watchFolder = CreateObject("System.Collections.ArrayList")
        Call ErrorHandler
    End Sub
    
    'When Object is Set to Nothing
    Private Sub Class_Terminate(  )
        'Termination code goes here
        Call ErrorHandler
    End Sub

end class

'---------TESTING-----------
function test()

    WScript.Echo "[TESTING] The script is run from: " & activefolder

end function

'--------------INCLUDE OTHER FILES----------------
Private Sub Include( scriptName )
    WScript.Echo "[LOADING]: "&scriptName
    Dim sScript
    Dim oStream
    With CreateObject( "Scripting.FileSystemobject" )
        Set oStream = .OpenTextFile(scriptName)
    End With
    sScript = oStream.ReadAll()
    oStream.Close
    ExecuteGlobal sScript
    WScript.Echo "[LOADED]: "&scriptName
    Call ErrorHandler
End Sub

Private Function ErrorHandler()
    If Err.Number <> 0 Then
        ' MsgBox or whatever. You may want to display or log your error there
        Call Err.Raise(vbObjectError + 10, "File Manipulation", "Unknow error with description: "&Err.Description)
        WScript.Echo("[ERROR] {File Manipulation} Unknow error with description: "&Err.Description)
        Err.Clear
    End If
end function
