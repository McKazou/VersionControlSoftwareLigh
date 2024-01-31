
#Get Started : 
#https://stackoverflow.com/questions/10137146/is-there-a-way-to-make-a-powershell-script-work-by-double-clicking-a-ps1-file
#Create a shortcut with : powershell.exe -command "& 'C:\A path with spaces\MyScript.ps1' -MyArguments blah"

try{
Get-ExecutionPolicy
Write-Host -NoNewLine 'Press any key to continue...';
}
catch {
    "An error occured"
}
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');

#First thing first need to check the powershell version
#$PSVersionTable.PSVersion this will give the version as a table (source: https://learn.microsoft.com/en-us/training/modules/introduction-to-powershell/3-exercise-powershell)


#Powershell is using "Get-Command" "get" is the action "Command" is the noun

#   [] - Create a folder listener who triggers some function on "Create", "Delete" or "edit" a file
#   [] - Make the listener listen any folder recursively
#   [] - Create a class to create an icon in the icon tray
#       [] - (Optionnal) Get a  custom icon
#       [] - Right clic function
#           [] - Exit : will kill the program
#           [] - Status : 
#                   'Running' : clic on it = paused the programm and so changed contextual menu to reflect
#                   'Paused' : clic on it runs it and so changed contextual menu to reflect
#                   'Error' : show the errors throw : clic on it brings more info
#           [] - Parameters
#                   get an interface with :
#                       Two parts:
#                           - What i've selected ([DeleteButton] [EditButton] - Name - recursif or not - path): 
#                              [x][|] Stel09 - Récursif - \\stccwp0015\Worksresearsh$\130_STELLANTIS\10_CAD\09 - Stellantis 48v - 2x6s1p - logfilePath
#                              [x][|] Chiro1 - non-récursif - \\stccwp0015\Worksresearsh$\180_PHEV Cherokee\03-3D\00 - RFI\Version 1 - logfilePath
#                               [|] add a folder
#                           - What folder the programm see:
#                               (Optional)> insert maybe a way to filter with (*mask*) the path column<
#                               >insert all possible folder name - path<
#                   [Button] Save Folders - Store the list (what about multiple users?)
#                   [Option] to register what happen in a log file
#                   [Button] Open program log file
#                   [Option] Regroup files save within 'x' minutes
#                   [Button] Save Parameters
#
#                               
