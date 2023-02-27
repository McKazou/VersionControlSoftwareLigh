# VersionControlSoftwareLigh
Version Control App, managing a specified folder, and writing change into Excel file so it's light, easy to use for anyone.

The idea behind this is to have a small programm running in the background.

Manual version : 
Looking into one or multiple folder. 
  On a time based trigger it will start looking into each folder.
    In each folder it will list very single file in it.
    Each time a file is saved (last save date > last reccorded save date)
      It will store the modified filed into a dictionnary based on the saved date

At the end of the query, it will prompt a notification if a file has been modified.
(1) It will prompt a message "(5) files have been modified, do you want to specified Commit message now ?"
Clicking on it will show a UI listing any modified files
   | you will be able to add a commit message to specified "why"
   | you will be able to ignore this modification
   | you will be able to save the list with a general commit message
  "<FileName>                         <insert a commit message>
     |<Date1> has been modified :"    <insert a commit message>
     |<Date2> has been modified :"    <insert a commit message> 
     |<Date...> has been modified :"  <insert a commit message>" 
Each commited files will be set as read only
Each commit message will be added to an excel file with : 
<miniature> <fileName> <dateModification> <commitMessage> <dateCommit>

Git version:
Git is a powerfull version control tool but it handles binary file really badly : https://stackoverflow.com/questions/4697216/is-git-good-with-binary-files
It mean Git have to ignore merge cmd

What can be done with git : 
  - Run an app in background running the bast command to use GIT
  - GIT : Have a list of the modified file since the last log  and print it to a file
    based cmd: >-- git log<
    | list of last modified log 
    | print the log to a csv file       => git log --pretty=format:%h,%an,%ae,%s > /path/to/file.csv  (source:https://stackoverflow.com/questions/39253307/export-git-log-into-an-excel-file)
   - Have an app read the CSV file
    If modified files exist withing the time period
     | Go to (1)
     modify&format the log using excel => powerquery
