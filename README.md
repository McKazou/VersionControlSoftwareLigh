# VersionControlSoftwareLigh
Version Control App, managing a specified folder, and writing change into Excel file so it's light, easy to use for anyone.

## Use Case
The software will be able to track change done to files and summarize them into a more readable way. 
1. I want a program running in the background
2. I want the program to ask, whenever i save a .prt file so i can add a commit message of why i did the modification.
- If it's only a sauvegarde on top of the saved part; simply update the last modified date and add the commit message to the previous commit message added.
- If i "Save As" to add a new version of a file it should be tracked as an new "children" of a previous one
         => so i can have a list of all version of a specified part
         It should set the all previous version to a readonly state
3. I want the software to summarize all committed message into different .csv files; one for each parts (I should have a way to ignore some files)
```
<fileNamePath>,<partName>,<dateModification>,<commitMessage>,<dateCommit>
```

On the excel:
1. A way to see:
   - A list of what changed compared between two dates
   - A list of each revision for each part
2. An excel will read those csv files to print them into an excel. The Excel should containt: 
      - The part miniature
      - The Part name
      - The last version (AA_AA06)
      - Status ("Archived"(marron), "Work in Progress"(yellow), "Studies"(blue), "invalid"(red)
      - The commit message
      - A go to link
3. Can it have the tree ?

## Tracked files names :
1. I want to track different files base on the same syntaxe;
   a. Syntaxe 1 : BAT110000013-AA_AA76_K0_X250-Battery_Pack_48V_Assembly_230302 => BAT\d{9}-([A-Za-z]{2})_([A-Za-z]{2})\d{2}_.*_\d{6}(_.*)?
      See : https://regex101.com/delete/2ZeCuWFbc2wwsC2XXlE6k4XQ
   b. Syntaxe 2 : STEL9_IN_CONTEXT_230302 => \w*\d*_.*_\d{6}
      See : https://regex101.com/delete/yJrQqk5Nxv6DKhhO5TW5TPWG
      
I've found a script based on windows log : https://stackoverflow.com/questions/10041057/how-to-monitoring-folder-files-by-vbs that does the watching part.
Another : https://social.technet.microsoft.com/Forums/scriptcenter/en-US/6064064e-d273-41c1-9135-cc883735a465/vbscript-monitor-a-folder-and-sub-sub-folders?forum=ITCG

## Manual version : 
1. [x] Look into one or multiple given folder(s). 
2. [ ] List all folders in the targeted folder(s)
3. [ ] On a time based trigger it will update the folder list
4. [ ] Add a listener on each folder
5. [ ] On event start get what file have been modify or created
6. [ ] It will store the modified filed into a dictionnary based on the saved date
7. [ ] Prompt a message to add a commit message; into the systemTray (https://omen999.developpez.com/tutoriels/vbs/DynamicWrapperX/) for 30s then hide
   - if it's a saved file on itself => add the commit message to the previous one
   - if it's a new file (new indice) => set the previous file as read only
8. [ ] It will prompt a message "(5) files have been modified, do you want to specified Commit message now ?"
9. [ ] Clicking on it will show a UI listing any modified files
   - you will be able to add a commit message to specified "why"
   - you will be able to ignore this file
   - you will be able to save the list with a general commit message
  ```
  "<FileName>                         <insert a commit message>
     |<Date1> has been modified :"    <insert a commit message>
     |<Date2> has been modified :"    <insert a commit message> 
     |<Date...> has been modified :"  <insert a commit message>" 
  ```

