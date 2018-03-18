'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'			CLEAN DESKTOP, Folder/Directory Organising Script By Ben Cawley
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

'Clean Desktop is a script that allows people a quick and easy way to organise there messy folders
'Common examples of folders that get messy are Desktops, Documents, Cloud Drives and Downloads folders

'To Use
'Just drop the script into a directory/folder that you want organising and run the script.
'It will sort all your files and documents in to organised sub folders
'The folders that the files will go into is dependent on there file extension
'or in the case of apps and shortcuts it will organise them in to category's based on the app
'If the app is unknown what category it fits into then it is put into the app folder
'To reorganise how the files are sorted just modify the multi dimensional arrays bellow

'Also Now supports Hybrid Directory's (directory's that are made up of two or more other directory's) like desktop
'But you will have to populate the Hybrid Array yourself. It has Desktop there as a example


'TODO:
'	Make organising apps more efficient by searching for a keyword in its name rather then the whole name itself

'Changelog:
'	1.0.1: Fixed issue where if there is a . in the name file it throws the file into the unknown folder
'
'	1.0.2: Fixed issue where some files do not move because the directory is a hybrid directory and those files are in a different location
			'for example Desktop is a hybrid directory made up of C:\User\Username\Desktop and C:\User\Public\Destop.
			'the script currently only organises files in one of the directory's

'	1.0.3: 'Fixed issue when the script tries to move a file to a folder that already has a file with the same name


'MD array of Files and where they fit----------------------------------------------------------------------------------------------------------------
'Modify this array (arrayOfTypes) to add more types and extensions
	'Array() is a new category
	'The first index (index 0) is the name of the category
	'The rest of the index's are extensions, in lower case, of files that fit in to that category
	'the " _ " at the end of each line is just how you Break and Combine Statements in VBS so the MD array is easer to read

'Modafy the 2nd MD array (arrayOfApps) to organise apps in to there own folders

					'This Array is to organise folders by extension
					'Array("Folder Title", "file ext.","file ext."), _
arrayOfTypes = Array( _
					Array("Image","bmp", "jpg", "jpeg", "svg", "png", "tif", "tiff", "psd", "ai", "gif"), _
					Array("Video", "h264", "h265", "vob", "mp4", "mid", "mov", "avi", "mkv", "flv", "ogg", "mp2", "m4v", "3gp", "wmv", "rv"), _
					Array("Doc","doc", "docx", "odt", "ods", "exl", "xls", "xlxs", "txt", "ppt", "pptx", "rtf", "pdf"), _
					Array("Music", "mp3", "aac", "midi", "aiff", "m4b", "wma", "raw", "wav", "exs", "aa", "aax"), _
					Array("Script","ino", "py", "js", "php", "vb", "vbs", "c", "h", "c++", "cpp", "cs", "java", "class", "sh", "bat"), _
					Array("Web", "html", "css", "xml", "asp", "jsp", "rss", "xhtml"), _
					Array("App","exe", "run", "jar", "swf", "lnk"), _
					Array("Compressed","pkg", "zip", "rar", "shar", "iso", "mar", "tar", "bz2", "gz", "sfark", "7z", "dmg", "rev" ), _
					Array("Mobile","apk", "azw2", "swift"), _
					Array("CompiledFile", "bin", "dll", "o"), _
					Array("Installer", "msi"), _
					Array("Database", "db", "sql", "csv", "mdb", "dat"), _
					Array("Font", "fnt", "fon", "otf", "ttf") _
					)
					'This Array is to organise Apps by Name (and future keywords in name)
					'Array("Folder Title", "app name","app name"), _
arrayOfApps = Array( _
					Array("Media App","spotify", "itunes", "vlc media player" ), _
					Array("Admin Tools","msi afterburner", "putty (64-bit)", "putty", "geforce experience", "minitool partition wizard", "windirstat"), _
					Array("Productivity", "photoshop","open broadcaster software", "fl studio 12(64bit)", "fl studio 12", "fraps"), _
					Array("Social", "discord"), _
					Array("Anti Virus", "malwarebytes", "avast free antivirus", "avast safeZone browser","avast_free_antivirus_setup_online"), _
					Array("IDE", "arduino","notepad++", "git bash", "unity", "eclipse"), _
					Array("Web Browser", "google chrome","chromium", "edge", "microsoft edge", "brave", "firefox", "opera", "safari", "internet explorer") _
					)

					'This Array is grab the directory's that make up a known hybrid directory
					'Array("Hybrid Folder Name","Directory to other folder", "Directory to other folder")
ArrayOfHybridFolders = Array( _
					Array("Desktop","C:\Users\Public\Desktop") _
					)

					'=================================================================
'-------------------------------------- No need to edit past here ------------------------------------
					'=================================================================
' Main Body
'Sets up a file system object and finds out what directory we are in
Set FSO = CreateObject("Scripting.FileSystemObject")
dir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

'Check if its a hybrid directory
Dim isHybrid : isHybrid = CheckIfHybridFolder(FSO.GetBaseName(dir))
'if it is not a Hybrid folder then just clean the directory
If IsEmpty(isHybrid) Then
	Clean(FSO.GetFolder(dir).Files)
Else
'Else it is a Hybrid so go through all the directory's and clean
	Clean(FSO.GetFolder(dir).Files) 'Clean the current part of the Hybrid directory first
	Dim hybridArray : hybridArray = isHybrid 'This is only just to make it less confusing to read
	Dim i : i = 1	'setting the index to 1 as 0 is the name of the directory not its path
	For i=1 to UBound(hybridArray)					'for i to the length of the hybrid array (the number of known directory's that make up this hybrid directory)
		Clean(FSO.GetFolder(hybridArray(i)).Files)	'clean the directory
	Next
End If

'The main sorting function
Function Clean(files)
	'For each file in the folder get the ext of the file then compare it with the arrays above to work out what category of file it is then move the file into a folder of that category name
	For Each file in files
		If Not file.Name = WScript.ScriptName Then 'This is to prevent moving this script during the process
			'Unknown
			Dim ext : ext = LCase(StripExt(file.Name))		'Finds the ext of the current file (also converts the text to lower case so that we don't have to put the same ext twice in the arrays like HTML and html for example)
			Dim fType : fType = FileType(ext,arrayOfTypes)		'Works out the type of file by compeering its ext to the data in the Multidimensional array above if its a match it will return the first value of the 2nd array as the type
			If fType = "App" Then
				Dim fName : fName = LCase(FSO.GetBaseName(file))
				Dim appType : appType = FileType(fName,arrayOfApps)
				If appType = "Unknown" Then
					appType = "App"
				End If
				CheckForFolder(appType)
				MoveFile file, dir+"\"+appType+"\"
				'file.Move(dir+"\"+appType+"\")
			Else
				CheckForFolder(fType)					'Checks to see if the folder to move the file to exists if not create it
				MoveFile file, dir+"\"+fType+"\"
				'file.Move(dir+"\"+fType+"\")			'Move the file to the folder
			End If
		End If
	Next
End Function

'Function to check if the folder is a known hybrid directory and return the object if it is --------------------------------------
Function CheckIfHybridFolder(folderName)
	Dim isHybrid
	For Each hybridFolder in ArrayOfHybridFolders
		If folderName = hybridFolder(0) Then
			isHybrid = hybridFolder
		End If
	Next
	CheckIfHybridFolder = isHybrid
End Function

'Function to check if folder exists and creates it if false--------------------------------------------------
Function CheckForFolder(folderName)
	Dim exists : exists = FSO.FolderExists(dir+"\"+folderName)
	If Not exists Then
		fso.CreateFolder(dir+"\"+folderName)
	End If
End Function

'Function To check for a file ----------------------------------------------------------------------------------
Function CheckForFile(targit, destination)
	Dim fileInfo : fileInfo = Array("False", "FileName", "Dir")
    For Each file in FSO.GetFolder(destination).Files
		if file.Name = targit.Name Then
			fileInfo(0) = "True"
			fileInfo(1) = file.Name
			fileInfo(2) = destination
		End If
	Next
	CheckForFile = fileInfo
End Function

'Function to strip file name from extension -----------------------------------------
Function StripExt(fileName)
	Dim tempExt : tempExt = "NULL"
	If InStr(fileName,".") Then
		tempExt=TRIM(Right(fileName,Len(fileName) - InStrRev(fileName,".")))
	End If
	StripExt = tempExt
End Function

'Function to Work out File Type or App ---------------------------------------------------------
Function FileType(fileExt, arrayGroup)
	Dim tempType : tempType = "Unknown"

	If fileExt = "NULL" Then
		tempType = "None"
	Else
		Dim typeFound : typeFound = False
		Dim isIn : isIn = False
		Dim i : i = 0
		For i=0 to UBound(arrayGroup)
			Dim j : j = 0
			For j=0 to UBound(arrayGroup(i))
				If arrayGroup(i)(j) = fileExt Then
					isIn = True
					tempType = arrayGroup(i)(0)
					typeFound = True
					Exit For
				End If
			Next
			If typeFound Then
				Exit For
			End If
		Next
	End If
	FileType = tempType
End Function

'Function My own move file function
'Not to reinvent the wheel but because VBS built in Method to move files dose not handle overwriting files
'and the copy file method needs to know before hand if the user wants to overwrite the file or not
' -----------------------------------------------
Function MoveFile(targit, destination)
	Dim fileExists : fileExists = CheckForFile(targit, destination)' remove last / from destanation
	Dim msgBoxAns
	If fileExists(0) = "True" Then
		msgBoxAns =	Msgbox ("WARNING!" & vbcrlf & vbcrlf & "There is already a file called " & fileExists(1) & " In the folder " & vbcrlf & fileExists(2) & vbcrlf & "Do you want to overwrite it?" ,20, "Ooh no")
		If msgBoxAns = 6 Then
			FSO.CopyFile targit, destination, True
			FSO.DeleteFile(targit)
		End IF
	Else
		FSO.CopyFile targit, destination, True
		FSO.DeleteFile(targit)
	End If
End Function
