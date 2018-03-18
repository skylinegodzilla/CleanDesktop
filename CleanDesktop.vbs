'-------------------------------------------------------------------------------------------------------------------------------------------------------------
'			CLEAN DESKTOP, Folder/Directorey Organising Script By Ben Cawley
'-------------------------------------------------------------------------------------------------------------------------------------------------------------

'TODO:
	'Fix issue when the script tryes to move a file to a folder that allreadey has a file with the same name
	'Fix issue where some files do not move because the directorey is achaley a hybrid directory and those files are in a different location
		'for example Desktop is a hybrid Dorectorey made up of C:\User\Username\Desktop and C:\User\Public\Destop. 
		'the script currentley onley organises files in one of the directoreys
'	Make organising apps more efficent by searching for a keyword in its name rather then the whole name itself
'Changelog:
'	1.0.1: Fix issue where if there is a . in the name file it throws the file into the unknowen folder

'MD array of Files and where they fit----------------------------------------------------------------------------------------------------------------
'Modafy this array (arrayOfTypes) to add more types and extentions
	'Array() is a new catagory
	'The first index (index 0) is the name of the catagorey
	'The rest of the indexs are extentions, in lower case, of files that fit in to that catagorey
	'the " _ " at the end of eatch line is just how you Break and Combine Statements in VBS so the MD array is easer to read

'Modafy the 2nd MD array (arrayOfApps) to organise apps in to there own folders
					
					'Array("Folder Title", "file ext","file ext"), _ 
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
					
					
					'=================================================================
'-------------------------------------- No need to edit past here ------------------------------------
					'=================================================================
' Main Body
'Sets up a file system object and finds out what directory we are in
Set FSO = CreateObject("Scripting.FileSystemObject")	
dir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

'Gets infomation about the current folder and its files
Set folder = FSO.GetFolder(dir)
Set files = folder.Files

'For eatch file in the folder get the ext of the file then compear it with the arrays above to work out what catagorey of file it is then move the file into a folder of that catagory name
For Each file in files
	If Not file.Name = WScript.ScriptName Then 'This is to prevent moving this script during the process
		'Unknown
		Dim ext : ext = LCase(StripExt(file.Name))		'Finds the ext of the current file (also converts the text to lower case so that we dont have to put the same ext twice in the arrays like HTML and html for example) 
		Dim fType : fType = FileType(ext,arrayOfTypes)		'Works out the type of file by compearing its ext to the data in the Muitidimentional array above if its a match it will return the first value of the 2nd array as the type
		If fType = "App" Then
			Dim fName : fName = LCase(FSO.GetBaseName(file))
			Dim appType : appType = FileType(fName,arrayOfApps)
			If appType = "Unknown" Then 
				appType = "App"
			End If
			CheckForFolder(appType)
			file.Move(dir+"\"+appType+"\")
		Else
			CheckForFolder(fType)					'Checks to see if the folder to move the file to exists if not create it
			file.Move(dir+"\"+fType+"\")			'Move the file to the folder
		End If
	End If	
Next

'Function to check if folder exsists--------------------------------------------------
Function CheckForFolder(folderName)
	Dim exists : exists = FSO.FolderExists(dir+"\"+folderName)
	If Not exists Then
		fso.CreateFolder(dir+"\"+folderName)
	End If
End Function


'Function to strip file name from extention -----------------------------------------
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