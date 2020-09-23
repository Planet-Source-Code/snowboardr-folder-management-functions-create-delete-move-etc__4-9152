<div align="center">

## Folder Management Functions \(Create, Delete, Move etc\.\.\.\)


</div>

### Description

Functions to Create folder, Delete Folder, Move Folder, Copy Folder, Check if File Exists, and folder size in bytes
 
### More Info
 
Server.MapPath() to get the current directory so you would run the Create Folder function Like so

CreateFolder(Server.MapPath("\test"))

DeleteFolder Will remove everything within that folder... and becarefull you dont delete the wrong directory.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[snowboardr](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/snowboardr.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/snowboardr-folder-management-functions-create-delete-move-etc__4-9152/archive/master.zip)





### Source Code

```
<%
'#######################################
'# Folder Management Functions
'# Written By Jason
'# '#####################################
'# Site San Diego
'# sitesd.com
'#######################################
	Dim objFso
	'# CREATE FOLDER
	Function CreateFolder(fdirCRE)
			Set objFso = CreateObject("Scripting.FileSystemObject")
			If Not objFso.FolderExists(fdirCRE) Then
				objFso.CreateFolder(fdirCRE)
					Else
				Exit Function
			End If
			set objFso = nothing
	End Function
	'# DELETE FOLDER
	Function DeleteFolder(fdirDEL)
		Set objFso = CreateObject("Scripting.FileSystemObject")
		If objFso.FolderExists(fdirDEL) Then
			 objFso.DeleteFolder(fdirDEL)
		End If
		set objFso = nothing
	End Function
	'# MOVE FOLDER
	Function MoveFolder(fdirM1, fdirM2)
		Set objFso = CreateObject("Scripting.FileSystemObject")
		objFso.MoveFolder fdirM1, fdirM2
		set objFso = nothing
	End Function
	'# COPY FOLDER
	Function CopyFolder(fdirC1, fdirC2, fdirC3)
		If fdirC3 = "" then fdirC3 = False
		Set objFso = CreateObject("Scripting.FileSystemObject")
		objFso.CopyFolder fdirC1, fdirC2, fdirC3
		set objFso = nothing
	End Function
	'# CHECK IF FILE EXSISTS
	Function FileHere(fdirHere)
			Set objFso = CreateObject("Scripting.FileSystemObject")
			If objFso.FileExists(fdirHere) Then
				FileHere = True
					Else
				FileHere = False
			End If
			set objFso = nothing
	End Function
		'# Folder Size
		Function FolderSize(filespec)
		  Set fso = CreateObject("Scripting.FileSystemObject")
			  Dim fso, f, s
			If fso.FolderExists(filespec) Then
				  Set f = fso.GetFolder(filespec)
				  s = f.size
				  If s = "" then s = "0"
			  Else
				  s = "0"
			 End If
			  FolderSize = CLNG(s)
		  Set fso=nothing
		End Function
%>
```

