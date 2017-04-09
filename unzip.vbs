'check for facets error report.  if found, unzip and copy error summary.txt to ERROR_SUMMARY folder

set fso = CreateObject("Scripting.FileSystemObject")
vZipCheck = 0

'create date string
vDay = DatePart("d", Date()-1)
  if len(vDay = 1) then vDay = "0" & vDay end if
vMonth = DatePart("m", Date()-1)
  if len(vMonth = 1) then vMonth = "0" & vMonth end if
vYear = DatePart("yyyy", Date()-1)
vFullDate = vMonth & "-" & vDay & "-" & vYear

'define path locations
vZipPath = "J:\NM BSA\NPM_Facets\PROD\PIMS To FACETS\2017\" & vFullDate & "_Smart File\TRZ Feedback\"
vExtractPath = vZipPath & "extract\"
vDropPath = "J:\NM BSA\NPM_Facets\PROD\PIMS To FACETS\2017\ERROR_SUMMARY\"

'check if folder exists
if not fso.FolderExists(vZipPath) then
   msgbox "ERROR:  '" & vZipPath & "' does not exist"
   wscript.quit
end if

'loop thru prior day folder to find zip file
for each f In fso.GetFolder(vZipPath).files
	if instr(f.name,".zip") > 0 then
		vZipFile = vZipPath & f.name
		vZipCheck = vZipCheck + 1
	end if
next

'check if zip file exists
if vZipCheck = 0 then
	msgbox "ERROR:  no zip file found at " & vZipPath
	wscript.quit
end if

'extract zip file contents
if not fso.FolderExists(vExtractPath) then
   fso.CreateFolder(vExtractPath)
end if

set objShell = CreateObject("Shell.Application")
set FilesInZip = objShell.NameSpace(vZipFile).items
objShell.NameSpace(vExtractPath).CopyHere(FilesInZip)

'copy error_summary.txt
for each f in fso.GetFolder(vExtractPath).files
	if f.name = "ERROR_SUMMARY.TXT" then
		fso.copyfile f.path, vDropPath & vFullDate & "_ERROR_SUMMARY.TXT"
	end if
next

set fso = nothing
set objShell = nothing