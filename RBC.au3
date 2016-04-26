#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compression=4
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <string.au3>
#include <Array.au3>
#include <File.au3>
#include <TaskMasterBuild.au3>


## Declare paths as variables
#  Location of inkject mount point 
#  Should be a UNC path as a string.

$inkjetPath = "\\stash\inkjet\"

if $CmdLine[0] <> "" then
	$inFile = $CmdLine[1] ; If file dragged onto app, use that as VJ001 file.
else
	$inFile = FileOpenDialog("RBC VJ001 CSV","\\kciapp1\data\JOB LISTS\","CSV (*.csv) | All (*.*)") ; prompt user for code file to use
EndIf
dim $path_drive,$path_folder,$path_file,$path_ext
_PathSplit($infile,$path_drive,$path_folder,$path_file,$path_ext)

$continue = false
$jobnumber = ""
$prompt = ""
$inkjet = ""

while Not $continue
	$jobnumber = InputBox("Job Number",$prompt & "Please enter the Job Number:",$jobnumber) ; Prompt user for total number of codes to place into the 1UP file.
	$inkjet = $inkjetPath & StringRight($jobnumber,5) & ".1UP"
	$inkjetR = $inkjetPath & StringRight($jobnumber,5) & "R.1UP"
	if Not FileExists($inkjet) then
		$continue = True
	else
		$prompt = "The Job Number entered already exists. " & @CRLF
	EndIf
WEnd

$importTemplate = "RBC - VJ001.csv"
$inkjetFormat = "RBC VJ001 [no sort] - Inkjet"
$inkjetRevFormat = "RBC VJ001 [no sort] - Inkjet - Reverse"


$canada = MsgBox(4, "Canadian?", "Is this job Canadian?")
if ( $canada = 6 ) then
	$importTemplate = "RBC - VJ001.csv CANADA"
	$inkjetFormat = "RBC VJ001 [no sort] - Inkjet - CAN"
	$inkjetRevFormat = "RBC VJ001 [no sort] - Inkjet - CAN - Rev"
EndIf


$listpath ="\\kciapp2\apps\bcc\mm2010\lists\" & $jobnumber & " RBC.dbf"
$mjbname = 'zzz RBC VJ001 Automation'
$mjbpath = '\\kciapp2\APPS\BCC\MM2010\Jobs\' & $mjbname & '.mjb'
$jobname = $jobnumber & ' RBC'

buildtask()
$mailman = Runwait('\\kciapp2\APPS\BCC\MM2010\MailMan.exe -p -j "' & $mjbname & '" -u KYLE' ,"c:\")
ProcessWaitClose($mailman)
printForColleen()


func buildtask()
	ConsoleWrite( _FileCreate($mjbpath) )
	clearFile($mjbpath)
	NewList($listpath,$mjbpath,"   RBC - VJ001 [no sort]")
	Import($listpath,$infile,$mjbpath,$importTemplate)
	nonpresortedlabels($listpath,$mjbpath,$inkjetFormat, $inkjet)
	nonpresortedlabels($listpath,$mjbpath,$inkjetRevFormat, $inkjetR)
	PackageEnd($listpath,$mjbpath)
endfunc


func printForColleen()
	dim $listofcodes[10000]
	_FileReadToArray($inFile,$listOfCodes) ; reads entire file, breaks the file into an array with each element broken by a CRLF
	$totalCodes = $listOfCodes[0] ; the first element of the array returned from _FileReadToArray is the number of elements in the array
	$printout = FileOpen(@TempDir & "\temp.txt",10)
	FileWriteLine($printout,$jobnumber & " - RBC File Info")
	FileWriteLine($printout,"Original File Name: "& $path_file & $path_ext)
	FileWriteLine($printout,"")
	FileWriteLine($printout,"10 line file")
	FileWriteLine($printout,"")
	FileWriteLine($printout,"Layout:")
	FileWriteLine($printout,"line 1   :PRESORT SEQ")
	FileWriteLine($printout,"line 2   :ENTITY ID")
     FileWriteLine($printout,"line 3   :LABEL ADDRESS 1")
	FileWriteLine($printout,"line 4   :LABEL ADDRESS 2")
	FileWriteLine($printout,"line 5   :LABEL ADDRESS 3")
	FileWriteLine($printout,"line 6   :LABEL ADDRESS 4")
	FileWriteLine($printout,"line 7   :LABEL ADDRESS 5")
	FileWriteLine($printout,"line 8   :LABEL ADDRESS 6")
	FileWriteLine($printout,"line 9   :IMB")
	FileWriteLine($printout,"line 10  :OCR LINE (EntityScanID)")
	if ( $canada = 6 ) then
		FileWriteLine($printout,"line 11  :Bundle Number")
		FileWriteLine($printout,"line 12  :Container Number")
	EndIf
	FileWriteLine($printout,"")
	FileWriteLine($printout,"Inkjet File: "& StringRight($jobnumber,5) & ".1UP")
	FileWriteLine($printout,"Reversed Inkjet File: "& StringRight($jobnumber,5) & "R.1UP")
	FileWriteLine($printout,$totalCodes & " Total Records.")

	FileClose($Printout)

	Run("notepad.exe " & @TempDir & "\temp.txt")
EndFunc
