;~ Please note that this library uses a $linesCounter and a $TasksCounter to maintain TASKMASTER compatability, please do not alias these in your code.
;~ Please note that because of the way the library currently handles file writing we are unable to write tasks out of order, this means that you will need to call the write functions in the order in which you wish to have them preformed in Mail Manager

;INCLUDES
#include <File.au3>
#include <ComboConstants.au3>
#include<string.au3>
#include<Array.au3>
#include <GuiConstantsEx.au3>
#include <AVIConstants.au3>
#include <TreeViewConstants.au3>
#include <Debug.au3>
#include <INet.au3>

$LinesCounter = 1
$TasksCounter = 1
Func ClearFile($_filepath)
	$xyzc = 1
    while $xyzc < 1000
		FileWriteLine($_filepath,"")
		$xyzc += 1
	WEnd
    $xyzc = 1
	while $xyzc < 1000
		_FileWriteToLine($_filepath,$xyzc,"" & @CRLF ,1)
		$xyzc += 1
	WEnd
EndFunc
Func NewList($_listpath,$_mjbpath,$_template = "!n-c-a-a-csz")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[NEWLISTTEMPLATE-' & $TasksCounter & ']' & @CRLF &'SETTINGS="' & $_template & '"' & @CRLF &'FILENAME=' & '"' & $_listpath & '"' & @CRLF &'OVERWRITE=Y' & @CRLF &'PARALLEL=N' & @CRLF & 'USEINDEXES=Y' & @CRLF,1)
	$LinesCounter = $LinesCounter + 6
	$TasksCounter = $TasksCounter + 1
EndFunc
Func GRPressDBF($_listpath,$_filepath,$_mjbpath,$_listlayout)
	_FileWriteToLine($_mjbpath,$LinesCounter,'[IMPORT-' & $TasksCounter & ']' & @CRLF & 'LIST=' & '"' & $_listpath & '"' & @CRLF &'SETTINGS=' & $_listlayout & @CRLF &'FILENAME='  & '"' & $_filepath & '"' & @CRLF &'PARALLEL=N' & @CRLF &'WAIT=DEFAULT' & @CRLF &'STARTTIME=DEFAULT' & @CRLF &'SUPPRESSERRORS=N' & @CRLF ,1)
	$LinesCounter = $LinesCounter + 8
	$TasksCounter = $TasksCounter + 1
EndFunc
Func Encode($_listpath, $_mjbpath, $_jobname,$_selectivity="NONE")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[ENCODE-' & $TasksCounter & ']' & @CRLF & 'LIST="' & $_listpath & '"' & @CRLF & 'SELECTIVITY=' & $_selectivity & @CRLF & 'ADDRESSGROUPS="MAIN"' & @CRLF & 'SWAP=Y' & @CRLF & 'STANDARDIZEADDRESS=Y' & @CRLF & 'STANDARDIZECITY=Y' & @CRLF & 'ABBREVIATECITY=N' & @CRLF & 'IGNORENONUSPS=N' & @CRLF & 'EXTENDEDMATCHING=N' & @CRLF & 'CASE="MIXED"' & @CRLF & 'FIRMASIS=Y' & @CRLF & 'ZIP5CHECKDIGIT=N' & @CRLF & 'SUMMARYPAGE=N' & @CRLF & 'NDIREPORT=N' & @CRLF & 'PRINTER="\\kcimail2\HP LaserJet 5Si MX"' & @CRLF & 'FILENAME="\\kciapp1\DATA\LISTS\GR Press\' & $_jobname & '-PS3553.prn"' & @CRLF & 'OVERWRITE=N' & @CRLF & 'COPIES=1' & @CRLF & 'PARALLEL=N' & @CRLF & 'WAIT=DEFAULT' & @CRLF & 'STARTTIME=DEFAULT' & @CRLF & 'SUPPRESSERRORS=N' & @CRLF,1)
	$LinesCounter = $LinesCounter + 23
	$TasksCounter = $TasksCounter + 1
EndFunc
Func NCOA($_listpath,$_mjbpath)
	_FileWriteToLine($_mjbpath,$LinesCounter,'[DATASERVICES-' & $TasksCounter & ']' & @CRLF & 'LIST="' & $_listpath & '"' & @CRLF & 'PRINTER="\\kylec8\Adobe PDF"' & @CRLF & 'PROCESSES="FSP"' & @CRLF & 'PREPAID=Y' & @CRLF & 'EXTENDEDMATCHING=Y' & @CRLF & 'PAFELECTRONIC=N' & @CRLF & 'BYPASSNAMEMISSING=Y' & @CRLF & 'BYPASSADDRESSMISSING=Y' & @CRLF & 'BYPASSSTATEINVALID=Y' & @CRLF & 'BYPASSNONUSPS=Y' & @CRLF & 'JOBPASSWORD=120000005439F4CD618534AF280A4DE731197DC9568EE87EB769C138E35130F19F3FAC1C811F51ED47E4AA0221C3C960FB9F14657756F1778AA4A22748BCD9AD0C7E730F42BB6A105BB2484E' & @CRLF & 'ADDRESSGROUPS="MAIN"' & @CRLF & 'DESTINATIONFOLDER="\\kciapp1\DATA\LISTS\kyle\NCOA\auto"' & @CRLF & 'DATAQUALITYOVERRIDE=Y' & @CRLF & 'ORDERTERMSACCEPTED=Y' & @CRLF & 'CASE="MIXED"' & @CRLF & 'PRINTPROCESSINGCERTIFICATE=Y' & @CRLF & 'COAAUDITEXPORT=Y' & @CRLF & 'PRINTCOAAUDIT=Y' & @CRLF & 'HIDERETURNCODES="23 26 27 28"' & @CRLF & 'PARALLEL=N' & @CRLF & 'WAIT=DEFAULT' & @CRLF & 'STARTTIME=DEFAULT' & @CRLF & 'SUPPRESSERRORS=N' & @CRLF & 'LISTOWNER=LISTPROCESSOR' & @CRLF & 'CLASSOFMAIL="A"' & @CRLF & 'MAILINGZIPCODE="49501"' & @CRLF,1)
	$LinesCounter = $LinesCounter + 28
	$TasksCounter = $TasksCounter + 1
EndFunc
Func Load($_mjbpath,$_taskmasterjob)
	_FileWriteToLine($_mjbpath,$LinesCounter,'[LOAD-' & $TasksCounter & ']' & @CRLF & _
	'JOB="' & $_taskmasterjob & '"' & @CRLF,1)
	$LinesCounter = $LinesCounter + 2
	$TasksCounter = $TasksCounter + 1
EndFunc
Func Presortwithreports($_listpath,$_mjbpath,$_jobname,$_inkjet,$_settings,$_promptlabellayout="",$_presortname = "Scripted Presort")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[PRESORT-' & $TasksCounter & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'ADDRESSGROUP="MAIN"' & @CRLF & _
	'SELECTIVITY=NONE' & @CRLF & _
	'SETTINGS=' & $_settings & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT'  & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[QUALIFICATIONREPORT-' & $TasksCounter + 1 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'PRINTER="\\kcimail2\HP LaserJet 5Si MX"' & @CRLF & _
	'INCLUDEPSID=Y' & @CRLF & _
	'PSID="' & $_jobname & '"' & @CRLF & _
	'MAILID="112114"' & @CRLF & _
	'INCLUDEBATCHCODE=Y' & @CRLF & _
	'MANIFESTGRPBYRATE=N' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[CONTAINERTAGS-' & $TasksCounter + 2 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'INFOLINE=Y' & @CRLF & _
	'INFOLINEFORMAT="%T %N %K %S %F %R"' & @CRLF & _
	'HEADERLABEL=Y' & @CRLF & _
	'NEWPAGESTREAM=Y' & @CRLF & _
	'TAGTYPE="IMTRAYLABEL"' & @CRLF & _
	'PRINTER="\\Kciapp1\TRAY"' & @CRLF & _
	'PAPER="Custom Size"' & @CRLF & _
	'ORIENTATION="PORTRAIT"' & @CRLF & _
	'PAPERWIDTH=3.25' & @CRLF & _
	'PAPERLENGTH=2' & @CRLF & _
	'MARGINLEFT=0' & @CRLF & _
	'MARGINTOP=0' & @CRLF & _
	'LABELWIDTH=3.25' & @CRLF & _
	'LABELHEIGHT=2' & @CRLF & _
	'LABELHORZPITCH=3.25' & @CRLF & _
	'LABELVERTPITCH=2' & @CRLF & _
	'ACROSS=1' & @CRLF & _
	'DOWN=1' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[PRESORTEDLABELS-' & $TasksCounter + 3 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS=' & $_promptlabellayout & '"Inkjet - multi"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'FILENAME="\\stash\Inkjet\' & $_inkjet & '.1UP"' & @CRLF & _
	'OVERWRITE=Y' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF ,1)


	$LinesCounter += 64
	$TasksCounter += 4
EndFunc
Func GreatlandRONNCOAReports($_listpath,$_mjbpath)
	_FileWriteToLine($_mjbpath,$LinesCounter,'[ENCODE-' & $TasksCounter & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SELECTIVITY=NONE' & @CRLF & _
	'ADDRESSGROUPS="MAIN"' & @CRLF & _
	'SWAP=Y' & @CRLF & _
	'STANDARDIZEADDRESS=Y' & @CRLF & _
	'STANDARDIZECITY=Y' & @CRLF & _
	'ABBREVIATECITY=N' & @CRLF & _
	'IGNORENONUSPS=N' & @CRLF & _
	'EXTENDEDMATCHING=N' & @CRLF & _
	'CASE="MIXED"' & @CRLF & _
	'FIRMASIS=Y' & @CRLF & _
	'ZIP5CHECKDIGIT=N' & @CRLF & _
	'SUMMARYPAGE=N' & @CRLF & _
	'NDIREPORT=N' & @CRLF & _
	'COPIES=0' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=n' & @CRLF & _
	'[EXPORT-' & $TasksCounter + 1 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS="NCOA DO NOT MAIL"' & @CRLF & _
	'OVERWRITE=Y' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=N' & @CRLF & _
	'[EXPORT-' & $TasksCounter + 2 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS="NCOA GOOD"' & @CRLF & _
	'OVERWRITE=Y' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=N' & @CRLF & _
	'[EXPORT-' & $TasksCounter + 3 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS="NCOA Questionable"' & @CRLF & _
	'OVERWRITE=Y' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=N' & @CRLF & _
	'[DISTRIBUTIONREPORT-' & $TasksCounter + 4 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS="State"' & @CRLF & _
	'SELECTIVITY=NONE' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=N',1)


	$LinesCounter += 53
	$TasksCounter += 5

EndFunc
Func Presortreports($_listpath,$_mjbpath,$_jobname,$_inkjet,$_promptlabellayout="",$_presortname = "Scripted Presort")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[QUALIFICATIONREPORT-' & $TasksCounter & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'PRINTER="\\kcimail2\HP LaserJet 5Si MX"' & @CRLF & _
	'INCLUDEPSID=Y' & @CRLF & _
	'PSID="' & $_jobname & '"' & @CRLF & _
	'INCLUDEBATCHCODE=Y' & @CRLF & _
	'MANIFESTGRPBYRATE=N' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[CONTAINERTAGS-' & $TasksCounter + 1 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'MAILERID="106788"' & @CRLF & _
	'INFOLINE=Y' & @CRLF & _
	'INFOLINEFORMAT="%T %N %K %S %F %R"' & @CRLF & _
	'HEADERLABEL=Y' & @CRLF & _
	'NEWPAGESTREAM=Y' & @CRLF & _
	'TAGTYPE="IMTRAYLABEL"' & @CRLF & _
	'PRINTER="\\Kciapp1\TRAY"' & @CRLF & _
	'PAPER="Custom Size"' & @CRLF & _
	'ORIENTATION="PORTRAIT"' & @CRLF & _
	'PAPERWIDTH=3.25' & @CRLF & _
	'PAPERLENGTH=2' & @CRLF & _
	'MARGINLEFT=0' & @CRLF & _
	'MARGINTOP=0' & @CRLF & _
	'LABELWIDTH=3.25' & @CRLF & _
	'LABELHEIGHT=2' & @CRLF & _
	'LABELHORZPITCH=3.25' & @CRLF & _
	'LABELVERTPITCH=2' & @CRLF & _
	'ACROSS=1' & @CRLF & _
	'DOWN=1' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[PRESORTEDLABELS-' & $TasksCounter + 2 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS=' & $_promptlabellayout & '"Inkjet - multi"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'FILENAME="\\stash\Inkjet\' & $_inkjet & '.1UP"' & @CRLF & _
	'OVERWRITE=Y' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF ,1)


	$LinesCounter += 53
	$TasksCounter += 3
EndFunc
Func GreatlandPresort($_listpath,$_mjbpath,$_jobname,$_presortname = "Scripted Presort",$_presort = "GREATLAND RON")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[PRESORT-' & $TasksCounter & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'ADDRESSGROUP="MAIN"' & @CRLF & _
	'SELECTIVITY=NONE' & @CRLF & _
	'SETTINGS=' & $_presort & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT'  & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF,1)

	$LinesCounter += 11
	$TasksCounter += 3
EndFunc
Func GreatlandPresortreports($_listpath,$_mjbpath,$_jobname,$_presortname = "Scripted Presort",$_filename = "e:\code\greatland\Greatland-RON\Data\BCC in to ACCDB.csv")
	$writeString = '[QUALIFICATIONREPORT-' & $TasksCounter & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'PRINTER="\\kcimail2\HP LaserJet 5Si MX"' & @CRLF & _
	'INCLUDEPSID=Y' & @CRLF & _
	'PSID="' & $_jobname & '"' & @CRLF & _
	'INCLUDEBATCHCODE=Y' & @CRLF & _
	'MANIFESTGRPBYRATE=N' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[CONTAINERTAGS-' & $TasksCounter + 1 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'MAILERID="106788"' & @CRLF & _
	'INFOLINE=Y' & @CRLF & _
	'INFOLINEFORMAT="%T %N %K %S %F %R"' & @CRLF & _
	'HEADERLABEL=Y' & @CRLF & _
	'NEWPAGESTREAM=Y' & @CRLF & _
	'TAGTYPE="IMTRAYLABEL"' & @CRLF & _
	'PRINTER="\\Kciapp1\TRAY"' & @CRLF & _
	'PAPER="Custom Size"' & @CRLF & _
	'ORIENTATION="PORTRAIT"' & @CRLF & _
	'PAPERWIDTH=3.25' & @CRLF & _
	'PAPERLENGTH=2' & @CRLF & _
	'MARGINLEFT=0' & @CRLF & _
	'MARGINTOP=0' & @CRLF & _
	'LABELWIDTH=3.25' & @CRLF & _
	'LABELHEIGHT=2' & @CRLF & _
	'LABELHORZPITCH=3.25' & @CRLF & _
	'LABELVERTPITCH=2' & @CRLF & _
	'ACROSS=1' & @CRLF & _
	'DOWN=1' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[PRESORTEDLABELS-' & $TasksCounter + 2 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS="GREATLAND"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'FILENAME="' & $_filename & '"' & @CRLF & _
	'OVERWRITE=N' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF

	_FileWriteToLine($_mjbpath,$LinesCounter,$writeString ,1)

	$LinesCounter += 62
	$TasksCounter += 3

	return $writeString
EndFunc
Func NelcoPresortreports($_listpath,$_mjbpath,$_jobname,$_presortname = "Scripted Presort")
	$writeString = '[QUALIFICATIONREPORT-' & $TasksCounter & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'PRINTER="\\kcimail2\HP LaserJet 5Si MX"' & @CRLF & _
	'INCLUDEPSID=Y' & @CRLF & _
	'PSID="' & $_jobname & '"' & @CRLF & _
	'INCLUDEBATCHCODE=Y' & @CRLF & _
	'MANIFESTGRPBYRATE=N' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[CONTAINERTAGS-' & $TasksCounter + 1 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'MAILERID="106788"' & @CRLF & _
	'INFOLINE=Y' & @CRLF & _
	'INFOLINEFORMAT="%T %N %K %S %F %R"' & @CRLF & _
	'HEADERLABEL=Y' & @CRLF & _
	'NEWPAGESTREAM=Y' & @CRLF & _
	'TAGTYPE="IMTRAYLABEL"' & @CRLF & _
	'PRINTER="\\Kciapp1\TRAY"' & @CRLF & _
	'PAPER="Custom Size"' & @CRLF & _
	'ORIENTATION="PORTRAIT"' & @CRLF & _
	'PAPERWIDTH=3.25' & @CRLF & _
	'PAPERLENGTH=2' & @CRLF & _
	'MARGINLEFT=0' & @CRLF & _
	'MARGINTOP=0' & @CRLF & _
	'LABELWIDTH=3.25' & @CRLF & _
	'LABELHEIGHT=2' & @CRLF & _
	'LABELHORZPITCH=3.25' & @CRLF & _
	'LABELVERTPITCH=2' & @CRLF & _
	'ACROSS=1' & @CRLF & _
	'DOWN=1' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[PRESORTEDLABELS-' & $TasksCounter + 2 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS="GREATLAND"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'FILENAME="\\kylec7\c$\gitcode\psl\Nelco RON\Data\BCC in to ACCDB.csv"' & @CRLF & _
	'OVERWRITE=N' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF

	_FileWriteToLine($_mjbpath,$LinesCounter,$writeString ,1)



	$LinesCounter += 62
	$TasksCounter += 3

	return $writeString
EndFunc
Func PresortreportsPSN($_listpath,$_mjbpath,$_jobname,$_inkjet,$_promptlabellayout="",$_presortname = "Scripted Presort")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[QUALIFICATIONREPORT-' & $TasksCounter & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'PRINTER="HP 5Si Digital Printing"' & @CRLF & _
	'INCLUDEPSID=Y' & @CRLF & _
	'PSID="' & $_jobname & '"' & @CRLF & _
	'MAILID="112114"' & @CRLF & _
	'INCLUDEBATCHCODE=Y' & @CRLF & _
	'MANIFESTGRPBYRATE=N' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[CONTAINERTAGS-' & $TasksCounter + 1 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'INFOLINE=Y' & @CRLF & _
	'INFOLINEFORMAT="%T %N %K %S %F %R"' & @CRLF & _
	'HEADERLABEL=Y' & @CRLF & _
	'NEWPAGESTREAM=Y' & @CRLF & _
	'TAGTYPE="IMTRAYLABEL"' & @CRLF & _
	'PRINTER="\\Kciapp1\TRAY"' & @CRLF & _
	'PAPER="Custom Size"' & @CRLF & _
	'ORIENTATION="PORTRAIT"' & @CRLF & _
	'PAPERWIDTH=3.25' & @CRLF & _
	'PAPERLENGTH=2' & @CRLF & _
	'MARGINLEFT=0' & @CRLF & _
	'MARGINTOP=0' & @CRLF & _
	'LABELWIDTH=3.25' & @CRLF & _
	'LABELHEIGHT=2' & @CRLF & _
	'LABELHORZPITCH=3.25' & @CRLF & _
	'LABELVERTPITCH=2' & @CRLF & _
	'ACROSS=1' & @CRLF & _
	'DOWN=1' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[PRESORTEDLABELS-' & $TasksCounter + 2 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS=' & $_promptlabellayout & '"Inkjet - multi w/psn"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'FILENAME="\\stash\Inkjet\' & $_inkjet & '.1UP"' & @CRLF & _
	'OVERWRITE=Y' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF ,1)


	$LinesCounter += 54
	$TasksCounter += 3
EndFunc
Func PresortCSV($_listpath,$_mjbpath,$_CSV,$_promptlabellayout="",$_presortname = "Scripted Presort")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[PRESORTEDLABELS-' & $TasksCounter & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS=' & $_promptlabellayout & '"CSV w/Opt End and Barcode - Multi"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'FILENAME="\\dpserver\c$\gdi\' & $_CSV & '.csv"' & @CRLF & _
	'OVERWRITE=Y' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF ,1)
	$LinesCounter += 12
	$TasksCounter += 1
EndFunc
Func inkjetPSN($_listpath,$_mjbpath,$_inkjet,$_promptlabellayout="",$_presortname = "Scripted Presort")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[PRESORTEDLABELS-' & $TasksCounter & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS=' & $_promptlabellayout & '"Inkjet - multi w/psn"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'FILENAME="\\stash\Inkjet\' & $_inkjet & '.1UP"' & @CRLF & _
	'OVERWRITE=Y' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF ,1)
	$LinesCounter += 12
	$TasksCounter += 1
EndFunc
Func inkjetonly($_listpath,$_mjbpath,$_jobname,$_inkjet,$_promptlabellayout="",$_presortname = "Scripted Presort")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[PRESORTEDLABELS-' & $TasksCounter + 2 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS=' & $_promptlabellayout & '"Inkjet - multi"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'FILENAME="\\stash\Inkjet\' & $_inkjet & '.1UP"' & @CRLF & _
	'OVERWRITE=Y' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF ,1)

	$LinesCounter += 12
	$TasksCounter += 1
EndFunc
Func Presortflatwithreports($_listpath,$_mjbpath,$_jobname,$_inkjet,$_settings,$_promptlabellayout="",$_presortname = "Scripted Presort")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[PRESORT-' & $TasksCounter & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'ADDRESSGROUP="MAIN"' & @CRLF & _
	'SELECTIVITY=NONE' & @CRLF & _
	'SETTINGS=' & $_settings & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT'  & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[QUALIFICATIONREPORT-' & $TasksCounter + 1 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'PRINTER="\\kcimail2\HP LaserJet 5Si MX"' & @CRLF & _
	'INCLUDEPSID=Y' & @CRLF & _
	'PSID="' & $_jobname & '"' & @CRLF & _
	'MAILID="112114"' & @CRLF & _
	'INCLUDEBATCHCODE=Y' & @CRLF & _
	'MANIFESTGRPBYRATE=N' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[CONTAINERTAGS-' & $TasksCounter + 2 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'INFOLINE=Y' & @CRLF & _
	'INFOLINEFORMAT="%T %N %K %S %F %R"' & @CRLF & _
	'HEADERLABEL=Y' & @CRLF & _
	'NEWPAGESTREAM=Y' & @CRLF & _
	'TAGTYPE="IMTRAYLABEL"' & @CRLF & _
	'PRINTER="\\Kciapp1\TRAY"' & @CRLF & _
	'PAPER="Custom Size"' & @CRLF & _
	'ORIENTATION="PORTRAIT"' & @CRLF & _
	'PAPERWIDTH=3.25' & @CRLF & _
	'PAPERLENGTH=2' & @CRLF & _
	'MARGINLEFT=0' & @CRLF & _
	'MARGINTOP=0' & @CRLF & _
	'LABELWIDTH=3.25' & @CRLF & _
	'LABELHEIGHT=2' & @CRLF & _
	'LABELHORZPITCH=3.25' & @CRLF & _
	'LABELVERTPITCH=2' & @CRLF & _
	'ACROSS=1' & @CRLF & _
	'DOWN=1' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[CONTAINERTAGS-' & $TasksCounter + 3 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'INFOLINE=Y' & @CRLF & _
	'HEADERLABEL=Y' & @CRLF & _
	'NEWPAGESTREAM=Y' & @CRLF & _
	'TAGTYPE="PALLET"' & @CRLF & _
	'PRINTER="\\kcimail2\HP LaserJet 5Si MX"' & @CRLF & _
	'COPIES=2' & @CRLF & _
	'COLLATE=Y' & @CRLF & _
	'PAPER="Letter"' & @CRLF & _
	'PAPERSOURCE="Automatically Select"' & @CRLF & _
	'ORIENTATION="LANDSCAPE"' & @CRLF & _
	'MARGINLEFT=0.5' & @CRLF & _
	'MARGINTOP=0.5' & @CRLF & _
	'LABELWIDTH=10' & @CRLF & _
	'LABELHEIGHT=7.5' & @CRLF & _
	'LABELHORZPITCH=10' & @CRLF & _
	'LABELVERTPITCH=7.5' & @CRLF & _
	'ACROSS=1' & @CRLF & _
	'DOWN=1' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF & _
	'[PRESORTEDLABELS-' & $TasksCounter + 4 & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS=' & $_promptlabellayout & '"Inkjet - multi"' & @CRLF & _
	'PRESORTNAME="' & $_presortname & '"' & @CRLF & _
	'STREAMLIST=" ALL;ALL ALL"' & @CRLF & _
	'ABSOLUTECONTAINERNUMBERS=N' & @CRLF & _
	'FILENAME="\\stash\Inkjet\' & $_inkjet & '.1UP"' & @CRLF & _
	'OVERWRITE=Y' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=Y' & @CRLF ,1)


	$LinesCounter += 91
	$TasksCounter += 5
EndFunc
Func PackageEnd($_listpath,$_mjbpath)
	_FileWriteToLine($_mjbpath,$LinesCounter,'[PACKAGE-' & $TasksCounter & ']' & @CRLF & 'LIST=' & '"' & $_listpath & '"' & @CRLF & 'DESTINATIONFOLDER="\\kciapp1\DATA\MMPACKAGE\"' & @CRLF & 'OVERWRITE=Y' & @CRLF & 'PARALLEL=N' & @CRLF & 'WAIT=DEFAULT' & @CRLF & 'STARTTIME=DEFAULT' & @CRLF & 'SUPPRESSERRORS=N' & @CRLF,1)
	$TasksCounter += 1
	$LinesCounter += 8
	_FileWriteToLine($_mjbpath,$LinesCounter,'[TERMINATE-' & $TasksCounter  & ']' & @CRLF,1)
	$TasksCounter += 1
	$LinesCounter += 1
EndFunc
Func Import($_listpath,$_filepath,$_mjbpath,$_importTemplate)
	_FileWriteToLine($_mjbpath,$LinesCounter,'[IMPORT-' & $TasksCounter & ']' & @CRLF & 'LIST=' & '"' & $_listpath & '"' & @CRLF &'SETTINGS="' & $_importTemplate & '"' & @CRLF &'FILENAME='  & '"' & $_filepath & '"' & @CRLF &'PARALLEL=N' & @CRLF &'WAIT=DEFAULT' & @CRLF &'STARTTIME=DEFAULT' & @CRLF &'SUPPRESSERRORS=N' & @CRLF ,1)
	$LinesCounter = $LinesCounter + 8
	$TasksCounter = $TasksCounter + 1
EndFunc
Func Sublist($_listpath,$_mjbpath,$_sublistpath,$_description = "NONE")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[SUBLIST-' & $TasksCounter & ']' & _
		'LIST="' & $_sublistpath & '"' & @CRLF & _
		'INDEX="NOTICES"' & @CRLF & _
		'FILENAME="' & $_listpath & '"' & @CRLF & _
		'HIDERECORDS=N' & @CRLF & _
		'AUTOOPENLIST=N' & @CRLF & _
		'PARALLEL=N' & @CRLF & _
		'WAIT=DEFAULT' & @CRLF & _
		'STARTTIME=DEFAULT' & @CRLF & _
		'SUPPRESSERRORS=Y',1)
	$LinesCounter += 13
	$TasksCounter += 1
EndFunc
Func Export($_listpath,$_mjbpath,$_exportName,$_selectivity = "NONE")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[EXPORT-' & $TasksCounter & ']' & @CRLF & 'LIST="' & $_listpath & '"' & @CRLF & 'SETTINGS="' & $_exportName & '"' & @CRLF & 'SELECTIVITY="' & $_selectivity & '' & @CRLF & 'INDEX=NONE' & @CRLF & 'PARALLEL=N' & @CRLF & 'WAIT=DEFAULT' & @CRLF & 'STARTTIME=DEFAULT' & @CRLF & 'SUPPRESSERRORS=Y' & @CRLF,1)
	$LinesCounter = $LinesCounter + 9
	$TasksCounter += 1
EndFunc
Func hide($_listpath,$_mjbpath,$_selectivity="No state")
	_FileWriteToLine($_mjbpath,$LinesCounter,'[HIDE-' & $TasksCounter & ']' & @CRLF & _
		'LIST="' & $_listpath & '"' & @CRLF & _
		'SELECTIVITY="' & $_selectivity & '"' & @CRLF & _
		'PARALLEL=N' & @CRLF & _
		'WAIT=DEFAULT' & @CRLF & _
		'STARTTIME=DEFAULT' & @CRLF & _
		'SUPPRESSERRORS=N' & @CRLF,1)
	$LinesCounter = $LinesCounter + 7
	$TasksCounter = $TasksCounter + 1
EndFunc

Func modify($_listpath,$_mjbpath,$_modify)
	_FileWriteToLine($_mjbpath,$LinesCounter,'[MODIFY-' & $TasksCounter & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS="' & $_modify & '"' & @CRLF & _
	'PARALLEL=N' & @CRLF & _
	'WAIT=DEFAULT' & @CRLF & _
	'STARTTIME=DEFAULT' & @CRLF & _
	'SUPPRESSERRORS=N' & @CRLF,1)
	$LinesCounter = $LinesCounter + 7
	$TasksCounter = $TasksCounter + 1
EndFunc
Func nonpresortedlabels($_listpath,$_mjbpath,$_labelLayout, $_inkjet)
	_FileWriteToLine($_mjbpath,$LinesCounter,'[NONPRESORTEDLABELS-' & $TasksCounter & ']' & @CRLF & _
	'LIST="' & $_listpath & '"' & @CRLF & _
	'SETTINGS="' & $_labelLayout & '"'  & @CRLF & _
	'INDEX=NONE'                                             & @CRLF & _
	'SELECTIVITY=NONE'                                       & @CRLF & _
	'FILENAME="' & $_inkjet & '"'  & @CRLF & _
	'OVERWRITE=Y'                                            & @CRLF & _
	'PARALLEL=N'                                             & @CRLF & _
	'WAIT=DEFAULT'                                           & @CRLF & _
	'STARTTIME=DEFAULT'                                      & @CRLF & _
	'SUPPRESSERRORS=N' & @CRLF,1)
	$LinesCounter = $LinesCounter + 11
	$TasksCounter = $TasksCounter + 1
EndFunc
