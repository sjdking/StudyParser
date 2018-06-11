#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance force

;--- Function StrX by user SKAN --------------------------
;--- http://www.autohotkey.com/forum/viewtopic.php?t=51354
StrX( H,  BS="",BO=0,BT=1,   ES="",EO=0,ET=1,  ByRef N="" ) { ;    | by Skan | 19-Nov-2009
Return SubStr(H,P:=(((Z:=StrLen(ES))+(X:=StrLen(H))+StrLen(BS)-Z-X)?((T:=InStr(H,BS,0,((BO
 <0)?(1):(BO))))?(T+BT):(X+1)):(1)),(N:=P+((Z)?((T:=InStr(H,ES,0,((EO)?(P+1):(0))))?(T-P+Z
 +(0-ET)):(X+P)):(X)))-P) ; v1.0-196c 21-Nov-2009 www.autohotkey.com/forum/topic51354.html
}
;===========================================================================================

;FileSelectFolder, FolderPath, C:\temp\AutoHotKeyScripts\XML_DataSetParser\Studies\TEST
;FileSelectFolder, FolderPath, \\ANHB-SLP-DATA\NASSHARE\2017 Grad Dip Students\Nov29_ScoringSession
FileSelectFolder, FolderPath, J:\Pulmonary Physiology\SleepScience\SCG\Labs\Q Sleep\2018 QSleep
	
outputFile := % FolderPath . "\" . A_Now . "ScoringComparison.xlsx"
;outputText := % FolderPath . "\" . A_Now . "TextCheck.csv"

;FileAppend, Scorer`#Event`#Time`#Duration`n, %outputText%

objExcel := ComObjCreate("Excel.Application")
objExcel.Visible := true
newComparison := objExcel.Workbooks.Add()
objExcel.ActiveWorkbook.SaveAs(outputFile)

;Initialise spreadsheet
objExcel.Worksheets(1).Range("A1").Value := "Scorer"
objExcel.Worksheets(1).Range("B1").Value := "Sleep Efficiency"
objExcel.Worksheets(1).Range("C1").Value := "TST"
objExcel.Worksheets(1).Range("D1").Value := "N1%"
objExcel.Worksheets(1).Range("E1").Value := "N2%"
objExcel.Worksheets(1).Range("F1").Value := "N3%"
objExcel.Worksheets(1).Range("G1").Value := "REM%"
objExcel.Worksheets(1).Range("H1").Value := "N1 minutes"
objExcel.Worksheets(1).Range("I1").Value := "N2 minutes"
objExcel.Worksheets(1).Range("J1").Value := "N3 minutes"
objExcel.Worksheets(1).Range("K1").Value := "REM minutes"
objExcel.Worksheets(1).Range("L1").Value := "WASO minutes"
objExcel.Worksheets(1).Range("M1").Value := "SOL"
objExcel.Worksheets(1).Range("N1").Value := "ROL"
objExcel.Worksheets(1).Range("O1").Value := "AHI sleep"
objExcel.Worksheets(1).Range("P1").Value := "OA index"
objExcel.Worksheets(1).Range("Q1").Value := "CA index"
objExcel.Worksheets(1).Range("R1").Value := "MA index"
objExcel.Worksheets(1).Range("S1").Value := "Hyp index"
objExcel.Worksheets(1).Range("T1").Value := "RERA index"
objExcel.Worksheets(1).Range("U1").Value := "Arousal index"
objExcel.Worksheets(1).Range("V1").Value := "Limb index"
objExcel.Worksheets(1).Range("W1").Value := "Num OA"
objExcel.Worksheets(1).Range("X1").Value := "Num CA"
objExcel.Worksheets(1).Range("Y1").Value := "Num MA"
objExcel.Worksheets(1).Range("Z1").Value := "Num Hyp"

	
loop Files, %FolderPath%\*.xml
{
	k := % A_Index + 1

	;eventArr[] holds values [evName,evStartTime,evEpoch,evDuration,evDesat,evDesatMin,evDesatOffset]
	eventArr := []
	;sleepArr[] holds value [sleepStage], index is the epoch number
	sleepArr := []
	countArr := []

	countLOn := 0
	countWake := 0
	countN1 := 0
	countN2 := 0
	countN3 := 0
	countREM := 0
	countOther := 0
	countWASO := 0
	totalSleepTime := 0
	HypopCount := 0
	HypopCountSleep := 0
	OACount := 0
	OACountSleep := 0
	CACount := 0
	CACountSleep := 0
	MACount := 0
	MACountSleep := 0
	ArousalCount := 0
	ArousalCountSleep := 0
	RERACount := 0
	RERACountSleep := 0
	PLMCount := 0
	PLMCountSleep := 0
	LimbCount := 0
	LimbCountSleep := 0
	
	N := 1 ;reset counter for parsing individual XML files

	;FileRead, xml, %inputFile%
	FileRead, xml, %FolderPath%\%A_LoopFileName%

	;Initialise arrays for each region and sleep stage

	Loop, 138
	{
		i := A_Index
		If i = 1
		{
			loopSleepStage = 1
		}
		Else If i = 2
		{
			loopSleepStage = 2
		}
		Else If i = 3
		{
			loopSleepStage = 3
		}
		Else If i = 5
		{
			loopSleepStage = 5
		}
		Else If i = 10
		{
			loopSleepStage = 10
		}
		Else If i = 138
		{
			loopSleepStage = 138
		}
		Else
			Continue
		countArr[1,loopSleepStage] := 0
		countArr[2,loopSleepStage] := 0
		countArr[3,loopSleepStage] := 0
		countArr[4,loopSleepStage] := 0
		countArr[5,loopSleepStage] := 0
		countArr[6,loopSleepStage] := 0
		countArr[7,loopSleepStage] := 0
	}
	
	
	CreatedOn := StrX( xml, "<CREATEDON>",1,11,"</CREATEDON>",1,12)
	LastModifiedBy := StrX( xml, "<LASTMODIFIEDBY>",1,16,"</LASTMODIFIEDBY>",1,17)
	LastModifiedOn := StrX( xml, "<LASTMODIFIEDON>",1,16,"</LASTMODIFIEDON>",1,17)
	Mode := StrX( xml, "<MODE>",1,6,"</MODE>",1,7)
	strComments := StrX( xml, "<COMMENTS>",0,10,"</COMMENTS>",1,11)
	strCommentsLength := StrLen(strComments)

	If (strCommentsLength > 8)
		{
		InputBox, DataSetID, Enter ID, Score Data Set ID for %A_LoopFileName% is empty. Please enter a valid ID.
		strComments := % DataSetID
		}	

	Gui, Add, Progress, w500 Range0-8000 vMyProgress
	Gui, Add, Text, vEventIndex wp
	Gui, Add, Text, vEventTime wp
	Gui, Show, x100 y100


	While Item := StrX( xml, "<SCOREDEVENT>",N,0,"</SCOREDEVENT>",1,0, N)
		{
		i = % A_Index
		
		scoredEventName := StrX( Item, "<NAME>",1,6,"</NAME>",1,7)
		scoredEventStartTime := StrX( Item, "<TIME>",1,6,"</TIME>",1,7)
		scoredEventEpoch := Floor(1 + ScoredEventStartTime/30)
		scoredEventDuration := StrX( Item, "<DURATION>",1,10,"</DURATION>",1,11)
		scoredEventDesat := StrX( Item, "<PARAM1>",1,8,"</PARAM1>",1,9)
		scoredEventDesatMin .= StrX( Item, "<PARAM2>",1,8,"</PARAM2>",1,9)
		scoredEventDesatOffset .= StrX( Item, "<PARAM3>",1,8,"</PARAM3>",1,9)
		
		eventArr.Insert([scoredEventName,scoredEventStartTime,scoredEventEpoch,scoredEventDuration,scoredEventDesat,scoredEventDesatMin,scoredEventDesatOffset])
		
		GuiControl,, MyProgress, %i%
		GuiControl,, EventIndex, Now processing event %i%
		GuiControl,, EventTime, Scored event epoch is %scoredEventEpoch%
		
;		FileAppend, %strComments%`#%ScoredEventName%`#%ScoredEventStartTime%`#%ScoredEventDuration%`n, %outputText%

		}
		
	Gui Destroy

	Gui, Add, Progress, w500 Range0-1000 vMyProgress
	Gui, Add, Text, vEpochNumber wp
	Gui, Show, x100 y100

	N := 1
	
	sleepOnsetMarker := 0
	REMOnsetMarker := 0

	While Item := StrX( xml, "<SLEEPSTAGE>",N,0,"</SLEEPSTAGE>",1,0, N )
		{
		j = % A_Index
		SleepStage := StrX( Item, "<SLEEPSTAGE>",1,12,"</SLEEPSTAGE>",1,13)
		SetFormat, float, 0.1
		;GuiControl,, NapNumber, 0
		If (SleepStage = "138")
			{
			SleepStage = 138
			countLOn := % countLOn + 0.5
			}
		Else If (SleepStage = "10")
			{
			SleepStage = 10
			If (sleepOnsetMarker = 0)
				countWake := % countWake + 0.5
			Else
				{
				countWake := % countWake + 0.5
				countWASO := % countWASO + 0.5
				}
			}
		Else If (SleepStage = "5")
			{
			SleepStage = 5
			countREM := % countREM + 0.5
			If (sleepOnsetMarker = 0)
				sleepOnsetMarker := % A_Index / 2
			If (REMOnsetMarker = 0)
				REMOnsetMarker := % (A_Index / 2) - sleepOnsetMarker
			}
		Else If (SleepStage = "1")
			{
			SleepStage = 1
			countN1 := % countN1 + 0.5
			If (sleepOnsetMarker = 0)
				sleepOnsetMarker := % A_Index / 2
			}
		Else If (SleepStage = "2")
			{
			SleepStage = 2
			countN2 := countN2 + 0.5
			If (sleepOnsetMarker = 0)
				sleepOnsetMarker := % A_Index / 2
			}
		Else If (SleepStage = "3")
			{
			SleepStage = 3
			countN3 := countN3 + 0.5
			If (sleepOnsetMarker = 0)
				sleepOnsetMarker := % A_Index / 2
			}
		Else
			countOther++
		
		
		sleepArr.Insert(SleepStage)
		GuiControl,, MyProgress, %j%
		GuiControl,, EpochNumber, Now processing sleep stage for epoch %j%
		;Sleep Stage Codes
		;138 = LIGHTS ON
		;10  = WAKE
		;1   = N1
		;2   = N2
		;3   = N3
		;5   = R

		}

	SetFormat, float, 0.3
	
	NREMTime := % countN1 + countN2 + countN3
	totalSleepTime := % NREMTime + countREM
	totalStudyTime := % totalSleepTime + countWake
	sleepEff := % 100*totalSleepTime/totalStudyTime

	HypopCount := 0
	HypopCountSleep := 0
	OACount := 0
	OACountSleep := 0
	CACount := 0
	CACountSleep := 0
	MACount := 0
	MACountSleep := 0
	respEventCountSleep := 0
	ArousalCount := 0
	ArousalCountSleep := 0
	RERACount := 0
	RERACountSleep := 0
	PLMCount := 0
	PLMCountSleep := 0
	LimbCount := 0
	LimbCountSleep := 0
	totalApneaSleep := 0
	HypIndexSleep := 0
	OAIndexSleep := 0
	CAIndexSleep := 0
	MAIndexSleep := 0
	ApneaIndexSleep := 0
	AHISleep := 0
	ArousalIndexSleep := 0
	RERAIndexSleep := 0
	LimbIndexSleep := 0
	hypopDuration := 0
	OADuration := 0
	CADuration := 0
	MADuration := 0
	RERADuration := 0
	ArousalDuration := 0
;	MsgBox, 
;	(
;	Wake duration: %countWake%
;	N1 duration: %countN1%
;	N2 duration: %countN2%
;	N3 duration: %countN3%
;	REM duration: %countREM%
;	NREM duration: %NREMTime%
;	TST: %totalSleepTime%
;	WASO: %countWASO%
;	Sleep Onset Latency: %sleepOnsetMarker%
;	REM Onset Latency: %REMOnsetMarker%
;	Sleep Efficiency := %sleepEff%
;	)

	Gui, Destroy

	for Index, Value in eventArr
	{
		eventName := % Value.1
		eventTime := % Value.2
		eventEpoch := % Value.3
		eventDuration := % Value.4
		eventDesat := % Value.5
		eventDesatMin := % Value.6
		eventDesatOffset := % Value.7
		eventStage := sleepArr[eventEpoch]
		
		If(InStr(eventName,"Hypopnea"))
		{
			HypopCount++
			countArr[4,eventStage] := countArr[4,eventStage] + 1
			hypopDuration := hypopDuration + eventDuration
			If (eventStage > 0 and eventStage <=5)
				HypopCountSleep++
		}
		Else If(InStr(eventName,"Obstructive Apnea"))
		{
			OACount++
			countArr[1,eventStage] := countArr[1,eventStage] + 1
			OADuration := OADuration + eventDuration
			If (eventStage > 0 and eventStage <=5)
				OACountSleep++
		}
		Else If(InStr(eventName,"Central Apnea"))
		{
			OACount++
			countArr[2,eventStage] := countArr[2,eventStage] + 1
			CADuration := CADuration + eventDuration
			If (eventStage > 0 and eventStage <=5)
				CACountSleep++
		}
		Else If(InStr(eventName,"Mixed Apnea"))
		{
			MACount++
			countArr[3,eventStage] := countArr[3,eventStage] + 1
			MADuration := MADuration + eventDuration
			If (eventStage > 0 and eventStage <=5)
				MACountSleep++
		}
		Else If(InStr(eventName,"RERA"))
		{
			RERACount++
			countArr[5,eventStage] := countArr[5,eventStage] + 1
			RERADuration := RERADuration + eventDuration
			If (eventStage > 0 and eventStage <=5)
				RERACountSleep++
		}
		Else If(InStr(eventName,"Arousal"))
		{
			ArousalCount++
			countArr[6,eventStage] := countArr[6,eventStage] + 1
			arousalDuration := arousalDuration + eventDuration
			If (eventStage > 0 and eventStage <=5)
				ArousalCountSleep++
		}
		Else If(InStr(eventName,"Limb movement"))
		{
			ArousalCount++
			countArr[7,eventStage] := countArr[7,eventStage] + 1
			limbDuration := limbDuration + eventDuration
			If (eventStage > 0 and eventStage <=5)
				LimbCountSleep++
		}
	}

	;calculate number of events
	numOAWake := % countArr[1,10]
	numOANREM := % countArr[1,1] + countArr[1,2] + countArr[1,3]
	numOAREM := % countArr[1,5]
	numOASleep := % numOANREM + numOAREM
	
	numCAWake := % countArr[2,10]
	numCANREM := % countArr[2,1] + countArr[2,2] + countArr[2,3]
	numCAREM := % countArr[2,5]
	numCASleep := % numCANREM + numCAREM
	
	numMAWake := % countArr[3,10]
	numMANREM := % countArr[3,1] + countArr[3,2] + countArr[3,3]
	numMAREM := % countArr[3,5]
	numMASleep := % numMANREM + numMAREM
	
	numHypWake := % countArr[4,10]
	numHypNREM := % countArr[4,1] + countArr[4,2] + countArr[4,3]
	numHypREM := % countArr[4,5]
	numHypSleep := % numHypNREM + numHypREM
	
	numRERAWake := % countArr[5,10]
	numRERANREM := % countArr[5,1] + countArr[5,2] + countArr[5,3]
	numRERAREM := % countArr[5,5]
	numRERASleep := % numRERANREM + numRERAREM
	
	numArousalWake := % countArr[6,10]
	numArousalNREM := % countArr[6,1] + countArr[6,2] + countArr[6,3]
	numArousalREM := % countArr[6,5]
	numArousalSleep := % numArousalNREM + numArousalREM
	
	numLimbWake := % countArr[7,10]
	numLimbNREM := % countArr[7,1] + countArr[7,2] + countArr[7,3]
	numLimbREM := % countArr[7,5]
	numLimbSleep := % numLimbNREM + numLimbREM
	
	respEventCountSleep := %  numOASleep + numCASleep + numMASleep + numHypSleep
	
	;calculate index for each event type
	OAIndexSleep := % numOASleep * 60 / totalSleepTime
	CAIndexSleep := % numCASleep * 60 / totalSleepTime
	MAIndexSleep := % numMASleep * 60 / totalSleepTime
	HypIndexSleep := % numHypSleep * 60 / totalSleepTime
	RERAIndexSleep := % numRERASleep * 60 / totalSleepTime
	ArousalIndexSleep := % numArousalSleep * 60 / totalSleepTime
	AHISleep := % respEventCountSleep * 60 / totalSleepTime
	LimbIndexSleep := % numLimbSleep * 60 / totalSleepTime
	percentN1 := countN1 * 100 / totalSleepTime
	percentN2 := countN2 * 100 / totalSleepTime
	percentN3 := countN3 * 100 / totalSleepTime
	percentREM := countREM * 100 / totalSleepTime
	
	
	objExcel.Worksheets(1).Range("A"k).Value := strComments
	objExcel.Worksheets(1).Range("B"k).Value := sleepEff
	objExcel.Worksheets(1).Range("C"k).Value := totalSleepTime
	objExcel.Worksheets(1).Range("D"k).Value := percentN1
	objExcel.Worksheets(1).Range("E"k).Value := percentN2
	objExcel.Worksheets(1).Range("F"k).Value := percentN3
	objExcel.Worksheets(1).Range("G"k).Value := percentREM
	objExcel.Worksheets(1).Range("H"k).Value := countN1
	objExcel.Worksheets(1).Range("I"k).Value := countN2
	objExcel.Worksheets(1).Range("J"k).Value := countN3
	objExcel.Worksheets(1).Range("K"k).Value := countREM
	objExcel.Worksheets(1).Range("L"k).Value := countWASO
	objExcel.Worksheets(1).Range("M"k).Value := sleepOnsetMarker
	objExcel.Worksheets(1).Range("N"k).Value := REMOnsetMarker
	objExcel.Worksheets(1).Range("O"k).Value := AHISleep
	objExcel.Worksheets(1).Range("P"k).Value := OAIndexSleep
	objExcel.Worksheets(1).Range("Q"k).Value := CAIndexSleep
	objExcel.Worksheets(1).Range("R"k).Value := MAIndexSleep
	objExcel.Worksheets(1).Range("S"k).Value := HypIndexSleep
	objExcel.Worksheets(1).Range("T"k).Value := RERAIndexSleep
	objExcel.Worksheets(1).Range("U"k).Value := ArousalIndexSleep
	objExcel.Worksheets(1).Range("V"k).Value := LimbIndexSleep
	objExcel.Worksheets(1).Range("W"k).Value := numOASleep
	objExcel.Worksheets(1).Range("X"k).Value := numCASleep
	objExcel.Worksheets(1).Range("Y"k).Value := numMASleep
	objExcel.Worksheets(1).Range("Z"k).Value := numHypSleep

}

objExcel.ActiveWorkbook.Save()


MsgBox, Data extraction complete!


objExcel.Application.Quit()
ExitApp

^Esc::
{
objExcel.Application.Quit()
ExitApp
}
