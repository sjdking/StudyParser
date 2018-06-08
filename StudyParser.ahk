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

;FileSelectFolder, FolderPath
;FileSelectFolder, FolderPath, \\ANHB-SLP-DATA\NASSHARE\2017 Grad Dip Students\Nov29_ScoringSession
FileSelectFolder, FolderPath, J:\Pulmonary Physiology\SleepScience\SCG\Labs\
	
outputFile := % FolderPath . "\" . A_Now . "ScoringComparison.xlsx"
outputText := % FolderPath . "\" . A_Now . "TextCheck.csv"

FileAppend, Scorer`#Event`#Time`#Duration`n, %outputText%

;MsgBox, %FileList%

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

ScoredEventNames := Object()
ScoredEventStartTimes := Object()
ScoredEventEpochs := Object()
ScoredEventDurations := Object()
ScoredEventDesats := Object()
ScoredEventDesatMins := Object()
ScoredEventDesatOffsets := Object() ;offset is the time from the end of the associated respiratory event to the termination of the marked SpO2 desat event
SleepStages := Object()
ScoredEventSleepStage := Object()
HypopArr := {stage:[5],duration:[],desat:[100]}
OAArr := {stage:[5],duration:[],desat:[100]}
CAArr := {stage:[5],duration:[],desat:[100]}
MAArr := {stage:[5],duration:[],desat:[100]}
RERAArr := {stage:[5],duration:[],desat:[100]}
LimbArr := {stage:[5],duration:[],desat:[100]}
ArousalArr := {stage:[5],type:[5]}

ClearArray(Name, Size)
	{
	Loop, % Size
	%Name%%A_Index% :=
	Return
	}
	
loop Files, %FolderPath%\*.xml
{
	k := % A_Index + 1

	
	ScoredEventNames := Object()
	ScoredEventStartTimes := Object()
	ScoredEventEpochs := Object()
	ScoredEventDurations := Object()
	ScoredEventDesats := Object()
	ScoredEventDesatMins := Object()
	ScoredEventDesatOffsets := Object() ;offset is the time from the end of the associated respiratory event to the termination of the marked SpO2 desat event
	SleepStages := Object()
	ScoredEventSleepStage := Object()

	ClearArray("ScoredEventNames", Array0)
	ClearArray("ScoredEventStartTimes", Array0)
	ClearArray("ScoredEventEpochs", Array0)
	ClearArray("ScoredEventDurations", Array0)
	ClearArray("ScoredEventStartTimes", Array0)
	ClearArray("ScoredEventDesats", Array0)
	ClearArray("ScoredEventDesatMins", Array0)
	ClearArray("SleepStages", Array0)
	ClearArray("ScoredEventSleepStage", Array0)
	ClearArray("HypopArr", Array0)
	ClearArray("OAArr", Array0)
	ClearArray("CAArr", Array0)
	ClearArray("MAArr", Array0)
	ClearArray("RERAArr", Array0)
	ClearArray("LimbArr", Array0)
	ClearArray("ArousalArr", Array0)

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

	CreatedOn := StrX( xml, "<CREATEDON>",1,11,"</CREATEDON>",1,12)
	LastModifiedBy := StrX( xml, "<LASTMODIFIEDBY>",1,16,"</LASTMODIFIEDBY>",1,17)
	LastModifiedOn := StrX( xml, "<LASTMODIFIEDON>",1,16,"</LASTMODIFIEDON>",1,17)
	Mode := StrX( xml, "<MODE>",1,6,"</MODE>",1,7)
	strComments := StrX( xml, "<COMMENTS>",0,10,"</COMMENTS>",1,11)
	strCommentsLength := StrLen(strComments)

	If (strCommentsLength > 8)
		{
		InputBox, DataSetID, Enter ID, Score Data Set ID for %A_LoopFileName% is %strComments%. Please enter a valid ID.
		strComments := % DataSetID
		}	

	Gui, Add, Progress, w500 Range0-5000 vMyProgress
	Gui, Add, Text, vEventIndex wp
	Gui, Add, Text, vEventTime wp
	Gui, Show, x100 y100


	While Item := StrX( xml, "<SCOREDEVENT>",N,0,"</SCOREDEVENT>",1,0, N)
		{
		i = % A_Index
		ScoredEventName := StrX( Item, "<NAME>",1,6,"</NAME>",1,7)
		ScoredEventNames.Insert(StrX( Item, "<NAME>",1,6,"</NAME>",1,7))
		ScoredEventStartTime := StrX( Item, "<TIME>",1,6,"</TIME>",1,7)
		ScoredEventStartTimes.Insert(ScoredEventStartTime)
		ScoredEventEpochs.Insert(Floor(1 + ScoredEventStartTime/30))
		ScoredEventDuration := StrX( Item, "<DURATION>",1,10,"</DURATION>",1,11)
		ScoredEventDurations.Insert(StrX( Item, "<DURATION>",1,10,"</DURATION>",1,11))
		ScoredEventDesats.Insert(StrX( Item, "<PARAM1>",1,8,"</PARAM1>",1,9))
		ScoredEventDesatMins.Insert(StrX( Item, "<PARAM2>",1,8,"</PARAM2>",1,9))
		ScoredEventDesatOffsets.Insert(StrX( Item, "<PARAM3>",1,8,"</PARAM3>",1,9))
		
		GuiControl,, MyProgress, %i%
		GuiControl,, EventIndex, Now processing event %i%
		GuiControl,, EventTime, Scored event time is %ScoredEventStartTime%
		
		FileAppend, %strComments%`#%ScoredEventName%`#%ScoredEventStartTime%`#%ScoredEventDuration%`n, %outputText%

		}
		
	Gui Destroy

	Gui, Add, Progress, w500 Range0-1000 vMyProgress
	Gui, Add, Text, vEpochNumber wp
	Gui, Show, x100 y100

	N := 1
	
	;Item := StrX( xml, "<SLEEPSTAGE>",1,12,"</SLEEPSTAGE>",1,13, N )
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
			SleepStage = 0
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
		SleepStages.Insert(SleepStage)
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
	ArousalCount := 0
	ArousalCountSleep := 0
	RERACount := 0
	RERACountSleep := 0
	PLMCount := 0
	PLMCountSleep := 0
	LimbCount := 0
	LimbCountSleep := 0
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

	For index, element in ScoredEventNames
		{
		EventEpoch := ScoredEventEpochs[index] - 1
		EventStage := SleepStages[EventEpoch]
		EventDuration := ScoredEventDurations[index]
		EventDesat := ScoredEventDesats[index] ; DESAT ASSOCIATION ALGORITHM REQUIRED
		
		If(InStr(element,"Hypopnea"))
			{
			HypopCount++
			HypopArr.Insert(element, {stage:EventStage,duration:EventDuration,desat:EventDesat})
			If (EventStage > 0 and EventStage <= 5)
				HypopCountSleep++
			}
		If(InStr(element,"Obstructive Apnea"))
			{
			OACount++
			OAArr.Insert(element, {stage:EventStage,duration:EventDuration,desat:EventDesat})
			If (EventStage > 0 and EventStage <= 5)
				OACountSleep++
			}
		If(InStr(element,"Central Apnea"))
			{
			CACount++
			CAArr.Insert(element, {stage:EventStage,duration:EventDuration,desat:EventDesat})
			If (EventStage > 0 and EventStage <= 5)
				CACountSleep++
			}
		If(InStr(element,"Mixed Apnea"))
			{
			MACount++
			MAArr.Insert(element, {stage:EventStage,duration:EventDuration,desat:EventDesat})
			If (EventStage > 0 and EventStage <= 5)
				MACountSleep++
			}
		If(InStr(element,"Arousal"))
			{
			ArousalCount++
			ArousalArr.Insert(element, {stage:EventStage,type:element})
			If (EventStage > 0 and EventStage <= 5)
				ArousalCountSleep++
			}
		If(InStr(element,"RERA"))
			{
			RERACount++
			RERAArr.Insert(element, {stage:EventStage,duration:EventDuration,desat:EventDesat})
			If (EventStage > 0 and EventStage <= 5)
				RERACountSleep++
			}
		If(InStr(element,"Limb Movement"))
			{
			LimbCount++
			LimbArr.Insert(element, {stage:EventStage,duration:EventDuration,desat:EventDesat})
			If (EventStage > 0 and EventStage <= 5)
				LimbCountSleep++
			}

		}
	
	;Clear previous indexes
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
	
	;Calculate indexes
	totalApneaSleep := OACountSleep + CACountSleep + MACountSleep
	
	HypIndexSleep := % (HypopCountSleep / totalSleepTime) * 60
	OAIndexSleep := % (OACountSleep / totalSleepTime) * 60
	CAIndexSleep := % (CACountSleep / totalSleepTime) * 60
	MAIndexSleep := % (MACountSleep / totalSleepTime) * 60
	ApneaIndexSleep := % (totalApneaSleep / totalSleepTime) * 60
	AHISleep := % ((totalApneaSleep + HypopCountSleep) / totalSleepTime)  * 60
	ArousalIndexSleep := % (ArousalCountSleep / totalSleepTime) * 60
	RERAIndexSleep := % (RERACountSleep / totalSleepTime) * 60
	LimbIndexSleep := % (LimbCountSleep / totalSleepTime) * 60
	
	;Calculate sleep percentages
	percentN1 := % (countN1 / totalSleepTime) * 100
	percentN2 := % (countN2 / totalSleepTime) * 100
	percentN3 := % (countN3 / totalSleepTime) * 100
	percentREM := % (countREM / totalSleepTime) * 100
	
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
	
	Gui, Destroy
	ScoredEventNames := ""
	ScoredEventStartTimes := ""
	ScoredEventEpochs := ""
	ScoredEventDurations := ""
	ScoredEventDesats := ""
	ScoredEventDesatMins := ""
	ScoredEventDesatOffsets := "" ;offset is the time from the end of the associated respiratory event to the termination of the marked SpO2 desat event
	SleepStages := ""
	ScoredEventSleepStage := ""
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
