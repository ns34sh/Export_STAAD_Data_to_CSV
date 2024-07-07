Option Explicit
Dim ProgressCounter As Long
Dim MaxProgressCounter As Double
Dim lPrimaryLoadCaseCount As Long
Dim lPrimaryLoadCaseNumbersArray() As Long
Dim lGetLoadCombinationCaseCount As Long
Dim lLoadCombinationCaseNumbersArray() As Long
Dim DlgItem$
Dim SuppValue&
Dim Action%

Dim ForceUnit As String
Dim MomentUnit As String
Dim DistanceUnit As String


Dim ExitFlag As Long

Dim StartTime As Double
Dim TimeRemaining As Double


Sub Main()

'MaxProgressCounter=1
Dim objOpenSTAAD As Object

'Establish link with currently opened STAAD file
Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")





'Get numbers of primary load cases (including combinations defined with repeat load case)
lPrimaryLoadCaseCount = objOpenSTAAD.Load.GetPrimaryLoadCaseCount




'Get Primary Load Case Numbers (including combinations defined with repeat load case)
ReDim lPrimaryLoadCaseNumbersArray (lPrimaryLoadCaseCount)
objOpenSTAAD.Load.GetPrimaryLoadCaseNumbers  lPrimaryLoadCaseNumbersArray



'Get number of  load combinations
lGetLoadCombinationCaseCount = objOpenSTAAD.Load.GetLoadCombinationCaseCount

'Get load combinations primary load combination numbers
ReDim lLoadCombinationCaseNumbersArray (lGetLoadCombinationCaseCount)
objOpenSTAAD.Load. GetLoadCombinationCaseNumbers lLoadCombinationCaseNumbersArray ()



Dim lLoadCase As Long
Dim strLoadCaseName() As String

' Get names of primary load cases and load combinations
ReDim strLoadCaseName(1 To lPrimaryLoadCaseCount+lGetLoadCombinationCaseCount) As String
For lLoadCase=1 To lPrimaryLoadCaseCount+lGetLoadCombinationCaseCount
	If lLoadCase<=lPrimaryLoadCaseCount Then
	strLoadCaseName(lLoadCase) = objOpenSTAAD.Load.GetLoadCaseTitle(lPrimaryLoadCaseNumbersArray(lLoadCase-1))
    strLoadCaseName(lLoadCase) = lPrimaryLoadCaseNumbersArray(lLoadCase-1) & " : " & strLoadCaseName(lLoadCase)
    Else
	strLoadCaseName(lLoadCase) = objOpenSTAAD.Load.GetLoadCaseTitle(lLoadCombinationCaseNumbersArray(lLoadCase-lPrimaryLoadCaseCount-1))
    strLoadCaseName(lLoadCase) = lLoadCombinationCaseNumbersArray(lLoadCase-lPrimaryLoadCaseCount-1) & " : " & strLoadCaseName(lLoadCase)
	End If
Next lLoadCase

'Get base unit information from STAAD (1="English" or 2="Metric")
Dim BaseUnits As Long
BaseUnits=objOpenSTAAD.GetBaseUnit ()



'Display output units in dialog box based on base units set in STAAD
Select Case BaseUnits
Case Is=1
ForceUnit="kip"
MomentUnit="kip-in"
DistanceUnit="in"
Case Is=2
ForceUnit="kN"
MomentUnit="kN-m"
DistanceUnit="m"
End Select

'Prompt user to start and end load combination numbers and structure name
    Begin Dialog UserDialog 505,410,"Export Structure Data",.MyDialogFunc
        Text 30,10,400,15,"Enter Start Load Combination Number:"
        DropListBox 30,25,450,60,strLoadCaseName(),.list1
        Text 30,50,400,15,"Enter End Load Combination Number:"
        DropListBox 30,65,450,60,strLoadCaseName(),.list2
        Text 30,100,300,15,"Enter Structure Name:"
        TextBox 30,120,180,20,.Text
        GroupBox 30,155,200,100,"Output Units"
		Text 40,180,180,20,"Force Units: " & vbTab & ForceUnit
        Text 40,205,180,20,"Moment Units: " & vbTab & MomentUnit
        Text 40,230,180,20,"Distance Units: " & vbTab & DistanceUnit
'
		GroupBox 250,155,210,75,"Export Data For:"
		    OptionGroup .options
            OptionButton 260,180,180,15,"All Members",.Option1
            OptionButton 260,205,180,15,"Selected Members",.Option2
            'OptionButton 260,230,180,15,"Option &2"

'
        GroupBox 30,270,430,100,"Progress",.GroupBox1
		Text 40,295,180,15,"Estimated Time Remaining: "
		Text 220,295,180,15,"",.TimeRem
		Text 40,315,150,15,"Percent Completed: "
        Text 220,315,180,15,"",.PercComp
        GroupBox 30,330,430,40,"",.GroupBox2
		Text 40,345,410,20,"",.ProgText
        'OKButton 125,270,80,20
		PushButton 125,380,80,20,"Run",.Run
        CancelButton 300,380,80,20
    End Dialog
    Dim dlg As UserDialog
    Dim dialogres As Long
    dialogres=Dialog(dlg,1) ' show dialog (wait for ok)



    If dialogres<>1  Then Exit Sub
	'If the user presses "Cancel" or "x" then exit sub
     'If (dialogres<>-1 And dialogres<>1) Or dlg.Text="" Then Exit Sub

End Sub
Sub Process(StructureName,CaseStart,CaseEnd)

'
Dim objOpenSTAAD As Object
Set objOpenSTAAD = GetObject(, "StaadPro.OpenSTAAD")
'
Dim lBeamCnt As Long

If DlgValue("options")=0 Then
		'Get Beam list
		 lBeamCnt = objOpenSTAAD.GEOMETRY.GetMemberCount

ElseIf DlgValue("options")=1 Then
		'Get Beam list
        lBeamCnt=objOpenSTAAD.Geometry.GetNoOfSelectedBeams
                If lBeamCnt=0 Then
					MsgBox("No members selected! Please check.",vbInformation,"Select Members")
					End
				End If
End If

Dim StartLoadCaseNo As Long
Dim EndLoadCaseNo As Long


'Check whether the selected case number is a Repeat load case or Load Combination
If CaseStart<=lPrimaryLoadCaseCount-1 Then
	StartLoadCaseNo=lPrimaryLoadCaseNumbersArray(CaseStart)
ElseIf CaseStart>lPrimaryLoadCaseCount-1 And CaseStart<=lPrimaryLoadCaseCount+lGetLoadCombinationCaseCount-1 Then
	StartLoadCaseNo=lLoadCombinationCaseNumbersArray(CaseStart-lPrimaryLoadCaseCount)
End If

'Check whether the selected case number is a Repeat load case or Load Combination
If CaseEnd<=lPrimaryLoadCaseCount-1 Then
	EndLoadCaseNo=lPrimaryLoadCaseNumbersArray(CaseEnd)
ElseIf CaseEnd>lPrimaryLoadCaseCount-1 And CaseEnd<=lPrimaryLoadCaseCount+lGetLoadCombinationCaseCount-1 Then
	EndLoadCaseNo=lLoadCombinationCaseNumbersArray(CaseEnd-lPrimaryLoadCaseCount)
End If

'Ensure that the end load case number is greater than the start load case number
'If EndLoadCaseNo<StartLoadCaseNo Then
	'MsgBox("Start Load Case Number must be less than End Load Case Number!", 64,"Check Load Cases")
	'Exit Sub
'End If

Dim LoadCaseNos() As Long
Dim i,j As Long
Dim Counter As Long

'ReDim LoadCaseNos(1 To EndLoadCaseNo-StartLoadCaseNo+1)

'Find the actual number of cases which exist in STAAD between Start Load Case Number and End Load Case Numbers
For i= StartLoadCaseNo To EndLoadCaseNo
	For j=1 To lPrimaryLoadCaseCount+lGetLoadCombinationCaseCount
		If j<=lPrimaryLoadCaseCount Then
		 	If i= lPrimaryLoadCaseNumbersArray(j-1) Then
				Counter=Counter+1
				ReDim Preserve LoadCaseNos(1 To Counter) As Long
			 	LoadCaseNos(Counter)=i
			 	Exit For
			End If
	    ElseIf j<=lPrimaryLoadCaseCount+lGetLoadCombinationCaseCount Then
		 	If i= lLoadCombinationCaseNumbersArray(j-lPrimaryLoadCaseCount-1) Then
				Counter=Counter+1
				ReDim Preserve LoadCaseNos(1 To Counter) As Long
			 	LoadCaseNos(Counter)=i
				Exit For
			End If
		End If
	Next j

Next i


    Dim FileName
    'Use the user-entered structure name as the file name (as well as folder name)
	FileName=StructureName


Dim Shell
Set Shell = CreateObject("Shell.Application")
Dim folder
'Prompt user to select the folder path to store files
Set folder = Shell.BrowseForFolder(0, "Select Folder Or Enter Destination Path:",&H0030,&H0028)




'Start timer to measure code-execution time

StartTime = Timer

Dim fullpath As String


'Exit sub if the user selects nothing
If folder Is Nothing Then
DlgEnd 1000
Exit Sub
End If

Dim folderItem As Object

        Set folderItem = folder.Items.Item
		'Store the user-selected path
        fullpath = folderItem.Path

Set folder = Nothing

Set folderItem = Nothing
Set Shell = Nothing




		'Check if the path needs to be appended with "\". Also, append the path with new folder name (using the structure name entered by the user)
		If Right(fullpath,1)<>"\" Then
		fullpath= fullpath & "\" & FileName
		Else
		fullpath= fullpath & FileName
		End If

		'Create new folder based on user-entered structure name (if it doesn't already exist)
			On Error Resume Next
			MkDir fullpath
			If Dir(fullpath, vbDirectory) = "" Then
				MsgBox "Please ensure write permissions for the specifed path!"
			Exit Sub
			End If
			On Error GoTo 0







'Modify "fullpath" to be used for text file creation
fullpath= fullpath  & "\" & FileName

'Calculate parameters for progress updation
'Dim NoOfBeams As Long
'Dim NoOfNodes As Long
'Dim NoOfSupportNodes As Long

'Set objOpenSTAAD=Nothing





'

Dim lNodeCnt As Long
Dim BeamNumberArray() As Long
Dim NodeNumberArray() As Long
Dim iSupportCount As Integer
Dim ret As Long
'
Dim NodeA() As Long
Dim NodeB() As Long
'


If DlgValue("options")=0 Then

		 ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long
		ReDim NodeA(lBeamCnt-1) As Long
		ReDim NodeB(lBeamCnt-1) As Long
		 objOpenSTAAD.GEOMETRY.GetBeamList BeamNumberArray
		For i=0 To lBeamCnt-1
			objOpenSTAAD.GEOMETRY.GetMemberIncidence BeamNumberArray(i), NodeA(i), NodeB(i)
		Next i

		 lNodeCnt = objOpenSTAAD.Geometry.GetNodeCount()
		 ReDim NodeNumberArray(0 To (lNodeCnt-1)) As Long
		'Get node list
		 objOpenSTAAD.Geometry.GetNodeList NodeNumberArray

ElseIf DlgValue("options")=1 Then



		ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long
		ReDim NodeA(lBeamCnt-1) As Long
		ReDim NodeB(lBeamCnt-1) As Long

		ret=objOpenSTAAD.Geometry.GetSelectedBeams(BeamNumberArray,1)

		For i=0 To lBeamCnt-1
			objOpenSTAAD.GEOMETRY.GetMemberIncidence BeamNumberArray(i), NodeA(i), NodeB(i)
		Next i

		Dim MatchA As Long
		Dim MatchB As Long

		Counter=0
			 	 For i=0 To lBeamCnt-1

				 	 	'NodeNumberArray(Counter)=NodeA(i)
				 	 	MatchA=0
				 	 	MatchB=0
				 	 	If Counter=0 Then
					 	 	Counter=Counter+1
					 	 	ReDim Preserve NodeNumberArray(0 To Counter-1)
					 	 	NodeNumberArray(Counter-1)=NodeA(i)
				 	 	End If

				 	 	For j=0 To Counter-1
							If NodeNumberArray(j)=NodeA(i) Then
							MatchA=MatchA+1
							ElseIf NodeNumberArray(j)=NodeB(i) Then
							MatchB=MatchB+1
							End If
						Next j

						If MatchA=0 Then
							Counter=Counter +1
				 	 		ReDim Preserve NodeNumberArray(0 To Counter-1) As Long
				 	 		NodeNumberArray(Counter-1)=NodeA(i)
				 	 	End If
				 	 	If MatchB=0 Then
				 	 		Counter=Counter+1
				 	 		ReDim Preserve NodeNumberArray(0 To Counter-1) As Long
				 	 		NodeNumberArray(Counter-1)=NodeB(i)
				 	 	End If
			 	 Next i
				Dim Swap As Long
				'Sorting propeties in ascending order
				For i=0 To Counter-2
						For j=i+1 To Counter-1
							If NodeNumberArray(j)<NodeNumberArray(i) Then
									Swap=NodeNumberArray(j)
									NodeNumberArray(j)=NodeNumberArray(i)
									NodeNumberArray(i)=Swap
							End If
						Next j
				Next i
				lNodeCnt=UBound(NodeNumberArray)+1


		'lNodeCnt = objOpenSTAAD.Geometry.GetNoOfSelectedNodes
		'If lNodeCnt=0 Then
			'MsgBox("Use Geometry Cursor to select beams and nodes together.",vbInformation,"Select Geometry")
			'End
		'End If
	 	'ReDim NodeNumberArray(0 To (lNodeCnt-1)) As Long
		'Get node list
	 	'ret=objOpenSTAAD.Geometry.GetSelectedNodes(NodeNumberArray,1)

End If

'NoOfBeams = lBeamCnt
'NoOfNodes=lNodeCnt
iSupportCount=objOpenSTAAD.Support.GetSupportCount

MaxProgressCounter=lBeamCnt*(2*UBound(LoadCaseNos)+5)+lNodeCnt+iSupportCount



'Call individual sub-routines to create various text files
Call WriteForces(fullpath,objOpenSTAAD,LoadCaseNos(),ForceUnit,MomentUnit,lBeamCnt,BeamNumberArray,NodeA,NodeB)
'If DlgValue("options")=0 Then
	Call WriteBeams(fullpath,objOpenSTAAD,lBeamCnt,BeamNumberArray,NodeA,NodeB)
'ElseIf DlgValue("options")=1 Then
	'Call WriteBeamsSel(fullpath,objOpenSTAAD,lBeamCnt,BeamNumberArray,NodeA,NodeB)
'End If
Call WriteNodes(fullpath,objOpenSTAAD,DistanceUnit,lNodeCnt,NodeNumberArray)
Call WriteReleases(fullpath,objOpenSTAAD,lBeamCnt,BeamNumberArray)
Call WriteSpec(fullpath,objOpenSTAAD,lBeamCnt,BeamNumberArray)
Call WriteSupports(fullpath,objOpenSTAAD,iSupportCount)
Call WriteOffsets(fullpath,objOpenSTAAD,DistanceUnit,lBeamCnt,BeamNumberArray)
Call WriteProperties(fullpath,objOpenSTAAD,DistanceUnit,lBeamCnt,BeamNumberArray)


'Call WriteBeams
'Call WriteNodes
'Call WriteReleases
'Call WriteSpec
'Call WriteSupports
'Call WriteOffsets
'Call WriteProperties
'Call WriteForces(LoadCaseNos())


'Stop timer
'SecondsElapsed = Round(Timer - StartTime, 2)


 'MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation
'If ExitFlag<>1 Then
    Dim SecondsElapsed As Long
	If SecondsElapsed<60 Then
	    SecondsElapsed=Round(Timer-StartTime,0)
	    MsgBox("The code ran successfully in " & SecondsElapsed  & " sec.", 64,"Time Taken")
    ElseIf SecondsElapsed>=60 Then
		Dim TimeTakenMinutes As Long
		Dim TimeTakenSec As Long
		TimeTakenMinutes=Round(SecondsElapsed/60,0)
		TimeTakenSec= Round(SecondsElapsed-TimeTakenMinutes*60,0)
	    MsgBox("The code ran successfully in " & TimeTakenMinutes  & " min " & TimeTakenSec & "sec.", 64,"Time Elapsed")
	End If
DlgEnd 1000
'Break link with open STAAD file
'End If
Set objOpenSTAAD = Nothing
End Sub
Private Function MyDialogFunc(DlgItem$, Action%, SuppValue&) As Boolean
     'Static dlgCounter As Long
     'Static maxCounter As Long
     Dim gProgressBarWidth As Long
     'Dim processDone As Boolean
     Dim StructureName As String
     Dim CaseStart As Long
     Dim CaseEnd As Long
     Select Case Action%
     Case 1 ' Dialog box initialization
          'DlgEnable("OkButton", False)
          'DlgVisible("CopyButton", False)
          'DlgVisible("RunAgain", False)
          'DlgText("ProgText", "")
          'dlgCounter = 0
          'maxCounter = GetNumThingsToProcess() 'this returns an estimate of how many times ProcessStuff() will need to be called before it's done
          'ProcessStuff(True) 'initialize our working function
          'processDone = False

     Case 2 ' Value changing or button pressed

          If DlgItem$ = "Run"  Then 'reset everything
          		MaxProgressCounter=0
          		ProgressCounter=0
          		ExitFlag=0
               'DlgEnable("OkButton", False)
               'DlgVisible("CopyButton", False)

               'DlgText("ProgText", "")
               'dlgCounter = 0
               'maxCounter = GetNumThingsToProcess()
               'ProcessStuff(True) 'initialize our working function
               'InitOutputText()
               'DlgListBoxArray("OutputText", gOutputText)

				MyDialogFunc=True

               CaseStart=DlgValue("list1")
               CaseEnd=DlgValue("list2")

			    Dim StartLoadCaseNo As Long
				Dim EndLoadCaseNo As Long


			'Check whether the selected case number is a Repeat load case or Load Combination
			If CaseStart<=lPrimaryLoadCaseCount Then
			StartLoadCaseNo=lPrimaryLoadCaseNumbersArray(CaseStart)
			ElseIf CaseStart<=lPrimaryLoadCaseCount+lGetLoadCombinationCaseCount Then
			StartLoadCaseNo=lLoadCombinationCaseNumbersArray(CaseStart-lPrimaryLoadCaseCount)
			End If

			'Check whether the selected case number is a Repeat load case or Load Combination
			If CaseEnd<=lPrimaryLoadCaseCount Then
				EndLoadCaseNo=lPrimaryLoadCaseNumbersArray(CaseEnd)
		    ElseIf CaseEnd<=lPrimaryLoadCaseCount+lGetLoadCombinationCaseCount Then
				EndLoadCaseNo=lLoadCombinationCaseNumbersArray(CaseEnd-lPrimaryLoadCaseCount)
			End If



               StructureName=DlgText("Text")
               If StructureName="" Then
               		MsgBox "Please Enter Structure Name.", vbInformation
				ElseIf StartLoadCaseNo>EndLoadCaseNo  Then
					MsgBox("Start Load Case Number must be less than End Load Case Number!", 64,"Check Load Cases")

				Else
				'DlgEnable("Run", False)
				DlgEnable("list1", False)
				DlgEnable("list2", False)
				DlgEnable("Text", False)
'
				DlgEnable("Option1", False)
				DlgEnable("Option2", False)




               'Dim dlg As UserDialog
				'Dim dlg1,dlg2
               'dlg1=dlg.list1
               'dlg2=dlg.list2
               'MyDialogFunc = False
               Call Process(StructureName,CaseStart,CaseEnd)

               End If
               'processDone = False
          End If

          If DlgItem$ = "Cancel"  Then
          'MyDialogFunc = True
		  DlgEnd 1000
          'ExitFlag=1
          End If


     Case 3 ' TextBox or ComboBox text changed



     Case 4 ' Focus changed



     Case 5 ' Idle
     If MaxProgressCounter<>0 Then
          MyDialogFunc = True ' Continue getting idle actions
          'processDone = ProcessStuff() 'This is the function that's doing the real work (looping over elements, parsing data, etc)
          'If processDone Then
               'Wait 0.1 'so we don't take up 100% cpu time getting idle events...
          'End If
          'If Not processDone Then
                'dlgCounter = dlgCounter + 1
               'DlgListBoxArray("OutputText", gOutputText)
               'DoEvents

				If DlgFocus = "Cancel" Then
				'DlgEnd 1000
				'DlgEnable("Run", True)
				'MyDialogFunc = True
				ExitFlag=1
				MsgBox "Process aborted by user!",vbInformation
				End If
				'If ProgressCounter/MaxProgressCounter>0.75 Then
				'ExitFlag=1
				'End If
				Dim TimeMinutes As Long
				Dim TimeSeconds As Long
				TimeRemaining=(MaxProgressCounter-ProgressCounter)*(Timer-StartTime)/ProgressCounter
				If TimeRemaining<60 Then
				DlgText("TimeRem", Int(TimeRemaining) & " sec")
				Else
				TimeMinutes=Int(TimeRemaining/60)
				TimeSeconds=Int(TimeRemaining-TimeMinutes*60)
				DlgText("TimeRem", TimeMinutes & " min " & TimeSeconds & " sec")
				End If
				DlgText("PercComp", (Round(ProgressCounter/MaxProgressCounter*100,0) & " %"))
               gProgressBarWidth=0.4*385 * (ProgressCounter/MaxProgressCounter)
               DlgText("ProgText", String(gProgressBarWidth, "|"))
				End If
          'Else
               'DlgEnable("OkButton", True)
               'DlgVisible("CopyButton", True)
               'DlgVisible("RunAgain", True)
               'dlgCounter = maxCounter
          'End If
          'update output text and progress bar
     Case 6 ' Function key

     End Select
End Function

Sub WriteForces(fullpath As String,objOpenSTAAD As Object,LoadCaseNos() As Long,ForceUnit As String,MomentUnit As String,lBeamCnt As Long,BeamNumberArray() As Long,NodeA() As Long,NodeB() As Long)
'Sub WriteForces(LoadCaseNos() As Long)

Dim fullpathForces As String
'Specify file name for new text file
fullpathForces=fullpath &  "_Forces.txt"



'Create text file for Forces in the newly created folder
On Error Resume Next
Close #1
On Error GoTo 0
Open fullpathForces For Output As #1

Print #1, "Beam End Forces"
Write#1,
'Write units of forces
Print #1, "Force Unit:" & vbTab& ForceUnit
Print #1, "Moment Unit:" & vbTab& MomentUnit
Write#1,
'Column headers
Print #1, "Beam" & vbTab & 	"Node" & vbTab & 	"L/C" & vbTab & 	"FX" & vbTab & 	"FY" & vbTab & 	"FZ" & vbTab & 	"MX" & vbTab & 	"MY" & vbTab & 	"MZ"
Write#1,

'Dim lBeamCnt As Long



Dim i As Long
Dim j As Long
Dim k As Long
Dim ForceArray(0 To 5) As Double
'Dim NodeA As Long
'Dim NodeB As Long
Dim ret As Long

Dim SubString(1 To 9) As Variant


'Dim BeamNumberArray() As Long

'Get total number of beams in the structure/selected beams

'If DlgValue("options")=0 Then
		'Get Beam list


'		 lBeamCnt = objOpenSTAAD.GEOMETRY.GetMemberCount
	'	 ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long
	'	 objOpenSTAAD.GEOMETRY.GetBeamList BeamNumberArray

'		ElseIf DlgValue("options")=1 Then
'
 '       lBeamCnt=objOpenSTAAD.Geometry.GetNoOfSelectedBeams
	'	ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long
	'	ret=objOpenSTAAD.Geometry.GetSelectedBeams(BeamNumberArray,1)

'End If

Dim StringLine As String


Dim sBeamMaterialName As String

Dim BigString As String

Dim Position As Long


Dim lenSubString As Long
Dim l As Long




Position=1
BigString=Space(UBound(LoadCaseNos)*75)


'Dim StartTime2 As Double
'Dim ApproximateTime As Double
'StartTime2 = Timer

'For each beam in the structure, extract forces from STAAD for user-specified load cases
For i=0 To lBeamCnt-1
'DoEvents
'If DlgItem$="Cancel" Then
'Exit Sub
'Else
				'Action=5%
				'End If



		'If beam material is concrete then skip force extraction for such beams
		'sBeamMaterialName = objOpenSTAAD.Property.GetBeamMaterialName (BeamNumberArray(i))

	'	If sBeamMaterialName="CONCRETE" Then GoTo line10

		'objOpenSTAAD.GEOMETRY.GetMemberIncidence BeamNumberArray(i), NodeA, NodeB

		For j=0 To 1

			For k=1 To (UBound(LoadCaseNos))


			'For each user-specified load case, obtain force values from STAAD
				Erase ForceArray
				ret = objOpenSTAAD.Output.GetMemberEndForces(BeamNumberArray(i), j,LoadCaseNos(k), ForceArray)
				On Error Resume Next
				'Display Error message If analysis results are NOT available
				If IsError(ForceArray(0) )Then
					MsgBox "Please ensure that analysis results are available!"
				End
				End If
			    On Error GoTo 0

			'StringLine=Space(80)
			'StringLine=""
	     	If k =1 And j=0 Then
	     			'Mid$(StringLine,1,7)=Format(BeamNumberArray(i),"0")
	     			'Mid$(StringLine,8,14)=Format(NodeA,"0")
					'Mid$(StringLine,15,20)=Format(LoadCaseNos(k),"0")
					'Mid$(StringLine,21,30)=Format(ForceArray(0),"#.##")
					'Mid$(StringLine,31,40)=Format(ForceArray(1),"#.##")
					'Mid$(StringLine,41,50)=Format(ForceArray(2),"#.##")
					'Mid$(StringLine,51,60)=Format(ForceArray(3),"#.##")
					'Mid$(StringLine,61,70)=Format(ForceArray(4),"#.##")
					'Mid$(StringLine,71,80)=Format(ForceArray(5),"#.##")
				    'SubString(1)= BeamNumberArray(i)
				    'Print #1, vbTab;
				    'SubString(2)= NodeB
				    'Print #1, vbTab;
				    'SubString(3)= LoadCaseNos(k)
				    'Print #1, vbTab;
				    'SubString(4)= Round(ForceArray(0),2)
				    'Print #1, vbTab ;
				    'SubString(5)= Round(ForceArray(1),2)
				    'Print #1, vbTab ;
				    'SubString(6)=Round(ForceArray(2),2)
				    'Print #1, vbTab ;
				    'SubString(7)= Round(ForceArray(3),2)
				    'Print #1, vbTab ;
				    'SubString(8)= Round(ForceArray(4),2)
				    'Print #1, vbTab ;
				    'SubString(9)= Round(ForceArray(5),2)
				    If UBound(LoadCaseNos)=1 Then
					StringLine = BeamNumberArray(i) & vbTab & NodeA(i) & vbTab  & LoadCaseNos(k) & vbTab  & Round(ForceArray(0),2) & vbTab & Round(ForceArray(1),2) & vbTab & Round(ForceArray(2),2) & vbTab & Round(ForceArray(3),2) & vbTab & Round(ForceArray(4),2) & vbTab & Round(ForceArray(5),2)
						Else
				     StringLine = BeamNumberArray(i) & vbTab & NodeA(i) & vbTab  & LoadCaseNos(k) & vbTab  & Round(ForceArray(0),2) & vbTab & Round(ForceArray(1),2) & vbTab & Round(ForceArray(2),2) & vbTab & Round(ForceArray(3),2) & vbTab & Round(ForceArray(4),2) & vbTab & Round(ForceArray(5),2) & vbNewLine
					End If
ElseIf k=1 And j=1 Then
	     			'Mid$(StringLine,1,7)=Format(BeamNumberArray(i),"0")
	     			'Mid$(StringLine,8,14)=Format(NodeB,"0")
					'Mid$(StringLine,15,20)=Format(LoadCaseNos(k),"0")
					'Mid$(StringLine,21,30)=Format(ForceArray(0),"#.##")
					'Mid$(StringLine,31,40)=Format(ForceArray(1),"#.##")
					'Mid$(StringLine,41,50)=Format(ForceArray(2),"#.##")
					'Mid$(StringLine,51,60)=Format(ForceArray(3),"#.##")
					'Mid$(StringLine,61,70)=Format(ForceArray(4),"#.##")
					'Mid$(StringLine,71,80)=Format(ForceArray(5),"#.##")
				    'SubString(1)= BeamNumberArray(i)
				    'Print #1, vbTab;
				    'SubString(2)= NodeB
				    'Print #1, vbTab;
				    'SubString(3)= LoadCaseNos(k)
				    'Print #1, vbTab;
				    'SubString(4)= Round(ForceArray(0),2)
				    'Print #1, vbTab ;
				    'SubString(5)= Round(ForceArray(1),2)
				    'Print #1, vbTab ;
				    'SubString(6)=Round(ForceArray(2),2)
				    'Print #1, vbTab ;
				    'SubString(7)= Round(ForceArray(3),2)
				    'Print #1, vbTab ;
				    'SubString(8)= Round(ForceArray(4),2)
				    'Print #1, vbTab ;
				    'SubString(9)= Round(ForceArray(5),2)
					 If UBound(LoadCaseNos)=1 Then
				     StringLine =  vbTab & NodeB(i) & vbTab  & LoadCaseNos(k) & vbTab  & Round(ForceArray(0),2) & vbTab & Round(ForceArray(1),2) & vbTab & Round(ForceArray(2),2) & vbTab & Round(ForceArray(3),2) & vbTab & Round(ForceArray(4),2) & vbTab & Round(ForceArray(5),2)
					Else
					 StringLine =  vbTab & NodeB(i) & vbTab  & LoadCaseNos(k) & vbTab  & Round(ForceArray(0),2) & vbTab & Round(ForceArray(1),2) & vbTab & Round(ForceArray(2),2) & vbTab & Round(ForceArray(3),2) & vbTab & Round(ForceArray(4),2) & vbTab & Round(ForceArray(5),2) & vbNewLine
					End If
				Else



					'Mid$(StringLine,15,20)=Format$(LoadCaseNos(k),"0")
					'Mid$(StringLine,21,30)=Format$(ForceArray(0),"#.##")
					'Mid$(StringLine,31,40)=Format$(ForceArray(1),"#.##")
					'Mid$(StringLine,41,50)=Format$(ForceArray(2),"#.##")
					'Mid$(StringLine,51,60)=Format$(ForceArray(3),"#.##")
					'Mid$(StringLine,61,70)=Format$(ForceArray(4),"#.##")
					'Mid$(StringLine,71,80)=Format$(ForceArray(5),"#.##")
				    'SubString(1)= ""
				    'Print #1, vbTab;
				    'SubString(2)= ""
				    'Print #1, vbTab;
				    'SubString(3)= LoadCaseNos(k)
				    'Print #1, vbTab;
				    'SubString(4)= Round(ForceArray(0),2)
				    'Print #1, vbTab ;
				    'SubString(5)= Round(ForceArray(1),2)
				    'Print #1, vbTab ;
				    'SubString(6)=Round(ForceArray(2),2)
				    'Print #1, vbTab ;
				    'SubString(7)= Round(ForceArray(3),2)
				    'Print #1, vbTab ;
				    'SubString(8)= Round(ForceArray(4),2)
				    'Print #1, vbTab ;
				    'SubString(9)= Round(ForceArray(5),2)
					'If k=UBound(LoadCaseNos) And j=1 Then
					If k=UBound(LoadCaseNos) Then
				     StringLine =  vbTab &  vbTab  & LoadCaseNos(k) & vbTab  & Round(ForceArray(0),2) & vbTab & Round(ForceArray(1),2) & vbTab & Round(ForceArray(2),2) & vbTab & Round(ForceArray(3),2) & vbTab & Round(ForceArray(4),2) & vbTab & Round(ForceArray(5),2)
					Else
					 StringLine =  vbTab &  vbTab  & LoadCaseNos(k) & vbTab  & Round(ForceArray(0),2) & vbTab & Round(ForceArray(1),2) & vbTab & Round(ForceArray(2),2) & vbTab & Round(ForceArray(3),2) & vbTab & Round(ForceArray(4),2) & vbTab & Round(ForceArray(5),2) & vbNewLine
					End If
				End If
				lenSubString=Len(StringLine)
				Mid$(BigString,Position,lenSubString)=StringLine
				Position=Position+lenSubString

				'StringLine=Join(SubString,vbTab)
				'Print #1,StringLine
			Next k
			BigString=Trim$(BigString)
			Print #1,BigString
			If i=lBeamCnt-1 And j=1 Then GoTo line10
			Position=1
			BigString=Space(UBound(LoadCaseNos)*75)
		Next j


		'If (i>0 And (i Mod 100)=0) Or i=lBeamCnt Then

		'End If

'If i=19 And lBeamCnt>19 Then
'ApproximateTime=(Timer-StartTime2)*lBeamCnt/20
'MsgBox "Approximate time required is " & Round(ApproximateTime,2) & " seconds."

'End If

		line10:

				ProgressCounter=ProgressCounter+2*UBound(LoadCaseNos)
				If i Mod 10=0 Or i=lBeamCnt-1Then MyDialogFunc(DlgItem$, 5%, SuppValue&)
				'MyDialogFunc(DlgItem$, 5%, SuppValue&)
				If ExitFlag=1 Then Exit Sub
				'Set objOpenSTAAD=Nothing


	Next i

	Close #1


End Sub
Sub WriteBeamsAll(fullpath As String,objOpenSTAAD As Object,lBeamCnt As Long,BeamNumberArray() As Long)
'Sub WriteBeams

Dim fullpathBeams As String
'Create file name for new text file
fullpathBeams=fullpath &  "_Beams.txt"

'Create new text file at the user-specified path
On Error Resume Next
Close #1
On Error GoTo 0
Open fullpathBeams For Output As #1

Print #1, "Beam Incidences and Property Reference Numbers"
Write#1,
'Write column headings
Print #1, "Beam" & vbTab & 	"Node A" & vbTab & 	"Node B" & vbTab & "Ref.No."
Write#1,
'Dim lBeamCnt As Long

Dim i As Long

Dim NodeA As Long
Dim NodeB As Long
Dim ret As Long
Dim lBSPropertyTypeNo As Integer

'Get total number of beams in the structure/selected beams

'If DlgValue("options")=0 Then



'		 lBeamCnt = objOpenSTAAD.GEOMETRY.GetMemberCount
	'	 ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long
	'	 objOpenSTAAD.GEOMETRY.GetBeamList BeamNumberArray

'		ElseIf DlgValue("options")=1 Then

'        lBeamCnt=objOpenSTAAD.Geometry.GetNoOfSelectedBeams
	'	ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long
'		ret=objOpenSTAAD.Geometry.GetSelectedBeams(BeamNumberArray,1)

'End If


'For each beam in the structure, find member incidences and property reference numbers
For i=0 To lBeamCnt-1
    ProgressCounter=ProgressCounter+1
				If i Mod 10=0 Then MyDialogFunc(DlgItem$, 5%, SuppValue&)
				'MyDialogFunc(DlgItem$, 5%, SuppValue&)
				If ExitFlag=1 Then Exit Sub
	'Get Start and End Nodes

	objOpenSTAAD.GEOMETRY.GetMemberIncidence BeamNumberArray(i), NodeA, NodeB

	'Get Property
	lBSPropertyTypeNo = objOpenSTAAD.Property. GetBeamSectionPropertyRefNo (BeamNumberArray(i))
	Print #1, BeamNumberArray(i) & vbTab & NodeA & vbTab & NodeB & vbTab & lBSPropertyTypeNo
Next i
Close #1


End Sub
Sub WriteBeams(fullpath As String,objOpenSTAAD As Object,lBeamCnt As Long,BeamNumberArray() As Long,NodeA() As Long,NodeB() As Long)
'Sub WriteBeams

Dim fullpathBeams As String
'Create file name for new text file
fullpathBeams=fullpath &  "_Beams.txt"

'Create new text file at the user-specified path
On Error Resume Next
Close #1
On Error GoTo 0
Open fullpathBeams For Output As #1

Print #1, "Beam Incidences and Property Reference Numbers"
Write#1,
'Write column headings
Print #1, "Beam" & vbTab & 	"Node A" & vbTab & 	"Node B" & vbTab & "Ref.No."
Write#1,
'Dim lBeamCnt As Long

Dim i As Long

'Dim NodeA As Long
'Dim NodeB As Long
Dim ret As Long
Dim lBSPropertyTypeNo As Integer

'Get total number of beams in the structure/selected beams

'If DlgValue("options")=0 Then



'		 lBeamCnt = objOpenSTAAD.GEOMETRY.GetMemberCount
	'	 ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long
	'	 objOpenSTAAD.GEOMETRY.GetBeamList BeamNumberArray

'		ElseIf DlgValue("options")=1 Then

'        lBeamCnt=objOpenSTAAD.Geometry.GetNoOfSelectedBeams
	'	ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long
'		ret=objOpenSTAAD.Geometry.GetSelectedBeams(BeamNumberArray,1)

'End If


'For each beam in the structure, find member incidences and property reference numbers
For i=0 To lBeamCnt-1
    ProgressCounter=ProgressCounter+1
				If i Mod 10=0 Then MyDialogFunc(DlgItem$, 5%, SuppValue&)
				'MyDialogFunc(DlgItem$, 5%, SuppValue&)
				If ExitFlag=1 Then Exit Sub
	'Get Start and End Nodes

	'objOpenSTAAD.GEOMETRY.GetMemberIncidence BeamNumberArray(i), NodeA, NodeB

	'Get Property
	lBSPropertyTypeNo = objOpenSTAAD.Property. GetBeamSectionPropertyRefNo (BeamNumberArray(i))
	Print #1, BeamNumberArray(i) & vbTab & NodeA(i) & vbTab & NodeB(i) & vbTab & lBSPropertyTypeNo
Next i
Close #1


End Sub
Sub WriteNodes(fullpath As String,objOpenSTAAD As Object,DistanceUnit As String,lNodeCnt As Long,NodeNumberArray() As Long)
'Sub WriteNodes
If ExitFlag=1 Then Exit Sub
Dim fullpathNodes As String
'Create file name for the new text file
fullpathNodes=fullpath &  "_Nodes.txt"

'Create new text file
On Error Resume Next
Close #1
On Error GoTo 0
Open fullpathNodes For Output As #1
Print #1, "Node Coordinates"
Write#1,
Print #1, "Distance Unit:" & vbTab & DistanceUnit
Write#1,
Print #1, "Node" & vbTab & 	"X" & vbTab & 	"Y" & vbTab & "Z"
Write#1,

Dim ret As Long

' Dim lNodeCnt As Long
 'Dim NodeNumberArray() As Long


'Get Node Numbers
'If DlgValue("options")=0 Then
	' lNodeCnt = objOpenSTAAD.Geometry.GetNodeCount()
	 'ReDim NodeNumberArray(0 To (lNodeCnt-1)) As Long
	'Get node list
	 'objOpenSTAAD.Geometry.GetNodeList NodeNumberArray
'ElseIf DlgValue("options")=1 Then
	' lNodeCnt = objOpenSTAAD.Geometry.GetNoOfSelectedNodes
	 'ReDim NodeNumberArray(0 To (lNodeCnt-1)) As Long
	'Get node list
	 'ret=objOpenSTAAD.Geometry.GetSelectedNodes(NodeNumberArray,1)
 'End If


Dim CoordX As Double
Dim CoordY As Double
Dim CoordZ As Double


Dim i As Long

For i=1 To lNodeCnt
ProgressCounter=ProgressCounter+1
				If i Mod 10=0 Then MyDialogFunc(DlgItem$, 5%, SuppValue&)
				'MyDialogFunc(DlgItem$, 5%, SuppValue&)
				If ExitFlag=1 Then Exit Sub
'Get coordinates for each node in the structure

objOpenSTAAD.Geometry.GetNodeCoordinates NodeNumberArray(i-1), CoordX, CoordY, CoordZ

Print #1, NodeNumberArray(i-1) & vbTab & Round(CoordX,3) & vbTab & Round(CoordY,3) & vbTab & Round(CoordZ,3)

Next i

Close #1


End Sub
Sub WriteReleases(fullpath As String,objOpenSTAAD As Object,lBeamCnt As Long,BeamNumberArray() As Long)
'Sub WriteReleases
If ExitFlag=1 Then Exit Sub
'Dim lBeamCnt As Long
'Get number of beams in the structure
'lBeamCnt = objOpenSTAAD.GEOMETRY.GetMemberCount
'Dim BeamNumberArray() As Long
'ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long
'Get Beam list
'objOpenSTAAD.GEOMETRY.GetBeamList BeamNumberArray

Dim lReleaseArray(0 To 5) As Integer
Dim lSpringConstArray(0 To 5) As Double
Dim MPFactor As Double
Dim MPFactorArray(0 To 5) As Double

Dim ReleaseStart As Long
Dim ReleaseEnd As Long

Dim NodeA As Long
Dim NodeB As Long

Dim i As Long
Dim j As Integer
Dim k As Long

Dim Counter As Long
Counter=0

Dim ReleaseInfo() As String

'For each beam in the structure, check releases at start and end node
For i=0 To lBeamCnt-1
ProgressCounter=ProgressCounter+1
				If i Mod 10=0 Then MyDialogFunc(DlgItem$, 5%, SuppValue&)
				'MyDialogFunc(DlgItem$, 5%, SuppValue&)
				If ExitFlag=1 Then Exit Sub
	ReleaseStart=0
	ReleaseEnd=0
	objOpenSTAAD.GEOMETRY.GetMemberIncidence BeamNumberArray(i), NodeA, NodeB
'Check releases at each end of the beam
	For j=0 To 1
			Erase lReleaseArray
			Erase lSpringConstArray
			objOpenSTAAD.Property.GetMemberReleaseSpecEx (BeamNumberArray(i), j, lReleaseArray, lSpringConstArray,MPFactor,MPFactorArray)

			For k=0 To 5
				If lReleaseArray(k)>0 Or lSpringConstArray(k)>0  Or MPFactor>0 Or MPFactorArray(k)>0 Then
					Counter=Counter+1
					ReDim Preserve ReleaseInfo(1 To Counter) As String
					If j=0 Then ReleaseInfo(Counter)=BeamNumberArray(i) & vbTab & NodeA
					If j=1 Then ReleaseInfo(Counter)=BeamNumberArray(i) & vbTab & NodeB
					Exit For
				End If
			Next k

	Next j


Next i

'Create text file only if there are releases assigned in structure
	If Counter>0 Then
		Dim fullpathReleases As String
		'Create file name for the new text file
		fullpathReleases=fullpath &  "_Releases.txt"
	    'Create new text file
	On Error Resume Next
Close #1
On Error GoTo 0
	    Open fullpathReleases For Output As #1
        Print #1, "Beam End Releases"
        Write#1,
	    Print #1, "Only beams with releases provided at any of the ends are included in the list below."
		'Print #1, "In the second and third columns, 0 indicates no release"
		'Print #1, "while 1 indicates release/s applied to the beam end."
		Write#1,
		'Print #1, "Beam" & vbTab & 	"Start" & vbTab & 	"End"
		Print #1, "Beam" & vbTab & 	"Node"
		Write#1,

		For i=1 To Counter
		Print #1,ReleaseInfo(i)
		Next i

		Close #1
	End If

End Sub
Sub WriteSpec(fullpath As String,objOpenSTAAD As Object,lBeamCnt As Long,BeamNumberArray() As Long)
'Sub WriteSpec
If ExitFlag=1 Then Exit Sub

'Dim lBeamCnt As Long
'Get total number of beams in the structure
'lBeamCnt = objOpenSTAAD.GEOMETRY.GetMemberCount
'Dim BeamNumberArray() As Long
'ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long
'Get Beam list
'objOpenSTAAD.GEOMETRY.GetBeamList BeamNumberArray


Dim Counter As Long
Dim SpecInfo() As Long

Dim i As Long

Dim spec As Integer

Counter=0

'For each beam in the structure, obtain member specifications (if any) from STAAD
For i=0 To lBeamCnt-1
ProgressCounter=ProgressCounter+1
				If i Mod 10=0 Then MyDialogFunc(DlgItem$, 5%, SuppValue&)
				'MyDialogFunc(DlgItem$, 5%, SuppValue&)
				If ExitFlag=1 Then Exit Sub

	spec=0
	objOpenSTAAD.Property.GetMemberSpecCode BeamNumberArray(i), spec

	If spec>=0 Then
					Counter=Counter+1
					ReDim Preserve SpecInfo(1 To Counter) As Long
					SpecInfo(Counter)=BeamNumberArray(i)
					'Print #1,BeamNumberArray(i)
	End If


Next i

'Create text file only if there are member specifications in the structure
If Counter>0 Then
	Dim fullpathSpecifications As String
	'Create new text file name
    fullpathSpecifications=fullpath &  "_Specifications.txt"
    'Create new text file
On Error Resume Next
Close #1
On Error GoTo 0
    Open fullpathSpecifications For Output As #1
    Print #1, "Beam Specifications"
	Write #1,
	Print #1, "Only beams with assigned specifications are included in the list below."
    Write #1,
	Print #1,"Beam"
	Write #1,
	For i=1 To Counter
		Print #1,SpecInfo(i)
	Next i
	Close #1
End If

End Sub
Sub WriteSupports(fullpath As String,objOpenSTAAD As Object,iSupportCount As Integer)
'Sub WriteSupports
If ExitFlag=1 Then Exit Sub
Dim lSupportNodesArray() As Long
'Dim iSupportCount As Integer
'Get the application object --
'iSupportCount = objOpenSTAAD.Support.GetSupportCount
ReDim lSupportNodesArray(0 To (iSupportCount-1)) As Long
'Get Support Nodes

            objOpenSTAAD.Support.GetSupportNodes (lSupportNodesArray)

'Create text file only if there are member supports in the structure
If iSupportCount>0 Then
	Dim fullpathSupports As String
	'Create new text file name
    fullpathSupports=fullpath &  "_Supports.txt"
    'Create new text file
On Error Resume Next
Close #1
On Error GoTo 0
    Open fullpathSupports For Output As #1
    Print #1, "Supports"
	Write #1,
	Print #1, "Only nodes with assigned supports are included in the list below."
    Write #1,
	Print #1,"Node"
	Write #1,
	Dim i As Long
	For i=0 To iSupportCount-1
		ProgressCounter=ProgressCounter+1
						If i Mod 10=0 Then MyDialogFunc(DlgItem$, 5%, SuppValue&)
				'MyDialogFunc(DlgItem$, 5%, SuppValue&)
				If ExitFlag=1 Then Exit Sub
		Print #1,lSupportNodesArray(i)
	Next i
	Close #1
End If

End Sub
Sub WriteOffsets(fullpath As String,objOpenSTAAD As Object,DistanceUnit As String,lBeamCnt As Long,BeamNumberArray() As Long)
'Sub WriteOffsets
If ExitFlag=1 Then Exit Sub
'Dim lBeamCnt As Long
'Get total number of beams in the structure
'lBeamCnt = objOpenSTAAD.GEOMETRY.GetMemberCount
'Dim BeamNumberArray() As Long
'ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long
'Get Beam list
'objOpenSTAAD.GEOMETRY.GetBeamList BeamNumberArray


Dim StartX As Double
Dim StartY As Double
Dim StartZ As Double
Dim EndX As Double
Dim EndY As Double
Dim EndZ As Double

Dim Counter As Long
Counter =0
Dim OffSetInfo() As String


Dim i,j As Long

'Get offset information for each beam in the structure
For i=0 To lBeamCnt-1
ProgressCounter=ProgressCounter+1
				If i Mod 10=0 Then MyDialogFunc(DlgItem$, 5%, SuppValue&)
				'MyDialogFunc(DlgItem$, 5%, SuppValue&)
				If ExitFlag=1 Then Exit Sub
		StartX=0
		StartY=0
		StartZ=0
		EndX=0
		EndY=0
		EndZ=0
		'Get offset information for each end of a given beam
		For j=0 To 1
		'Get Property
		If j=0 Then
			objOpenSTAAD.Property.GetMemberGlobalOffSet (BeamNumberArray(i), j, StartX, StartY, StartZ)
		Else
			objOpenSTAAD.Property.GetMemberGlobalOffSet (BeamNumberArray(i), j, EndX, EndY, EndZ)
		End If
		Next j
		If StartX<>0 Or StartY<>0 Or StartZ<>0 Or EndX<>0 Or EndY<>0 Or EndZ<>0 Then
			Counter=Counter +1
			ReDim Preserve OffSetInfo(1 To Counter)
			OffSetInfo(Counter)= BeamNumberArray(i) & vbTab & Round(StartX,2) & vbTab & Round(StartY,2) & vbTab & Round(StartZ,2) & vbTab & Round(EndX,2) & vbTab & Round(EndY,2) & vbTab & Round(EndZ,2)
		End If
Next i

'Create text file only if there are offsets assigned in the structure
If Counter>0 Then
	Dim fullpathOffsets As String
    'Create new file name
	fullpathOffsets=fullpath &  "_Offsets.txt"
	'Create new text file
On Error Resume Next
Close #1
On Error GoTo 0
	Open fullpathOffsets For Output As #1
	Print #1, "Beam End Offsets"
	Write #1,
	Print #1, "Distance Unit:" & vbTab & DistanceUnit
	Write #1,
    Print #1, "Only beams with offsets provided at any of the ends are included in the list below."
	Write #1,
	Print #1, vbTab & "| Start OffSets  |" & vbTab & "|   End OffSets  |"
	Write #1,
	Print #1, "Beam"  & vbTab & "|X" & vbTab & "Y" & vbTab & "Z|" & vbTab & "|X" & vbTab & "Y" & vbTab & "Z|"
	Write #1,
	For i=1 To Counter
	Print #1, OffSetInfo(i)
	Next i
	Close #1
End If

End Sub
Sub WriteProperties(fullpath As String,objOpenSTAAD As Object,DistanceUnit As String,lBeamCnt As Long,BeamNumberArray() As Long)
'Sub WriteProperties
If ExitFlag=1 Then Exit Sub
Dim fullpathProperties As String
'Create new file name
fullpathProperties=fullpath &  "_Properties.txt"
'Create new text file
On Error Resume Next
Close #1
On Error GoTo 0
Open fullpathProperties For Output As #1
Print #1, "Beam Properties"
Write#1,
Print #1, "Ref.No." & vbTab & 	"Name" & vbTab & "Width" & vbTab & "Depth"
Print #1, vbTab & vbTab & DistanceUnit & vbTab &  DistanceUnit


Dim strPropertyName As String
'Dim lBeamCnt As Long
'Get total number of beam in the structure
'lBeamCnt = objOpenSTAAD.GEOMETRY.GetMemberCount

'Dim BeamNumberArray() As Long
'ReDim BeamNumberArray(0 To (lBeamCnt - 1)) As Long
'Get Beam list
'objOpenSTAAD.GEOMETRY.GetBeamList BeamNumberArray
'
'Dim sBeamMaterialName As String

Dim i As Long
Dim j As Long

Dim lBSPropertyTypeNo() As Integer
ReDim lBSPropertyTypeNos(0 To (lBeamCnt - 1)) As Integer


Dim dWidth()  As Double
Dim dDepth()  As Double
Dim dA_x  As Double
Dim dA_y  As Double
Dim dA_z  As Double
Dim dI_x As Double
Dim dI_y  As Double
Dim dI_z  As Double

ReDim dWidth(0 To (lBeamCnt - 1))  As Double
ReDim dDepth(0 To (lBeamCnt - 1))  As Double
'ReDim dA_x(0 To (lBeamCnt - 1))  As Double
'ReDim dA_y(0 To (lBeamCnt - 1))  As Double
'ReDim dA_z(0 To (lBeamCnt - 1))  As Double
'ReDim dI_x(0 To (lBeamCnt - 1))  As Double
'ReDim dI_y(0 To (lBeamCnt - 1))  As Double
'ReDim dI_z(0 To (lBeamCnt - 1))  As Double


Dim UniquePropertyNos() As Long
Dim Counter As Long
	For i=0 To lBeamCnt-1
	ProgressCounter=ProgressCounter+1
					If i Mod 10=0 Or i=lBeamCnt-1Then MyDialogFunc(DlgItem$, 5%, SuppValue&)

				'MyDialogFunc(DlgItem$, 5%, SuppValue&)
				If ExitFlag=1 Then Exit Sub
        'Get Property
		'sBeamMaterialName = objOpenSTAAD.Property.GetBeamMaterialName (BeamNumberArray(i))
		'If beam material is concrete then skip the property information for such beams
		'If sBeamMaterialName="CONCRETE" Then GoTo line20
		lBSPropertyTypeNos(i) = objOpenSTAAD.Property. GetBeamSectionPropertyRefNo (BeamNumberArray(i))
		'Get property information such as depth,width,etc.
		objOpenSTAAD.Property.GetBeamProperty (BeamNumberArray(i), dWidth(i), dDepth(i), dA_x, dA_y, dA_z, dI_x, dI_y, dI_z)

line20:
Next i



Dim Match As Long
Dim Index() As Long
Counter=0

'Find unique beam properties and sort them in ascending order
For i=0 To lBeamCnt-1

		If lBSPropertyTypeNos(i)=0 Then
			GoTo line30
		Else

			If Counter=0 Then
				Counter=Counter+1
				ReDim Preserve UniquePropertyNos(1 To Counter) As Long
				ReDim Preserve Index(1 To Counter) As Long
				UniquePropertyNos(Counter)=lBSPropertyTypeNos(i)
				Index(Counter)=i
			Else
				Match=0
				'Find unique property numbers
				For j=1 To Counter
					If UniquePropertyNos(j)=lBSPropertyTypeNos(i) Then Match=Match+1
				Next j

				If Match=0 Then

					Counter=Counter+1
					ReDim Preserve UniquePropertyNos(1 To Counter) As Long
					ReDim Preserve Index(1 To Counter) As Long
					UniquePropertyNos(Counter)=lBSPropertyTypeNos(i)
					Index(Counter)=i


				End If
			End If
		End If
line30:
Next i


Dim	Swap As Long
'Sorting propeties in ascending order
For i=1 To Counter-1
		For j=i+1 To Counter
			If UniquePropertyNos(j)<UniquePropertyNos(i) Then
					Swap=UniquePropertyNos(j)
					UniquePropertyNos(j)=UniquePropertyNos(i)
					UniquePropertyNos(i)=Swap
					Swap=Index(j)
				 	Index(j)=Index(i)
					Index(i)=Swap
			End If
		Next j
Next i

For i=1 To Counter
	'Get property names from property numbers
	objOpenSTAAD.Property.GetSectionPropertyName (UniquePropertyNos(i), strPropertyName)
	'Get Property
	Print #1,  UniquePropertyNos(i) & vbTab & strPropertyName & vbTab & Round(dWidth(Index(i)),2) & vbTab & Round(dDepth(Index(i)),2)


Next i

Close #1




End Sub
