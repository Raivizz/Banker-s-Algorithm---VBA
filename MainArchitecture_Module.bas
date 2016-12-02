Attribute VB_Name = "MainArchitecture_Module"
Public Sub BankerAlgorithmMain()
Dim NumberOfProcesses As Integer
Dim NumberOfResources As Integer

NumberOfProcesses = InputBox("Please enter the number of processes.", "Number of processes.")

If NumberOfProcesses = 1 Then
MsgBox "The algorithm cannot be used for a single process."
GoTo EndImmediately
End If

NumberOfResources = InputBox("Please enter the number of resources.", "Number of resources.")

If NumberOfResources = 0 Then
MsgBox "System cannot have 0 resources."
GoTo EndImmediately
End If

If IsNumeric(NumberOfProcesses) = False Or IsNumeric(NumberOfResources) = False Then
MsgBox "Variables must be numeric."
GoTo EndImmediately
End If

'Claim vector
'Unused
'MaxInput:
'                    MaxVector = InputBox("Please enter the maximum claim vector.", "Claim vector.")
'                    If MaxVector > NumberOfResources Then
'                    MsgBox "Maximum claim vector cannot exceed the number of resources."
'                    GoTo MaxInput
'                    End If

RequirementReached = 0
RequirementReachedNumOfProcess = 0

'For easier allocation, arrays are used instead of unique variables.
'Arrays also make for much easier-to-read cycles, as demonstrated below.
Dim AllocatedTo() As Integer, AllocatedToArrayEnd As Integer
AllocatedToArrayEnd = NumberOfProcesses
ReDim AllocatedTo(AllocatedToArrayEnd)

Dim MaximumDemand() As Integer, MaximumDemandArrayEnd As Integer
MaximumDemandArrayEnd = NumberOfProcesses
ReDim MaximumDemand(MaximumDemandArrayEnd)

NumberOfResourcesLeft = NumberOfResources

Dim ProcessResourcesRequired() As Integer, ProcessResourcesRequiredArrayEnd As Integer
ProcessResourcesRequiredArrayEnd = NumberOfProcesses
ReDim ProcessResourcesRequired(ProcessResourcesRequiredArrayEnd)

'1. Cycle
'Requests the user to input maximum resource demand for every process.
'Includes three error handlers.
For MaximumDemandCalculation = 1 To NumberOfProcesses
MaximumDemand(MaximumDemandCalculation) = InputBox("Maximum memory demand of process " & MaximumDemandCalculation, "Max demand of process " & MaximumDemandCalculation)

    If MaximumDemand(MaximumDemandCalculation) > NumberOfResources Then
    MsgBox "The process cannot demand more resources than the system has."
    GoTo EndImmediately
    End If
    
    If MaximumDemand(MaximumDemandCalculation) = 0 Then
    MsgBox "The process cannot demand zero resources."
    GoTo EndImmediately
    End If
    
    If IsNumeric(MaximumDemand(MaximumDemandCalculation)) = False Then
    MsgBox "Demand must be numeric."
    GoTo EndImmediately
    End If

Next
MaximumDemandCalculation = MaximumDemandCalculation - 1

'2. Cycle
'Requests the user to input resource allocation for every process.
'Includes three error handlers.
For ProcessAllocationCalculation = 1 To NumberOfProcesses
AllocatedTo(ProcessAllocationCalculation) = InputBox("How many resources are allocated to process " & ProcessAllocationCalculation, "Allocation for process " & ProcessAllocationCalculation)
Next
ProcessAllocationCalculation = ProcessAllocationCalculation - 1

    If AllocatedTo(MaximumDemandCalculation) > MaximumDemand(MaximumDemandCalculation) Then
    MsgBox "Resource allocation higher than maximum memory demand of current process. Unsafe."
    Verdict = "Unsafe"
    GoTo GenerateTable
    End If
    
    If AllocatedTo(MaximumDemandCalculation) = 0 Then
    MsgBox "Cannot allocate 0 resources."
    GoTo EndImmediately
    End If
    
    If IsNumeric(AllocatedTo(MaximumDemandCalculation)) = False Then
    MsgBox "Allocation must be numeric."
    GoTo EndImmediately
    End If

'3. Cycle
'Calculates, how many resources are left free after all processes have been allocated their resources.
For NumberOfResourcesLeftCalculation = 1 To NumberOfProcesses
NumberOfResourcesLeft = NumberOfResourcesLeft - AllocatedTo(NumberOfResourcesLeftCalculation)
Next
NumberOfResourcesLeftCalculation = NumberOfResourcesLeftCalculation - 1

'4. Cycle
'Calculates how many resources each process needs.
For ProcessResourcesRequiredCalculation = 1 To NumberOfProcesses
ProcessResourcesRequired(ProcessResourcesRequiredCalculation) = MaximumDemand(ProcessResourcesRequiredCalculation) - AllocatedTo(ProcessResourcesRequiredCalculation)
Next
ProcessResourcesRequiredCalculation = ProcessResourcesRequiredCalculation - 1

'5. Cycle
'Searches for a process to allocate the free resources to.
Dim RRNOP As Integer
For CheckIfRequiredFits = 1 To NumberOfProcesses
If ProcessResourcesRequired(CheckIfRequiredFits) <= NumberOfResourcesLeft Then
RequirementReached = RequirementReached + 1
Exit For
End If
Next
CheckIfRequiredFits = CheckIfRequiredFits - 1
RRNOP = CheckIfRequiredFits

'Generates a verdict based on if the requirement for safe work was reached.
If RequirementReached > 0 Then
GoTo FinalSteps:
Else: MsgBox "Not enough resources to allocate. Not safe. "
Verdict = "Unsafe"
GoTo GenerateTable
End If

'The final calculation steps.
'Calculates, if the freed process can "feed" other processes with it's resources.
'If at least one process can be "fed", the algorithm is considered a success and a "Safe" verdict is generated.
FinalSteps:
For FinalResourceAllocation = 1 To NumberOfProcesses
Calc01F = Calc01F + ProcessResourcesRequired(FinalResourceAllocation)
Next

Calc02F = Calc01F - ProcessResourcesRequired(RRNOP)

If MaximumDemand(RRNOP) <= Calc02F Then
MsgBox "System has enough resources for stable process work. Safe."
Verdict = "Safe"
GoTo GenerateTable
Else
MsgBox "Not enough resources to allocate. Not safe."
Verdict = "Unsafe"
End If

GenerateTable:
    ActiveSheet.UsedRange.Clear

    Range("A1").Value = "Processes"
    Range("A1").ColumnWidth = 29
    Range("A1").Font.FontStyle = "Bold"
    Range("B1").Value = "Resources Allocated"
    Range("B1").ColumnWidth = 20
    Range("B1").Font.FontStyle = "Bold"
    Range("C1").Value = "Maximum Resource Demand"
    Range("C1").ColumnWidth = 29
    Range("C1").Font.FontStyle = "Bold"
    Range("D1").Value = "Process Resource Requirement"
    Range("D1").ColumnWidth = 29
    Range("D1").Font.FontStyle = "Bold"
    
    Range("A" & NumberOfProcesses + 3).Value = "Amount of System Resources"
    Range("B" & NumberOfProcesses + 3).Value = NumberOfResources

    Range("A" & NumberOfProcesses + 4).Value = "Number of Resources Free"
    Range("B" & NumberOfProcesses + 4).Value = NumberOfResourcesLeft
    
        If NumberOfResourcesLeft <= 0 Then
        Range("B" & NumberOfProcesses + 4).Interior.Color = RGB(255, 0, 0)
        End If
        If NumberOfResourcesLeft = NumberOfResources Then
        Range("B" & NumberOfProcesses + 4).Value = "Error in generation."
        Range("B" & NumberOfProcesses + 4).Font.Color = vbRed
        End If

    Range("A" & NumberOfProcesses + 6).Value = "Algorithm Verdict"
    Range("B" & NumberOfProcesses + 6).Value = Verdict
    
    If Verdict = "Safe" Then
    Range("B" & NumberOfProcesses + 6).Font.Color = vbGreen
    Else
    Range("B" & NumberOfProcesses + 6).Font.Color = vbRed
    End If
    
    Columns("A:D").HorizontalAlignment = xlCenter
    
    Range("A1", "D" & NumberOfProcesses + 1).BorderAround (xlContinuous)

    Range("A1").Activate

    For A1 = 1 To NumberOfProcesses
    ActiveCell.Offset(A1, 0).Value = "Process" & A1
    Next

    For A2 = 1 To NumberOfProcesses
    ActiveCell.Offset(A2, 1).Value = AllocatedTo(A2)
    Next

    For A3 = 1 To NumberOfProcesses
    ActiveCell.Offset(A3, 2).Value = MaximumDemand(A3)
    Next

    For A4 = 1 To NumberOfProcesses
    ActiveCell.Offset(A4, 3).Value = ProcessResourcesRequired(A4)
    Next

EndImmediately:
End Sub
