'FalsePassInterceptor
'Author: Radu Jorda
'=================================================================================================================
' This script was developed to prevent a critical bug encountered in live production
' within a test automation system.
'=================================================================================================================
' Issue Summary:
' An incorrect modification of a configuration .xml file caused the test application to
' skip actual testing, yet still report a successful "PASS" both visually and in the traceability
' system. As a result, the system would also print an "OK" label — despite no test being performed.
'
' This script also includes an input validation mechanism that ensures item names only contain
' safe characters ([A-Z0-9_;,-/]) — protecting against human error or corrupted input from upstream
' design tools which could otherwise break the test system.
' =================================================================================================
' Original Variable Mapping (for reference):
' [ReceivedItems]  - Input string: Names of items declared by the upstream configuration system
' [VerifiedItems]  - Input string: Names of items that were actually processed and validated
'
' [ScriptStatus]      - Output numeric: 0 = mismatch detected / 1 = validation OK
' [UnverifiedItems]   - Output string: List of received items that had no corresponding verification
'
' Originally stored in numeric memory slots via a proprietary COM interface.
' =================================================================================================

OPTION EXPLICIT

' --- SIMULATED VALUES USING A DICTIONARY (mock COM-style variable slots) ---
dim variableStore : Set variableStore = CreateObject("Scripting.Dictionary")

sub SetVariable(name, value)
    if variableStore.Exists(name) then
        variableStore(name) = value
    else
        variableStore.Add name, value
    end if
end sub

function GetVariable(name)
    if variableStore.Exists(name) then
        GetVariable = variableStore(name)
    else
        GetVariable = ""
    end if
end function

SetVariable "receivedItems", "ITEM001,ITEM002,ITEM003"
SetVariable "testedItems",   "ITEM001,ITEM002"

dim delimiter
dim foundMatch
dim splitReceivedItems
dim splitTestedItems
dim arrayReceivedItems()
dim arrayTestedItems()
dim receivedItem
dim verifiedItem
dim reservedReceivedItem
dim reservedVerifiedItem
dim containedItem
dim checkArray
dim unmatchedItems

' Entry point (used by the test system to trigger this script)
sub main

delimiter = ","

SetVariable "receivedItems", Replace(GetVariable("receivedItems"), ";", delimiter)
SetVariable "testedItems", Replace(GetVariable("testedItems"), ";", delimiter)

filterAndCorrect
duplicationRemover

splitReceivedItems = Split(GetVariable("receivedItems"), delimiter)
splitTestedItems = Split(GetVariable("testedItems"), delimiter)

prepareArrays

if sorting(arrayTestedItems,arrayReceivedItems) then
    SetVariable "scriptStatus", 0
    SetVariable "unverifiedItems", "Unmatched Items: " & Join(unmatchedItems, ",")
else
    SetVariable "scriptStatus", 1
    SetVariable "unverifiedItems", ""
end if

WScript.Echo "ScriptStatus: " & GetVariable("scriptStatus")
WScript.Echo GetVariable("unverifiedItems")

end sub

function sorting(arrayTestedItems,arrayReceivedItems)
    unmatchedItems = Array()

    for each receivedItem in arrayReceivedItems
        foundMatch = false
        reservedReceivedItem = receivedItem
        for each verifiedItem in arrayTestedItems
            reservedVerifiedItem = verifiedItem
           if verifiedItem = receivedItem then
                foundMatch = true
            end if
        next

        if not foundMatch then
            if reservedReceivedItem = "MissingReceivedItem" then
                containedItem = reservedVerifiedItem
            elseif reservedVerifiedItem = "MissingTestedItem" then
                containedItem = reservedReceivedItem
            end if
            if UBound(unmatchedItems) = -1 then
                ReDim unmatchedItems(0)
            else
                ReDim Preserve unmatchedItems(UBound(unmatchedItems) + 1)
            end if
            checkArray = isarray(unmatchedItems)
            checkArray = empty
            unmatchedItems(UBound(unmatchedItems)) = containedItem
        end if
    next
    sorting = (UBound(unmatchedItems) >= 0)
end function

sub prepareArrays
    dim i
    ReDim preserve arrayReceivedItems(ubound(splitReceivedItems))
    for i = 0 to ubound(splitReceivedItems)
        arrayReceivedItems(i) = splitReceivedItems(i)
    next
    ReDim preserve arrayTestedItems(ubound(splitTestedItems))
    for i = 0 to ubound(splitTestedItems)
        arrayTestedItems(i) = splitTestedItems(i)
    next
    If UBound(arrayReceivedItems) > UBound(arrayTestedItems) Then
        ReDim preserve arrayTestedItems(ubound(arrayReceivedItems))
        for i = ubound(arrayTestedItems) to ubound(arrayReceivedItems)
            arrayTestedItems(i)="MissingTestedItem"
        next
    Elseif UBound(arrayReceivedItems) < UBound(arrayTestedItems) then
        ReDim preserve arrayReceivedItems(ubound(arrayTestedItems))
        for i = ubound(arrayReceivedItems) to ubound(arrayTestedItems)
            arrayReceivedItems(i)="MissingReceivedItem"
        next
    End If
End Sub

sub duplicationRemover
    Dim checkDuplicates
    dim itemDR
    dim itemDT
    dim j
    dim key
    dim uniqueReceived()
    dim uniqueTested()
    dim tempReceived
    dim tempTested
    Set checkDuplicates = CreateObject("Scripting.Dictionary")
    tempReceived = split(GetVariable("receivedItems"),delimiter)
    tempTested = split(GetVariable("testedItems"),delimiter)

    For Each itemDR in tempReceived
        If Not checkDuplicates.Exists(itemDR) Then
            checkDuplicates.Add itemDR, itemDR
        End If
    Next
    Redim uniqueReceived(checkDuplicates.Count - 1)
    j = 0
    For Each key in checkDuplicates.Keys
        uniqueReceived(j) = key
        j = j + 1
    Next

    checkDuplicates.RemoveAll

    For Each itemDT in tempTested
        If Not checkDuplicates.Exists(itemDT) Then
            checkDuplicates.Add itemDT, itemDT
        End If
    Next
    Redim uniqueTested(checkDuplicates.Count - 1)
    j = 0
    For Each key in checkDuplicates.Keys
        uniqueTested(j) = key
        j = j + 1
    Next

    SetVariable "receivedItems", join(uniqueReceived, ",")
    SetVariable "testedItems", join(uniqueTested, ",")
end sub

sub filterAndCorrect
    dim re
    set re = New RegExp
    re.Pattern = "^[A-Z0-9_;,-/]*$"
    
    if GetVariable("receivedItems") = "" Then
        SetVariable "scriptStatus", 0
        SetVariable "unverifiedItems", "receivedItems is empty!"
        WScript.Quit(1)
    elseif NOT re.test(GetVariable("receivedItems")) then
        SetVariable "scriptStatus", 0
        SetVariable "unverifiedItems", "receivedItems contains invalid characters"
        WScript.Quit(1)
    else
        if Mid(GetVariable("receivedItems"), Len(GetVariable("receivedItems")), 1) = ";" OR Mid(GetVariable("receivedItems"), Len(GetVariable("receivedItems")), 1) = "," Then
            SetVariable "receivedItems", Mid(GetVariable("receivedItems"), 1, Len(GetVariable("receivedItems")) - 1)
        end if
    end if

    if GetVariable("testedItems") = "" Then
        SetVariable "scriptStatus", 0
        SetVariable "unverifiedItems", "testedItems is empty!"
        WScript.Quit(1)
    elseif NOT re.test(GetVariable("testedItems")) then
        SetVariable "scriptStatus", 0
        SetVariable "unverifiedItems", "testedItems contains invalid characters"
        WScript.Quit(1)
    else
        if Mid(GetVariable("testedItems"), Len(GetVariable("testedItems")), 1) = ";" OR Mid(GetVariable("testedItems"), Len(GetVariable("testedItems")), 1) = "," Then
            SetVariable "testedItems", Mid(GetVariable("testedItems"), 1, Len(GetVariable("testedItems")) - 1)
        end if
    end if
end sub

main
