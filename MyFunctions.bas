Attribute VB_Name = "MyFunctions"
Public Enum inListStatus
    inListUnknown
    inList
    notInlist
End Enum

Function ISLEAPYEAR(year) As Variant

'Leap years occur every four years, unless the year is divisible by 100
'then we skip the leap year, unless the year is also divisible by 400 then we don't skip.
'For example,  1996, 2000, and 2004 were all leap years but 2100 will not be.
'See https://en.wikipedia.org/wiki/Leap_year (viewed, 2021 February 9)

'This Excel VBA function takes the year as an input at returns TRUE if that year is a leap year,
'otherwise, it returns FALSE


'First check to see if the value passed to the function is an integer,
'if not, exit the function and return the #NUM error

If year <> Int(year) Then
    ISLEAPYEAR = CVErr(xlErrNum)
    Exit Function
End If

'Check to see if the value passed to the function is divisible by 4, 100, and 400. Depending on the results
'return TRUE if the value would be a leap year, otherwise return FALSE.

    If year Mod 4 = 0 Then
        If year Mod 100 = 0 Then
                If year Mod 400 = 0 Then
                    ISLEAPYEAR = True
                Else
                    ISLEAPYEAR = False
                End If
         Else
            ISLEAPYEAR = True
         End If
    Else
        ISLEAPYEAR = False
    End If
    
End Function



Function uniqueRandNumber(myArray As Variant, minNumber As Long, maxNumber As Long) As Variant

Dim newUniqueRandNumber As Long
Dim uniqueRandNumberStatus As inListStatus

uniqueRandNumberStatus = inListUnknown

Do While uniqueRandNumberStatus <> notInlist
    newUniqueRandNumber = WorksheetFunction.RandBetween(minNumber, maxNumber)
    
            For myLoop = LBound(myArray) To UBound(myArray)
            
                If newUniqueRandNumber = myArray(myLoop) Then
                    
                    uniqueRandNumberStatus = inList
                    Else
                    
                End If
            
            Next myLoop


    If uniqueRandNumberStatus = inListUnknown Then
       'if the status is inLIstUnknown after the For loop, setting to Not in list to end the loop
       uniqueRandNumberStatus = notInlist
    
    Else
        'Reseting to unknown for the next loop
        uniqueRandNumberStatus = inListUnknown
          
    End If
    
Loop


uniqueRandNumber = newUniqueRandNumber

End Function

Sub callTest()
 
Dim myX As Long
'Dim myRange As Range



myX = uniqueRandNumber(Array(100, 101, 102), 100, 106)

Debug.Print "My New Unique Random Number" & myX

End Sub



