Attribute VB_Name = "RecentWorks"
Option Explicit
Dim A As New cls2000Registry

Function CleanRecent()
Dim TempData() As String
Dim B As Long, C As Long
Dim Count As Long
Dim Ret1 As Boolean
Dim Ret2 As String

'This loop will search until he found the count of recent files
Do
Count = Count + 1 ' the number of files
Ret1 = A.ValueNameExists(HKEY_USERS, ".DEFAULT\Software\Microsoft\Visual Basic\6.0\RecentFiles", Val(Count)) ' ask if key exist
Loop Until Ret1 = False ' loop until the key does not exist
Count = Count - 1
ReDim TempData(Count)

'Getting data from registery and invalid data will be delete here
For B = 1 To Count  ' creating loop based on the count of files
Ret2 = A.getValue(HKEY_USERS, ".DEFAULT\Software\Microsoft\Visual Basic\6.0\RecentFiles", Val(B)) ' getting path
A.deleteValueName HKEY_USERS, ".DEFAULT\Software\Microsoft\Visual Basic\6.0\RecentFiles", Val(B) ' Delete the value name bequess we have the data in ret2
If Ret2 = "" Or Dir(Ret2) = "" Then GoTo 1 ' check if file exist if not then skip the saving of data
C = C + 1
TempData(C) = Ret2
1:
Next B

'Filetering the empty strings and get a new count
For B = 1 To Count  ' creating loop based on the count of files
If TempData(B) = "" Then ' if the data is empty then set the count to the data before and exit this count based loop
Count = B - 1
Exit For
End If
Next B

'Setting data back to registery
For B = 1 To Count  ' creating loop based on the count of files
A.setValue HKEY_USERS, ".DEFAULT\Software\Microsoft\Visual Basic\6.0\RecentFiles", Val(B), TempData(B), REG_SZ ' Delete the value name bequess we have the data in ret2
Next B

'end
End Function

Function DeleteRecent()
Dim B As Long
Dim Count As Long
Dim Ret1 As Boolean

'This loop will search until he found the count of recent files
Do
Count = Count + 1 ' the number of files
Ret1 = A.ValueNameExists(HKEY_CURRENT_USER, "Software\microsoft\Visual Basic\6.0\RecentFiles", Val(Count)) ' ask if key exist
Loop Until Ret1 = False ' loop until the key does not exist
Count = Count - 1
ReDim TempData(Count)

'Deleting all recent data from registery
For B = 1 To Count  ' creating loop based on the count of files
A.deleteValueName HKEY_CURRENT_USER, "Software\microsoft\Visual Basic\6.0\RecentFiles", Val(B) ' Delete the value name
Next B
End Function
