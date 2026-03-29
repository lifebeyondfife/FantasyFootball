Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
'
'   Copyright (C) Iain McDonald 2012
'
'   This program is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
Dim bestScore As Double
Dim playerOffset As Integer


Sub searchForPlayers()
    bestScore = 0
    playerOffset = 18
  
    Dim availablePlayerList As Collection
    Set availablePlayerList = New Collection
    Dim chosenPlayerList(1 To 15) As Integer
    
    For i = 1 To Range("G1").Value
        availablePlayerList.Add (i)
    Next i
    
    Range("A1:E15").Clear
        
    recurseForSolution chosenPlayerList, availablePlayerList, 1
    
    
End Sub


Sub recurseForSolution(chosenPlayerList, availablePlayerList As Collection, depth As Integer)

    If depth = 16 Then
        printSolution
    Else
        
        Do While availablePlayerList.Count > 0
            Dim playerNumber As Integer
            Dim conViolated As Boolean
            
            playerNumber = availablePlayerList.Item(1)
            availablePlayerList.Remove (1)
            
            copyRow depth, playerNumber
            
            conViolated = constraintsViolated(depth)
            If Not conViolated Then
                Dim availablePlayerListCopy As Collection
                Set availablePlayerListCopy = New Collection
                            
                chosenPlayerList(depth) = playerNumber
                For Each playerNumberObject In availablePlayerList
                    availablePlayerListCopy.Add (playerNumberObject)
                Next
                
                recurseForSolution chosenPlayerList, availablePlayerListCopy, depth + 1
            End If
        Loop
        
        clearRow depth
    End If

End Sub

Sub clearRow(depth As Integer)
    Dim clearCell As Range
    Set clearCell = Range("A1").Offset(depth - 1, 0)
    
    For i = 0 To 4
        clearCell.Offset(0, i).Clear
    Next i
End Sub

Sub printSolution()
    If Range("G3").Value > bestScore Then
        bestScore = Range("G3").Value
            
        Dim copyToCell As Range
        Dim copyFromCell As Range
        
        Set copyToCell = Range("I1")
        Set copyFromCell = Range("A1")
        
        For j = 0 To 14
            For i = 0 To 4
                copyToCell.Offset(j, i).Value = copyFromCell.Offset(j, i).Value
            Next i
        Next j
    End If
        
End Sub

Sub copyRow(copyTo As Integer, rowNumber As Integer)
    Dim copyToCell As Range
    Dim copyFromCell As Range
    
    Set copyToCell = Range("A1").Offset(copyTo - 1, 0)
    Set copyFromCell = Range("A1").Offset(rowNumber + playerOffset, 0)
        
    For i = 0 To 4
        copyToCell.Offset(0, i).Value = copyFromCell.Offset(0, i).Value
    Next i
    
End Sub

Function constraintsViolated(depth As Integer)
    Dim returnValue As Boolean
    Dim adjustment As Double
    returnValue = False
    adjustment = (15 - depth) * 4.5
    
    If Range("G4").Value > (100 - adjustment) Then
        returnValue = True
    End If
    
    If Range("G8").Value > 3 Then
        returnValue = True
    End If
    
    If Range("G7").Value > 5 Then
        returnValue = True
    End If
    
    If Range("G6").Value > 5 Then
        returnValue = True
    End If
    
    If Range("G5").Value > 2 Then
        returnValue = True
    End If
    
    If Range("G30").Value > 3 Then
        returnValue = True
    End If
    
    constraintsViolated = returnValue
End Function

