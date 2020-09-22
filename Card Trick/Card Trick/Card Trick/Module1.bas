Attribute VB_Name = "Module1"
Public Type cCard
    cType As CardTypes
    cValue As CardValues
End Type

Public CardDeck(1 To 52) As cCard

Public Sub Shuffle()
Dim cT As CardTypes, cV As CardValues, rNum As Integer
    Randomize Timer
    Erase CardDeck
    For cT = Spades To Hearts
        For cV = Ace To King
            Do
                rNum = Int(Rnd * 52) + 1
            Loop Until CardDeck(rNum).cType = 0
            CardDeck(rNum).cType = cT
            CardDeck(rNum).cValue = cV
        Next cV
    Next cT
End Sub

