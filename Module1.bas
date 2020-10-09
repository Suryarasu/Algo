Attribute VB_Name = "Module1"
Function marketmake(timeremaining)
    'RIT API Initialization
    Dim API As RIT2.API
    Set API = New RIT2.API
    'Verify Trading is happening within specified window
    If timeremaining < 296 And timeremaining > 6 Then
        'Verify Open Order Book is Empty
       If Not IsNumeric(Range("order2ID")) = True And IsNumeric(Range("order1ID")) = True Then
            'Verify Open Order is cancelled
            API.cancelorder (Range("order1ID"))
            'Ensure maintaining balanced position and not trading over limits
            If Range("current_position") > 10000 And Range("current_position") > -20000 Then
                'Selling
                OrderID = API.AddOrder("ALGO", Range("quantitytraded"), Range("algo_ask") + Range("reqspread"), -1, 1)
            'Ensure maintaining balanced position and not trading over limits
            ElseIf Range("current_position") < -10000 And Range("current_position") < 20000 Then
                'Buying
                OrderID = API.AddOrder("ALGO", Range("quantitytraded"), Range("algo_bid") - Range("reqspread"), 1, 1)
            End If
        'Submit Pair orders - buy and sell while verifying within trade limits
        ElseIf Not IsNumeric(Range("order1ID")) = True And Range("current_position") < 20000 And Range("current_position") > -20000 Then
            OrderID = API.AddOrder("ALGO", Range("quantitytraded"), Range("Algo_Bid") - Range("reqspread"), 1, 1)
            OrderID = API.AddOrder("ALGO", Range("quantitytraded"), Range("Algo_ask") + Range("Reqspread"), -1, 1)
            'Cancel pair orders if not executed timely
        ElseIf Range("timeelapsed") > Range("tick2") + 7 Then
            API.cancelorder (Range("order1ID"))
            API.cancelorder (Range("order2ID"))
        
        End If

      
End If

End Function



Function PARSERTD(str As String) As Variant
'This function is used to parse the Open Order output into a table
 Dim Rows() As String
 Dim Cols() As String
 Dim NoR As Integer
 Dim NoC As Integer
 If Len(Trim(str)) = 0 Then
     ReDim Res(0, 0) As String
     PARSERTD = Res
 Else
     Rows = Split(str, ";")
     Cols = Split(Rows(0), ",")
     NoR = UBound(Rows)
     NoC = UBound(Cols)
     ReDim Res(NoR + 1, NoC) As String
     For I = 0 To NoR
         Cols = Split(Rows(I), ",")
         For j = 0 To NoC
             Res(I, j) = Cols(j)
         Next j
     Next I
     PARSERTD = Res
 End If
End Function

Sub info()

    Dim API As RIT2.API
    Set API = New RIT2.API
    Dim status As Variant
    status = API.GetTickerInfo("ALGO")
    Range("L2", "T2") = status

End Sub
