Public Function joursOuvrés(ByVal DateDebut As Date, ByVal DateFin As Date, ByVal DemiJourDébut As String, ByVal DemiJourFin As String, ByVal Region As Boolean) As Double

    Dim CurrentDate As Date
    Dim jour As Integer
    Dim DemiJour As String
    Dim JourException As String
    Dim Cn As Object
    Dim Rst As Object
    Dim SelectExceptions As String
    Dim SelectRessources As String
    Dim ConnectionString As String

    For CurrentDate = DateDebut To DateFin
        If Weekday(CurrentDate) > 1 And Weekday(CurrentDate) < 7 Then
            If Not EstJourFérié(CurrentDate, Region) Then
                If CurrentDate = DateDebut Then
                    DemiJour = DemiJourDébut
                ElseIf CurrentDate = DateFin Then
                    DemiJour = DemiJourFin
                Else
                    DemiJour = "entier"
                End If
                JourException = ExceptionContractuelle(CurrentDate)

                Select Case JourException
                    Case "M"
                        Select Case DemiJour
                            Case "matin"
                                JourConge = JourConge + 0
                            Case "après-midi"
                                JourConge = JourConge + 0.5
                            Case "entier"
                                JourConge = JourConge + 0.5
                        End Select
                    Case "A"
                        Select Case DemiJour
                            Case "matin"
                                JourConge = JourConge + 0.5
                            Case "après-midi"
                                JourConge = JourConge + 0
                            Case "entier"
                                JourConge = JourConge + 0.5
                        End Select
                    Case "E"
                        JourConge = JourConge + 0
                    Case "N"
                        JourConge = JourConge + 1
                End Select
            End If
        End If
    Next
    joursOuvrés = JourConge
End Function
