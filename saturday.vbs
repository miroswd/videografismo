''''' Saturday ''''' 

''' Manual Date '''

'actualDay = 28
'actualMonth = 02
'actualYear = 2019
 
''' Actual Date '''

actualDay = day(now)
actualMonth = month(now)
actualYear = year(now)
 

''' Validations '''

'' weekday name

days = array("sunday","monday","tuesday","wednesday","thursday","friday","saturday")
week = weekday(date) - 1
weekName = days(week)
 

'' last day month
thirtyOne = array(1,3,5,7,8,10,12)
thirty = array(4,6,9,11)
 

'Identifies the last day of month

For each x in filter(thirtyOne,actualMonth)
    lastDay = 31
Next 

For each x in filter(thirty,actualMonth)
    lastDay = 30
Next



'' leap year or not

If actualMonth = 02 then
divFour = actualYear mod 4          'true if == 0
divHundred = actualYear mod 100     'false if == 0
divFourHundred = actualYear mod 400 'true if == 0

    If divFour = 0 or divHundred <> 0 and divFourHundred = 0 then
        lastDay = 29
    Else
        lastDay = 28
    End If
End IF

Select Case weekName
                
    Case "monday"
        'monday = 5
        saturday = actualDay + 5
   
        If saturday > lastDay then
            saturday = saturday - lastDay
            actualMonth = actualmonth + 1          

            If actualMonth = 13 then
                actualMonth = 1
            End If

            If actualMonth = 01 then
                actualYear = actualYear + 1
            End If
        End If

    Case "tuesday"
        'tuesday = 4
        saturday = actualDay + 4

        If saturday > lastDay then
            saturday = saturday - lastDay
            actualMonth = actualmonth + 1
            
            If actualMonth = 13 then
                actualMonth = 1
            End If

            If actualMonth = 01 then
                actualYear = actualYear + 1
            End If
        End If

    Case "wednesday"
        'wednesday = 3
        saturday = actualDay + 3

        If saturday > lastDay then
            saturday = saturday - lastDay
            actualMonth = actualmonth + 1

            If actualMonth = 13 then
                actualMonth = 1
            End If

            If actualMonth = 13 then
                actualYear = actualYear + 1
            End If
        End If

    Case "thursday"
        'thursday = 2
        saturday = actualDay + 2

        If saturday > lastDay then
            saturday = saturday - lastDay
            actualMonth = actualmonth + 1

            If actualMonth = 13 then
                actualMonth = 1
            End If

            If actualMonth = 01 then
                actualYear = actualYear + 1
            End If
        End If

    Case "friday"
        'friday = 1
         saturday = actualDay + 1

        If saturday > lastDay then
            saturday = saturday - lastDay
            actualMonth = actualmonth + 1

            If actualMonth = 13 then
                actualMonth = 1
            End If

            If actualMonth = 01 then
                actualYear = actualYear + 1
            End If
        End If

End Select         


If saturday < 10 then
    saturday = 0 & saturday
End If

If actualMonth < 10 then
    actualMonth = 0 & actualMonth
End If


saturday = saturday & "/" & actualMonth & "/" & actualYear

'msgbox saturday

ActiveCanvas.Scene.GetTemplate("Dia").text = saturday