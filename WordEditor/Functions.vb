Public Class Functions

    Public Shared Function fncLastDay(ByVal bvDate As Date) As Date

        Dim dReturnValue As Date

        dReturnValue = New Date(year:=bvDate.Year,
                           month:=bvDate.Month,
                           day:=System.DateTime.DaysInMonth(bvDate.Year, bvDate.Month))

        Return dReturnValue

    End Function


End Class
