<div align="center">

## Last day of the month


</div>

### Description

Returns the last day of a specified month. Takes into account leap years.
 
### More Info
 
Month (optional), Year (optional)

Last day of the month


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Karabunga](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/karabunga.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/karabunga-last-day-of-the-month__1-12343/archive/master.zip)





### Source Code

```
Function LastDay(Optional MyMonth As Integer, Optional MyYear As Integer) As Integer
  ' Returns the last day of the month. Takes into account leap years
  ' Usage: LastDay(Month, Year)
  ' Example: LastDay(12,2000) or LastDay(12) or Lastday
  If MyMonth = 0 Then MyMonth = Month(Date)
  Select Case MyMonth
    Case 1, 3, 5, 7, 8, 10, 12
      LastDay = 31
    Case 4, 6, 9, 11
      LastDay = 30
    Case 2
      If MyYear = 0 Then MyYear = Year(Date)
      If IsDate(MyYear & "-" & MyMonth & "-" & "29") Then LastDay = 29 Else LastDay = 28
    Case Else
      LastDay = 0
  End Select
End Function
```

