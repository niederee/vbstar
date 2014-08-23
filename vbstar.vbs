'***** VBSTAR
'*****
'*********************************************

function to_date(dateval,conv_str)
dateval = cdate(dateval)
conv_str = ucase(conv_str)

 yearcnt = 0
 daycnt = 0
 monthcnt = 0

yearcnt = CharCntr(conv_str, "Y")
daycnt = CharCntr(conv_str, "D")
monthcnt = CharCntr(conv_str, "M")

 yearval = string(yearcnt,"Y")
 montval = string(monthcnt,"M")
 dayval = string(daycnt,"D")
     Select case yearval 
          case "YYYY"  yearval = datepart("yyyy",dateval)
          case "YY"    yearval = datepart("yy",dateval)
          Case else yearval = ""
     end select 
     Select case montval 
          case "MM"  montval = Lpad(datepart("m",dateval),2,"0")
          case "M"   montval = datepart("m",dateval)
          Case else montval = ""
     end select 
      Select case dayval 
          case "DD"  dayval = Lpad(datepart("d",dateval),2,"0")
          case "D"   dayval = datepart("d",dateval)
          Case else dayval = ""
     end select 

 to_date = replace(conv_str,string(yearcnt,"Y"),yearval)
 to_date = replace(to_date,string(monthcnt,"M"),montval)
 to_date = replace(to_date,string(monthcnt,"D"),dayval)

end function

Function LPad(s, l, c)
  Dim n : n = 0
  If l > Len(s) Then n = l - Len(s)
  LPad = String(n, c) & s
End Function

Function CharCntr( StringVal,  Charcnt) 
If Len(Charcnt) Then
    CharCntr = UBound(Split(StringVal, Charcnt))
End If
End Function
