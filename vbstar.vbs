'***** VBSTAR
'*****
'*********************************************

function to_date(dateval,conv_str)
if not isdate(dateval) then
to_date = "Invalid Date: " & dateval
exit function
end if 
dateval = cdate(dateval)
conv_str = ucase(conv_str)

 yearcnt = 0
 daycnt = 0
 monthcnt = 0
 hourcnt = 0

yearcnt = CharCntr(conv_str, "Y")
daycnt = CharCntr(conv_str, "D")
monthcnt = CharCntr(conv_str, "M")
hourcnt = CharCntr(conv_str, "H")
minutecnt =  CharCntr(conv_str, "N")
secondcnt =  CharCntr(conv_str, "S")

 yearval = string(yearcnt,"Y")
 montval = string(monthcnt,"M")
 dayval = string(daycnt,"D")
 hourval = string(hourcnt,"H")
 minuteval = string(minutecnt,"N")
 secondval = string(secondcnt,"S")
     Select case yearval 
          case "YYYY"  yearval = datepart("yyyy",dateval)
          case "YY"    yearval = mid(datepart("yyyy",dateval),3,2)
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
    Select case hourval 
          case "HH"  hourval = Lpad(datepart("h",dateval),2,"0")
          case "H"   hourval = datepart("h",dateval)
          Case else hourval = ""
     end select 
    Select case minuteval 
          case "NN"  minuteval = Lpad(datepart("n",dateval),2,"0")
          case "N"   minuteval = datepart("n",dateval)
          Case else minuteval = ""
     end select 
    Select case secondval 
          case "SS"  secondval = Lpad(datepart("s",dateval),2,"0")
          case "S"   secondval = datepart("s",dateval)
          Case else secondval = ""
     end select 
 to_date = replace(conv_str,string(yearcnt,"Y"),yearval)
 to_date = replace(to_date,string(monthcnt,"M"),montval)
 to_date = replace(to_date,string(monthcnt,"D"),dayval)
 to_date = replace(to_date,string(monthcnt,"H"),hourval)
 to_date = replace(to_date,string(monthcnt,"N"),minuteval)
 to_date = replace(to_date,string(monthcnt,"S"),secondval)

 set dateval = nothing
 set yearcnt = nothing
 set montval = nothing
 set dayval = nothing
 set hourval = nothing
 set minuteval = nothing
 set secondval = nothing
 set yearval = nothing
 set monthcnt = nothing

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

Function fileToMem(file)
set fso = CreateObject("Scripting.FileSystemObject")
ForReading=1
ForWriting=2
ForAppending=8
Unicode=-1
ASCII=0
set fsoFile = fso.OpenTextFile(file,ForReading,ASCII)
fileToMem = fsoFile.ReadAll
set fsoFile = nothing
End Function

Function memToFile(mem,file)
set fso = CreateObject("Scripting.FileSystemObject")
set fsoFile = fso.CreateTextFile(file,True)
fsoFile.Write mem
wscript.echo "File Created"
set fsoFile = nothing
End Function

Function UTC(dateval)
if not isdate(dateval) then
UTC = "Invalid Date: " & dateval
exit function
end if 
Set dateTime = CreateObject("WbemScripting.SWbemDateTime")  
dateval = cdate(dateval)
dateTime.SetVarDate (dateval)
UTC = dateTime.GetVarDate (false)
set dateval = nothing
set dateTime = nothing
End Function

Function UTCtoLocal(dateval)
if not isdate(dateval) then
UTC = "Invalid Date: " & dateval
exit function
end if 
Set dateTime = CreateObject("WbemScripting.SWbemDateTime")  
dateval = cdate(dateval)
dateTime.SetVarDate dateval,false
UTCtoLocal = dateTime.GetVarDate (true)
set dateval = nothing
set dateTime = nothing
End Function