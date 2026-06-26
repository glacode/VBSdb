<%
'***********************************************
'***********************************************
function funStrInputControlDateFormat( strDateFieldName , objVbsDb )
'***********************************************
'***********************************************
  dim datValue
  datValue = objVbsDb( "EditTableRecordSet" ).fields( strDateFieldName ).value
  funStrInputControlDateFormat = VbsDbGetDateFormat( objVbsDb )
  funStrInputControlDateFormat = replace( funStrInputControlDateFormat , "YYYY" , year( datValue ) )
  funStrInputControlDateFormat = replace( funStrInputControlDateFormat , "MM" , month( datValue ) )
  funStrInputControlDateFormat = replace( funStrInputControlDateFormat , "DD" , day( datValue ) )
  funStrInputControlDateFormat = replace( funStrInputControlDateFormat , "hh" , hour( datValue ) )
  funStrInputControlDateFormat = replace( funStrInputControlDateFormat , "mm" , minute( datValue ) )
  funStrInputControlDateFormat = replace( funStrInputControlDateFormat , "ss" , second( datValue ) )
  if ( hour( datValue ) = 0 ) and ( minute( datValue ) = 0 ) and ( second( datValue ) = 0 ) then
    funStrInputControlDateFormat = _
      left( funStrInputControlDateFormat , inStr( funStrInputControlDateFormat , " " ) - 1 )
  end if
end function

'***********************************************
function funStrDateFormatMayBeWithMeridianDiscriminator( strDateExample )
'***********************************************
  funStrDateFormatMayBeWithMeridianDiscriminator = strDateExample
  funStrDateFormatMayBeWithMeridianDiscriminator = replace( funStrDateFormatMayBeWithMeridianDiscriminator , "3333" , "YYYY" )
  funStrDateFormatMayBeWithMeridianDiscriminator = replace( funStrDateFormatMayBeWithMeridianDiscriminator , "11" , "MM" )
  funStrDateFormatMayBeWithMeridianDiscriminator = replace( funStrDateFormatMayBeWithMeridianDiscriminator , "22" , "DD" )
  funStrDateFormatMayBeWithMeridianDiscriminator = replace( funStrDateFormatMayBeWithMeridianDiscriminator , "10" , "hh" )
  funStrDateFormatMayBeWithMeridianDiscriminator = replace( funStrDateFormatMayBeWithMeridianDiscriminator , "44" , "mm" )
  funStrDateFormatMayBeWithMeridianDiscriminator = replace( funStrDateFormatMayBeWithMeridianDiscriminator , "55" , "ss" )
end function

'***********************************************
function funStrDateFormat( strDateExample )
'***********************************************
  funStrDateFormat = funStrDateFormatMayBeWithMeridianDiscriminator( strDateExample )
  funStrDateFormat = left( funStrDateFormat , 19 )
end function

'***********************************************
function funStrStandardDateSeparatorLocal( strDateExample )
'***********************************************
  dim intStartYearPosition
  intStartYearPosition = inStr( strDateExample , "3333" )
  if intStartYearPosition = 1 then
    ' the year is at the beginning of the date
    funStrStandardDateSeparatorLocal = mid( strDateExample , 5 , 1 )
  else
    ' the year is not at the beginning of the date
    funStrStandardDateSeparatorLocal = mid( strDateExample , intStartYearPosition - 1 , 1 )
  end if
end function

'***********************************************
function funStrDatePart( strDateExample )
'***********************************************
  funStrDatePart = left( strDateExample , 10 )
end function

'***********************************************
function funStrTimePart( strDateExample )
'***********************************************
  funStrTimePart = mid( strDateExample , 12 )
end function

'***********************************************
function funStrValidDateSeparatorTry( strTryChar , strStandardDateSeparator , strDateExample )
'***********************************************
  dim strDatePart , strTimePart , strDateExampleReplaced
  strDatePart = funStrDatePart( strDateExample )
  strTimePart = funStrTimePart( strDateExample )
  strDateExampleReplaced = replace( strDatePart , strStandardDateSeparator , strTryChar ) & " " & strTimePart
  if isDate( strDateExampleReplaced ) then
    funStrValidDateSeparatorTry = strTryChar
  else
    funStrValidDateSeparatorTry = ""
  end if
end function

'***********************************************
function funStrValidDateSeparatorsLocal( strDateExample )
'***********************************************
  dim strStandardDateSeparator
  strStandardDateSeparator = funStrStandardDateSeparatorLocal( strDateExample )
  funStrValidDateSeparatorsLocal = "" & funStrValidDateSeparatorTry( "/" , strStandardDateSeparator , strDateExample )
  funStrValidDateSeparatorsLocal = funStrValidDateSeparatorsLocal & funStrValidDateSeparatorTry( "-" , strStandardDateSeparator , strDateExample )
  funStrValidDateSeparatorsLocal = funStrValidDateSeparatorsLocal & funStrValidDateSeparatorTry( "." , strStandardDateSeparator , strDateExample )
  funStrValidDateSeparatorsLocal = funStrValidDateSeparatorsLocal & funStrValidDateSeparatorTry( ":" , strStandardDateSeparator , strDateExample )  
end function

'***********************************************
function funStrValidTimeSeparatorTry( strTryChar , strStandardTimeSeparator , strDateExample )
'***********************************************
  dim strDatePart , strTimePart , strDateExampleReplaced
  strDatePart = funStrDatePart( strDateExample )
  strTimePart = funStrTimePart( strDateExample )
  strDateExampleReplaced = strDatePart & " " & replace( strTimePart , strStandardTimeSeparator , strTryChar )
  if isDate( strDateExampleReplaced ) then
    funStrValidTimeSeparatorTry = strTryChar
  else
    funStrValidTimeSeparatorTry = ""
  end if
end function

'***********************************************
function funStrValidTimeSeparatorsLocal( strTimeExample )
'***********************************************
  dim strStandardTimeSeparator
  strStandardTimeSeparator = mid( strTimeExample , 14 , 1 )
  funStrValidTimeSeparatorsLocal = "" & funStrValidTimeSeparatorTry( "/" , strStandardTimeSeparator , strTimeExample )
  funStrValidTimeSeparatorsLocal = funStrValidTimeSeparatorsLocal & funStrValidTimeSeparatorTry( "-" , strStandardTimeSeparator , strTimeExample )
  funStrValidTimeSeparatorsLocal = funStrValidTimeSeparatorsLocal & funStrValidTimeSeparatorTry( "." , strStandardTimeSeparator , strTimeExample )
  funStrValidTimeSeparatorsLocal = funStrValidTimeSeparatorsLocal & funStrValidTimeSeparatorTry( ":" , strStandardTimeSeparator , strTimeExample )  
end function

'***********************************************
'***********************************************
sub vbsDbSetVbsDbGetDateFormatRelatedProperties( objVbsDb )
'***********************************************
'***********************************************
  dim strDateExample
  strDateExample = cStr( dateSerial( 3333 , 11 , 22 ) & " " & timeSerial( 10 , 44 , 55 ) )
  objVbsDb( "GlobalDateFormat" ) = funStrDateFormat( strDateExample )
  objVbsDb( "GlobalDateFormatValidDateSeparators" ) = funStrValidDateSeparatorsLocal( strDateExample )
  objVbsDb( "GlobalDateFormatValidTimeSeparators" ) = funStrValidTimeSeparatorsLocal( strDateExample )
end sub

'***********************************************
'***********************************************
function funStrStandardDateSeparator( objVbsDb )
'***********************************************
'***********************************************
  dim intStartYearPosition
  intStartYearPosition = inStr( VbsDbGetDateFormat( objVbsDb ) , "YYYY" )
  if intStartYearPosition = 1 then
    ' the year is at the beginning of the date
    funStrStandardDateSeparator = mid( VbsDbGetDateFormat( objVbsDb ) , 5 , 1 )
  else
    ' the year is not at the beginning of the date
    funStrStandardDateSeparator = mid( VbsDbGetDateFormat( objVbsDb ) , intStartYearPosition - 1 , 1 )
  end if
end function

'***********************************************
'***********************************************
function funStrStandardTimeSeparator( objVbsDb )
'***********************************************
'***********************************************
  funStrStandardTimeSeparator = mid( VbsDbGetDateFormat( objVbsDb ) , 14 , 1 )
end function

'***********************************************
'***********************************************
function funStrValidDateSeparators( objVbsDb )
'***********************************************
'***********************************************
  funStrValidDateSeparators = objVbsDb( "GlobalDateFormatValidDateSeparators" )
end function

'***********************************************
'***********************************************
function funStrValidTimeSeparators( objVbsDb )
'***********************************************
'***********************************************
  funStrValidTimeSeparators = objVbsDb( "GlobalDateFormatValidTimeSeparators" )
end function

'***********************************************
'***********************************************
function VbsDbGetDateFormat( objVbsDb )
'***********************************************
'***********************************************
  VbsDbGetDateFormat = objVbsDb( "GlobalDateFormat" )
end function
%>