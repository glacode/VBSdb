<%
'***********************************************
'***********************************************
function funStrPrimarySeparator( strPropertyName , objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "DictionaryGlobalSeparators" ).exists( strPropertyName ) then
    ' custom separators have been defined for the property named strPropertyName
    funStrPrimarySeparator = objVbsDb( "DictionaryGlobalSeparators" )( strPropertyName )( "Primary" )
  else
    ' custom separators have been defined for the property named strPropertyName
    funStrPrimarySeparator = constStrDefaultPrimarySeparator
  end if
end function

'***********************************************
'***********************************************
function funStrSecondarySeparator( strPropertyName , objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "DictionaryGlobalSeparators" ).exists( strPropertyName ) then
    ' custom separators have been defined for the property named strPropertyName
    funStrSecondarySeparator = objVbsDb( "DictionaryGlobalSeparators" )( strPropertyName )( "Secondary" )
  else
    ' custom separators have been defined for the property named strPropertyName
    funStrSecondarySeparator = constStrDefaultSecondarySeparator
  end if
end function

'***********************************************
sub vbsDbCheckErrorPropertyNameForGlobalSeparators( strPossiblePropertyName , objVbsDb )
'***********************************************
  if not funBolIsAPropertyName( strPossiblePropertyName , objVbsDb ) then
    drawError "GlobalSeparators" & " bad assignment.<br><br>" & _
      strPossiblePropertyName & " is not a property name." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbSetGlobalSeparatorsCheckErrorNotDivisibleByThree( arrStrList , objVbsDb )
'***********************************************
  if ( uBound( arrStrList ) + 1 ) mod 3 <> 0 then
    drawError "GlobalSeparators" & " bad assignment.<br><br>" & _
      "The number of items assigned is not divisible by 3." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbSetGlobalSeparatorsLoop( byRef objVbsDb )
'***********************************************
  dim arrStrList , intArrayIndex , strCouple , arrStrCouple
  arrStrList = split( trim( objVbsDb( "GlobalSeparators" ) ) , " " )
  vbsDbSetGlobalSeparatorsCheckErrorNotDivisibleByThree arrStrList , objVbsDb
  for intArrayIndex = 0 to uBound( arrStrList ) step 3
    vbsDbCheckErrorPropertyNameForGlobalSeparators arrStrList( intArrayIndex ) , objVbsDb
    set objVbsDB( "DictionaryGlobalSeparators" )( arrStrList( intArrayIndex ) ) = _
      server.createObject( "scripting.dictionary" )
    objVbsDB( "DictionaryGlobalSeparators" )( arrStrList( intArrayIndex ) ).compareMode = 1
    objVbsDB( "DictionaryGlobalSeparators" )( arrStrList( intArrayIndex ) ).add "Primary" , arrStrList( intArrayIndex + 1 )
    objVbsDB( "DictionaryGlobalSeparators" )( arrStrList( intArrayIndex ) ).add "Secondary" , arrStrList( intArrayIndex + 2 )
  next
end sub

'***********************************************
'***********************************************
sub vbsDbSetGlobalSeparators( byRef objVbsDb )
'***********************************************
'***********************************************
  set objVbsDb( "DictionaryGlobalSeparators" ) = server.createObject( "scripting.dictionary" )
  objVbsDb( "DictionaryGlobalSeparators" ).compareMode = 1
  vbsDbSetGlobalSeparatorsLoop objVbsDb
end sub
%>