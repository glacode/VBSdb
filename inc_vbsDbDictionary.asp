<%
'***********************************************
sub vbsDbCheckErrorsForNameFieldForDictionary( arrStrCouple , strPropertyName , byRef objVbsDb )
'***********************************************
  if uBound( arrStrCouple ) >= 0 then
    if objVbsDB( "Dictionary" & strPropertyName ).exists( uCase( arrStrCouple( 0 ) ) ) then
      drawError strPropertyName & " bad assignment.<br><br>The field " & arrStrCouple( 0 ) & _
              " has been assigned twice." , objVbsDb
    end if
  end if
end sub

'***********************************************
sub vbsDbSetDictionaryObjectFromListWithSeparatorCharacters( strPrimarySeparator , strSecondarySeparator , strPropertyName , byRef objVbsDb )
'***********************************************
  dim arrStrList , intArrIndex , strCouple , arrStrCouple
  arrStrList = split( objVbsDb( strPropertyName ) , strPrimarySeparator )
  for intArrIndex = 0 to uBound( arrStrList )
    strCouple = arrStrList( intArrIndex )
    arrStrCouple = split( strCouple , strSecondarySeparator )
    vbsDbCheckErrorsForNameFieldForDictionary arrStrCouple , strPropertyName , objVbsDb
    'vbsDbCheckErrorRepeatedNameFieldForDictionary arrStrCouple , strPropertyName , objVbsDb
    if uBound( arrStrCouple ) >=1 then
      objVbsDB( "Dictionary" & strPropertyName ).add uCase( arrStrCouple( 0 ) ) , arrStrCouple( 1 )
      'objDictionary.add arrStrCouple( 0 ) , arrStrCouple( 1 )
    elseif uBound( arrStrCouple ) >= 0 then
      objVbsDB( "Dictionary" & strPropertyName ).add uCase( arrStrCouple( 0 ) ) , arrStrCouple( 0 )
      'objDictionary.add arrStrCouple( 0 ) , ""
    end if
  next
end sub

'***********************************************
'***********************************************
sub vbsDbSetDictionaryObjectFromList( strPropertyName , byRef objVbsDb )
'***********************************************
'***********************************************
  dim strPrimarySeparator , strSecondarySeparator
  'strPrimarySeparator = constStrDefaultMainSeparator
  'strSecondarySeparator = constStrDefaultSecondarySeparator
  strPrimarySeparator = funStrPrimarySeparator( strPropertyName , objVbsDb )
  strSecondarySeparator = funStrSecondarySeparator( strPropertyName , objVbsDb )
  set objVbsDb( "Dictionary" & strPropertyName ) = server.createObject( "scripting.dictionary" )
  objVbsDb( "Dictionary" & strPropertyName ).compareMode = 1
  vbsDbSetDictionaryObjectFromListWithSeparatorCharacters strPrimarySeparator , strSecondarySeparator , strPropertyName , objVbsDb
end sub

'***********************************************
sub vbsDbSetDictionaryFromListForNItemsCheckErrorFieldNameAlreadyExists( strPropertyName , arrStrItemValues , byRef objVbsDb )
'***********************************************
  if objVbsDB( "Dictionary" & strPropertyName ).exists( uCase( arrStrItemValues( 0 ) ) ) then
    drawError strPropertyName & " bad assignment.<br><br>The field " & arrStrItemValues( 0 ) & _
              " has been assigned twice." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbSetDictionaryFromListForNItemsCheckErrorWrongNumberOfItemValues( strPropertyName , arrStrItemNames , arrStrItemValues , byRef objVbsDb )
'***********************************************
  if uBound( arrStrItemNames ) + 1 <> uBound( arrStrItemValues ) then
    drawError strPropertyName & " bad assignment.<br><br>The field " & arrStrItemValues( 0 ) & _
              " has been assigned a wrong number of item values." , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbSetDictionaryFromListForNItemsCheckError( strPropertyName , arrStrItemNames , arrStrItemValues , byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryFromListForNItemsCheckErrorFieldNameAlreadyExists strPropertyName , arrStrItemValues , objVbsDb
  vbsDbSetDictionaryFromListForNItemsCheckErrorWrongNumberOfItemValues strPropertyName , arrStrItemNames , arrStrItemValues , objVbsDb
end sub

'***********************************************
sub vbsDbSetDictionaryObjectFromListForNItemsAddItemObject( strItemCompositValue , strSecondarySeparator , strPropertyName , strItemNames , byRef objVbsDb )
'***********************************************
  dim arrStrItemNames , arrStrItemValues , intArrIndex , objDictionaryItem
  arrStrItemNames = split( strItemNames , ";" )   ' attention: ";" is correct!! Here you don't have to use strPrimarySeparator
  arrStrItemValues = split( strItemCompositValue , strSecondarySeparator )
  vbsDbSetDictionaryFromListForNItemsCheckError strPropertyName , arrStrItemNames , arrStrItemValues , objVbsDb
  set objDictionaryItem = server.createObject( "scripting.dictionary" )
  objDictionaryItem.CompareMode = 1
	objDictionaryItem.add "strName" , arrStrItemValues( 0 )
  for intArrIndex = 0 to uBound( arrStrItemNames )
    objDictionaryItem.add arrStrItemNames( intArrIndex ) , arrStrItemValues( intArrIndex + 1 )
  next
'on error resume next
  objVbsDb( "Dictionary" & strPropertyName ).add arrStrItemValues( 0 ) , objDictionaryItem
'if err.number<> 0 then
  'response.end
'end if
end sub

'***********************************************
sub vbsDbSetDictionaryObjectFromListForNItemsWithSeparatorCharacters( strPrimarySeparator , strSecondarySeparator , strPropertyName , strItemNames , byRef objVbsDb )
'***********************************************
  dim strDictionaryName , arrStrList , intArrIndex
  arrStrList = split( objVbsDb( strPropertyName ) , strPrimarySeparator )
  for intArrIndex = 0 to uBound( arrStrList )
    vbsDbSetDictionaryObjectFromListForNItemsAddItemObject arrStrList( intArrIndex ) , strSecondarySeparator , strPropertyName , strItemNames , objVbsDb
  next
end sub

'***********************************************
'***********************************************
sub vbsDbSetDictionaryObjectFromListForNItems( strPropertyName , strItemNames , byRef objVbsDb )
'***********************************************
'***********************************************
  dim strPrimarySeparator , strSecondarySeparator
  'strPrimarySeparator = constStrDefaultMainSeparator
  'strSecondarySeparator = constStrDefaultSecondarySeparator
  strPrimarySeparator = funStrPrimarySeparator( strPropertyName , objVbsDb )
  strSecondarySeparator = funStrSecondarySeparator( strPropertyName , objVbsDb )
  set objVbsDb( "Dictionary" & strPropertyName ) = server.createObject( "scripting.dictionary" )
  objVbsDb( "Dictionary" & strPropertyName ).compareMode = 1
  vbsDbSetDictionaryObjectFromListForNItemsWithSeparatorCharacters _
    strPrimarySeparator , strSecondarySeparator , strPropertyName , strItemNames , objVbsDb
end sub

'***********************************************
'***********************************************
sub vbsDbDictionaryClose( byRef objDictionary )
'***********************************************
'***********************************************
	dim strObjDictionaryItemKey
	for each strObjDictionaryItemKey in objDictionary
		if isObject( objDictionary( strObjDictionaryItemKey ) ) then
			' l'elemento attuale, a sua volta, e' un objDictionary
			vbsDbDictionaryClose objDictionary( strObjDictionaryItemKey )
		end if
	next
	objDictionary.removeAll
	set objDictionary = nothing
end sub
%>