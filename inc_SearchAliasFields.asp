<%
'***********************************************
function funBolWasSearchField( strFieldName , objVbsDb )
'***********************************************
  funBolWasSearchField = ( request.form( funStrVbsDbHiddenFieldTypeName( strFieldName , objVbsDb ) ) <> "" )
end function

'***********************************************
sub vbsDbCheckSearchAliasFieldErrors( strFieldNameToDisplay , strFieldName , objVbsDb )
'***********************************************
  dim strErrorMessage
  if not funBolWasSearchField( strFieldName , objVbsDb ) then
    drawError "SearchAliasFields property is wrong: " & strFieldNameToDisplay & " was not a field in the search screen" , objVbsDb
  end if
end sub

'***********************************************
'***********************************************
sub vbsDbCheckSearchAliasFieldsErrors( objVbsDb )
'***********************************************
'***********************************************
  dim strFieldName , strFieldNameToDisplay
  for each strFieldName in objVbsDb( "DictionarySearchAliasFields" )
    strFieldNameToDisplay = funStrFieldNameToDisplay( strFieldName , "DictionarySearchAliasFields" , objVbsDb )
    vbsDbCheckSearchAliasFieldErrors strFieldNameToDisplay , strFieldName , objVbsDb
  next
end sub

'***********************************************
'***********************************************
sub vbsDbSetDictionarySearchAliasFields( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectFromList "SearchAliasFields" , objVbsDb
end sub

'***********************************************
'***********************************************
function funStrSearchAliasField( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "DictionarySearchAliasFields" )( strFieldName ) = "" then
    ' strFieldName is not a field alias
    funStrSearchAliasField = strFieldName
  else
    ' objVbsDb( "DictionarySearchAliasFields" )( uCase( strFieldName ) ) is the field of which strFieldName is the alias
    funStrSearchAliasField = objVbsDb( "DictionarySearchAliasFields" )( strFieldName )
  end if
end function
%>