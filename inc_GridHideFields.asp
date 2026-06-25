<%
'***********************************************
sub vbsDbSetDictionaryObjectForGridHideFields( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectFromList "GridHideFields" , objVbsDb
end sub

'***********************************************
sub vbsDbRemoveGridHideFieldsFromDictionaryGridFields( byRef objVbsDb )
'***********************************************
  dim strFieldName
  for each strFieldName in objVbsDb( "DictionaryGridFields" )
    if objVbsDb( "DictionaryGridHideFields" ).exists( strFieldName ) then
      ' strFieldName is both in DictionaryGridFields and in DictionaryGridHideFields
      objVbsDb( "DictionaryGridFields" ).remove( strFieldName )
    end if
  next
end sub

'***********************************************
'***********************************************
sub vbsDbHandleGridHideFields( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectForGridHideFields objVbsDb
  vbsDbRemoveGridHideFieldsFromDictionaryGridFields objVbsDb
end sub
%>