<%
'***********************************************
sub vbsDbSetDictionaryObjectForSearchHideFields( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectFromList "SearchHideFields" , objVbsDb
end sub

'***********************************************
sub vbsDbRemoveSearchHideFieldsFromDictionarySearchFields( byRef objVbsDb )
'***********************************************
  dim strFieldName
  for each strFieldName in objVbsDb( "DictionarySearchFields" )
    if objVbsDb( "DictionarySearchHideFields" ).exists( strFieldName ) then
      ' strFieldName is both in DictionarySearchFields and in DictionarySearchHideFields
      objVbsDb( "DictionarySearchFields" ).remove( strFieldName )
    end if
  next
end sub

'***********************************************
'***********************************************
sub vbsDbHandleSearchHideFields( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectForSearchHideFields objVbsDb
  vbsDbRemoveSearchHideFieldsFromDictionarySearchFields objVbsDb
end sub
%>