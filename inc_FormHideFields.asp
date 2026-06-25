<%
'***********************************************
sub vbsDbSetDictionaryObjectForFormHideFields( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectFromList "FormHideFields" , objVbsDb
end sub

'***********************************************
sub vbsDbRemoveFormHideFieldsFromDictionaryFormFields( byRef objVbsDb )
'***********************************************
  dim strFieldName
  for each strFieldName in objVbsDb( "DictionaryFormFields" )
    if objVbsDb( "DictionaryFormHideFields" ).exists( strFieldName ) then
      ' strFieldName is both in DictionaryFormFields and in DictionaryFormHideFields
      objVbsDb( "DictionaryFormFields" ).remove( strFieldName )
    end if
  next
end sub

'***********************************************
'***********************************************
sub vbsDbHandleFormHideFields( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectForFormHideFields objVbsDb
  vbsDbRemoveFormHideFieldsFromDictionaryFormFields objVbsDb
end sub
%>