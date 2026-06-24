<%
'***********************************************
'***********************************************
function funBolEditFieldIsHideField( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  funBolEditFieldIsHideField = objVbsDb( "DictionaryEditHideFields" ).exists( strFieldName )
end function

'***********************************************
sub vbsDbSetDictionaryObjectForEditHideFields( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectFromList "EditHideFields" , objVbsDb
end sub

'***********************************************
sub vbsDbRemoveEditHideFieldsFromDictionaryEditFields( byRef objVbsDb )
'***********************************************
  dim strFieldName
  for each strFieldName in objVbsDb( "DictionaryEditFields" )
    if objVbsDb( "DictionaryEditHideFields" ).exists( strFieldName ) then
      ' strFieldName is both in DictionaryEditFields and in DictionaryEditHideFields
      objVbsDb( "DictionaryEditFields" ).remove( strFieldName )
    end if
  next
end sub

'***********************************************
'***********************************************
sub vbsDbHandleEditHideFields( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectForEditHideFields objVbsDb
  vbsDbRemoveEditHideFieldsFromDictionaryEditFields objVbsDb
end sub
%>