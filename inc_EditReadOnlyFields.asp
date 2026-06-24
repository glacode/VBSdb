<%
'***********************************************
'***********************************************
function funBolEditFieldIsReadOnly( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "EditTableName" ) = "" then
    ' there is no edit field at all
    funBolEditFieldIsReadOnly = false
  else
    funBolEditFieldIsReadOnly = ( objVbsDb( "DictionaryEditReadOnlyFields" ).exists( uCase( strFieldName ) ) )
  end if
end function

'***********************************************
'***********************************************
sub vbsDbSetDictionaryEditReadOnlyFields( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectFromList "EditReadOnlyFields" , objVbsDb
  'vbsDbCheckEditFieldsErrors objVbsDb
end sub

'***********************************************
function funBolEditKeyFieldIsToBeSetToReadOnly( objColumn , objVbsDb )
'***********************************************
  funBolEditKeyFieldIsToBeSetToReadOnly = _
    objVbsDb( "DictionaryEditKeyFields" ).exists( uCase( objColumn.name ) ) and _
    ( ( VbsDbGetScreenType( objVbsDb ) = "Update" ) or ( objVbsDb( "RequestVBSdbClickClass" ) = "VBSdbApplyUpdate" ) )
end function

'***********************************************
function funBolEditFieldHasToBeSetToReadOnly( objColumn , objVbsDb )
'***********************************************
  funBolEditFieldHasToBeSetToReadOnly = ( _
    ( VbsDbGetScreenType( objVbsDb ) = "Delete" ) or _
    ( funBolAdoxFieldAutoincrement( objColumn , objVbsDb ) ) or _
    funBolAdoxEditFieldIsOleObjectType( objColumn , objVbsDb ) or _
    funBolEditKeyFieldIsToBeSetToReadOnly( objColumn , objVbsDb ) )
end function

'***********************************************
sub vbsDbSetDictionaryEditReadOnlyField( strFieldName , byRef objVbsDb )
'***********************************************
  if not objVbsDb( "DictionaryEditReadOnlyFields" ).exists( uCase( strFieldName ) ) then
    ' the field with name strFieldName is not in DictionaryEditReadOnlyFields yet
    objVbsDb( "DictionaryEditReadOnlyFields" ).add uCase( strFieldName ) , strFieldName
  end if
end sub

'***********************************************
sub vbsDbSetDictionaryEditReadOnlyFieldsAddAdditionalColumn( objColumn , byRef objVbsDb )
'***********************************************
  if funBolEditFieldHasToBeSetToReadOnly( objColumn , objVbsDb ) then
    ' either we are in a delete screen
    ' or the field is not updatable (as an example, it is an autonumber type field)
    ' or the field is an ole object type filed
    ' or the field is an edit key field and we either are in an update screen or we are executing a table update
    vbsDbSetDictionaryEditReadOnlyField objColumn.name , objVbsDb
  end if
end sub

'***********************************************
'***********************************************
sub vbsDbAddAdditionalFieldsToDictionaryEditReadOnlyFields( byRef objVbsDb )
'***********************************************
'***********************************************
  dim objColumn
  'for each objField in objVbsDb( "EditTableRecordSet" ).fields
  for each objColumn in funObjAdoxEditTableColumns( objVbsDb )
    vbsDbSetDictionaryEditReadOnlyFieldsAddAdditionalColumn objColumn , objVbsDb
  next
end sub
%>