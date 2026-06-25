<%
'***********************************************
'***********************************************
sub vbsDbSetDictionaryEditUpdateFieldDefaults( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectFromList "EditUpdateFieldDefaults" , objVbsDb
end sub

'***********************************************
'***********************************************
function funStrDbEditFieldsRequestValuesListForEditUpdateFieldDefaults( objVbsDb )
'***********************************************
'***********************************************
  dim strEditFieldName
  funStrDbEditFieldsRequestValuesListForEditUpdateFieldDefaults = ""
  for each strEditFieldName in objVbsDb( "DictionaryEditUpdateFieldDefaults" )
    if not funBolEditableEditField( strEditFieldName , objVbsDb ) then
      ' strEditFieldName is an edit hide field or it is a read only field
      funStrDbEditFieldsRequestValuesListForEditUpdateFieldDefaults = _
        funStrDbEditFieldsRequestValuesListForEditUpdateFieldDefaults & _
        funVarDbEditFieldFormattedValue( strEditFieldName , objVbsDb( "DictionaryEditUpdateFieldDefaults" )( strEditFieldName ) , objVbsDb ) & constRequestValueListDelimiter
    end if
  next
end function

'***********************************************
'***********************************************
function funStrEditFieldNamesListForEditUpdateFieldDefaults( objVbsDb )
'***********************************************
'***********************************************
  dim strEditFieldName
  funStrEditFieldNamesListForEditUpdateFieldDefaults = ""
  for each strEditFieldName in objVbsDb( "DictionaryEditUpdateFieldDefaults" )
    if not funBolEditableEditField( strEditFieldName , objVbsDb ) then
      ' strEditFieldName is an edit hide field or it is a read only field
  	  funStrEditFieldNamesListForEditUpdateFieldDefaults = _
  	    funStrEditFieldNamesListForEditUpdateFieldDefaults & _
  	    funStrGlobalDbTypeFieldName( strEditFieldName , objVbsDb ) & " , "
  	end if
  next
end function
%>