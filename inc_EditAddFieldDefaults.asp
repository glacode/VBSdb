<%
'***********************************************
'***********************************************
sub vbsDbSetDictionaryEditAddFieldDefaults( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectFromList "EditAddFieldDefaults" , objVbsDb
end sub

'***********************************************
'***********************************************
function funStrDbEditFieldsRequestValuesListForEditAddFieldDefaults( objVbsDb )
'***********************************************
'***********************************************
  dim strEditFieldName
  funStrDbEditFieldsRequestValuesListForEditAddFieldDefaults = ""
  for each strEditFieldName in objVbsDb( "DictionaryEditAddFieldDefaults" )
    if not funBolEditableEditField( strEditFieldName , objVbsDb ) then
      ' strEditFieldName is an edit hide field or it is a read only field
      funStrDbEditFieldsRequestValuesListForEditAddFieldDefaults = _
        funStrDbEditFieldsRequestValuesListForEditAddFieldDefaults & _
        funVarDbEditFieldFormattedValue( strEditFieldName , objVbsDb( "DictionaryEditAddFieldDefaults" )( strEditFieldName ) , objVbsDb ) & constRequestValueListDelimiter
    end if
  next
end function

'***********************************************
'***********************************************
function funStrEditFieldNamesListForEditAddFieldDefaults( objVbsDb )
'***********************************************
'***********************************************
  dim strEditFieldName
  funStrEditFieldNamesListForEditAddFieldDefaults = ""
  for each strEditFieldName in objVbsDb( "DictionaryEditAddFieldDefaults" )
    if not funBolEditableEditField( strEditFieldName , objVbsDb ) then
      ' strEditFieldName is an edit hide field or it is a read only field
  	  funStrEditFieldNamesListForEditAddFieldDefaults = _
  	    funStrEditFieldNamesListForEditAddFieldDefaults & _
  	    funStrGlobalDbTypeFieldName( strEditFieldName , objVbsDb ) & " , "
  	end if
  next
end function
%>