<%
'***********************************************
function funBolVbsDbValidateRequiredHasToBeDrawn( objVbsDb )
'***********************************************
  funBolVbsDbValidateRequiredHasToBeDrawn = _
    ( ( ( VbsDbGetScreenType( objVbsDb ) = "Add" ) or ( VbsDbGetScreenType( objVbsDb ) = "Update" ) ) and _
      ( objVbsDb( "EditValidateRequiredFields" ) <> "" ) )
end function

'***********************************************
sub drawValidateRequiredFunctionCallLine( strEditFieldName , objVbsDb )
'***********************************************
  if not funBolEditFieldIsReadOnly( strEditFieldName , objVbsDb ) then
    ' the edit field is not read only and it is a required edit field
%>
    message = message + vbsDbValidateRequired(form.<%=funStrFormFieldName( strEditFieldName , objVbsDb )%>,"<%=objVbsDb( "DictionaryEditValidateRequiredFields" )( strEditFieldName )%>")
<%
  end if
end sub

'***********************************************
'***********************************************
sub drawValidateFunctionCallLinesForEditValidateRequiredFieldsActually( objVbsDb )
'***********************************************
'***********************************************
  dim strEditFieldName
  for each strEditFieldName in objVbsDb( "DictionaryEditValidateRequiredFields" )
    drawValidateRequiredFunctionCallLine strEditFieldName , objVbsDb
  next
end sub

'***********************************************
'***********************************************
sub drawValidateFunctionCallLinesForEditValidateRequiredFields( objVbsDb )
'***********************************************
'***********************************************
  if funBolVbsDbValidateRequiredHasToBeDrawn( objVbsDb ) then
    ' vbsDbValidateRequired javascript funcion has to be drawn
    drawValidateFunctionCallLinesForEditValidateRequiredFieldsActually objVbsDb
  end if
end sub

'***********************************************
'***********************************************
sub drawVbsDbValidateRequiredActually( objVbsDb )
'***********************************************
'***********************************************
%>
  function vbsDbValidateRequired( objField , strMessage )
  {
	var msg_addition = ""
	//new_fieldname = fieldname
	if (objField.value == "")
	  return strMessage + "\n"
	else
	  return ""
  }
<%
end sub

'***********************************************
'***********************************************
sub drawVbsDbValidateRequired( objVbsDb )
'***********************************************
'***********************************************
  if funBolVbsDbValidateRequiredHasToBeDrawn( objVbsDb ) then
    ' vbsDbValidateRequired javascript funcion has to be drawn
    drawVbsDbValidateRequiredActually objVbsDb
  end if
end sub

'***********************************************
sub vbsDbSetDictionaryObjectForFieldsValidateRequired( byRef objVbsDb )
'***********************************************
  vbsDbSetDictionaryObjectFromList "EditValidateRequiredFields" , objVbsDb
end sub

'***********************************************
'***********************************************
sub vbsDbSetDictionaryEditValidateRequiredFields( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectForFieldsValidateRequired objVbsDb
end sub
%>