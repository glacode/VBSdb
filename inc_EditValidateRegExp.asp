<%
'***********************************************
function funBolVbsDbValidateRegExpHasToBeDrawn( objVbsDb )
'***********************************************
  funBolVbsDbValidateRegExpHasToBeDrawn = _
    ( ( ( VbsDbGetScreenType( objVbsDb ) = "Add" ) or ( VbsDbGetScreenType( objVbsDb ) = "Update" ) ) and _
      ( objVbsDb( "EditValidateRegExp" ) <> "" ) )
end function

'***********************************************
sub drawValidateRegExpFunctionCallLine( strEditFieldName , objVbsDb )
'***********************************************
  if not funBolEditFieldIsReadOnly( strEditFieldName , objVbsDb ) then
    ' the edit field is not read only and it is a regular expression edit field
%>
    message = message + vbsDbValidateRegExp(form.<%=uCase( funStrFormFieldName( strEditFieldName , objVbsDb ) )%>,/<%=objVbsDb( "DictionaryEditValidateRegExp" )( strEditFieldName )( "regExp" )%>/,"<%=objVbsDb( "DictionaryEditValidateRegExp" )( strEditFieldName )( "errorMessage" )%>")
<%
  end if
end sub

'***********************************************
sub drawValidateFunctionCallLinesForEditValidateRegExpActually( objVbsDb )
'***********************************************
  dim strEditFieldName
  for each strEditFieldName in objVbsDb( "DictionaryEditValidateRegExp" )
    drawValidateRegExpFunctionCallLine strEditFieldName , objVbsDb
  next
end sub

'***********************************************
'***********************************************
sub drawValidateFunctionCallLinesForEditValidateRegExp( objVbsDb )
'***********************************************
'***********************************************
  if funBolVbsDbValidateRegExpHasToBeDrawn( objVbsDb ) then
    ' vbsDbValidateRegExp javascript funcion has to be drawn
    drawValidateFunctionCallLinesForEditValidateRegExpActually objVbsDb
  end if
end sub

'***********************************************
sub drawVbsDbValidateRegExpActually( objVbsDb )
'***********************************************
%>
  function vbsDbValidateRegExp( objField , strRegExp , strMessage )
  {
	var arrRegExpItems;
	arrRegExpItems = strRegExp.exec( objField.value );
	if ( ( arrRegExpItems == null ) || ( arrRegExpItems.input != arrRegExpItems[ 0 ] ) )
	  // no match
	  return strMessage + "\n";
	else
	  // good match
	  return "";
  }
<%
end sub

'***********************************************
'***********************************************
sub drawVbsDbValidateRegExp( objVbsDb )
'***********************************************
'***********************************************
  if funBolVbsDbValidateRegExpHasToBeDrawn( objVbsDb ) then
    ' vbsDbValidateRequired javascript funcion has to be drawn
    drawVbsDbValidateRegExpActually objVbsDb
  end if
end sub

'***********************************************
'***********************************************
sub vbsDbSetDictionaryEditValidateRegExp( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectFromListForNItems "EditValidateRegExp" , "regExp;errorMessage" , objVbsDb
end sub
%>