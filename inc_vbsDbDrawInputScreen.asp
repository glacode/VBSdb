<%
'***********************************************
sub drawJavaScriptFunctions( objVbsDb )
'***********************************************
  drawValidateJavascriptFunction( objVbsDb )
  drawJavaScriptFunctionForIntMaxEditCharsForMemoFields objVbsDb
end sub

'***********************************************
sub vbsDbDrawInputScreenDrawTitle( objVbsDb )
'***********************************************
%>
  <center>
  	<div id="InputTitle">
  	  <%=objVbsDb( "ScreenTitle" )%>
  	</div>
  </center>
<%
end sub

'***********************************************
function funStrInputScreenFormBegin( objVbsDb )
'***********************************************
  funStrInputScreenFormBegin = _
    "<form name='vbsDbForm_" & objVbsDb( "GlobalId" ) & "' action='" & _
    funStrVbsDbGlobalQuerystringPreserve( request.serverVariables( "url" ) , objVbsDb ) & _
    "' method='post'>"
end function

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyTableBegin( objVbsDb )
'***********************************************
  vbsDbBeginDrawForBorders "InputBorders"
%>
  <table <%=objVbsDb( "GridTableTag" )%> id="InputForm">
<%
end sub

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyTableHeadersRow( byRef objVbsDb )
'***********************************************
%>
  <th class="InputHeaders">
  	<%=objVbsDb( "InputFormLeftHeader" )%>
  </th>
  <th class="InputHeaders">
  	<%=objVbsDb( "InputFormRightHeader" )%>
  </th>	
<%
end sub

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyTableFieldsRowHeader( strEditFieldName , objVbsDb )
'***********************************************
%>
  <td class="InputLeftColumn">
  	<%=funStrFieldHeader( objVbsDb( "DictionaryInputFields" )( strEditFieldName ) , objVbsDb )%>
  </td>
<%
end sub

'***********************************************
function funLngFieldDefinedSize( strFieldName , objVbsDb )
'***********************************************
  dim lngFieldDefinedSize
  if VbsDbGetScreenType( objVbsDb ) = "Search" then
    funLngFieldDefinedSize = funVarSqlRecordSetFieldDefinedSize( strFieldName , objVbsDb )
  else
    ' the screen is an add or an update or a delete
    on error resume next
    lngFieldDefinedSize = funLngAdoxFieldDefinedSize( strFieldName , objVbsDb )
    'lngFieldDefinedSize = objVbsDb( "EditTableRecordSet" ).fields( strFieldName ).definedSize
    if err.number <> 0 then
      ' user has written a bad edit field name
      drawError strFieldName & " is not a field in the edit table" , objVbsDb
    else
      ' the edit field name is correct
      if funBolAdoxEditFieldIsString( strFieldName , objVbsDb ) then
        ' the edit field is a string
        funLngFieldDefinedSize = lngFieldDefinedSize
      else
        ' the edit field is not of string type
        funLngFieldDefinedSize = constDefaultMaxLenght
      end if
    end if
  end if
end function

'***********************************************
function funBolInputFieldIsReadOnly( strFieldName , objVbsDb )
'***********************************************
  if VbsDbGetScreenType( objVbsDb ) = "Search" then
    ' we are in a search screen
    funBolInputFieldIsReadOnly = false
  else
    ' we are in an edit screen
    funBolInputFieldIsReadOnly = funBolEditFieldIsReadOnly( strFieldName , objVbsDb )
  end if
end function

'***********************************************
function funStrVbsDbDrawInputElementDrawEditTextFieldType( strFieldName , objVbsDb )
'***********************************************
  if funBolFieldIsOfPasswordType( strFieldName , objVbsDb ) then
    funStrVbsDbDrawInputElementDrawEditTextFieldType = "password"
  else
    funStrVbsDbDrawInputElementDrawEditTextFieldType = "text"
  end if
end function

'***********************************************
function funStrVbsDbDrawInputElementDrawEditTextField( strFieldName , objVbsDb )
'***********************************************
  if not funBolInputFieldIsReadOnly( strFieldName , objVbsDb ) then
    ' the field is not set to read only
    funStrVbsDbDrawInputElementDrawEditTextField = _
      "<input name='" & funStrFormFieldName( strFieldName , objVbsDb ) & _
      "' type='" & funStrVbsDbDrawInputElementDrawEditTextFieldType( strFieldName , objVbsDb ) & "' " & _
      "maxlength=" & funLngFieldDefinedSize( strFieldName , objVbsDb ) & _
      " size='25' value=""" & funFieldDefaultValue( strFieldName , objVbsDb ) & """>"
  else
    ' the field is set to read only
    funStrVbsDbDrawInputElementDrawEditTextField = funFieldDefaultValue( strFieldName , objVbsDb )
  end if
end function

'***********************************************
function funStrVbsDbDrawInputElement( strEditFieldName , objVbsDb )
'***********************************************
  if funBolIsSelectField( strEditFieldName , objVbsDb ) then
  	' the field is a select
  	funStrVbsDbDrawInputElement = funStrVbsDbDrawInputSelectField( strEditFieldName , objVbsDb )
  elseif funBolIsInputEnumeratedField( strEditFieldName , objVbsDb ) then
  	' the field is a select
  	funStrVbsDbDrawInputElement = funStrVbsDbDrawInputEnumeratedField( strEditFieldName , objVbsDb )
  elseif ( funBolFieldIsMemo( strEditFieldName , objVbsDb ) and _
         ( VbsDbGetScreenType( objVbsDb ) <> "Search" ) ) then
  	' the field is a memo
  	funStrVbsDbDrawInputElement = funStrVbsDbDrawEditMemoField( strEditFieldName , objVbsDb )
  elseif funBolInputFieldIsBoolean( strEditFieldName , objVbsDb ) then
    ' the field is boolean
  	funStrVbsDbDrawInputElement = funStrVbsDbDrawBooleanInputField( strEditFieldName , objVbsDb )
  	'vbsDbDrawBooleanInputField strEditFieldName , objVbsDb
  else
  	' the field is not a select, it is not a memo field and it is not a boolean
  	funStrVbsDbDrawInputElement = _
  	  funStrVbsDbDrawInputElementDrawEditTextField( strEditFieldName , objVbsDb )
  end if
end function

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyTableFieldsRowData( strEditFieldName , objVbsDb )
'***********************************************
%>
  <td class="InputRightColumn">
    <%=funStrVbsDbDrawInputElement( strEditFieldName , objVbsDb )%>
  </td>
<%
end sub

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyTableFieldsRow( strEditFieldName , objVbsDb )
'***********************************************
%>
  <tr>
<%
  vbsDbDrawInputScreenDrawFormBodyTableFieldsRowHeader strEditFieldName , objVbsDb
  vbsDbDrawInputScreenDrawFormBodyTableFieldsRowData strEditFieldName , objVbsDb
%>
  </tr>
<%
end sub

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyTableFieldsRows( byRef objVbsDb )
'***********************************************
  dim strEditFieldName
  for each strEditFieldName in objVbsDb( "DictionaryInputFields" )
  	vbsDbDrawInputScreenDrawFormBodyTableFieldsRow strEditFieldName , objVbsDb
  next
end sub

'***********************************************
function funStrEditSubmitValue( byRef objVbsDb )
'***********************************************
  select case VbsDbGetScreenType( objVbsDb )
  	case "Search"
  	  funStrEditSubmitValue = "VBSdbApplyFilter"
  	case "Add"
  	  funStrEditSubmitValue = "VBSdbApplyAdd"
  	case "Update"
  	  funStrEditSubmitValue = "VBSdbApplyUpdate"
  	case "Delete"
  	  funStrEditSubmitValue = "VBSdbApplyDelete"
  end select 
end function

'***********************************************
function funStrHtmlForCancelButton( objVbsDb )
'***********************************************
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    funStrHtmlForCancelButton = objVbsDb( "InputCancelContent" )
  else
    ' the developer wants to use images for buttons
    funStrHtmlForCancelButton = "<img border='0' VALIGN='middle' SRC='" & objVbsDb( "GlobalImageDir" ) & "cancel.gif" & "'>"
  end if
end function

'***********************************************
function funStrVbsDbDrawInputCancel( objVbsDb )
'***********************************************
  dim strHref
  strHref = request.serverVariables( "url" ) & "?VBSdbClickClass_" & objVbsDb( "GlobalId" ) & "=VBSdbCancel&VBSdbIndex_" & _
    objVbsDb( "GlobalId" ) & "=1"
  funStrVbsDbDrawInputCancel = "<a title='" & objVbsDb( "InputCancelTitle" ) & _
    "' href='" & funStrVbsDbGlobalQuerystringPreserve( strHref , objVbsDb ) & "' class='InputCancel'>" & _
    funStrHtmlForCancelButton( objVbsDb ) & "</a>"
end function

'***********************************************
function funStrVbsDbDrawInputSubmit( objVbsDb )
'***********************************************
  if ( funBolJavascriptValidation( objVbsDb ) ) then
    ' the form must be client side validated before submit
    funStrVbsDbDrawInputSubmit = "<input class='submit' type='button' name='submit_dummy' value='" & _
      objVbsDb( "InputSubmitValue" ) & "' onclick='EditValidate(form)'>"
  else
    ' the form must be submitted without client side validation
    funStrVbsDbDrawInputSubmit = "<input type='submit' name='submit_dummy' value='" & _
      objVbsDb( "InputSubmitValue" ) & "'>"
  end if
end function

'***********************************************
'***********************************************
function funStrVbsDbDrawInputReset( objVbsDb )
'***********************************************
'***********************************************
  funStrVbsDbDrawInputReset = "<input type='reset' value='" & _
    objVbsDb( "InputResetValue" ) & "' id=reset1 name=reset1>"
end function

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyTableSubmitRowDrawSubmitAndReset( objVbsDb )
'***********************************************
  response.write funStrVbsDbDrawInputSubmit( objVbsDb ) & "<br>" & chr( 10 )
  'vbsDbDrawInputScreenDrawFormBodyTableSubmitRowDrawSubmit objVbsDb
  if VbsDbGetScreenType( objVbsDb ) <> "Delete" then
    ' it is an add screen or an update screen
    response.write funStrVbsDbDrawInputReset( objVbsDb ) & chr( 10 )
  end if
end sub

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyTableSubmitRow( byRef objVbsDb )
'***********************************************
%>
  <td valign="middle" class="InputCancel">
    <%=funStrVbsDbDrawInputCancel( objVbsDb )%>
  </td>
  <td valign="middle" id="InputSubmitReset">
<%
    vbsDbDrawInputScreenDrawFormBodyTableSubmitRowDrawSubmitAndReset objVbsDb
%>
  </td>
<%
end sub

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyTableBody( byRef objVbsDb )
'***********************************************
%>
  <tr>
<%
	  vbsDbDrawInputScreenDrawFormBodyTableHeadersRow objVbsDb
%>
	</tr>
<%
	  vbsDbDrawInputScreenDrawFormBodyTableFieldsRows objVbsDb
%>
	<tr align="center">
<%
	  vbsDbDrawInputScreenDrawFormBodyTableSubmitRow objVbsDb
%>
	</tr>
<%
end sub

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyTableEnd()
'***********************************************
%>
	</table>
<%
  vbsDbEndDrawForBorders
end sub

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyTable( objVbsDb )
'***********************************************
%>
  <p>
<%
  vbsDbDrawInputScreenDrawFormBodyTableBegin objVbsDb
  vbsDbDrawInputScreenDrawFormBodyTableBody objVbsDb
  vbsDbDrawInputScreenDrawFormBodyTableEnd
end sub

'***********************************************
function funStrVbsDbHiddenFieldTypesForField( strInputFieldName , objVbsDb )
'***********************************************
  funStrVbsDbHiddenFieldTypesForField = _
    "<input type='hidden' name='" & funStrVbsDbHiddenFieldTypeName( strInputFieldName , objVbsDb ) & "' " & _
    "value='" & funVarSqlRecordSetFieldType( strInputFieldName , objVbsDb ) & "'>"
end function

'***********************************************
function funStrVbsDbHiddenFieldTypesActually( objVbsDb )
'***********************************************
  dim strInputFieldName
  funStrVbsDbHiddenFieldTypesActually = ""
  for each strInputFieldName in objVbsDb( "DictionaryInputFields" )
  	funStrVbsDbHiddenFieldTypesActually = funStrVbsDbHiddenFieldTypesActually & _
  	  funStrVbsDbHiddenFieldTypesForField( strInputFieldName , objVbsDb )
  next
end function

'***********************************************
function funStrVbsDbHiddenFieldTypes( objVbsDb )
'***********************************************
  if VbsDbGetScreenType( objVbsDb ) = "Search" then
    funStrVbsDbHiddenFieldTypes = funStrVbsDbHiddenFieldTypesActually( objVbsDb )
  else
    funStrVbsDbHiddenFieldTypes = ""
  end if
end function

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyOtherHiddenFields( byRef objVbsDb )
'***********************************************
%>
  <input type="hidden" name="VBSdbClickClass_<%=objVbsDb( "GlobalId" )%>" value="<%=funStrEditSubmitValue( objVbsDb )%>">
  <input type="hidden" name="VBSdbEditWhere_<%=objVbsDb( "GlobalId" )%>" value="<%=objVbsDb( "RequestVBSdbEditWhere" )%>">
<%
end sub

'***********************************************
function funStrVbsDbOtherHiddenFields( objVbsDb )
'***********************************************
  dim strEditWhereClause
  strEditWhereClause = replace( objVbsDb( "RequestVBSdbEditWhere" ) , """" , constStrReplacementForQuotesInEditWhereConditions )
  funStrVbsDbOtherHiddenFields = _
    "<input type='hidden' name='VBSdbClickClass_" & objVbsDb( "GlobalId" ) & "' value=""" & _
    funStrEditSubmitValue( objVbsDb ) & """>" & chr( 10 ) & _
    "<input type='hidden' name='VBSdbEditWhere_" & objVbsDb( "GlobalId" ) & "' value=""" & _
    strEditWhereClause & """>"
end function

'***********************************************
function funStrVbsDbHiddenFields( objVbsDb )
'***********************************************
  funStrVbsDbHiddenFields = funStrVbsDbHiddenFieldTypes( objVbsDb ) & funStrVbsDbOtherHiddenFields( objVbsDb )
end function

'***********************************************
sub vbsDbDrawInputScreenDrawFormBodyHiddenFields( byRef objVbsDb )
'***********************************************
  response.write funStrVbsDbHiddenFields( objVbsDb )
  'response.write funStrVbsDbHiddenFieldTypes( objVbsDb )
  'vbsDbDrawInputScreenDrawFormBodyOtherHiddenFields objVbsDb
end sub

'***********************************************
sub vbsDbDrawInputScreenDrawFormBody( byRef objVbsDb )
'***********************************************
  vbsDbDrawInputScreenDrawFormBodyTable objVbsDb
  vbsDbDrawInputScreenDrawFormBodyHiddenFields objVbsDb
end sub

'***********************************************
sub vbsDbDrawInputScreenDrawForm( byRef objVbsDb )
'***********************************************
%>
  <center>
<%
    response.write funStrInputScreenFormBegin( objVbsDb )
	  vbsDbDrawInputScreenDrawFormBody objVbsDb
%>
  	</form>
  </center>
<%
end sub

'***********************************************
'***********************************************
sub vbsDbDrawStandardInputScreen( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbDrawInputScreenDrawTitle objVbsDb
  vbsDbDrawInputScreenDrawForm objVbsDb
end sub

'***********************************************
'***********************************************
sub vbsDbDrawInputScreen( byRef objVbsDb )
'***********************************************
'***********************************************
  drawJavaScriptFunctions objVbsDb
  if funBolInputTemplate( objVbsDb ) then
    ' the user wants to display a custom input screen
    vbsDbDrawTemplateInputScreen objVbsDb
  else
    ' the user wants to display the standard input screen
    vbsDbDrawStandardInputScreen objVbsDb
  end if
end sub
%>