<%
'***********************************************
'***********************************************
function funBolSqlFieldIsBoolean( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  'funBolSqlFieldIsBoolean = ( objVbsDb( "SqlRecordSet" ).fields( strFieldName ).type = vbsDbAdBoolean )
  funBolSqlFieldIsBoolean = ( funVarSqlRecordSetFieldType( strFieldName , objVbsDb ) = vbsDbAdBoolean )
end function

'***********************************************
'***********************************************
function funBolInputFieldIsBoolean( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  if VbsDbGetScreenType( objVbsDb ) = "Search" then
    ' current screen is a search screen
    funBolInputFieldIsBoolean = ( funVarSqlRecordSetFieldType( strFieldName , objVbsDb ) = vbsDbAdBoolean )
  else
    ' current screen is an screen
    funBolInputFieldIsBoolean = funBolAdoxColumnIsBoolean( strFieldName , objVbsDb )
  end if
end function

'***********************************************
'***********************************************
function funStrDbEditBooleanFieldFormattedValue( strFieldValue , objVbsDb )
'***********************************************
'***********************************************
  funStrDbEditBooleanFieldFormattedValue = funStrGlobalDbTypeBooleanFieldFormattedValue( strFieldValue , objVbsDb )
end function

'***********************************************
function funStrVbsDbDrawBooleanSelectBegin( strFieldName , objVbsDb )
'***********************************************
  funStrVbsDbDrawBooleanSelectBegin = _
    "<select name='" & funStrFormFieldName( strFieldName , objVbsDb ) & "'>" & chr( 10 )
end function

'***********************************************
function funStVbsDbDrawBooleanSelectEmptyOption()
'***********************************************
  funStVbsDbDrawBooleanSelectEmptyOption = "<option></option>" & chr( 10 )
end function

'***********************************************
function funStrVbsDbDrawBooleanSelectTrueOption( strFieldName , objVbsDb )
'***********************************************
  funStrVbsDbDrawBooleanSelectTrueOption = _
    "<option value='true'>" & funStrTranslate( "booleanTrue" , objVbsDb ) & "</option>" & chr( 10 )
end function

'***********************************************
function funStrVbsDbDrawBooleanSelectFalseOption( strFieldName , objVbsDb )
'***********************************************
  funStrVbsDbDrawBooleanSelectFalseOption = _
    "<option value='false'>" & funStrTranslate( "booleanFalse" , objVbsDb ) & "</option>" & chr( 10 )
end function

'***********************************************
function funStrVbsDbDrawBooleanSelectOptions( strFieldName , objVbsDb )
'***********************************************
  funStrVbsDbDrawBooleanSelectOptions = _
    funStVbsDbDrawBooleanSelectEmptyOption() & _
    funStrVbsDbDrawBooleanSelectTrueOption( strFieldName , objVbsDb ) & _
    funStrVbsDbDrawBooleanSelectFalseOption( strFieldName , objVbsDb )
end function

'***********************************************
function funStrVbsDbDrawBooleanSelectEnd( strFieldName , objVbsDb )
'***********************************************
  funStrVbsDbDrawBooleanSelectEnd = "</select>"
end function

'***********************************************
function funStrVbsDbDrawBooleanSelect( strFieldName , objVbsDb )
'***********************************************
  funStrVbsDbDrawBooleanSelect = _
    funStrVbsDbDrawBooleanSelectBegin( strFieldName , objVbsDb ) & _
    funStrVbsDbDrawBooleanSelectOptions( strFieldName , objVbsDb ) & _
    funStrVbsDbDrawBooleanSelectEnd( strFieldName , objVbsDb )
end function

'***********************************************
function funBolEditFieldValueIsTrue( strFieldName , objVbsDb )
'***********************************************
  if ( VbsDbGetScreenType( objVbsDb ) = "Add" ) then
    funBolEditFieldValueIsTrue = _
      ( funFieldDefaultValue( strFieldName , objVbsDb ) = "true" )
  else
    ' current screen is an update screen
    funBolEditFieldValueIsTrue = _
      objVbsDb( "EditTableRecordSet" ).fields( strFieldName ).value
  end if
end function

'***********************************************
function funBolBrowserIsInternetExplorer()
'***********************************************
  funBolBrowserIsInternetExplorer = _
    ( instr( 1, request.serverVariables("HTTP_USER_AGENT") ,"MSIE" ) > 0 )
end function

'***********************************************
function funStrClassForInternetExplorer()
'***********************************************
  if funBolBrowserIsInternetExplorer() then
    ' the user is using an Internet Explorer browser
    funStrClassForInternetExplorer = "class=""InputCheckBoxForIE"""
  else
    ' the user is not using an Internet Explorer browser
    funStrClassForInternetExplorer = "class=""InputCheckBoxForNonIE"""
  end if
end function

'***********************************************
function funStrDrawBooleanCheckboxChecked( strFieldName , objVbsDb )
'***********************************************
  if funBolEditFieldValueIsTrue( strFieldName , objVbsDb ) then
    ' either we are in add screen and the default value is true
    ' or we are in update screen and the default value is true
    funStrDrawBooleanCheckboxChecked = "checked"
  else
    ' either we are in add screen and the default value is false
    ' or we are in update screen and the default value is false
    funStrDrawBooleanCheckboxChecked = ""
  end if
end function

'***********************************************
function funStrVbsDbDrawBooleanCheckbox( strFieldName , objVbsDb )
'***********************************************
  funStrVbsDbDrawBooleanCheckbox = _
    "<input " & funStrClassForInternetExplorer & " type='checkbox' " & _
    "name='" & funStrFormFieldName( strFieldName , objVbsDb ) & "' " & _
    funStrDrawBooleanCheckboxChecked( strFieldName , objVbsDb ) & ">"
end function

'***********************************************
function funStrVbsDbDrawBooleanInputNonReadOnlyField( strFieldName , objVbsDb )
'***********************************************
  if VbsDbGetScreenType( objVbsDb ) = "Search" then
    funStrVbsDbDrawBooleanInputNonReadOnlyField = funStrVbsDbDrawBooleanSelect( strFieldName , objVbsDb )
  else
    ' current screen is an add screen or an update screen
    funStrVbsDbDrawBooleanInputNonReadOnlyField = funStrVbsDbDrawBooleanCheckbox( strFieldName , objVbsDb )
  end if
end function

'***********************************************
function funStrReadOnlyBooleanEditAddField( strFieldName , objVbsDb )
'***********************************************
  select case funFieldDefaultValue( strFieldName , objVbsDb )
    case "true"
      ' current field is boolean, it is read only and it has to default to 'true'
      funStrReadOnlyBooleanEditAddField = objVbsDb( "GlobalTrueText" )
    case "false"
      ' current field is boolean, it is read only and it has to default to 'false'
      funStrReadOnlyBooleanEditAddField = objVbsDb( "GlobalFalseText" )
    case else
      ' current field is boolean, it is read only and it has no default value (it seems to be no sense, a readonly field with no default value, and probably it is...)
      funStrReadOnlyBooleanEditAddField = ""
  end select
end function

'***********************************************
function funStrReadOnlyBooleanEditUpdateOrDeleteField( strFieldName , objVbsDb )
'***********************************************
  if objVbsDb( "EditTableRecordSet" ).fields( strFieldName ) then
    ' current sql record set boolean field is a true valued
    funStrReadOnlyBooleanEditUpdateOrDeleteField = funStrTranslate( "booleanTrue" , objVbsDb )
  else
    ' current sql record set boolean field is a false valued
    funStrReadOnlyBooleanEditUpdateOrDeleteField = funStrTranslate( "booleanFalse" , objVbsDb )
  end if
end function

'***********************************************
function funStrReadOnlyBooleanEditField( strFieldName , objVbsDb )
'***********************************************
  if VbsDbGetScreenType( objVbsDb ) = "Add" then
    funStrReadOnlyBooleanEditField = funStrReadOnlyBooleanEditAddField( strFieldName , objVbsDb )
  else
    ' VbsDbGetScreenType( objVbsDb ) is either "Update" or "Delete"
    funStrReadOnlyBooleanEditField = funStrReadOnlyBooleanEditUpdateOrDeleteField( strFieldName , objVbsDb )
  end if
end function

'***********************************************
'***********************************************
function funStrVbsDbDrawBooleanInputField( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  if not funBolInputFieldIsReadOnly( strFieldName , objVbsDb ) then
    ' the field is not set to read only
    'vbsDbDrawBooleanSelect strFieldName , objVbsDb
    funStrVbsDbDrawBooleanInputField = funStrVbsDbDrawBooleanInputNonReadOnlyField( strFieldName , objVbsDb )
  else
    ' the field is set to read only
    funStrVbsDbDrawBooleanInputField = funStrReadOnlyBooleanEditField( strFieldName , objVbsDb )
  end if
end function

'***********************************************
'***********************************************
function funStrGlobalBooleanText( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  if funVarSqlRecordSetFieldValue( strFieldName , objVbsDb ) then
    ' current sql record set boolean field is a true valued
    funStrGlobalBooleanText = funStrTranslate( "booleanTrue" , objVbsDb )
  else
    ' current sql record set boolean field is a false valued
    funStrGlobalBooleanText = funStrTranslate( "booleanFalse" , objVbsDb )
  end if
end function
%>