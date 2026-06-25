<%
'***********************************************
'***********************************************
function funBolIsInputEnumeratedField( strFieldName , objVbsDb )
'***********************************************
'***********************************************
	funBolIsInputEnumeratedField = _
	  ( ( isObject( objVbsDb( "DictionaryInputEnumeratedFields" )( strFieldName ) ) ) or _
	    ( ( inStr( "**Add**Update**Delete" , VbsDbGetScreenType( objVbsDb ) ) > 0 ) and _
	      ( isObject( objVbsDb( "DictionaryInputEnumeratedFields" )( objVbsDb( "EditTableName" ) & "." & strFieldName ) ) or _
	        isObject( objVbsDb( "DictionaryInputEnumeratedFields" )( funStrGlobalDbTypeEditTableName( objVbsDb ) & "." & strFieldName ) ) ) ) )
end function

'***********************************************
sub vbsDbSetDictionaryInputEnumeratedFieldsAddInputEnumeratedFieldObjectAddOptionItem( intArrayIndex , arrStrList , byRef objVbsDb )
'***********************************************
  set objVbsDb( "DictionaryInputEnumeratedFields" )( arrStrList( 0 ) )( "option_" & intArrayIndex ) = _
    server.createObject( "scripting.dictionary" )
  objVbsDb( "DictionaryInputEnumeratedFields" )( arrStrList( 0 ) )( "option_" & intArrayIndex ).add _
    "value" , arrStrList( ( intArrayIndex  - 1 ) * 2 + 1 )
  objVbsDb( "DictionaryInputEnumeratedFields" )( arrStrList( 0 ) )( "option_" & intArrayIndex ).add _
    "content" , arrStrList(intArrayIndex * 2 )
end sub

'***********************************************
sub vbsDbSetTestedDictionaryInputEnumeratedFieldsAddInputEnumeratedFieldObjectActually( arrStrList , byRef objVbsDb )
'***********************************************
  dim intArrayIndex
  set objVbsDb( "DictionaryInputEnumeratedFields" )( arrStrList( 0 ) ) = server.createObject( "scripting.dictionary" )
  objVbsDb( "DictionaryInputEnumeratedFields" )( arrStrList( 0 ) ).compareMode = 1
  for intArrayIndex = 1 to uBound( arrStrList ) / 2
    vbsDbSetDictionaryInputEnumeratedFieldsAddInputEnumeratedFieldObjectAddOptionItem _
      intArrayIndex , arrStrList , objVbsDb
  next
end sub

'***********************************************
sub vbsDbSetDictionaryInputEnumeratedFieldsAddInputEnumeratedFieldObjectActually( strInputEnumeratedFieldItem , byRef objVbsDb )
'***********************************************
  dim arrStrList
  arrStrList = split( strInputEnumeratedFieldItem , "|" )
  if objVbsDb( "DictionaryInputEnumeratedFields" ).exists( arrStrList( 0 ) ) then
    drawError "InputEnumeratedFields bad assignment.<br><br>The field " & arrStrList( 0 ) & _
              " has been assigned twice." , objVbsDb
  elseif funBolIsOdd( uBound( arrStrList ) ) then
    drawError "InputEnumeratedFields bad assignment.<br><br>The field " & arrStrList( 0 ) & _
              " has been assigned with an odd number of parameters." , objVbsDb
  else
    vbsDbSetTestedDictionaryInputEnumeratedFieldsAddInputEnumeratedFieldObjectActually _
      arrStrList , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbSetDictionaryInputEnumeratedFieldsAddInputEnumeratedFieldObject( strInputEnumeratedFieldItem , byRef objVbsDb )
'***********************************************
  if trim( strInputEnumeratedFieldItem ) <> "" then
    ' this is not a (nosense) empty item
    vbsDbSetDictionaryInputEnumeratedFieldsAddInputEnumeratedFieldObjectActually _
      strInputEnumeratedFieldItem , objVbsDb
  end if
end sub

'***********************************************
'***********************************************
sub vbsDbSetDictionaryInputEnumeratedFields( byRef objVbsDb )
'***********************************************
'***********************************************
  dim arrStrList , intArrIndex
  set objVbsDb( "DictionaryInputEnumeratedFields" ) = server.createObject( "scripting.dictionary" )
  objVbsDb( "DictionaryInputEnumeratedFields" ).compareMode = 1
  arrStrList = split( objVbsDb( "InputEnumeratedFields" ) , ";" )
  for intArrIndex = 0 to uBound( arrStrList )
    vbsDbSetDictionaryInputEnumeratedFieldsAddInputEnumeratedFieldObject arrStrList( intArrIndex ) , objVbsDb
  next
end sub

'***********************************************
function funStrVbsDbDrawInputEnumeratedFieldDrawBeginSelect( strFieldName , objVbsDb )
'***********************************************
  if not funBolInputFieldIsReadOnly( strFieldName , objVbsDb ) then
    ' the field is not set to read only
    funStrVbsDbDrawInputEnumeratedFieldDrawBeginSelect = _
      "<select name='" & funStrFormFieldName( strFieldName , objVbsDb ) & "'>"
  else
    ' the field is set to read only
    funStrVbsDbDrawInputEnumeratedFieldDrawBeginSelect = ""
  end if
end function

'***********************************************
sub vbsDbDrawInputEnumeratedFieldDrawOptionsDrawBlankOption( strInputFieldName , objVbsDb )
'***********************************************
  if not funBolInputFieldIsReadOnly( strInputFieldName , objVbsDb ) then
    ' the field is not set to read only
%>
    <option value=''></option>
<%
  end if
end sub

'***********************************************
function funStrInputEnumeratedFieldNameForEditField( strFieldName , objVbsDb )
'***********************************************
	if isObject( objVbsDb( "DictionaryInputEnumeratedFields" )( strFieldName ) ) then
	  funStrInputEnumeratedFieldNameForEditField = strFieldName
	elseif isObject( objVbsDb( "DictionaryInputEnumeratedFields" )( objVbsDb( "EditTableName" ) & "." & strFieldName ) ) then
	  funStrInputEnumeratedFieldNameForEditField = objVbsDb( "EditTableName" ) & "." & strFieldName
	else
	  ' it must be isObject( objVbsDb( "DictionaryInputEnumeratedFields" )( funStrGlobalDbTypeEditTableName( objVbsDb ) & "." & strFieldName ) )
	  funStrInputEnumeratedFieldNameForEditField = funStrGlobalDbTypeEditTableName( objVbsDb ) & "." & strFieldName
	end if
end function

'***********************************************
function funBolInputEnumeratedFieldIsOptionSelected( strInputFieldName , objInputEnumeratedFieldOption , objVbsDb )
'***********************************************
  funBolInputEnumeratedFieldIsOptionSelected = _
    ( cStr( funFieldDefaultValue( strInputFieldName , objVbsDb ) ) = _
      cStr( objInputEnumeratedFieldOption( "value" ) ) )
end function

'***********************************************
function funStrInputEnumeratedFieldIsOptionSelected( strInputFieldName , objInputEnumeratedFieldOption , objVbsDb )
'***********************************************
	if ( funBolInputEnumeratedFieldIsOptionSelected( strInputFieldName , objInputEnumeratedFieldOption , objVbsDb ) ) then
	  ' the current option element refers to the field value actually present in the record to be updated
	  funStrInputEnumeratedFieldIsOptionSelected = "selected"
	else
	  ' the current option element does not refer to the field value actually present in the record to be updated
	  funStrInputEnumeratedFieldIsOptionSelected = ""
  end if
end function

'***********************************************
function funStrVbsDbDrawInputEnumeratedFieldDrawOptionDrawBeginOption( strInputFieldName , objInputEnumeratedFieldOption , objVbsDb )
'***********************************************
  funStrVbsDbDrawInputEnumeratedFieldDrawOptionDrawBeginOption = _
    "<option " & funStrInputEnumeratedFieldIsOptionSelected( strInputFieldName , objInputEnumeratedFieldOption , objVbsDb ) & _
	  " value='" & objInputEnumeratedFieldOption( "value" ) & "'>"
end function

'***********************************************
function funStrVbsDbDrawInputEnumeratedFieldDrawOptionDrawOptionContent( strInputFieldName , objInputEnumeratedFieldOption , objVbsDb )
'***********************************************
	funStrVbsDbDrawInputEnumeratedFieldDrawOptionDrawOptionContent = objInputEnumeratedFieldOption( "content" )
end function

'***********************************************
function funStrVbsDbDrawInputEnumeratedFieldDrawOptionDrawEndOption()
'***********************************************
  funStrVbsDbDrawInputEnumeratedFieldDrawOptionDrawEndOption = "</option>" & chr( 10 )
end function

'***********************************************
function funStrVbsDbDrawInputEnumeratedFieldDrawOption( strInputFieldName , objInputEnumeratedFieldOption , objVbsDb )
'***********************************************
  if not funBolInputFieldIsReadOnly( strInputFieldName , objVbsDb ) then
    ' the field is not set to read only
    funStrVbsDbDrawInputEnumeratedFieldDrawOption = _
  	  funStrVbsDbDrawInputEnumeratedFieldDrawOptionDrawBeginOption( strInputFieldName , _
  	    objInputEnumeratedFieldOption , objVbsDb ) & _
  	  funStrVbsDbDrawInputEnumeratedFieldDrawOptionDrawOptionContent( strInputFieldName , _
  	    objInputEnumeratedFieldOption , objVbsDb ) & _
  	  funStrVbsDbDrawInputEnumeratedFieldDrawOptionDrawEndOption()
  else
    ' the field is set to read only
    if funBolInputEnumeratedFieldIsOptionSelected( strInputFieldName , _
      objInputEnumeratedFieldOption , objVbsDb ) then
      ' the current option element refers to the field value actually present in the record to be updated
  	  funStrVbsDbDrawInputEnumeratedFieldDrawOption = _
  	    funStrVbsDbDrawInputEnumeratedFieldDrawOptionDrawOptionContent( strInputFieldName , _
  	      objInputEnumeratedFieldOption , objVbsDb )
    end if
  end if
end function

'***********************************************
function funStrVbsDbDrawInputEnumeratedFieldDrawOptions( strInputFieldName , objVbsDb )
'***********************************************
  dim strEnumeratedOption , strInputEnumeratedFieldNameForEditField
  strInputEnumeratedFieldNameForEditField = _
    funStrInputEnumeratedFieldNameForEditField( strInputFieldName , objVbsDb )
  funStrVbsDbDrawInputEnumeratedFieldDrawOptions = ""
  for each strEnumeratedOption in objVbsDb( "DictionaryInputEnumeratedFields" )( strInputEnumeratedFieldNameForEditField )
    funStrVbsDbDrawInputEnumeratedFieldDrawOptions = funStrVbsDbDrawInputEnumeratedFieldDrawOptions & _
      funStrVbsDbDrawInputEnumeratedFieldDrawOption( strInputFieldName , _
        objVbsDb( "DictionaryInputEnumeratedFields" )( strInputEnumeratedFieldNameForEditField )( strEnumeratedOption ) , objVbsDb )
  next
end function

'***********************************************
function funStrVbsDbDrawInputEnumeratedFieldDrawEndSelect( strInputFieldName , objVbsDb )
'***********************************************
  if not funBolInputFieldIsReadOnly( strInputFieldName , objVbsDb ) then
    ' the field is not set to read only
    funStrVbsDbDrawInputEnumeratedFieldDrawEndSelect = "</select>" & chr( 10 )
  else
    ' the field is set to read only
    funStrVbsDbDrawInputEnumeratedFieldDrawEndSelect = ""
  end if
end function

'***********************************************
'***********************************************
function funStrVbsDbDrawInputEnumeratedField( strInputFieldName , objVbsDb )
'***********************************************
'***********************************************
	funStrVbsDbDrawInputEnumeratedField = _
	  funStrVbsDbDrawInputEnumeratedFieldDrawBeginSelect( strInputFieldName , objVbsDb )
	funStrVbsDbDrawInputEnumeratedField = funStrVbsDbDrawInputEnumeratedField & _
	  funStrVbsDbDrawInputEnumeratedFieldDrawOptions( strInputFieldName , objVbsDb )
	funStrVbsDbDrawInputEnumeratedField = funStrVbsDbDrawInputEnumeratedField & _
	  funStrVbsDbDrawInputEnumeratedFieldDrawEndSelect( strInputFieldName , objVbsDb )
end function
%>