<%
'***********************************************
sub drawJavaScriptFunctionForIntMaxEditCharsForMemoFieldsDrawActually()
'***********************************************
%>
  <script language="javascript">
  	function textCounter(field, countfield, maxlimit)
  	{
  	  if (field.value.length > maxlimit)
  	  	// if too long...trim it!
  	  	field.value = field.value.substring(0, maxlimit);
  	  else
        // otherwise, update 'characters left' counter
  	  {
  	  	countfield.value = maxlimit - field.value.length;
  	  }
  	}
  	// End -->
  </script>
<%
end sub

'***********************************************
'***********************************************
sub drawJavaScriptFunctionForIntMaxEditCharsForMemoFields( objVbsDb )
'***********************************************
'***********************************************
  if ( ( VbsDbGetScreenType( objVbsDb ) = "Add" ) or ( VbsDbGetScreenType( objVbsDb ) = "Update" ) ) and _
     objVbsDb( "EditMemoFieldMaxCharExists" ) then
    ' there is at least a field to be edited, with a maximum number of char to be edited
    drawJavaScriptFunctionForIntMaxEditCharsForMemoFieldsDrawActually
  end if
end sub

'***********************************************
sub vbsDbNormalizeEditMemoFieldsToSemiColons( byRef objVbsDb )
'***********************************************
  if inStr( objVbsDb( "EditMemoFields" ) , "|" ) then
    ' probably, the user, mistakely, used "," instead of ";"
    objVbsDb( "EditMemoFields" ) = replace( objVbsDb( "EditMemoFields" ) , "," , ";" )
  end if
end sub

'***********************************************
sub vbsDbNormalizeEditMemoFieldsToPipes( byRef objVbsDb )
'***********************************************
  if inStr( objVbsDb( "EditMemoFields" ) , ";" ) then
    ' probably, the user, mistakely, used "," instead of "|"
    objVbsDb( "EditMemoFields" ) = replace( objVbsDb( "EditMemoFields" ) , "," , "|" )
  end if
end sub

'***********************************************
'***********************************************
sub vbsDbNormalizeEditMemoFields( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbNormalizeEditMemoFieldsToSemiColons objVbsDb
  vbsDbNormalizeEditMemoFieldsToPipes objVbsDb
end sub

'***********************************************
sub vbsDbSetDictionaryObjectForEditMemoFieldsAddMemoObjectCreateObject( strMemoItemName , intMemoItemEditModeNumRows , intMemoItemEditModeNumCols , intMemoItemEditModeMaxChars , byRef objMemoItem )
'***********************************************
  set objMemoItem = server.createObject( "scripting.dictionary" )
  objMemoItem.CompareMode = 1
	objMemoItem( "strName" ) = strMemoItemName
	objMemoItem( "intEditModeNumRows" ) = intMemoItemEditModeNumRows
	objMemoItem( "intEditModeNumCols" ) = intMemoItemEditModeNumCols
	objMemoItem( "intEditModeMaxChars" ) = intMemoItemEditModeMaxChars
end sub

'***********************************************
sub vbsDbSetDictionaryObjectForEditMemoFieldsAddMemoObjectGetParameter( intNumParameter , arrMemoItem , byRef strMemoItemParameter )
'***********************************************
  if uBound( arrMemoItem ) >= intNumParameter then
    ' the user, for this memo field, defined at least intNumParameter parameters
    if trim( arrMemoItem( intNumParameter ) ) <> "" then
      ' the user, for this memo field, defined a non empty intNumParameter parameter
      strMemoItemParameter = arrMemoItem( intNumParameter )
    end if
  end if
end sub

'***********************************************
sub vbsDbSetDictionaryObjectForEditMemoFieldsAddMemoObjectGetParameters( strMemoItem , byRef strMemoItemName , byRef intMemoItemEditModeNumRows , byRef intMemoItemEditModeNumCols , byRef intMemoItemEditModeMaxChars )
'***********************************************
  dim arrMemoItem
  intMemoItemEditModeNumRows = constIntDefaultMemoItemEditModeNumRows
  intMemoItemEditModeNumCols = constIntDefaultMemoItemEditModeNumCols
  intMemoItemEditModeMaxChars = 0
  arrMemoItem = split( strMemoItem , "|" )
  strMemoItemName = arrMemoItem( 0 )
  vbsDbSetDictionaryObjectForEditMemoFieldsAddMemoObjectGetParameter 1 , arrMemoItem , intMemoItemEditModeNumRows
  vbsDbSetDictionaryObjectForEditMemoFieldsAddMemoObjectGetParameter 2 , arrMemoItem , intMemoItemEditModeNumCols
  vbsDbSetDictionaryObjectForEditMemoFieldsAddMemoObjectGetParameter 3 , arrMemoItem , intMemoItemEditModeMaxChars
end sub

'***********************************************
sub vbsDbCheckErrorRepeatedNameMemoFieldForDictionary( strMemoItemName , byRef objVbsDb )
'***********************************************
  if objVbsDB( "DictionaryEditMemoFields" ).exists( strMemoItemName ) then
    drawError "EditMemoFields bad assignment.<br><br>The field " & strMemoItemName & _
              " has been assigned twice." , objVbsDb
  end if
end sub

'***********************************************
sub setEditMemoFieldsEditModeMaxChar( objMemoItem , byRef objVbsDb )
'***********************************************
  if isNumeric( objMemoItem( "intEditModeMaxChars" ) ) then
    ' the current field is a memo field with a maximum number of characters to be edited
     objVbsDb( "EditMemoFieldMaxCharExists" ) = ( objMemoItem( "intEditModeMaxChars" ) > 0 )
  end if
end sub

'***********************************************
sub vbsDbSetDictionaryObjectForEditMemoFieldsAddMemoObjectActually( strMemoItem , byRef objVbsDb )
'***********************************************
  dim strMemoItemName , intMemoItemViewModeMaxChars , intMemoItemEditModeMaxChars , _
      intMemoItemEditModeNumRows , intMemoItemEditModeNumCols , objMemoItem
  vbsDbSetDictionaryObjectForEditMemoFieldsAddMemoObjectGetParameters strMemoItem , strMemoItemName , intMemoItemEditModeNumRows , intMemoItemEditModeNumCols , intMemoItemEditModeMaxChars
  vbsDbSetDictionaryObjectForEditMemoFieldsAddMemoObjectCreateObject strMemoItemName , intMemoItemEditModeNumRows , intMemoItemEditModeNumCols , intMemoItemEditModeMaxChars , objMemoItem
  vbsDbCheckErrorRepeatedNameMemoFieldForDictionary strMemoItemName , objVbsDb
  objVbsDb( "DictionaryEditMemoFields" ).add strMemoItemName , objMemoItem
  setEditMemoFieldsEditModeMaxChar objMemoItem , objVbsDb
end sub

'***********************************************
sub vbsDbSetDictionaryObjectForEditMemoFieldsAddMemoObject( strMemoItem , byRef objVbsDb )
'***********************************************
  if trim( strMemoItem ) <> "" then
    ' this is not a (nosense) empty item
    vbsDbSetDictionaryObjectForEditMemoFieldsAddMemoObjectActually strMemoItem , objVbsDb
  end if
end sub

'***********************************************
'***********************************************
sub vbsDbSetDictionaryObjectForEditMemoFields( byRef objVbsDb )
'***********************************************
'***********************************************
  dim arrStrList , intArrIndex
  set objVbsDb( "DictionaryEditMemoFields" ) = server.createObject( "scripting.dictionary" )
  objVbsDb( "DictionaryEditMemoFields" ).compareMode = 1
  arrStrList = split( objVbsDb( "EditMemoFields" ) , ";" )
  for intArrIndex = 0 to uBound( arrStrList )
    vbsDbSetDictionaryObjectForEditMemoFieldsAddMemoObject arrStrList( intArrIndex ) , objVbsDb
  next
end sub

'***********************************************
'***********************************************
function funBolFieldIsMemo( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  if isObject( objVbsDb( "DictionaryEditMemoFields" ) ) then
    ' some edit memo field was defined
    funBolFieldIsMemo = objVbsDb( "DictionaryEditMemoFields" ).exists( strFieldName )
  else
    ' no memo field was defined
	funBolFieldIsMemo = false
  end if
end function

'***********************************************
function funStrTextCounterEventAttributes( strEditFieldName , intEditModeMaxChars , objVbsDb )
'***********************************************
	funStrTextCounterEventAttributes = "textCounter(this.form." & _
	  funStrFormFieldName( strEditFieldName , objVbsDb ) & _
	  ",this.form." & funStrFormFieldName( strEditFieldName , objVbsDb ) & "_maxChars," & intEditModeMaxChars & ");"
end function

'***********************************************
function funStrDrawTextCounterEventAttributes( strEditFieldName , objVbsDb )
'***********************************************
  if objVbsDb( "DictionaryEditMemoFields" )( strEditFieldName )( "intEditModeMaxChars" ) > 0 then
	  ' the current field is a memo field and the web master wants to limit the number of characters in the corresponding text area
	  funStrDrawTextCounterEventAttributes = _
	    "onKeyDown='" & _
	    funStrTextCounterEventAttributes( strEditFieldName , _
	      objVbsDb( "DictionaryEditMemoFields" )( strEditFieldName )( "intEditModeMaxChars" ) , objVbsDb ) & _
	      "' onKeyUp='" & _
	    funStrTextCounterEventAttributes( strEditFieldName , _
	      objVbsDb( "DictionaryEditMemoFields" )( strEditFieldName )( "intEditModeMaxChars" ) , objVbsDb ) & _
	      "'"
  else
	  ' either the current field is not a memo field or the web master doesn't wanto to limit the number of characters in the corresponding text area
	  funStrDrawTextCounterEventAttributes = ""
  end if
end function

'***********************************************
function funStrDrawTextCounterControl( strEditFieldName , objVbsDb )
'***********************************************
  if objVbsDb( "DictionaryEditMemoFields" )( strEditFieldName )( "intEditModeMaxChars" ) > 0 then
	  ' the current field is a memo field and the web master wants to limit the number of characters in the corresponding text area
	  funStrDrawTextCounterControl = _
	    "<input readonly type=text name='" & funStrFormFieldName( strEditFieldName , objVbsDb ) & _
	     "_maxChars' size=3 maxlength=3 value='" & _
	    objVbsDb( "DictionaryEditMemoFields" )( strEditFieldName )( "intEditModeMaxChars" ) - len( funFieldDefaultValue( strEditFieldName , objVbsDb ) ) & _
	    "'><span class='InputRightColumn'>" & funStrTranslate( "editMemoCharactersLeft" , objVbsDb ) & "</span>"
  end if
end function

'***********************************************
function funStrVbsDbDrawEditMemoFieldDrawEditWithStrTextAreaContent( strEditFieldName , strTextAreaContent , objVbsDb )
'***********************************************
  funStrVbsDbDrawEditMemoFieldDrawEditWithStrTextAreaContent = _
    "<table cellPadding='0' cellSpacing='0'>" & chr( 10 ) & _
    "  <td>" & chr( 10 ) &  chr( 10 ) & _
    "    <textarea name='" & funStrFormFieldName( strEditFieldName , objVbsDb ) & "'" & chr( 10 ) & _
    "      " & funStrDrawTextCounterEventAttributes( strEditFieldName , objVbsDb ) &  chr( 10 ) & _
    "      rows=" & objVbsDb( "DictionaryEditMemoFields" )( strEditFieldName )( "intEditModeNumRows" ) &  chr( 10 ) & _
    "      cols=" & objVbsDb( "DictionaryEditMemoFields" )( strEditFieldName )( "intEditModeNumCols" ) & ">" & strTextAreaContent & "</textarea>" &  chr( 10 ) & _
    "  </td>" &  chr( 10 ) & _
    "  <tr>" &  chr( 10 ) & _
    "  <td>" &  chr( 10 ) & _
    "    " & funStrDrawTextCounterControl( strEditFieldName , objVbsDb ) &  chr( 10 ) & _
    "  </td>" &  chr( 10 ) & _
    "  </tr>" &  chr( 10 ) & _
    "</table>"
end function

'***********************************************
function funStrVbsDbDrawEditMemoFieldDrawEdit( strEditFieldName , objVbsDb )
'***********************************************
  dim strTextAreaContent
  'strTextAreaContent = funStrHtmlForCarriageReturn( funFieldDefaultValue( strEditFieldName , objVbsDb ) )
  strTextAreaContent = funFieldDefaultValue( strEditFieldName , objVbsDb )
  funStrVbsDbDrawEditMemoFieldDrawEdit = _
    funStrVbsDbDrawEditMemoFieldDrawEditWithStrTextAreaContent( strEditFieldName , strTextAreaContent , objVbsDb )
end function

'***********************************************
function funStrVbsDbDrawEditMemoFieldDrawReadOnlyField( strEditFieldName , objVbsDb )
'***********************************************
  funStrVbsDbDrawEditMemoFieldDrawReadOnlyField = _
    replace( funFieldDefaultValue( strEditFieldName , objVbsDb ) , chr( 10 ) , "<br>" ) 
end function

'***********************************************
'***********************************************
function funStrVbsDbDrawEditMemoField( strEditFieldName , objVbsDb )
'***********************************************
'***********************************************
  if not funBolEditFieldIsReadOnly( strEditFieldName , objVbsDb ) then
    ' the field is not set to read only
    funStrVbsDbDrawEditMemoField = funStrVbsDbDrawEditMemoFieldDrawEdit( strEditFieldName , objVbsDb )
  else
    ' the field is set to read only
    funStrVbsDbDrawEditMemoField = funStrVbsDbDrawEditMemoFieldDrawReadOnlyField( strEditFieldName , objVbsDb )
  end if
end function
%>