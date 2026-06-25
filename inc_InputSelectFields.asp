<%
'***********************************************
'***********************************************
function funBolIsSelectField( strFieldName , objVbsDb )
'***********************************************
'***********************************************
	funBolIsSelectField = _
	  ( objVbsDb( "DictionaryInputSelectFields" )( strFieldName ) <> "" ) or _
	  ( objVbsDb( "DictionaryInputSelectFields" )( objVbsDb( "EditTableName" ) & "." & strFieldName ) <> "" ) or _
	  ( objVbsDb( "DictionaryInputSelectFields" )( funStrGlobalDbTypeEditTableName( objVbsDb ) & "." & strFieldName ) <> "" )
end function

'***********************************************
function strInputSelectFieldNameForEditField( strFieldName , objVbsDb )
'***********************************************
	if objVbsDb( "DictionaryInputSelectFields" )( strFieldName ) <> "" then
	  strInputSelectFieldNameForEditField = strFieldName
	elseif objVbsDb( "DictionaryInputSelectFields" )( objVbsDb( "EditTableName" ) & "." & strFieldName ) <> "" then
	  strInputSelectFieldNameForEditField = objVbsDb( "EditTableName" ) & "." & strFieldName
	else
	  ' it must be objVbsDb( "DictionaryInputSelectFields" )( funStrGlobalDbTypeEditTableName( objVbsDb ) & "." & strFieldName ) <> ""
	  strInputSelectFieldNameForEditField = funStrGlobalDbTypeEditTableName( objVbsDb ) & "." & strFieldName
	end if
end function

'***********************************************
function funStrVbsDbDrawInputSelectFieldDrawBeginSelect( strFieldName , objVbsDb )
'***********************************************
  if not funBolInputFieldIsReadOnly( strFieldName , objVbsDb ) then
    ' the field is not set to read only
    funStrVbsDbDrawInputSelectFieldDrawBeginSelect = _
      "<select name='" & funStrFormFieldName( strFieldName , objVbsDb ) & "'>" & chr( 10 )
  else
    funStrVbsDbDrawInputSelectFieldDrawBeginSelect = ""
  end if
end function

'***********************************************
function funStrVbsDbDrawInputSelectFieldDrawOptionsDrawBlankOption( strInputFieldName , objVbsDb )
'***********************************************
  if not funBolInputFieldIsReadOnly( strInputFieldName , objVbsDb ) then
    ' the field is not set to read only
    funStrVbsDbDrawInputSelectFieldDrawOptionsDrawBlankOption = _
      "  <option value=''></option>" & chr( 10 )
  else
    funStrVbsDbDrawInputSelectFieldDrawOptionsDrawBlankOption = ""
  end if
end function

'***********************************************
sub vbsDbDrawInputSelectFieldDrawOptionsGetObjRs( strInputFieldName , objVbsDb , byRef objRs )
'***********************************************
  dim strSql , strErrorMessage
  strSql = objVbsDb( "DictionaryInputSelectFields" )( strInputSelectFieldNameForEditField( strInputFieldName , objVbsDb ) )
  set objRs = Server.CreateObject("ADODB.Recordset")
  on error resume next
  objRs.Open strSql , objVbsDb( "Connection" ) , vbsDbAdOpenKeySet , vbsDbAdLockReadOnly , vbsDbAdCmdText
  if err.number<>0 then
    strErrorMessage = "InputSelectFields is assigned this wrong Sql select statement:<br><br>" & _
      strSql & "<br><br>" & _
      "assigned to this field: " & strInputFieldName & "<br><br>" & _ 
      "Returned error description is: '" & err.description & "'<br><br>" & _
      "Please, perform the following debug steps:" & _
      "<div align=left>" & _
      "<ul><li>control none of the tables involved in the query is exclusively opened for database managment</li>" & _
      "<li>try the Sql query displayed above directly against your database</li>" & _
      "<li>check the InputSelectFields property assignment, for the field " & strInputFieldName & "</li>" & _
      "</ul>" & _
      "</div>" & _
      "If none of the above solves, please, " & _
      "contact VBSdb support, providing " &_
      "them with this message and the debugging values below. Thanks."
    'drawErrorWithDebugging "InputSelectFields is assigned a wrong select statement: " & strSql  , objVbsDb
    drawErrorWithDebugging strErrorMessage  , objVbsDb
  end if
end sub

'***********************************************
function funBolSelectFieldIsOptionSelected( strFieldName , objVbsDb , objRs )
'***********************************************
  funBolSelectFieldIsOptionSelected = ( cStr( funFieldDefaultValue( strFieldName , objVbsDb ) ) = cStr( objRs.fields.item( 0 ).value ) )
end function

'***********************************************
function funStrSelectFieldIsOptionSelected( strInputFieldName , objVbsDb , objRs )
'***********************************************
	if ( funBolSelectFieldIsOptionSelected( strInputFieldName , objVbsDb , objRs ) ) then
	  ' the current option element refers to the field value actually present in the record to be updated
	  funStrSelectFieldIsOptionSelected = "selected"
	else
	  ' the current option element does not refer to the field value actually present in the record to be updated
	  funStrSelectFieldIsOptionSelected = ""
  end if
end function

'***********************************************
function funStrVbsDbDrawInputSelectFieldDrawOptionDrawBeginOption( strInputFieldName , objVbsDb , objRs )
'***********************************************
  funStrVbsDbDrawInputSelectFieldDrawOptionDrawBeginOption = _
    "  <option " & funStrSelectFieldIsOptionSelected( strInputFieldName , objVbsDb , objRs ) & _
      " value='" & objRs.fields.item( 0 ).value & "'>"
end function

'***********************************************
function funStrVbsDbDrawInputSelectFieldDrawOptionDrawOptionContent( strInputFieldName , objVbsDb , objRs )
'***********************************************
  on error resume next
  funStrVbsDbDrawInputSelectFieldDrawOptionDrawOptionContent = objRs.fields.item( 1 ).value
  if err.number<>0 then
    strErrorMessage = "InputSelectFields is assigned this wrong Sql select statement:<br><br>" & _
      objVbsDb( "DictionaryInputSelectFields" )( strInputFieldName ) & "<br><br>" & _
      "assigned to this field: " & strInputFieldName & "<br><br>" & _ 
      "Returned error description is: '" & err.description & "'<br><br>" & _
      "Please, remember that each InputSelectFields select statement must be of the form <br><br>" & _
      "'select FieldA,FieldB from ...'<br><br>" & _
      "If this doesn't solve, please, " & _
      "contact VBSdb support, providing " &_
      "them with this message and the debugging values below. Thanks"
    'drawErrorWithDebugging "InputSelectFields is assigned a wrong select statement: " & strSql  , objVbsDb
    drawErrorWithDebugging strErrorMessage  , objVbsDb
  end if
end function

'***********************************************
function funStrVbsDbDrawInputSelectFieldDrawOptionDrawEndOption()
'***********************************************
  funStrVbsDbDrawInputSelectFieldDrawOptionDrawEndOption = "</option>" & chr( 10 )
end function

'***********************************************
function funStrvbsDbDrawInputSelectFieldDrawOption( strInputFieldName , objVbsDb , objRs )
'***********************************************
  if not funBolInputFieldIsReadOnly( strInputFieldName , objVbsDb ) then
    ' the field is not set to read only
  	funStrvbsDbDrawInputSelectFieldDrawOption = _
  	  funStrVbsDbDrawInputSelectFieldDrawOptionDrawBeginOption( strInputFieldName , objVbsDb , objRs )
  	funStrvbsDbDrawInputSelectFieldDrawOption = funStrvbsDbDrawInputSelectFieldDrawOption & _
  	  funStrVbsDbDrawInputSelectFieldDrawOptionDrawOptionContent( strInputFieldName , objVbsDb , objRs )
  	funStrvbsDbDrawInputSelectFieldDrawOption = funStrvbsDbDrawInputSelectFieldDrawOption & _
    	funStrVbsDbDrawInputSelectFieldDrawOptionDrawEndOption()
  else
    ' the field is set to read only
    if funBolSelectFieldIsOptionSelected( strInputFieldName , objVbsDb , objRs ) then
      ' the current option element refers to the field value actually present in the record to be updated
  	  funStrvbsDbDrawInputSelectFieldDrawOption = funStrvbsDbDrawInputSelectFieldDrawOption & _
  	    funStrVbsDbDrawInputSelectFieldDrawOptionDrawOptionContent( strInputFieldName , objVbsDb , objRs )
  	else
  	  funStrvbsDbDrawInputSelectFieldDrawOption = ""
    end if
  end if
end function

'***********************************************
function funStrVbsDbDrawInputSelectFieldDrawOptionsLoop( strInputFieldName , objVbsDb , byRef objRs )
'***********************************************
  funStrVbsDbDrawInputSelectFieldDrawOptionsLoop = ""
	while not objRs.eof
		funStrVbsDbDrawInputSelectFieldDrawOptionsLoop = funStrVbsDbDrawInputSelectFieldDrawOptionsLoop & _
		  funStrvbsDbDrawInputSelectFieldDrawOption( strInputFieldName , objVbsDb , objRs )
		objRs.moveNext
	wend
end function

'***********************************************
sub vbsDbDrawInputSelectFieldDrawOptionsCloseRecordSet( byRef objRs )
'***********************************************
	objRs.close
	set objRs = nothing
end sub

'***********************************************
function funStrVbsDbDrawInputSelectFieldDrawOptions( strInputFieldName , objVbsDb )
'***********************************************
	dim objRs
	funStrVbsDbDrawInputSelectFieldDrawOptions = _
	  funStrVbsDbDrawInputSelectFieldDrawOptionsDrawBlankOption( strInputFieldName , objVbsDb )
	vbsDbDrawInputSelectFieldDrawOptionsGetObjRs strInputFieldName , objVbsDb , objRs
	funStrVbsDbDrawInputSelectFieldDrawOptions = funStrVbsDbDrawInputSelectFieldDrawOptions & _
	  funStrVbsDbDrawInputSelectFieldDrawOptionsLoop( strInputFieldName , objVbsDb , objRs )
	vbsDbDrawInputSelectFieldDrawOptionsCloseRecordSet objRs
end function

'***********************************************
function funStrVbsDbDrawInputSelectFieldDrawEndSelect( strInputFieldName , objVbsDb )
'***********************************************
  if not funBolInputFieldIsReadOnly( strInputFieldName , objVbsDb ) then
    ' the field is not set to read only
    funStrVbsDbDrawInputSelectFieldDrawEndSelect = "</select>" & chr( 10 )
  else
    ' the field is set to read only
    funStrVbsDbDrawInputSelectFieldDrawEndSelect = ""
  end if
end function

'***********************************************
'***********************************************
function funStrVbsDbDrawInputSelectField( strInputFieldName , objVbsDb )
'***********************************************
'***********************************************
  'dim strInputFieldNameActually
  'strInputFieldNameActually = funStrInputFieldNameActually( strInputFieldName , objVbsDb )
  funStrVbsDbDrawInputSelectField = _
	  funStrVbsDbDrawInputSelectFieldDrawBeginSelect( strInputFieldName , objVbsDb ) & _
	  funStrVbsDbDrawInputSelectFieldDrawOptions( strInputFieldName , objVbsDb ) & _
	  funStrVbsDbDrawInputSelectFieldDrawEndSelect( strInputFieldName , objVbsDb )
end function

'***********************************************
'***********************************************
sub vbsDbSetDictionaryInputSelectFields( byRef objVbsDb )
'***********************************************
'***********************************************
  vbsDbSetDictionaryObjectFromList "InputSelectFields" , objVbsDb
end sub
%>