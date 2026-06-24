<%
'***********************************************
function funStrEditFieldNameForInputFieldName( strInputFieldName , objVbsDb )
'***********************************************
  dim intDotPosition
  intDotPosition = inStr( strInputFieldName , "." )
  if intDotPosition > 0 then
    ' the input field name is disambiguated by the table name prefix
    funStrEditFieldNameForInputFieldName = mid( strInputFieldName , intDotPosition + 1 )
  else
    funStrEditFieldNameForInputFieldName = strInputFieldName
  end if
  funStrEditFieldNameForInputFieldName = replace( funStrEditFieldNameForInputFieldName , "[" , "" )
  funStrEditFieldNameForInputFieldName = replace( funStrEditFieldNameForInputFieldName , "]" , "" )
end function

'***********************************************
function funBolAddIsGranted( objVbsDb )
'***********************************************
  funBolAddIsGranted = ( funBolInDbNavigationItem( "ADD" , objVbsDb ) or _
                         objVbsDb( "EditAddSingle" ) or _
                         uCase( objVbsDb( "GlobalStartScreen" ) ) = "ADD" )
end function

'***********************************************
function funStrCheckFormTamperingForAdd( objVbsDb )
'***********************************************
  if funBolAddIsGranted( objVbsDb ) then
    ' form tampering for a non-allowed add was not detected
    funStrCheckFormTamperingForAdd = ""
  else
    ' form tampering for a non-allowed add was detected
    funStrCheckFormTamperingForAdd= "An Add form has been submitted, but an Add operation is not granted."
  end if
end function

'***********************************************
function funBolUpdateIsGranted( objVbsDb )
'***********************************************
  funBolUpdateIsGranted = ( funBolInDbNavigationItem( "UPDATE" , objVbsDb ) or objVbsDb( "GridUpdateButtons" ) )
end function

'***********************************************
function funStrCheckFormTamperingForUpdate( objVbsDb )
'***********************************************
  if funBolUpdateIsGranted( objVbsDb ) then
    ' form tampering for a non-allowed update was not detected
    funStrCheckFormTamperingForUpdate = ""
  else
    ' form tampering for a non-allowed update was detected
    funStrCheckFormTamperingForUpdate= "An Update form has been submitted, but an Update operation is not granted."
  end if
end function

'***********************************************
function funBolDeleteIsGranted( objVbsDb )
'***********************************************
  funBolDeleteIsGranted = ( funBolInDbNavigationItem( "DELETE" , objVbsDb ) or objVbsDb( "GridDeleteButtons" ) )
end function

'***********************************************
function funStrCheckFormTamperingForDelete( objVbsDb )
'***********************************************
  if funBolDeleteIsGranted( objVbsDb ) then
    ' form tampering for a non-allowed update was not detected
    funStrCheckFormTamperingForDelete = ""
  else
    ' form tampering for a non-allowed update was detected
    funStrCheckFormTamperingForDelete= "A Delete form has been submitted, but a Delete operation is not granted."
  end if
end function

'***********************************************
function funStrEditWhereClauseForTamperingControl( objVbsDb )
'***********************************************
  dim objColumn
  ' edit table name prefix is added, to avoid field name ambiguities, in the query
  funStrEditWhereClauseForTamperingControl = objVbsDb( "RequestVBSdbEditWhere" )
  for each objColumn in objVbsDb( "AdoxCatalog" ).tables( objVbsDb( "EditTableName" ) ).columns
    funStrEditWhereClauseForTamperingControl = _
      replace( funStrEditWhereClauseForTamperingControl , _
        "(" & funStrGlobalDbTypeFieldName( objColumn.name , objVbsDb ) & " is null" , _
        "(" & funStrGlobalDbTypeEditTableName( objVbsDb ) & "." & _
        funStrGlobalDbTypeFieldName( objColumn.name , objVbsDb ) & " is null" , _
        1 , -1 , 1 )
    funStrEditWhereClauseForTamperingControl = _
      replace( funStrEditWhereClauseForTamperingControl , _
        "(" & funStrGlobalDbTypeFieldName( objColumn.name , objVbsDb ) & "=" , _
        "(" & funStrGlobalDbTypeEditTableName( objVbsDb ) & "." & _
        funStrGlobalDbTypeFieldName( objColumn.name , objVbsDb ) & "=" , _
        1 , -1 , 1 )
  next
end function

'***********************************************
function funStrSqlForSubsetFormTamperingControl( objVbsDb )
'***********************************************
  funStrSqlForSubsetFormTamperingControl = _
    funStrSqlInsertWhereClause( objVbsDb( "Sql" ) , _
      funStrEditWhereClauseForTamperingControl( objVbsDb ) , objVbsDb )
end function

'***********************************************
function funStrErrorMessageForWrongSqlInFormTamperingOutOfSubsetOk( strSql , objVbsDb )
'***********************************************
  funStrErrorMessageForWrongSqlInFormTamperingOutOfSubsetOk = _
    "Sql error trying to perform a 'Form Tampering' control. " & _
    "Returned error description is: '" & err.description & "'<br><br>" & _
    "The Form Tampering control Sql query is:<br><br>" & _
    strSql & "<br><br>" & _
    "The original Sql query is:<br><br>" & _
    objVbsDb( "Sql" ) & "<br><br>" & _
    "If you have not tried a Form Tampering, you should perform the following controls:" & _
    "<div align=left>" & _
    "<ul><li>control none of the tables involved in the query is exclusively opened for database managment</li>" & _
    "<li>try the Sql queries displayed above directly against your database</li>" & _
    "<li>check the Sql property assignment</li>" & _
    "<li>check the SearchAliasFields property if a filter has been applyed and an ambiguous field name is detected</li>" & _
    "</ul>" & _
    "</div>" & _
    "If none of the above solves, you can avoid this error setting the EditFormTamperingControl to false (but " & _
    "you should check the EditFormTamperingControl documentation to decide about this) and please, " & _
    "contact VBSdb support, providing " &_
    "them with this message and the debugging values below. Thanks."
end function

'***********************************************
function funStrFormTamperingOutOfSubsetOkActually( strSql , objVbsDb )
'***********************************************
  dim objRecordSet , strErrorMessage
  strSql = funStrSqlForSubsetFormTamperingControl( objVbsDb )
  set objRecordSet = Server.CreateObject("ADODB.Recordset")
  on error resume next
  objRecordSet.Open strSql , objVbsDb( "Connection" ) , vbsDbAdOpenKeySet , vbsDbAdLockReadOnly , vbsDbAdCmdText
  if err.number<>0 then
    ' this message should arise, only if the Sql property assignment is wrong, but it should never
    ' happen at this point (it should always be intercepted by the Sql error control, in previous VBSdb produced screens)
    strErrorMessage = funStrErrorMessageForWrongSqlInFormTamperingOutOfSubsetOk( strSql , objVbsDb )
    drawErrorWithDebugging strErrorMessage , objVbsDb
  end if
  if objRecordSet.eof then
    ' form tampering detected
    funStrFormTamperingOutOfSubsetOkActually = "The record to be edited results to be outside the set of records returned by the Sql property."
  else
    ' form tampering is not detected
    funStrFormTamperingOutOfSubsetOkActually = ""
  end if
  objRecordSet.close
  set objRecordSet = nothing
  objRecordSet = ""
end function

'***********************************************
function funStrFormTamperingOutOfSubsetOk( objVbsDb )
'***********************************************
  dim strUcaseSql
  strUcaseSql = uCase( objVbsDb( "Sql" ) )
  if inStr( strUcaseSql , "SELECT " ) > 0 and inStr( strUcaseSql , " FROM " ) > 0 then
    ' the Sql property contains a select statement, it is not a stored procedure call
    funStrFormTamperingOutOfSubsetOk = funStrFormTamperingOutOfSubsetOkActually( strUcaseSql , objVbsDb )
  else
    ' the Sql property contains a stored procedure call
    funStrFormTamperingOutOfSubsetOk = true
  end if
end function

'***********************************************
function funStrFindClauseForFormTamperingInputSelectFields( strFieldName , objRecordSet , objVbsDb )
'***********************************************
  dim strRecordSetFieldName , strFieldValue
  strRecordSetFieldName = objRecordSet.fields( 0 ).name
  strFieldValue = cStr( request.form( funStrFormFieldName( funStrEditFieldNameForInputFieldName( strFieldName , objVbsDb ) , objVbsDb ) ) )
  funStrFindClauseForFormTamperingInputSelectFields = _
    "(" & funStrFieldWhereClauseWithFieldValue( objRecordSet.fields( 0 ).type , _
    strRecordSetFieldName , strFieldValue , objVbsDb ) & ")"
end function

'***********************************************
function funStrFormTamperingInputSelectFieldsOkForFieldWithRecordSet( strFieldName , objRecordSet , objVbsDb )
'***********************************************
  objRecordSet.find( funStrFindClauseForFormTamperingInputSelectFields( strFieldName , objRecordSet , objVbsDb ) )
  if objRecordSet.eof then
    ' form tampering for InputSelectedFields was detected for the field strFieldName
    funStrFormTamperingInputSelectFieldsOkForFieldWithRecordSet = _
      "InputSelectFields form tampering: the form value for the field " & _
      strFieldName & " is " & cStr( request.form( funStrFormFieldName( funStrEditFieldNameForInputFieldName( strFieldName , objVbsDb ) , objVbsDb ) ) ) & _
      " and it is not in the drop down list for allowed values."
  else
    ' form tampering for InputSelectedFields was not detected for the field strFieldName
    funStrFormTamperingInputSelectFieldsOkForFieldWithRecordSet = ""
  end if
end function

'***********************************************
function funStrFormTamperingInputSelectFieldsOkForFieldActually( strFieldName , objVbsDb )
'***********************************************
  dim strSql , objRecordSet , strErrorMessage
  strSql = objVbsDb( "DictionaryInputSelectFields" )( strFieldName )
  set objRecordSet = Server.CreateObject("ADODB.Recordset")
  on error resume next
  objRecordSet.Open strSql , objVbsDb( "Connection" ) , vbsDbAdOpenKeySet , vbsDbAdLockReadOnly , vbsDbAdCmdText
  funStrFormTamperingInputSelectFieldsOkForFieldActually = _
    funStrFormTamperingInputSelectFieldsOkForFieldWithRecordSet( strFieldName , objRecordSet , objVbsDb )
  objRecordSet.close
  set objRecordSet = nothing
  objRecordSet = ""
end function

'***********************************************
function funStrFormTamperingInputSelectFieldsOkForField( strFieldName , objVbsDb )
'***********************************************
  if funBolEditFieldIsReadOnly( strFieldName , objVbsDb ) or _
    trim( cStr( request.form( funStrFormFieldName( funStrEditFieldNameForInputFieldName( strFieldName , objVbsDb ) , objVbsDb ) ) ) ) = "" then
    ' the field is read only or the visitor has chosen the empty option
    funStrFormTamperingInputSelectFieldsOkForField = ""
  else
    funStrFormTamperingInputSelectFieldsOkForField = _
      funStrFormTamperingInputSelectFieldsOkForFieldActually( strFieldName , objVbsDb )
  end if
end function

'***********************************************
function funStrFormTamperingInputSelectFieldsOk( objVbsDb )
'***********************************************
  dim strFieldName
  funStrFormTamperingInputSelectFieldsOk = ""
  for each strFieldName in objVbsDb( "DictionaryInputSelectFields" )
    if funStrFormTamperingInputSelectFieldsOk = "" then
      funStrFormTamperingInputSelectFieldsOk = _
        funStrFormTamperingInputSelectFieldsOkForField( strFieldName , objVbsDb )
    end if
  next
end function

'***********************************************
function funStrFormTamperingInputEnumeratedFieldsOkForField( strFieldName , objVbsDb )
'***********************************************
  dim bolFormTamperingInputEnumeratedFieldsOkForField , strFormFieldName , strEnumeratedOption
  bolFormTamperingInputEnumeratedFieldsOkForField = false
  strFormFieldName = funStrFormFieldName( funStrEditFieldNameForInputFieldName( strFieldName , objVbsDb ) , objVbsDb )
  for each strEnumeratedOption in objVbsDb( "DictionaryInputEnumeratedFields" )( strFieldName )
    bolFormTamperingInputEnumeratedFieldsOkForField = _
      ( bolFormTamperingInputEnumeratedFieldsOkForField or _
      ( objVbsDb( "DictionaryInputEnumeratedFields" )( strFieldName )( strEnumeratedOption )( "value" ) = _
        request.form( strFormFieldName ) ) )
  next
  if bolFormTamperingInputEnumeratedFieldsOkForField then
    ' form tampering for InputEnumeratedFields was not detected for field strFieldName
    funStrFormTamperingInputEnumeratedFieldsOkForField = ""
  else
    ' form tampering for InputEnumeratedFields was detected for field strFieldName
    funStrFormTamperingInputEnumeratedFieldsOkForField = "InputEnumeratedFields form tampering: the form value for the field " & _
      strFieldName & " is " & request.form( strFormFieldName ) & " and it is not in the drop down list for allowed values."
  end if
end function

'***********************************************
function funStrFormTamperingInputEnumeratedFieldsOkActually( objVbsDb )
'***********************************************
  dim strFieldName
  funStrFormTamperingInputEnumeratedFieldsOkActually = ""
  for each strFieldName in objVbsDb( "DictionaryInputEnumeratedFields" )
    if ( funStrFormTamperingInputEnumeratedFieldsOkActually = "" ) _
      and _
      ( not funBolEditFieldIsReadOnly( strFieldName , objVbsDb ) ) then
      funStrFormTamperingInputEnumeratedFieldsOkActually = funStrFormTamperingInputEnumeratedFieldsOkForField( strFieldName , objVbsDb )
    end if
  next
end function

'***********************************************
function funStrFormTamperingInputEnumeratedFieldsOk( objVbsDb )
'***********************************************
  if ( not funBolIsProVersion() ) then
    ' current version is VBSdb Free (and it doesn't support InputEnumeratedFields)
    funStrFormTamperingInputEnumeratedFieldsOk = ""
  else
    ' current version is VBSdb Pro (and it supports InputEnumeratedFields)
    funStrFormTamperingInputEnumeratedFieldsOk = funStrFormTamperingInputEnumeratedFieldsOkActually( objVbsDb )
  end if
end function

'***********************************************
function funStrFormTamperingDropDownOk( objVbsDb )
'***********************************************
  funStrFormTamperingDropDownOk = funStrFormTamperingInputSelectFieldsOk( objVbsDb )
  if funStrFormTamperingDropDownOk = "" then
    ' form tampering for InputSelectFields was not detected
    funStrFormTamperingDropDownOk = funStrFormTamperingInputEnumeratedFieldsOk( objVbsDb )
  end if
end function

'***********************************************
function funStrFormTamperingOkActually( objVbsDb )
'***********************************************
  funStrFormTamperingOkActually = ""
  if ( VbsDbGetLastAction( objVbsDb ) = "SubmitUpdate" or _
      VbsDbGetLastAction( objVbsDb ) = "SubmitDelete" ) then
    funStrFormTamperingOkActually = funStrFormTamperingOutOfSubsetOk( objVbsDb )
  end if
  if ( funStrFormTamperingOkActually = "" ) and _
    ( VbsDbGetLastAction( objVbsDb ) = "SubmitAdd" or _
      VbsDbGetLastAction( objVbsDb ) = "SubmitUpdate" ) then
    funStrFormTamperingOkActually = funStrFormTamperingDropDownOk( objVbsDb )
  end if
end function

'***********************************************
function funStrFormComplexFormTamperingControlIfRequested( objVbsDb )
'***********************************************
  if objVbsDb( "EditFormTamperingControl" ) then
    ' the developer asked for a Form Tampering control
    funStrFormComplexFormTamperingControlIfRequested = funStrFormTamperingOkActually( objVbsDb )
  else
    ' the developer asked not to perform a Form Tampering control
    funStrFormComplexFormTamperingControlIfRequested = ""
  end if
end function

'***********************************************
'***********************************************
function funStrFormTamperingControlForEditDb( objVbsDb )
'***********************************************
'***********************************************
  dim strLastAction
  strLastAction = VbsDbGetLastAction( objVbsDb )
  funStrFormTamperingControlForEditDb = ""
  if ( strLastAction = "SubmitAdd" ) then
    funStrFormTamperingControlForEditDb = funStrCheckFormTamperingForAdd( objVbsDb )
  end if
  if ( funStrFormTamperingControlForEditDb = "" ) and ( strLastAction = "SubmitUpdate" ) then
    funStrFormTamperingControlForEditDb = funStrCheckFormTamperingForUpdate( objVbsDb )
  end if
  if ( funStrFormTamperingControlForEditDb = "" ) and ( strLastAction = "SubmitDelete" ) then
    funStrFormTamperingControlForEditDb = funStrCheckFormTamperingForDelete( objVbsDb )
  end if
  if ( funStrFormTamperingControlForEditDb = "" ) then
    funStrFormTamperingControlForEditDb = funStrFormComplexFormTamperingControlIfRequested( objVbsDb )
  end if
end function
%>