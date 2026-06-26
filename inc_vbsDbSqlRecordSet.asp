<%
'***********************************************
'***********************************************
function funVarSqlRecordSetFieldValue( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  funVarSqlRecordSetFieldValue = _
    objVbsDb( "SqlRecordSet" ).fields( strFieldName ).value
end function

'***********************************************
'***********************************************
function funVarSqlRecordSetFieldType( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  funVarSqlRecordSetFieldType = _
    objVbsDb( "SqlRecordSet" ).fields( strFieldName ).type
end function

'***********************************************
function funBolSqlRecordSetFieldIsString( strFieldName , objVbsDb )
'***********************************************
  dim intFieldType
  intFieldType = funVarSqlRecordSetFieldType( strFieldName , objVbsDb )
  funBolSqlRecordSetFieldIsString = _
    ( intFieldType = vbsDbAdChar ) or _
    ( intFieldType = vbsDbAdVarChar ) or _
    ( intFieldType = vbsDbAdLongVarChar ) or _
    ( intFieldType = vbsDbAdWChar ) or _
    ( intFieldType = vbsDbAdVarWChar ) or _
    ( intFieldType = vbsDbAdLongVarWChar )
end function

'***********************************************
'***********************************************
function funVarSqlRecordSetFieldDefinedSize( strFieldName , objVbsDb )
'***********************************************
'***********************************************
  if funBolSqlRecordSetFieldIsString( strFieldName , objVbsDb ) then
    ' the field is a string
    funVarSqlRecordSetFieldDefinedSize = _
      objVbsDb( "SqlRecordSet" ).fields( strFieldName ).definedSize
  else
    ' the field is a not a string
    funVarSqlRecordSetFieldDefinedSize = constDefaultMaxLenght
  end if
end function

'***********************************************
'***********************************************
sub vbsDbSetSqlRecordSet( byRef objVbsDb )
'***********************************************
'***********************************************
  dim strSql , strErrorMessage
  strSql = objVbsDb( "SqlWithFilterWhereClause" )
  set objVbsDb( "SqlRecordSet" ) = Server.CreateObject("ADODB.Recordset")
  objVbsDb( "SqlRecordSet" ).cursorLocation = vbsDbAdUseClient
  on error resume next
  objVbsDb( "SqlRecordSet" ).Open strSql , objVbsDb( "Connection" ) , vbsDbAdOpenKeySet , vbsDbAdLockReadOnly , vbsDbAdCmdText
  if err.number<>0 then
    strErrorMessage = "Sql property not accepted. To fetch the Grid records, VBSdb is trying this Sql query:<br><br>" & _
      strSql & "<br><br>" & _
      "Returned error description is: '" & err.description & "'<br><br>" & _
      "Please, perform the following debug steps:" & _
      "<div align=left>" & _
      "<ul><li>control none of the tables involved in the query is exclusively opened for database managment</li>" & _
      "<li>try the Sql query displayed above directly against your database</li>" & _
      "<li>check the Sql property assignment</li>" & _
      "<li>check the SearchAliasFields property if a filter has been applyed and an ambiguous field name is detected</li>" & _
      "</ul>" & _
      "</div>"
    drawError strErrorMessage , objVbsDb
  end if
end sub
%>