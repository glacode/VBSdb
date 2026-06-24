<%
'***********************************************
sub vbsDbOpenObjRsForSqlServerAutoIncrement( objVbsDb , byRef objRs )
'***********************************************
  dim strSql
  strSql = _
    "select name from syscolumns "                        & _
    "where (status & 128) <> 0 "                          & _ 
       "and id = (select id from sysobjects "             & _
                  "where lower(name)='[" & lCase( objVbsDb( "EditTableName" ) ) & "]')"
  set objRs = Server.CreateObject("ADODB.Recordset")
  objRs.cursorLocation = vbsDbAdUseClient
  on error resume next
  objVbsDb( "SqlRecordSet" ).Open strSql , objVbsDb( "Connection" ) , vbsDbAdOpenKeySet , vbsDbAdLockReadOnly , vbsDbAdCmdText
  if err.number<>0 then
    funBolAdoxFieldAutoincrementNoAdoxSupportForSqlServer = false
  end if
end sub

'***********************************************
function funStrVbsDbSetEditSlqServerAutoIncrementWithObjRs( objVbsDb , objRs )
'***********************************************
  funStrVbsDbSetEditSlqServerAutoIncrementWithObjRs = ""
  on error resume next
  if not objRs.eof then
    ' objRs( "name" ) contains the Autoincrement field name
    funStrVbsDbSetEditSlqServerAutoIncrementWithObjRs = objRs( "name" )
  end if
end function

'***********************************************
function funStrVbsDbSetEditSlqServerAutoIncrementActually( objVbsDb )
'***********************************************
  dim objRs
  vbsDbOpenObjRsForSqlServerAutoIncrement objVbsDb , objRs
  funStrVbsDbSetEditSlqServerAutoIncrementActually =_
    funStrVbsDbSetEditSlqServerAutoIncrementWithObjRs( objRs , objVbsDb )
  on error resume next
  objRs.close
  set objRs = nothing
end function

'***********************************************
function funStrVbsDbSetEditSlqServerAutoIncrement( objVbsDb )
'***********************************************
  if uCase( objVbsDb( "GlobalDbType" ) ) = "SQLSERVER" then
    funStrVbsDbSetEditSlqServerAutoIncrement = _
      funStrVbsDbSetEditSlqServerAutoIncrementActually( objVbsDb )
  else
    funStrVbsDbSetEditSlqServerAutoIncrement = ""
  end if
end function

'***********************************************
'***********************************************
sub vbsDbSetEditSlqServerAutoIncrement( byRef objVbsDb )
'***********************************************
'***********************************************
  objVbsDb( "EditSlqServerAutoIncrement" ) = funStrVbsDbSetEditSlqServerAutoIncrement( objVbsDb )
end sub
%>