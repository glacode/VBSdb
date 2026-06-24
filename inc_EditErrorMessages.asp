<%
'***********************************************
sub vbsDbDefaultEditErrorMessageForAdd( strSql , objVbsDb )
'***********************************************
  dim strErrorMessage
  strErrorMessage = "Error executing this Add statement:<br><br>" & _
    strSql & "<br><br>" & _
    "Returned error number is: '" & err.number & "'<br><br>" & _
    "Returned error description is: '" & err.description & "'<br><br>" & _
    "Please, perform the following debug steps:" & _
    "<div align=left>" & _
    "<ul><li>control the adding table is not exclusively opened for database managment</li>" & _
    "<li>control user inserted values are not of the wrong type. To avoid this kind of error, just add " & _
    "Javascript validation constraints</li>" & _
    "<li>control if field constraints have been violated. To avoid this kind of error, just add " & _
    "Javascript validation constraints</li>" & _
    "<li>control if foreign key constraints have been violated. To avoid this kind of error, use the InputSelectFields property and, if necessary, add " & _
    "Javascript validation constraints</li>" & _
    "<li>to help you check the described errors, try the above Sql add statement directly against your database</li>" & _
    "<li>furthermore, if your server is Win NT or Win 2000, control the Internet Guest User has been given write permissions on the database folder</li>" & _
    "<li>if the case, you can consider to use the EditErrorMessages property</li>" & _
    "</ul>" & _
    "</div>" & _
    "If none of the above solves, please, " & _
    "contact VBSdb support, providing " &_
    "them with this message and the debugging values below. Thanks."
  drawErrorWithDebugging strErrorMessage , objVbsDb
end sub

'***********************************************
'***********************************************
sub vbsDbEditErrorMessageForAdd( strSql , objVbsDb )
'***********************************************
'***********************************************
  objVbsDb( "EditErrorNumber" ) = err.number
  objVbsDb( "EditErrorDescription" ) = err.description
  if ( uCase( objVbsDb( "EditErrorMessages" ) ) = "DEFAULT" ) then
    vbsDbDefaultEditErrorMessageForAdd strSql , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbDefaultEditErrorMessageForUpdate( strSql , objVbsDb )
'***********************************************
  dim strErrorMessage
  strErrorMessage = "Error executing this Update statement:<br><br>" & _
    strSql & "<br><br>" & _
    "Returned error description is: '" & err.description & "'<br><br>" & _
    "Please, perform the following debug steps:" & _
    "<div align=left>" & _
    "<ul><li>control the table to be updated is not exclusively opened for structure managment</li>" & _
    "<li>control user inserted values are not of the wrong type. To avoid this kind of error, just add " & _
    "Javascript validation constraints</li>" & _
    "<li>control if field constraints have been violated. To avoid this kind of error, just add " & _
    "Javascript validation constraints</li>" & _
    "<li>control if foreign key constraints have been violated. To avoid this kind of error, use the InputSelectFields property and, if necessary, add " & _
    "Javascript validation constraints</li>" & _
    "<li>to help you check the described errors, try the above Sql add statement directly against your database</li>" & _
    "<li>furthermore, if your server is Win NT or Win 2000, control the Internet Guest User has been given write permissions on the database folder</li>" & _
    "</ul>" & _
    "</div>" & _
    "If none of the above solves, please, " & _
    "contact VBSdb support, providing " &_
    "them with this message and the debugging values below. Thanks."
  drawErrorWithDebugging strErrorMessage , objVbsDb
end sub

'***********************************************
'***********************************************
sub vbsDbEditErrorMessageForUpdate( strSql , objVbsDb )
'***********************************************
'***********************************************
  objVbsDb( "EditErrorNumber" ) = err.number
  objVbsDb( "EditErrorDescription" ) = err.description
  if ( uCase( objVbsDb( "EditErrorMessages" ) ) = "DEFAULT" ) then
    vbsDbDefaultEditErrorMessageForUpdate strSql , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbDefaultEditErrorMessageForDelete( strSql , objVbsDb )
'***********************************************
  dim strErrorMessage
  strErrorMessage = "Error executing this Delete statement:<br><br>" & _
    strSql & "<br><br>" & _
    "Returned error description is: '" & err.description & "'<br><br>" & _
    "Please, perform the following debug steps:" & _
    "<div align=left>" & _
    "<ul><li>control the table to be updated is not exclusively opened for structure managment</li>" & _
    "<li>control if foreign key constraints have been violated</li>" & _
    "<li>to help you check the described errors, try the above Sql add statement directly against your database</li>" & _
    "<li>furthermore, if your server is Win NT or Win 2000, control the Internet Guest User has been given write permissions on the database folder</li>" & _
    "</ul>" & _
    "</div>" & _
    "If none of the above solves, please, " & _
    "contact VBSdb support, providing " &_
    "them with this message and the debugging values below. Thanks."
  drawErrorWithDebugging strErrorMessage , objVbsDb
end sub

'***********************************************
'***********************************************
sub vbsDbEditErrorMessageForDelete( strSql , objVbsDb )
'***********************************************
'***********************************************
  objVbsDb( "EditErrorNumber" ) = err.number
  objVbsDb( "EditErrorDescription" ) = err.description
  if ( uCase( objVbsDb( "EditErrorMessages" ) ) = "DEFAULT" ) then
    vbsDbDefaultEditErrorMessageForDelete strSql , objVbsDb
  end if
end sub

'***********************************************
sub vbsDbDrawVbsDbDrawEditDbErrorActually( strTextItemName , objVbsDb )
'***********************************************
%>
  <div class=ViewMessage>
    <%=funStrTranslate( strTextItemName , objVbsDb )%>
  </div>
<%
end sub

'***********************************************
'***********************************************
sub vbsDbDrawVbsDbDrawEditDbError( objVbsDb )
'***********************************************
'***********************************************
  select case VbsDbGetLastAction( objVbsDb )
    case "SubmitAdd"
      vbsDbDrawVbsDbDrawEditDbErrorActually "addErrorMessage" , objVbsDb
    case "SubmitUpdate"
      vbsDbDrawVbsDbDrawEditDbErrorActually "updateErrorMessage" , objVbsDb
    case "SubmitDelete"
      vbsDbDrawVbsDbDrawEditDbErrorActually "deleteErrorMessage" , objVbsDb
  end select
end sub
%>