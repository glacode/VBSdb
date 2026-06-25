<%
'***********************************************
sub vbsDbDrawGridEditColumnHeaderActually( objVbsDb )
'***********************************************
%>
  <th class="GridEditColumn">&nbsp;</th>
<%
end sub

'***********************************************
'***********************************************
sub vbsDbDrawGridEditColumnHeader( objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "GridUpdateButtons" ) or  objVbsDb( "GridDeleteButtons" ) then
    ' the grid edit column must be displayed
    vbsDbDrawGridEditColumnHeaderActually objVbsDb
  end if
end sub

'***********************************************
function funStrEditButtonClickClass( strEditType )
'***********************************************
  if strEditType = "update" then
    ' an update button has to be displayed
    funStrEditButtonClickClass = "VBSdbEditUpdate"
  else
    ' a delete button has to be displayed
    funStrEditButtonClickClass = "VBSdbEditDelete"
  end if
end function

'***********************************************
function funStrHrefForEditButton( strEditType , objVbsDb )
'***********************************************
  dim strClickClass
  strClickClass = funStrEditButtonClickClass( strEditType )
  funStrHrefForEditButton = request.serverVariables( "url" ) & "?VBSdbClickClass_" & objVbsDb( "GlobalId" ) & _
    "=" & strClickClass & "&amp;VBSdbEditWhere_" & objVbsDb( "GlobalId" ) & "=" & _
    funStrVbsDbEditWhereUrlEncoded( objVbsDb )
end function


'***********************************************
function funStrGridEditButtonActually( strHref , strTitle , strImageFileName , objVbsDb )
'***********************************************
  funStrGridEditButtonActually = "<a href='" & funStrVbsDbGlobalQuerystringPreserve( strHref , objVbsDb ) & "'><img title='" & _
    strTitle & "' border='0' align='absmiddle' hspace=1 width=14 height=14 SRC='" & _
    objVbsDb( "GlobalImageDir" ) & strImageFileName & "'></a>"
end function

'***********************************************
function funStrGridEditButton( strHref , strTitle , strImageFileName , objVbsDb )
'***********************************************
  if objVbsDb( "GridUpdateButtons" ) = "" then
    funStrGridEditButton = ""
  else
    funStrGridEditButton = funStrGridEditButtonActually( strHref , strTitle , strImageFileName , objVbsDb )
  end if
end function

'***********************************************
function funStrGridUpdateButton( objVbsDb )
'***********************************************
  funStrGridUpdateButton = funStrGridEditButton( _
    funStrHrefForEditButton( "update" , objVbsDb ) , _
    funStrTranslate( "gridUpdateButtonTitle" , objVbsDb ) , _
    "updateColumn.gif" , _
    objVbsDb )
end function

'***********************************************
function funStrGridDeleteButton( objVbsDb )
'***********************************************
  funStrGridDeleteButton = funStrGridEditButton( _
    funStrHrefForEditButton( "delete" , objVbsDb ) , _
    funStrTranslate( "gridDeleteButtonTitle" , objVbsDb ) , _
    "deleteColumn.gif" , _
    objVbsDb )
end function

'***********************************************
function funStrGridRowEditButtonsContent( objVbsDb )
'***********************************************
  if objVbsDb( "GridUpdateButtons" ) then
    funStrGridRowEditButtonsContent = funStrGridUpdateButton( objVbsDb )
  else
    funStrGridRowEditButtonsContent = ""
  end if
  if objVbsDb( "GridUpdateButtons" ) and objVbsDb( "GridDeleteButtons" ) then
    'funStrGridRowEditButtonsContent = funStrGridRowEditButtonsContent & "&nbsp;"
    funStrGridRowEditButtonsContent = funStrGridRowEditButtonsContent
  end if
  if objVbsDb( "GridDeleteButtons" ) then
    funStrGridRowEditButtonsContent = funStrGridRowEditButtonsContent & funStrGridDeleteButton( objVbsDb )
  end if
end function

'***********************************************
sub vbsDbDrawGridRowEditButtonsActually( objVbsDb )
'***********************************************
%>
  <td valign=center align=center class="<%=funStrDbDrawCSSClassGridRow( objVbsDb )%>"><%=funStrGridRowEditButtonsContent( objVbsDb )%></td>
<%
end sub

'***********************************************
'***********************************************
sub vbsDbDrawGridRowEditButtons( objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "GridUpdateButtons" ) or objVbsDb( "GridDeleteButtons" ) then
    ' the grid edit column must be displayed
    vbsDbDrawGridRowEditButtonsActually objVbsDb
  end if
end sub
%>