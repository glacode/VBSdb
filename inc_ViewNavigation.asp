<%
'***********************************************
function funBolFirstPage( byRef objVbsDb )
'***********************************************
  funBolFirstPage = ( objVbsDb( "SqlRecordSet" ).absolutePage = 1 )
end function

'***********************************************
sub vbsDbDrawBeginNavigationItem( objVbsDb )
'***********************************************
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
%>
	<td class="ViewNavigation">
<%
  else
    ' the developer wants to use images for buttons
%>
  <td>
<%
  end if
end sub

'***********************************************
sub vbsDbDrawEndNavigationItem()
'***********************************************
%>
  </td>
<%
end sub

'***********************************************
sub vbsDbDrawNavigationDisabledText( strText , objVbsDb )
'***********************************************
%>
  &nbsp;<%=strText%>&nbsp;
<%
end sub

'***********************************************
sub vbsDbDrawNavigationDisabledImage( strTitle , strImageFileName , objVbsDb )
'***********************************************
%>
  <img title="<%=strTitle%>" border="0" align="absmiddle" SRC="<%=objVbsDb( "GlobalImageDir" ) & strImageFileName%>">
<%
end sub

'***********************************************
sub vbsDbDrawNavigationEnabledImage( strHref , strTitle , strImageFileName , objVbsDb )
'***********************************************
%>
  <a href="<%=funStrVbsDbGlobalQuerystringPreserve( strHref , objVbsDb )%>">
    <img title="<%=strTitle%>" border="0" align="absmiddle" SRC="<%=objVbsDb( "GlobalImageDir" ) & strImageFileName%>"></a>
<%
end sub

'***********************************************
sub vbsDbDrawNavigationEnabledText( strHref , strTitle , strText , objVbsDb )
'***********************************************
%>
  &nbsp;<a href="<%=funStrVbsDbGlobalQuerystringPreserve( strHref , objVbsDb )%>"
            title="<%=strTitle%>" class="ViewNavigation"><%=strText%></a>&nbsp;
<%
end sub

'***********************************************
sub vbsDbDrawNavigationFirstFirstPage( objVbsDb )
'***********************************************
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationDisabledText funStrTranslate( "first" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationDisabledImage funStrTranslate( "firstAnchorTitleAlreadyFirst" , objVbsDb ) , "disabledFirst.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
sub vbsDbDrawNavigationFirstNotFirstPage( byRef objVbsDb )
'***********************************************
  dim strHref
  strHref = request.serverVariables( "url" ) & "?VBSdbClickClass_" & _
    objVbsDb( "GlobalId" ) & "=VBSdbGridFirst"
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationEnabledText strHref , funStrTranslate( "firstAnchorTitle" , objVbsDb ) , funStrTranslate( "first" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationEnabledImage strHref , funStrTranslate( "firstAnchorTitle" , objVbsDb ) , "first.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
sub vbsDbDrawNavigationFirst( byRef objVbsDb )
'***********************************************
  if objVbsDb( "ViewNavigationFirst" ) and ( not objVbsDb( "SqlRecordSet" ).eof ) then
  	' the developer wants to show the First button
  	if funBolFirstPage( objVbsDb ) then
  	  ' current page is the first one
  	  vbsDbDrawNavigationFirstFirstPage objVbsDb
  	else
  	  ' current page is not the first one
  	  vbsDbDrawNavigationFirstNotFirstPage objVbsDb			
  	end if
  end if
end sub

'***********************************************
sub vbsDbDrawNavigationPrevFirstPage( byRef objVbsDb )
'***********************************************
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationDisabledText funStrTranslate( "prev" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationDisabledImage funStrTranslate( "prevAnchorTitleAlreadyFirst" , objVbsDb ) , "disabledPrev.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
sub vbsDbDrawNavigationPrevNotFirstPage( byRef objVbsDb )
'***********************************************
  dim strHref
  strHref = request.serverVariables( "url" ) & "?VBSdbClickClass_" & objVbsDb( "GlobalId" ) & _
    "=VBSdbGridPrev&VBSdbIndex_" & objVbsDb( "GlobalId" ) & "=" & objVbsDb( "GridStart" )-1
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationEnabledText strHref , funStrTranslate( "previousAnchorTitle" , objVbsDb ) , funStrTranslate( "Prev" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationEnabledImage strHref , funStrTranslate( "previousAnchorTitle" , objVbsDb ) , "prev.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
sub vbsDbDrawNavigationPrev( byRef objVbsDb )
'***********************************************
  if objVbsDb( "ViewNavigationPrev" ) and ( not objVbsDb( "SqlRecordSet" ).eof ) then
  	' the developer wants to show the Prev button
  	if funBolFirstPage( objVbsDb ) then
  	  ' current page is the first one
  	  vbsDbDrawNavigationPrevFirstPage objVbsDb
  	else
  	  ' current page is not the first one
  	  vbsDbDrawNavigationPrevNotFirstPage objVbsDb			
  	end if
  end if
end sub

'***********************************************
sub vbsDbDrawNavigationNextLastPage( byRef objVbsDb )
'***********************************************
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationDisabledText funStrTranslate( "next" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationDisabledImage funStrTranslate( "nextAnchorTitleAlreadyLast" , objVbsDb ) , "disabledNext.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
sub vbsDbDrawNavigationNextNotLastPage( byRef objVbsDb )
'***********************************************
  dim strHref
  strHref = request.serverVariables( "url" ) & "?VBSdbClickClass_" & _
    objVbsDb( "GlobalId" ) & "=VBSdbGridNext&VBSdbIndex_" & _
    objVbsDb( "GlobalId" ) & "=" & objVbsDb( "GridStart" )-1
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationEnabledText strHref , funStrTranslate( "nextAnchorTitle" , objVbsDb ) , funStrTranslate( "next" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationEnabledImage strHref , funStrTranslate( "nextAnchorTitle" , objVbsDb ) , "next.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
function funBolLastPage( byRef objVbsDb )
'***********************************************
  funBolLastPage = ( objVbsDb( "SqlRecordSet" ).absolutePage = objVbsDb( "SqlRecordSet" ).pageCount )
end function

'***********************************************
sub vbsDbDrawNavigationNext( byRef objVbsDb )
'***********************************************
  if objVbsDb( "ViewNavigationNext" ) and ( not objVbsDb( "SqlRecordSet" ).eof ) then
  	' the developer wants to show the First button
  	if funBolLastPage( objVbsDb ) then
  		' current page is the last one
  		vbsDbDrawNavigationNextLastPage objVbsDb
  	else
  		' current page is not the last one
  		vbsDbDrawNavigationNextNotLastPage objVbsDb			
  	end if
  end if
end sub

'***********************************************
sub vbsDbDrawNavigationLastLastPage( byRef objVbsDb )
'***********************************************
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationDisabledText funStrTranslate( "last" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationDisabledImage funStrTranslate( "lastAnchorTitleAlreadyLast" , objVbsDb ) , "disabledLast.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
sub vbsDbDrawNavigationLastNotLastPage( byRef objVbsDb )
'***********************************************
  dim strHref
  strHref = request.serverVariables( "url" ) & "?VBSdbClickClass_" & _
    objVbsDb( "GlobalId" ) & "=VBSdbGridLast&VBSdbIndex_" & _
    objVbsDb( "GlobalId" ) & "=" & objVbsDb( "GridStart" )-1
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationEnabledText strHref , funStrTranslate( "lastAnchorTitle" , objVbsDb ) , funStrTranslate( "last" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationEnabledImage strHref , funStrTranslate( "lastAnchorTitle" , objVbsDb ) , "last.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
sub vbsDbDrawNavigationLast( byRef objVbsDb )
'***********************************************
  if objVbsDb( "ViewNavigationLast" ) and ( not objVbsDb( "SqlRecordSet" ).eof ) then
  	' the developer wants to show the Last button
  	if funBolLastPage( objVbsDb ) then
  	  ' current page is the last one
  	  vbsDbDrawNavigationLastLastPage objVbsDb
  	else
  	  ' current page is not the last one
  	  vbsDbDrawNavigationLastNotLastPage objVbsDb			
  	end if
  end if
end sub

'***********************************************
sub vbsDbDrawNavigationRemoveFilterActually( objVbsDb )
'***********************************************
  dim strHref
  strHref = request.serverVariables( "url" ) & _
    "?VBSdbClickClass_" & objVbsDb( "GlobalId" ) & "=VBSdbRemoveFilter"
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationEnabledText strHref , funStrTranslate( "removeFilterAnchorTitle" , objVbsDb ) , funStrTranslate( "removeFilter" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationEnabledImage strHref , funStrTranslate( "removeFilterAnchorTitle" , objVbsDb ) , "removeFilter.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
sub vbsDbDrawNavigationRemoveFilter( objVbsDb )
'***********************************************
  if objVbsDb( "ViewNavigationRemoveFilter" ) and ( objVbsDb( "FilterWhereClause" ) <> "" ) then
  	' the developer wants to show the RemoveFilter button, and a filter is currently applied
  	vbsDbDrawNavigationRemoveFilterActually objVbsDb
  end if
end sub

'***********************************************
sub vbsDbDrawNavigationSearchActually( byRef objVbsDb )
'***********************************************
  dim strHref
  strHref = request.serverVariables( "url" ) & _
    "?VBSdbClickClass_" & objVbsDb( "GlobalId" ) & "=VBSdbFilter"
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationEnabledText strHref , funStrTranslate( "searchAnchorTitle" , objVbsDb ) , funStrTranslate( "search" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationEnabledImage strHref , funStrTranslate( "searchAnchorTitle" , objVbsDb ) , "search.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
sub vbsDbDrawNavigationSearch( objVbsDb )
'***********************************************
  if objVbsDb( "ViewNavigationSearch" ) and ( ( not objVbsDb( "SqlRecordSet" ).eof ) or ( objVbsDb( "FilterWhereClause" ) <> "" ) ) then
  	' the developer wants to show the Search button, and it is not the case the query is empty and the user is not applying any filter
  	vbsDbDrawNavigationSearchActually objVbsDb
  end if
end sub

'***********************************************
sub vbsDbDrawNavigationAddActually( byRef objVbsDb )
'***********************************************
  dim strHref
  strHref = request.serverVariables( "url" ) & "?VBSdbClickClass_" & objVbsDb( "GlobalId" ) & "=VBSdbEditAdd"
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationEnabledText strHref , funStrTranslate( "addAnchorTitle" , objVbsDb ) , funStrTranslate( "add" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationEnabledImage strHref , funStrTranslate( "addAnchorTitle" , objVbsDb ) , "add.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
sub vbsDbDrawNavigationAdd( objVbsDb )
'***********************************************
  if objVbsDb( "ViewNavigationAdd" ) and ( ( not objVbsDb( "SqlRecordSet" ).eof ) or ( objVbsDb( "FilterWhereClause" ) = "" ) ) then
  	' the developer wants to show the Add button, and it is not the case the query is empty and the user is applying some filter
  	vbsDbDrawNavigationAddActually objVbsDb
  end if
end sub

'***********************************************
function funStrVbsDbEditWhereFieldValue( strFieldName , objVbsDb )
'***********************************************
  if not objVbsDb( "SqlRecordSet" ).eof then
    ' the query returned at least one record
    on error resume next
	funStrVbsDbEditWhereFieldValue = funVarSqlRecordSetFieldValue( strFieldName , objVbsDb )
	if err.number <> 0 then
      drawError strFieldName & " is defined as a key field, but it is not a field in the " &_
        "Sql property. Check these properties:<div align=left><ul><li><b>EditKeyFields</b></li><li><b>Sql</b></li></ul></div>" , objVbsDb
	end if
  end if
end function

'***********************************************
function funStrVbsDbEditWhereUrlEncoded( objVbsDb )
'***********************************************
  funStrVbsDbEditWhereUrlEncoded = server.urlEncode( funStrVbsDbEditWhere( objVbsDb ) )
end function

'***********************************************
sub vbsDbDrawNavigationUpdateActually( byRef objVbsDb )
'***********************************************
  dim strHref
  strHref = request.serverVariables( "url" ) & "?VBSdbClickClass_" & objVbsDb( "GlobalId" ) & _
    "=VBSdbEditUpdate&VBSdbEditWhere_" & objVbsDb( "GlobalId" ) & "=" & _
    funStrVbsDbEditWhereUrlEncoded( objVbsDb )
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationEnabledText strHref , funStrTranslate( "updateAnchorTitle" , objVbsDb ) , funStrTranslate( "update" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationEnabledImage strHref , funStrTranslate( "updateAnchorTitle" , objVbsDb ) , "update.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
sub vbsDbDrawNavigationUpdate( byRef objVbsDb )
'***********************************************
  if objVbsDb( "ViewNavigationUpdate" ) and ( not objVbsDb( "SqlRecordSet" ).eof ) then
  	' the developer wants to show the Update button
  	vbsDbDrawNavigationUpdateActually objVbsDb
  end if
end sub

'***********************************************
sub vbsDbDrawNavigationDeleteActually( byRef objVbsDb )
'***********************************************
  dim strHref
  strHref = request.serverVariables( "url" ) & "?VBSdbClickClass_" & objVbsDb( "GlobalId" ) & _
    "=VBSdbEditDelete&VBSdbEditWhere_" & objVbsDb( "GlobalId" ) & "=" & _
    funStrVbsDbEditWhereUrlEncoded( objVbsDb )
  vbsDbDrawBeginNavigationItem objVbsDb
  if objVbsDb( "GlobalImageDir" ) = "" then
    ' the developer doesn't want to use images for buttons
    vbsDbDrawNavigationEnabledText strHref , funStrTranslate( "deleteAnchorTitle" , objVbsDb ) , funStrTranslate( "delete" , objVbsDb ) , objVbsDb
  else
    ' the developer wants to use images for buttons
    vbsDbDrawNavigationEnabledImage strHref , funStrTranslate( "deleteAnchorTitle" , objVbsDb ) , "delete.gif" , objVbsDb
  end if
  vbsDbDrawEndNavigationItem
end sub

'***********************************************
sub vbsDbDrawNavigationDelete( byRef objVbsDb )
'***********************************************
  if objVbsDb( "ViewNavigationDelete" ) and ( not objVbsDb( "SqlRecordSet" ).eof ) then
  	' the developer wants to show the Delete button
  	vbsDbDrawNavigationDeleteActually objVbsDb
  end if
end sub

'***********************************************
sub vbsDbDrawNavigationButtons( byRef objVbsDb )
'***********************************************
  vbsDbDrawNavigationFirst objVbsDb
  vbsDbDrawNavigationPrev objVbsDb
  vbsDbDrawNavigationNext objVbsDb
  vbsDbDrawNavigationLast objVbsDb
  vbsDbDrawNavigationRemoveFilter objVbsDb
  vbsDbDrawNavigationSearch objVbsDb
  vbsDbDrawNavigationAdd objVbsDb
  vbsDbDrawNavigationUpdate objVbsDb
  vbsDbDrawNavigationDelete objVbsDb
end sub

'***********************************************
sub vbsDbDrawNavigationTable( byRef objVbsDb )
'***********************************************
%>
  <center>
  	<table cellspacing="2" cellpadding="1" border="0">
  	  <tr>
<%
  	  	vbsDbDrawNavigationButtons objVbsDb
%>
  	  </tr>
  	</table>
  </center>
<%
end sub

'***********************************************
sub vbsDbDrawNavigation( byRef objVbsDb )
'***********************************************
  if objVbsDb( "SqlRecordSet" ).recordCount > 0 then
    ' the query returned at least one record
	  objVbsDb( "SqlRecordSet" ).absolutePosition = objVbsDb( "FormAbsolutePosition" )
  end if
  vbsDbDrawNavigationTable objVbsDb
end sub

'***********************************************
'***********************************************
sub vbsDbDrawNavBar( strPosition , byRef objVbsDb )
'***********************************************
'***********************************************
  if ( objVbsDb( "ViewNavigationPosition" ) = "BOTH" ) or _
     ( objVbsDb( "ViewNavigationPosition" ) = uCase( strPosition ) ) then
    ' the navigation bar has to be written in this position
    vbsDbDrawNavigation objVbsDb
  end if
end sub
%>