<%
'***********************************************
sub vbsDbDrawVbsDbHiddenAnchors( byRef objVbsDb )
'***********************************************
  response.write "<a href=""http://www.vbsdb.com""></a>" & _
    funStrVbsDbGlobalPagePositioning_forAnchor( objVbsDb )
end sub

'***********************************************
sub vbsDbDrawVbsDbActually( byRef objVbsDb )
'***********************************************
  select case VbsDbGetScreenType( objVbsDb )
  	case "View"
  	  vbsDbDrawViewScreen objVbsDb
  	case "Search" , "Add" , "Update" , "Delete"
  	  vbsDbDrawInputScreen objVbsDb
  end select
end sub

'***********************************************
sub vbsDbDrawVbsDb( byRef objVbsDb )
'***********************************************
  if objVbsDb( "EditErrorNumber" ) = 0 then
    ' there has been no error while updating the database
    vbsDbDrawVbsDbActually objVbsDb
  else
    ' there has been an error while updating the database
    vbsDbDrawVbsDbDrawEditDbError objVbsDb
  end if
end sub

'***********************************************
sub vbsDbDraw( byRef objVbsDb )
'***********************************************
  vbsDbDrawCSS objVbsDb
  vbsDbDrawVbsDbHiddenAnchors objVbsDb
  vbsDbDrawVbsDb objVbsDb
end sub
%>