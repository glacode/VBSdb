<%
'***********************************************
sub vbsDbDrawViewScreenGridMode( byRef objVbsDb )
'***********************************************
%>
  <table>
  	<td valign="top">
<%
  	  vbsDbDrawGrid objVbsDb
%>
  </table>
<%
end sub

'***********************************************
sub vbsDbDrawViewScreenFormMode( byRef objVbsDb )
'***********************************************
%>
  <table>
  	<td valign="top">
<%
  	  vbsDbDrawForm objVbsDb
%>
  </table>
<%
end sub

'***********************************************
sub vbsDbDrawViewScreenBothHorizontalMode( byRef objVbsDb )
'***********************************************
%>
  <table>
    <tbody>
      <tr>
  	    <td valign="top">
<%
      	  vbsDbDrawGrid objVbsDb
%>
        </td>
  	    <td valign="top">
<%
  	      vbsDbDrawForm objVbsDb
%>
        </td>
      </tr>
    </tbody>
  </table>
<%
end sub

'***********************************************
sub vbsDbDrawViewScreenBothVerticalMode( byRef objVbsDb )
'***********************************************
%>
  <table>
    <tr>
    	<td valign="top" align="center">
<%
    	  vbsDbDrawGrid objVbsDb
%>
      </td>
    </tr>
    <tr>
    	<td valign="top" align="center">
<%
    	  vbsDbDrawForm objVbsDb
%>
      </td>
    </tr>
  </table>
<%
end sub

'***********************************************
sub vbsDbDrawViewScreenSelectModeActually( byRef objVbsDb )
'***********************************************
  select case objVbsDb( "ViewMode" )
    case "GRID"
      vbsDbDrawViewScreenGridMode objVbsDb
    case "FORM"
      vbsDbDrawViewScreenFormMode objVbsDb
    case "GRID-FORM"
      vbsDbDrawViewScreenBothHorizontalMode objVbsDb
    case "GRID|FORM"
      vbsDbDrawViewScreenBothVerticalMode objVbsDb
  end select
end sub

'***********************************************
sub vbsDbDrawViewScreenSelectMode( byRef objVbsDb )
'***********************************************
  if ( objVbsDb( "SqlRecordSet" ).recordCount = 0 ) then
    ' record set is empty
    vbsDbDrawViewEmptyMessage objVbsDb
  else
    ' record set is not empty
    vbsDbDrawViewScreenSelectModeActually objVbsDb
  end if
end sub

'***********************************************
sub vbsDbDrawViewScreen( byRef objVbsDb )
'***********************************************
%>
  <center>
<%
    vbsDbDrawNavBar "TOP" , objVbsDb
    vbsDbDrawViewScreenSelectMode objVbsDb
    vbsDbDrawNavBar "BOTTOM" , objVbsDb
    vbsDbDrawViewPosition objVbsDb
%>			
  </center>
<%
end sub
%>

	
