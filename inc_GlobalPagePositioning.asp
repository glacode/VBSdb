<%
'***********************************************
'***********************************************
function funStrVbsDbGlobalPagePositioning_forAnchor( objVbsDb )
'***********************************************
'***********************************************
  funStrVbsDbGlobalPagePositioning_forAnchor = _
    "<a name=""VBSdb_PagePositioning_" & objVbsDb( "GlobalId" ) & """></a>"
end function

'***********************************************
'***********************************************
function funStrVbsDbGlobalPagePositioning_forHRef( strUrl , objVbsDb )
'***********************************************
'***********************************************
  if uCase( objVbsDb( "GlobalPagePositioning" ) ) = "VBSDB" then
    ' the site developer asks to focus the page on the VBSdb object
    funStrVbsDbGlobalPagePositioning_forHRef = strUrl & "#VBSdb_PagePositioning_" & objVbsDb( "GlobalId" )
  else
    funStrVbsDbGlobalPagePositioning_forHRef = strUrl
  end if
end function
%>