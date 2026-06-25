<%
'***********************************************
sub vbsDbDrawCSSLink( objVbsDb )
'***********************************************
%>
  <link rel="stylesheet" type="text/css" href="<%=objVbsDb( "GlobalCSS" )%>">
<%
end sub

'***********************************************
function funStrGridEvenRowBgColor( objVbsDb )
'***********************************************
  if objVbsDb( "GridHorizontalStripeBGColor" ) <> "" then
  	' objVbsDb( "GridHorizontalStripeBGColor" ) is defined
  	funStrGridEvenRowBgColor = objVbsDb( "GridHorizontalStripeBGColor" )
  else
  	' objVbsDb( "GridHorizontalStripeBGColor" ) is not defined
  	funStrGridEvenRowBgColor = objVbsDb( "GlobalTableBGColor" )
  end if
end function

'***********************************************
sub vbsDbDrawCSSInside( objVbsDb )
'***********************************************
%>
  <STYLE type="text/css">
    <!--
/******************************************************************************/
/*                            Grid related attributes                         */
/******************************************************************************/
#GridBorders
{
    BACKGROUND-COLOR: #d3d3d3;
}
#Grid
{
    /*WIDTH: 200px; uncomment this to fix the Grid to a constant width */
}
.GridEditColumn
{
    BACKGROUND-COLOR: <%=objVbsDb( "GlobalHeaderBGColor" )%>;
    COLOR: <%=objVbsDb( "GlobalHeaderFGColor" )%>;
}
.GridIndex
{
    BACKGROUND-COLOR: <%=objVbsDb( "GlobalHeaderBGColor" )%>;
    COLOR: <%=objVbsDb( "GlobalHeaderFGColor" )%>;
}
TH.GridHeader
{
    COLOR: <%=objVbsDb( "GlobalHeaderFGColor" )%>;
    BACKGROUND-COLOR: <%=objVbsDb( "GlobalHeaderBGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    FONT-WEIGHT: 600;
}
A.GridHeader
{
    COLOR: <%=objVbsDb( "GlobalHeaderFGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    FONT-WEIGHT: 600;
    TEXT-DECORATION: underline;
}
A.GridHeader:HOVER
{
    COLOR: red;
}
A.GridHeader:VISITED
{
}
A.GridHeader:ACTIVE
{
}
.GridUnselectedIndex
{
    COLOR: <%=objVbsDb( "GridUnselectedIndexFGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 11px;
    FONT-WEIGHT: 600;
    TEXT-ALIGN: center;
    TEXT-DECORATION: underline;
}
.GridSelectedIndex
{
    COLOR: <%=objVbsDb( "GridSelectedIndexFGColor" )%>;
    BACKGROUND-COLOR: <%=objVbsDb( "GridSelectedIndexBGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 11px;
    FONT-WEIGHT: 600;
    TEXT-ALIGN: center;
    TEXT-DECORATION: underline;
}
.GridOddRowUnselected
{
    BACKGROUND-COLOR: <%=objVbsDb( "GlobalTableBGColor" )%>;
    COLOR: <%=objVbsDb( "GlobalTableFGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    FONT-WEIGHT: normal;
}
.GridOddRowSelected
{
    BACKGROUND-COLOR: <%=objVbsDb( "GlobalTableBGColor" )%>;
    COLOR: <%=objVbsDb( "GlobalTableFGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    FONT-WEIGHT: normal;
}
.GridEvenRowUnselected
{
    BACKGROUND-COLOR: <%=funStrGridEvenRowBgColor( objVbsDb )%>;
    COLOR: <%=objVbsDb( "GlobalTableFGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    FONT-WEIGHT: normal;
}
.GridEvenRowSelected
{
    BACKGROUND-COLOR: <%=funStrGridEvenRowBgColor( objVbsDb )%>;
    COLOR: <%=objVbsDb( "GlobalTableFGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    FONT-WEIGHT: normal;
}
#GridTemplate
{
    BORDER-COLLAPSE: collapse;
}

/******************************************************************************/
/*                          Form related attributes                           */
/******************************************************************************/
#FormBorders
{
    BACKGROUND-COLOR: #d3d3d3;
}
TABLE#Form
{
    /*WIDTH: 300px; uncomment this to fix the Form to a constant width */
}
TR.Form
{
}
.FormLeftColumn
{
    BACKGROUND-COLOR: <%=objVbsDb( "GlobalHeaderBGColor" )%>;
    COLOR: <%=objVbsDb( "GlobalHeaderFGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    FONT-WEIGHT: 600;
    TEXT-ALIGN: right;
    /* WIDTH: 100px; uncomment this to fix the Form left column to a constant width */
}
.FormRightColumn
{
    BACKGROUND-COLOR: <%=objVbsDb( "GlobalTableBGColor" )%>;
    COLOR: <%=objVbsDb( "GlobalTableFGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    FONT-WEIGHT: normal;
    TEXT-ALIGN: left;
}

/******************************************************************************/
/*                          View related attributes                           */
/******************************************************************************/
TD.ViewNavigation
{
    BACKGROUND-COLOR: <%=objVbsDb( "ViewNavigationBGColor" )%>;
    COLOR: <%=objVbsDb( "ViewNavigationDisabledFGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 11px;
    FONT-WEIGHT: normal;
    TEXT-ALIGN: center;
    VERTICAL-ALIGN: center;
}
A.ViewNavigation
{
    COLOR: <%=objVbsDb( "ViewNavigationFGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 11px;
    FONT-WEIGHT: normal;
}
#ViewPositionBorders
{
    BACKGROUND-COLOR: #d3d3d3;
}
#ViewPosition
{
    BACKGROUND-COLOR: #ffffff;
    COLOR: #000000;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 10px;
    FONT-WEIGHT: normal;
}
.ViewMessage
{
    BACKGROUND-COLOR: #ffffff;
    COLOR: #000000;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    TEXT-ALIGN: center;
    WIDTH: 100%;
}

/******************************************************************************/
/*                          Input related attributes                          */
/******************************************************************************/
#InputTitle
{
    COLOR: #000000;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 16px;
    FONT-WEIGHT: bold;
}
#InputBorders
{
    BACKGROUND-COLOR: #d3d3d3;
}
.InputHeaders
{
    BACKGROUND-COLOR: <%=objVbsDb( "GlobalHeaderBGColor" )%>;
    COLOR: <%=objVbsDb( "GlobalHeaderFGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    FONT-WEIGHT: 600;
    TEXT-ALIGN: center;
}
#InputForm
{
    /*WIDTH: 200px; uncomment this to fix the Input Form to a constant width */
}
.InputLeftColumn
{
    BACKGROUND-COLOR: <%=objVbsDb( "GlobalTableBGColor" )%>;
    COLOR: <%=objVbsDb( "GlobalHeaderFGColor" )%>;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    FONT-WEIGHT: 600;
    TEXT-ALIGN: right;
}
.InputRightColumn
{
    BACKGROUND-COLOR: <%=objVbsDb( "GlobalTableBGColor" )%>;
    COLOR: #000000;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    FONT-WEIGHT: normal;
    TEXT-ALIGN: left;
}
.InputCancel
{
    BACKGROUND-COLOR: <%=objVbsDb( "GlobalTableBGColor" )%>;
    COLOR: #000000;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    FONT-SIZE: 12px;
    FONT-WEIGHT: normal;
    TEXT-ALIGN: center;
}
.InputCheckBoxForIE
{
    MARGIN-LEFT: -4px;
}
.InputCheckBoxForNonIE
{
    MARGIN-LEFT: 0px;
}
#InputSubmitReset
{
    BACKGROUND-COLOR: <%=objVbsDb( "GlobalTableBGColor" )%>;
    COLOR: #000000;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    TEXT-ALIGN: center;
}
   -->
  </STYLE>
<%
end sub

'***********************************************
'***********************************************
sub vbsDbDrawCSS( objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "GlobalCSS" ) <> "" then
    vbsDbDrawCSSLink objVbsDb
  else
    vbsDbDrawCSSInside objVbsDb
  end if
end sub
%>