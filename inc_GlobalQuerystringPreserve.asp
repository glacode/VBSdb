<%
'***********************************************
function funStrVbsDbGlobalQuerystringPreserveUrlWithoutParameters( strUrl )
'***********************************************
  dim intQuestionMarkPosition
  intQuestionMarkPosition = inStr( strUrl , "?" )
  if intQuestionMarkPosition > 0 then
    ' strUrl contains parameters
    funStrVbsDbGlobalQuerystringPreserveUrlWithoutParameters = left( strUrl , intQuestionMarkPosition - 1 )
  else
    ' strUrl does not contain parameters
    funStrVbsDbGlobalQuerystringPreserveUrlWithoutParameters = strUrl
  end if
end function

'***********************************************
function funStrVbsDbGlobalQuerystringPreserveAddParameterToParameterString( strParameterName , strParameterString )
'***********************************************
  dim strSingleParameterString
  strSingleParameterString = strParameterName & "=" & server.urlEncode( request.queryString( strParameterName ) )
  if strParameterString = "" then
    funStrVbsDbGlobalQuerystringPreserveAddParameterToParameterString = strSingleParameterString
  else
    funStrVbsDbGlobalQuerystringPreserveAddParameterToParameterString = _
      strParameterString & "&" & strSingleParameterString
  end if
end function

'***********************************************
function funStrVbsDbGlobalQuerystringPreserveParameterStringCombinationHandleParameter( strParameterName , strParameterString )
'***********************************************
  if inStr( 1 , "&" & strParameterString , "&" & strParameterName & "=" , 1 ) = 0 then
    ' strParameterName is not a parameter name in strParameterString
    funStrVbsDbGlobalQuerystringPreserveParameterStringCombinationHandleParameter = _
      funStrVbsDbGlobalQuerystringPreserveAddParameterToParameterString( strParameterName , strParameterString )
  else
    ' strParameterName is a parameter name in strParameterString
    funStrVbsDbGlobalQuerystringPreserveParameterStringCombinationHandleParameter = strParameterString
  end if
end function

'***********************************************
function funStrVbsDbGlobalQuerystringPreserveParameterStringCombination( strParameterString , byRef objVbsDb )
'***********************************************
  dim strParameterName
  funStrVbsDbGlobalQuerystringPreserveParameterStringCombination = strParameterString
  for each strParameterName in request.queryString()
    if uCase( strParameterName ) <> uCase( "VBSdbIndex_" & objVbsDb( "GlobalId" ) ) and _
       uCase( strParameterName ) <> uCase( "VBSdbGridSort_" & objVbsDb( "GlobalId" ) ) and _
       uCase( strParameterName ) <> uCase( "VBSdbEditWhere_" & objVbsDb( "GlobalId" ) ) and _
       uCase( strParameterName ) <> uCase( "VBSdbClickClass_" & objVbsDb( "GlobalId" ) ) then
      ' the current queryString parameter was not introduced by VBSdb 
      funStrVbsDbGlobalQuerystringPreserveParameterStringCombination = _
        funStrVbsDbGlobalQuerystringPreserveParameterStringCombinationHandleParameter( strParameterName , _
          funStrVbsDbGlobalQuerystringPreserveParameterStringCombination )
    end if
  next
end function

'***********************************************
function funStrVbsDbGlobalQuerystringPreserveParameterString( strUrl , byRef objVbsDb )
'***********************************************
  dim intQuestionMarkPosition
  intQuestionMarkPosition = inStr( strUrl , "?" )
  if intQuestionMarkPosition > 0 then
    ' strUrl contains parameters
    funStrVbsDbGlobalQuerystringPreserveParameterString = _
      funStrVbsDbGlobalQuerystringPreserveParameterStringCombination( _
        mid( strUrl , intQuestionMarkPosition + 1 ) , objVbsDb )
  else
    ' strUrl does not contain parameters
    funStrVbsDbGlobalQuerystringPreserveParameterString = _
      funStrVbsDbGlobalQuerystringPreserveParameterStringCombination( "" , objVbsDb )
  end if
  if funStrVbsDbGlobalQuerystringPreserveParameterString <> "" then
    funStrVbsDbGlobalQuerystringPreserveParameterString = "?" & _
      funStrVbsDbGlobalQuerystringPreserveParameterString
  end if
end function

'***********************************************
function funStrVbsDbGlobalQuerystringPreserveActually( strUrl , objVbsDb )
'***********************************************
  funStrVbsDbGlobalQuerystringPreserveActually = _
    funStrVbsDbGlobalQuerystringPreserveUrlWithoutParameters( strUrl ) & _
    funStrVbsDbGlobalQuerystringPreserveParameterString( strUrl , objVbsDb )
end function

'***********************************************
'***********************************************
function funStrVbsDbGlobalQuerystringPreserve( strUrl , objVbsDb )
'***********************************************
'***********************************************
  if objVbsDb( "GlobalQuerystringPreserve" ) then
    ' the site developer asked to preserve the querystring value
    funStrVbsDbGlobalQuerystringPreserve = funStrVbsDbGlobalQuerystringPreserveActually( strUrl , objVbsDb )
  else
    funStrVbsDbGlobalQuerystringPreserve = strUrl
  end if
  funStrVbsDbGlobalQuerystringPreserve = _
    funStrVbsDbGlobalPagePositioning_forHRef( funStrVbsDbGlobalQuerystringPreserve , objVbsDb )
end function
%>