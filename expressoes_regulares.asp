<%
function regEX(strOriginalString, strPattern, varIgnoreCase)
    ' Function matches pattern, returns true or false
    ' varIgnoreCase must be TRUE (match is case insensitive) or FALSE (match is case sensitive)
    dim objRegExp : set objRegExp = new RegExp
    with objRegExp
		.Pattern = strPattern
		.IgnoreCase = varIgnoreCase
		.Global = True
    end with
    regEX = objRegExp.test(strOriginalString)
    set objRegExp = nothing
end function
function regEX_replace(strOriginalString, strPattern, strReplacement, varIgnoreCase)
    ' Function replaces pattern with replacement
    ' varIgnoreCase must be TRUE (match is case insensitive) or FALSE (match is case sensitive)
    dim objRegExp : set objRegExp = new RegExp
    with objRegExp
        .Pattern = strPattern
        .IgnoreCase = varIgnoreCase
        .Global = True
    end with
    regEX_replace = objRegExp.replace(strOriginalString, strReplacement)
    set objRegExp = nothing
end function

if regEX(strOriginalString, "[^a-f0-9\s]", True) = True then
    response.write "String is not hexadecimal."
else
    response.write "String is hexadecimal."
end if

strOriginalString = "192.168.0.1"
strOriginalString = regEX_replace(strOriginalString, "([^0-9])([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})\.[0-9]{1,3}([^0-9])", "$1$2.$3.$4.***$5", True)

%>
