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
%>
