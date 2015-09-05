<%
class Person

private isConstructed
private name

private sub Class_Initialize
isConstructed = false
name = null
end sub
public default function construct(pName)
name = pName
set construct = me
isConstructed = true
end function

function getName()
if (not isConstructed) then
call err.raise(60000, "ObjectNotConstructedException", "Person is not constructed")
end if
getName = name
end function

function setName(pName)
if (not isConstructed) then
call err.raise(60000, "ObjectNotConstructedException", "Person is not constructed")
end if
name = pName
end function

end class

'exemplo'
set batman = (new Person)("Bruce Wayne")
response.write batman.getName() 
%>