// Strings.inc

frame Left(dword ptrString, dword Length) as string;
	lstrcpyn(ptrString,ptrString,Length+1);
	return(ptrString);
end;

frame Mid(dword ptrString, dword Start, dword Length) as string;
	lstrcpyn(ptrString,ptrString+(Start-1),Length+1);
	return(ptrString);
end;

frame Right(dword ptrString, dword Length) as string;
	ptrString = ptrString + lstrlen(ptrString) - Length;
	return(lstrcpy(ptrString,ptrString));
end;

frame UCase(dword ptrString) as string;
	return(CharUpper(ptrString));
end;

frame LCase(dword ptrString) as string;
	return(CharLower(ptrString));
end;

frame Len(dword ptrString) as string;
	return(lstrlen(@ptrString));

end;