//Conversion.inc

import StrToInt alias StrToIntA lib "shlwapi.dll",1;

string nStr;

frame Str2Int(dword ptrString) as dword;
	return(StrToInt(ptrString));
end;

frame Int2Str(dword Value) as string;
	Format(nStr,"%d",Value);
	return(nStr);
end;

frame CStr(dword Value) as string;
	return(Int2Str(Value));
end;

frame CInt(dword ptrString) as dword;
	return(Str2Int(ptrString));
end;

frame LoWord(dword Value) as dword;
	return(Value & $0000FFFF);
end;

frame HiWord(dword Value) as dword;
	return(Value >> 16);
end;
