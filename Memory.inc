//Memory.inc

frame MemAlloc(dword Size);
	dword hVA;
	hVA = VirtualAlloc(0,Size,$1000,4);
	if (hVA = NULL) {
		return(FALSE); 
	} else {
		return(hVA);
	}
end;

frame MemFree(dword Handle);
	VirtualFree(Handle,0,$8000);
end;

frame Space(dword Size) as string;
	local dword tmp;
	tmp = MemAlloc(Size+1);
	return(@tmp);
end;