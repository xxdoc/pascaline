// Console.inc

const STD_INPUT_HANDLE = -10;
const STD_OUTPUT_HANDLE = -11;
const STD_ERROR_HANDLE = -12;

const FOREGROUND_BLUE = $1;
const FOREGROUND_GREEN = $2;
const FOREGROUND_RED = $4;
const FOREGROUND_INTENSITY = $8;
const BACKGROUND_BLUE = $10;
const BACKGROUND_GREEN = $20;
const BACKGROUND_RED = $40;
const BACKGROUND_INTENSITY = $80;

const ENABLE_LINE_INPUT = $2;
const ENABLE_ECHO_INPUT = $4;
const ENABLE_MOUSE_INPUT = $10;
const ENABLE_PROCESSED_INPUT = $1;
const ENABLE_WINDOW_INPUT = $8;

const ENABLE_PROCESSED_OUTPUT = $1;
const ENABLE_WRAP_AT_EOL_OUTPUT = $2;

dword hConsoleIn;
dword hConsoleOut;
dword hConsoleErr;

type COORDINATE {
	word x;
	word y;
}

frame Console.Init(string sTitle[256]);
	AllocConsole();
	SetConsoleTitle(sTitle);
	hConsoleIn = GetStdHandle(STD_INPUT_HANDLE);
	hConsoleOut = GetStdHandle(STD_OUTPUT_HANDLE);
	hConsoleErr = GetStdHandle(STD_ERROR_HANDLE);
end;

frame Console.Position(dword X, dword Y);
	COORDINATE coord;
	coord.x = X;
	coord.y = Y;
	SetConsoleCursorPosition(hConsoleOut,^coord);	
end;

frame Console.Color(dword dwColor);
	SetConsoleTextAttribute(hConsoleOut,dwColor);
end;

frame Console.Write(string sOut[256]);
	dword BW;
	WriteConsole(hConsoleOut,sOut,lstrlen(sOut),BW,0);
end;

frame Console.Read() as string;
	local dword BR;
	local string sUserInput[256];
	ReadConsole(hConsoleIn,sUserInput,lstrlen(@sUserInput),BR,0);
	lstrcpyn(sUserInput,sUserInput,lstrlen(sUserInput)-1);
	return(sUserInput);
end;

frame Console.Free();
	FreeConsole();
end;