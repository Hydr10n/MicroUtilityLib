# MicroUtilityLib

## In Source File "MicroUtilityLib.c"
### Functions
```C
BOOL WINAPI DownloadFileFromInternetW(LPCWSTR lpcwUrl, LPCWSTR lpcwNewFileName, BOOL bFailIfFileExists);

//This function attempts to find all files/folders which can be accessible on a disk, but might fail in some cases.
//The description of LPFIND_DISK_DATA_ROUTINEW can be found in Header File "MicroUtilityLib.h"
BOOL WINAPI FindDiskDataW(LPCWSTR lpcwPathName, LPVOID lpData, LPFIND_DISK_DATA_ROUTINEW lpFindDiskData_Routine);

//This function attempts to get product version of a Windows executable file, but might fail in some cases.
BOOL WINAPI GetFileProductVersionW(LPCWSTR lpcwFileName, LPWSTR lpwFileProductVersionBuffer, DWORD cchFileProductVersionBuffer);

//Check if current process is run as Administrator
BOOL WINAPI IsRunAsAdministrator(VOID);

//This function uses Windows API StrCmpLogicalW which works better than string compare functions like strcmp to sort strings.
//Strings are sorted in ascending.
BOOL WINAPI SortStringsLogicalW(LPWSTR *lpwStrings, DWORD dwNumberOfStrings);

//This function attempts to find a specified string in strings and return the index.
//If the string is not found, this funtcion will return (DWORD)-1.
//Strings must be sorted in ascending because this function uses binary search algorithm. This function is not case-sensitive.
DWORD WINAPI FindStringInSortedStringsLogicalW(LPCWSTR lpcwStringToFind, LPWSTR *lpwStrings, DWORD dwLowerBound, DWORD dwUpperBound);

//This functions attempts to format a string from a system error code, but might fail in some cases.
//The description of LPGET_SYSTEM_ERROR_MESSAGE_ROUTINEW can be found in Header File "MicroUtilityLib.h"
DWORD WINAPI GetSystemErrorMessageW(DWORD dwErrorCode, LPGET_SYSTEM_ERROR_MESSAGE_ROUTINEW lpGetSystemErrorMessageRoutine);

DWORD WINAPI GetFileCRC32W(LPCWSTR lpcwFileName);

INT WINAPI GetCurrentDPI(AXIS Axis);

//This function attempts to get the base address of a process, but might fail if no access rights.
PVOID WINAPI GetProcessBaseAddress(DWORD dwProcessID);

//This function attempts to delete the executable file of current process, but might fail in some cases.
VOID WINAPI DeleteSelf(VOID);
```

## In Header File "MicroUtilityLib.h"
### Macros
```C
//This macro converts the case of a character into lower
ToLowerCaseW(ch);
ToLowerCaseA(ch);

//This macro converts the case of a character into upper
ToUpperCaseW(ch);
ToUpperCaseA(ch);

//This macro combines two WORD integers into one DWORD integer
MAKEDWORD(HIGH_WORD, LOW_WORD);

//This macro combines two DWORD integers into one DWORD64 integer
MAKEDWORD64(HIGH_DWORD, LOW_DWORD);

//This macro scales standard width/height (in pixels) to fit current DPI
SCALE(iPixels, iCurrentDPI) MulDiv(iPixels, iCurrentDPI, USER_DEFAULT_SCREEN_DPI)

//This macro creates a window which fits current DPI
DPIAware_CreateWindowExW(iCurrentDPI, dwExStyle, lpClassName, lpWindowName, dwStyle, iX, iY, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam);
```

### Data Types
```C
typedef enum _AXIS { AxisX = LOGPIXELSX, AxisY = LOGPIXELSY } AXIS;

//This is a callback function used for FindDiskDataW.
//The parameters in this function contain most information of current file/folder.
//return FALSE to stop function FindDiskDataW
typedef BOOL(CALLBACK* LPFIND_DISK_DATA_ROUTINEW)(LPWIN32_FIND_DATAW lpWin32_FindData, LPWSTR lpwPathName, LPVOID lpData);
LPFIND_DISK_DATA_ROUTINEW lpFindDiskData_Routine;

//This is a callback function used for GetSystemErrorMessageW.
//The parameter in this function contains a string of a system error message.
typedef VOID(CALLBACK* LPGET_SYSTEM_ERROR_MESSAGE_ROUTINEW)(LPCWSTR lpcwSystemErrorMessage);
```

## Minimum Supported Client
Microsoft Windows Vista [Desktop Only]
