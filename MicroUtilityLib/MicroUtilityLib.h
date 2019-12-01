/*
Header File: MicroUtilityLib.h
Last Update: 2019/12/01
Minimum Supported Client: Microsoft Windows 7 [Desktop Only]

This project is hosted on https://github.com/Hydr10n/MicroUtilityLib
Copyright (C) 2018 - 2019 Programmer-Yang_Xun@outlook.com. All Rights Reserved.
*/

#pragma once

#ifdef __cplusplus
#define EXTERN extern "C"
#else
#define EXTERN extern
#endif

#pragma region Enable Win32 Visual Styles
#ifdef ENABLE_WIN32_VISUAL_STYLES
#ifdef _M_IX86
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='x86' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_IA64
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='ia64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#elif defined _M_X64
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='amd64' publicKeyToken='6595b64144ccf1df' language='*'\"")
#else
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#endif
#endif
#pragma endregion

#include <Windows.h>

#define MAX_DRIVE_NUMBER 26
#define MAX_UNICODE_PATH UNICODE_STRING_MAX_CHARS

#define ToLowerCaseW(ch) (((wchar_t)(ch) >= L'A' && (wchar_t)(ch) <= L'Z') ? ((wchar_t)(ch) | 0x20) : (wchar_t)(ch))
#define ToLowerCaseA(ch) (char)ToLowerCaseW(ch)
#define ToUpperCaseW(ch) (((wchar_t)(ch) >= L'a' && (wchar_t)(ch) <= L'z') ? ((wchar_t)(ch) & 0xdf) : (wchar_t)(ch))
#define ToUpperCaseA(ch) (char)ToUpperCaseW(ch)
#define MAKEDWORD(HIGH_WORD, LOW_WORD) ((DWORD)((((DWORD)(HIGH_WORD)) << 0x10) | (DWORD)(LOW_WORD)))
#define MAKEDWORD64(HIGH_DWORD, LOW_DWORD) ((DWORD64)((((DWORD64)(HIGH_DWORD)) << 0x20) | (DWORD64)(LOW_DWORD)))

#define ScaleX(iX, iCurrentDpi_X) MulDiv(iX, iCurrentDpi_X, USER_DEFAULT_SCREEN_DPI)
#define ScaleY(iY, iCurrentDpi_Y) MulDiv(iY, iCurrentDpi_Y, USER_DEFAULT_SCREEN_DPI)
#define DPIAware_CreateWindowExW(iCurrentDpi_X, iCurrentDpi_Y, dwExStyle, lpClassName, lpWindowName, dwStyle, iX, iY, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam) CreateWindowExW(dwExStyle, lpClassName, lpWindowName, dwStyle, ScaleX(iX, iCurrentDpi_X), ScaleY(iY, iCurrentDpi_Y), ScaleX(nWidth, iCurrentDpi_X), ScaleY(nHeight, iCurrentDpi_Y), hWndParent, hMenu, hInstance, lpParam)

typedef enum _AXIS { AxisX = LOGPIXELSX, AxisY = LOGPIXELSY } AXIS;

//LPFIND_DISK_DATA_ROUTINEW: return FALSE to stop function FindDiskDataW
typedef BOOL(CALLBACK* LPFIND_DISK_DATA_ROUTINEW)(LPWIN32_FIND_DATAW lpWin32_FindData, LPWSTR lpwPathName, LPVOID lpData);
EXTERN BOOL WINAPI FindDiskDataW(LPCWSTR lpcwPathName, LPVOID lpData, LPFIND_DISK_DATA_ROUTINEW lpFindDiskData_Routine);

typedef VOID(CALLBACK* LPGET_SYSTEM_ERROR_MESSAGE_ROUTINEW)(LPCWSTR lpcwSystemErrorMessage);
EXTERN DWORD WINAPI GetSystemErrorMessageW(DWORD dwErrorCode, LPGET_SYSTEM_ERROR_MESSAGE_ROUTINEW lpGetSystemErrorMessageRoutine);

EXTERN BOOL WINAPI DownloadFileFromInternetW(LPCWSTR lpcwUrl, LPCWSTR lpcwNewFileName, BOOL bFailIfFileExists);
EXTERN BOOL WINAPI GetFileProductVersionW(LPCWSTR lpcwFileName, LPWSTR lpwFileProductVersionBuffer, DWORD cchFileProductVersionBuffer);
EXTERN BOOL WINAPI IsRunAsAdministrator(VOID);
EXTERN BOOL WINAPI SortStringsLogicalW(LPWSTR* lpwStrings, DWORD dwNumberOfStrings);
EXTERN DWORD WINAPI FindStringInSortedStringsLogicalW(LPCWSTR lpcwStringToFind, LPWSTR* lpwStrings, DWORD dwLowerBound, DWORD dwUpperBound);
EXTERN DWORD WINAPI GetFileCRC32W(LPCWSTR lpcwFileName);
EXTERN INT WINAPI GetCurrentDPI(AXIS Axis);
EXTERN PVOID WINAPI GetProcessBaseAddress(DWORD dwProcessID);
EXTERN VOID WINAPI DeleteSelf(VOID);