//Copyright (C) 2018 Programmer-Yang_Xun@outlook.com

#define MICROUTILITYLIB_API __declspec(dllexport)
#include "MicroUtilityLib.h"
#include <Windows.h>
#include <ShellAPI.h>
#include <ShlObj.h>
#include <Shlwapi.h>
#pragma comment(lib, "Kernel32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Shlwapi.lib")
#pragma comment(lib, "User32.lib")

BOOL APIENTRY DllMain(HMODULE hModule, DWORD ul_reason_for_call, LPVOID lpReserved)
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}

BOOL FindDiskDataW(WCHAR wchDriveLetter, LPVOID lpData, LPFIND_DISK_DATA_ROUTINEW lpFindDiskData_Routine)
{
	HANDLE hHeap = HeapCreate(HEAP_NO_SERIALIZE, 0, 0);
	if (hHeap)
	{
		LPWSTR lpwNewFileName = (LPWSTR)HeapAlloc(hHeap, HEAP_NO_SERIALIZE, MAX_FILENAME * sizeof(WCHAR)),
			lpwPathName = (LPWSTR)HeapAlloc(hHeap, HEAP_NO_SERIALIZE, MAX_FILENAME * sizeof(WCHAR));
		if (lpwNewFileName == NULL || lpwPathName == NULL)
		{
			HeapDestroy(hHeap);
			return FALSE;
		}
		wnsprintfW(lpwPathName, MAX_FILENAME, L"\\\\?\\%lc:", wchDriveLetter);
		struct _DATA {
			INT iLengthOfPathName;
			HANDLE hFindFile;
			struct _DATA *Previous, *Next;
		} FindData, *pFindData = &FindData;
		FindData.Previous = NULL;
		WIN32_FIND_DATAW Win32_FindData;
		*Win32_FindData.cFileName = 0;
		Win32_FindData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY;
		while (TRUE)
		{
			if (*Win32_FindData.cFileName != L'.' || (Win32_FindData.cFileName[1] && (Win32_FindData.cFileName[1] != L'.' || Win32_FindData.cFileName[2])))
			{
				if (Win32_FindData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
				{
					if ((pFindData->Next = (struct _DATA*)HeapAlloc(hHeap, HEAP_NO_SERIALIZE | HEAP_ZERO_MEMORY, sizeof(FindData))) == NULL)
					{
						DWORD dwLastErrorCode = GetLastError();
						while (TRUE)
						{
							FindClose(pFindData->hFindFile);
							if ((pFindData = pFindData->Previous) == &FindData)
							{
								HeapDestroy(hHeap);
								SetLastError(dwLastErrorCode);
								return TRUE;
							}
						}
					}
					pFindData->Next->Previous = pFindData;
					pFindData->Next->iLengthOfPathName = wnsprintfW(lpwPathName, MAX_FILENAME, L"%ls%ls\\*.*", lpwPathName, Win32_FindData.cFileName) - 3;
					pFindData->Next->hFindFile = FindFirstFileW(lpwPathName, &Win32_FindData);
					lpwPathName[pFindData->Next->iLengthOfPathName] = 0;
					pFindData = pFindData->Next;
					if (pFindData->Previous == &FindData)
					{
						continue;
					}
					if (lpFindDiskData_Routine && lpFindDiskData_Routine(&Win32_FindData, lpwPathName, lpData) == PROGRESS_STOP)
					{
						while (TRUE)
						{
							FindClose(pFindData->hFindFile);
							if ((pFindData = pFindData->Previous) == &FindData)
							{
								HeapDestroy(hHeap);
								return TRUE;
							}
						}
					}
				}
				else
				{
					wnsprintfW(lpwNewFileName, MAX_FILENAME, L"%ls%ls", lpwPathName, Win32_FindData.cFileName);
					if (lpFindDiskData_Routine && lpFindDiskData_Routine(&Win32_FindData, lpwNewFileName, lpData) == PROGRESS_STOP)
					{
						while (TRUE)
						{
							FindClose(pFindData->hFindFile);
							if ((pFindData = pFindData->Previous) == &FindData)
							{
								HeapDestroy(hHeap);
								return TRUE;
							}
						}
					}
				}
			}
			while (!FindNextFileW(pFindData->hFindFile, &Win32_FindData))
			{
				FindClose(pFindData->hFindFile);
				if ((pFindData = pFindData->Previous) == &FindData)
				{
					HeapDestroy(hHeap);
					return TRUE;
				}
				HeapFree(hHeap, HEAP_NO_SERIALIZE, pFindData->Next);
				lpwPathName[pFindData->iLengthOfPathName] = 0;
			}
		}
	}
	else
	{
		return FALSE;
	}
}

DWORD GetFileCRC32W(LPCWSTR lpcwFileName)
{
	DWORD dwCRC32 = 0;
	HANDLE hFile = CreateFileW(lpcwFileName, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
	if (hFile != INVALID_HANDLE_VALUE)
	{
		DWORD dwCRC32_Table[256];
		for (DWORD dwIndex = 0; dwIndex < 256; dwIndex++)
		{
			dwCRC32_Table[dwIndex] = dwIndex;
			for (BYTE byteIndex = 0; byteIndex < 8; byteIndex++)
			{
				if (dwCRC32_Table[dwIndex] & 0x00000001)
				{
					dwCRC32_Table[dwIndex] = (dwCRC32_Table[dwIndex] >> 1) ^ 0xEDB88320;
				}
				else
				{
					dwCRC32_Table[dwIndex] >>= 1;
				}
			}
		}
		DWORD dwNumberOfBytesRead;
		BYTE byteCRC32Buffer[8 * 1024];
		while (ReadFile(hFile, byteCRC32Buffer, sizeof(byteCRC32Buffer), &dwNumberOfBytesRead, NULL))
		{
			dwCRC32 ^= 0xffffffff;
			DWORD dwBytes = dwNumberOfBytesRead;
			PBYTE pCRC32Buffer = byteCRC32Buffer;
			while (dwBytes--)
			{
				dwCRC32 = (dwCRC32 >> 8) ^ dwCRC32_Table[(dwCRC32 & 0xff) ^ *(pCRC32Buffer++)];
			}
			dwCRC32 ^= 0xffffffff;
			if (dwNumberOfBytesRead != sizeof(byteCRC32Buffer))
			{
				break;
			}
		}
	}
	return dwCRC32;
}

INT GetCurrentDpiX(VOID)
{
	INT iCurrentDpiX = USER_DEFAULT_SCREEN_DPI;
	HDC hDC = GetDC(NULL);
	SetProcessDPIAware();
	if (hDC)
	{
		GetDeviceCaps(hDC, LOGPIXELSX);
		iCurrentDpiX = GetDeviceCaps(hDC, LOGPIXELSX);
		ReleaseDC(NULL, hDC);
	}
	return iCurrentDpiX;
}

INT GetCurrentDpiY(VOID)
{
	INT iCurrentDpiY = USER_DEFAULT_SCREEN_DPI;
	HDC hDC = GetDC(NULL);
	SetProcessDPIAware();
	if (hDC)
	{
		GetDeviceCaps(hDC, LOGPIXELSX);
		iCurrentDpiY = GetDeviceCaps(hDC, LOGPIXELSY);
		ReleaseDC(NULL, hDC);
	}
	return iCurrentDpiY;
}

VOID DeleteSelf(VOID)
{
	WCHAR szCommandLine[MAX_PATH + 20], szProgramFileName[MAX_PATH];
	GetModuleFileNameW(NULL, szProgramFileName, _countof(szProgramFileName));
	SetPriorityClass(GetCurrentProcess(), REALTIME_PRIORITY_CLASS);
	SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_TIME_CRITICAL);
	SHChangeNotify(SHCNE_DELETE, SHCNF_PATH, szProgramFileName, NULL);
	wnsprintfW(szCommandLine, _countof(szCommandLine), L"/c del /q \"%ls\"", szProgramFileName);
	ShellExecuteW(NULL, L"open", L"cmd.exe", szCommandLine, NULL, SW_HIDE);
	ExitProcess(0);
}