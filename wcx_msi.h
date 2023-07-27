/*****************************************************************************/
/* wcx_msi.h                              Copyright (c) Ladislav Zezula 2023 */
/*---------------------------------------------------------------------------*/
/* Main header file for the wcx_msi plugin                                   */
/*---------------------------------------------------------------------------*/
/*   Date    Ver   Who  Comment                                              */
/* --------  ----  ---  -------                                              */
/* 22.07.23  1.00  Lad  Created                                              */
/*****************************************************************************/

#ifndef __WCX_MSI_H__
#define __WCX_MSI_H__

#pragma warning (disable: 4091)                 // 4091: 'typedef ': ignored on left of 'tagDTI_ADTIWUI' when no variable is declared
#include <tchar.h>
#include <stdio.h>
#include <string.h>
#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <strsafe.h>

#include <string>
#include <vector>

#include "utils.h"                              // Utility functions
#include "TMsi.h"                               // MSI classes

//-----------------------------------------------------------------------------
// Defines

/* Error codes returned to calling application */
#define E_END_ARCHIVE     10                    // No more files in archive
#define E_NO_MEMORY       11                    // Not enough memory
#define E_BAD_DATA        12                    // Data is bad
#define E_BAD_ARCHIVE     13                    // CRC error in archive data
#define E_UNKNOWN_FORMAT  14                    // Archive format unknown
#define E_EOPEN           15                    // Cannot open existing file
#define E_ECREATE         16                    // Cannot create file
#define E_ECLOSE          17                    // Error closing file
#define E_EREAD           18                    // Error reading from file
#define E_EWRITE          19                    // Error writing to file
#define E_SMALL_BUF       20                    // Buffer too small
#define E_EABORTED        21                    // Function aborted by user
#define E_NO_FILES        22                    // No files found
#define E_TOO_MANY_FILES  23                    // Too many files to pack
#define E_NOT_SUPPORTED   24                    // Function not supported

/* flags for ProcessFile */
enum PROCESS_FILE_OPERATION
{
    PK_SKIP = 0,                               // Skip this file
    PK_TEST,
    PK_EXTRACT
};

/* Flags passed through PFN_CHANGE_VOL */
#define PK_VOL_ASK          0                   // Ask user for location of next volume
#define PK_VOL_NOTIFY       1                   // Notify app that next volume will be unpacked

/* Flags for packing */

/* For PackFiles */
#define PK_PACK_MOVE_FILES  1                   // Delete original after packing
#define PK_PACK_SAVE_PATHS  2                   // Save path names of files

/* Returned by GetPackCaps */
#define PK_CAPS_NEW         1                   // Can create new archives
#define PK_CAPS_MODIFY      2                   // Can modify existing archives
#define PK_CAPS_MULTIPLE    4                   // Archive can contain multiple files
#define PK_CAPS_DELETE      8                   // Can delete files
#define PK_CAPS_OPTIONS    16                   // Has options dialog
#define PK_CAPS_MEMPACK    32                   // Supports packing in memory
#define PK_CAPS_BY_CONTENT 64                   // Detect archive type by content
#define PK_CAPS_SEARCHTEXT 128                  // Allow searching for text in archives created with this plugin}
#define PK_CAPS_HIDE       256                  // Show as normal files (hide packer icon), open with Ctrl+PgDn, not Enter

/* Flags for packing in memory */
#define MEM_OPTIONS_WANTHEADERS 1               // Return archive headers with packed data

#define MEMPACK_OK          0                   // Function call finished OK, but there is more data
#define MEMPACK_DONE        1                   // Function call finished OK, there is no more data

#define INVALID_SIZE_T          (size_t)(-1)

//-----------------------------------------------------------------------------
// Definitions of callback functions

// Ask to swap disk for multi-volume archive
typedef int (WINAPI * PFN_CHANGE_VOLUMEA)(LPCSTR szArcName, int nMode);
typedef int (WINAPI * PFN_CHANGE_VOLUMEW)(LPCWSTR szArcName, int nMode);

// Notify that data is processed - used for progress dialog
typedef int (WINAPI * PFN_PROCESS_DATAA)(LPCSTR szFileName, int nSize);
typedef int (WINAPI * PFN_PROCESS_DATAW)(LPCWSTR szFileName, int nSize);

//-----------------------------------------------------------------------------
// Structures

struct DOS_FTIME
{
    unsigned ft_tsec  : 5;                  // Two second interval
    unsigned ft_min   : 6;                  // Minutes
    unsigned ft_hour  : 5;                  // Hours
    unsigned ft_day   : 5;                  // Days
    unsigned ft_month : 4;                  // Months
    unsigned ft_year  : 7;                  // Year
};

struct THeaderData
{
    char      ArcName[260];
    char      FileName[260];
    DWORD     Flags;
    DWORD     PackSize;
    DWORD     UnpSize;
    DWORD     HostOS;
    DWORD     FileCRC;
    DOS_FTIME FileTime;
    DWORD     UnpVer;
    DWORD     Method;
    DWORD     FileAttr;
    char    * CmtBuf;
    DWORD     CmtBufSize;
    DWORD     CmtSize;
    DWORD     CmtState;
};

#pragma pack(push, 4)
struct THeaderDataEx
{
    char      ArcName[1024];
    char      FileName[1024];
    DWORD     Flags;
    ULONGLONG PackSize;
    ULONGLONG UnpSize;
    DWORD     HostOS;
    DWORD     FileCRC;
    DOS_FTIME FileTime;
    DWORD     UnpVer;
    DWORD     Method;
    DWORD     FileAttr;
    char    * CmtBuf;
    DWORD     CmtBufSize;
    DWORD     CmtSize;
    DWORD     CmtState;
    char Reserved[1024];
};

struct THeaderDataExW
{
    WCHAR ArcName[1024];
    WCHAR FileName[1024];
    DWORD Flags;
    ULONGLONG PackSize;
    ULONGLONG UnpSize;
    DWORD HostOS;
    DWORD FileCRC;
    DOS_FTIME FileTime;
    DWORD UnpVer;
    DWORD Method;
    DWORD FileAttr;
    char* CmtBuf;
    DWORD CmtBufSize;
    DWORD CmtSize;
    DWORD CmtState;
    char Reserved[1024];
};
#pragma pack(pop, 4)

//-----------------------------------------------------------------------------
// Open archive information

#define PK_OM_LIST          0
#define PK_OM_EXTRACT       1

struct TOpenArchiveData
{
    union
    {
        LPCSTR szArchiveNameA;              // Archive name to open
        LPCWSTR szArchiveNameW;             // Archive name to open
    };

    int    OpenMode;                        // Open reason (See PK_OM_XXXX)
    int    OpenResult;
    char * CmtBuf;
    int    CmtBufSize;
    int    CmtSize;
    int    CmtState;
};

typedef struct
{
	int   size;
	DWORD PluginInterfaceVersionLow;
	DWORD PluginInterfaceVersionHi;
	char  DefaultIniName[MAX_PATH];
} TPackDefaultParamStruct;

//-----------------------------------------------------------------------------
// Configuration structures

struct TConfiguration
{
    DWORD dwDummy;                          // Dummy
};

//-----------------------------------------------------------------------------
// Global variables

extern TConfiguration g_cfg;                // Plugin configuration
extern HINSTANCE g_hInst;                   // Our DLL instance
extern HANDLE g_hHeap;                      // Process heap
extern TCHAR g_szIniFile[MAX_PATH];         // Packer INI file

//-----------------------------------------------------------------------------
// Global functions

// Configuration functions
int SetDefaultConfiguration();
int LoadConfiguration();
int SaveConfiguration();

//-----------------------------------------------------------------------------
// Dialogs

INT_PTR SettingsDialog(HWND hParent);

#endif // __WCX_MSI_H__
