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
#include <msiquery.h>                           // MSI.dll functions

#include <string>
#include <vector>

#include "utils.h"                              // Utility functions

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

#define MSI_MAGIC_SIGNATURE  0x434947414D49534D // "MSIMAGIC"

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
// Information about MSI database

typedef std::vector<std::tstring> MSI_STRING_LIST;

struct TMsiDatabase;

typedef enum MSI_TYPE
{
    MsiTypeUnknown = 0,
    MsiTypeInteger,
    MsiTypeString,
    MsiTypeStream,
};

struct MSI_BLOB
{
    MSI_BLOB()
    {
        pbData = NULL;
        cbData = 0;
    }

    ~MSI_BLOB()
    {
        if(pbData != NULL)
            HeapFree(g_hHeap, 0, pbData);
        pbData = NULL;
        cbData = 0;
    }

    DWORD Reserve(DWORD cbSize)
    {
        // Must not be allocated before
        assert(pbData == NULL);

        // Allocate zero-filled buffer
        if((pbData = (LPBYTE)HeapAlloc(g_hHeap, HEAP_ZERO_MEMORY, cbSize)) == NULL)
            return ERROR_NOT_ENOUGH_MEMORY;

        cbData = cbSize;
        return ERROR_SUCCESS;
    }

    LPBYTE pbData;
    DWORD cbData;
};

struct TMsiColumn
{
    TMsiColumn(LPCTSTR szName, LPCTSTR szType);

    std::tstring m_strName;                 // Name of the column as LPTSTR
    std::tstring m_strType;                 // Type of the column as LPTSTR
    MSI_TYPE m_Type;                        // Type of the column
    size_t m_Size;                          // Length of the column, if it's string
};

struct TMsiTable
{
    TMsiTable(TMsiDatabase * pMsiDb, const std::tstring & strName, MSIHANDLE hMsiView);
    ~TMsiTable();

    DWORD AddRef();
    DWORD Release();

    DWORD Load();
    DWORD LoadColumns();

    const std::vector<TMsiColumn> & Columns()   { return m_Columns; }
    MSIHANDLE MsiView()                         { return m_hMsiView; }
    LPCTSTR Name()                              { return m_strName.c_str(); }

    std::vector<TMsiColumn> m_Columns;      // List of columns
    std::tstring m_strName;                 // Table name
    TMsiDatabase * m_pMsiDb;                // Pointer to the parent database
    LIST_ENTRY m_Entry;                     // Links to other tables
    MSIHANDLE m_hMsiView;                   // MSI handle to the database view
    size_t m_nStreamColumn;                 // Index of the stream column. -1 if none
    size_t m_nNameColumn;                   // Index of the name column. -1 if none
    DWORD m_bIsStreamsTable;                // TRUE if this is the "_Streams" table
    DWORD m_dwRefs;
};

struct TMsiFile
{
    TMsiFile(TMsiTable * pMsiTable);
    ~TMsiFile();

    DWORD AddRef();
    DWORD Release();

    DWORD SetBinaryFile(TMsiDatabase * pMsiDb, MSIHANDLE hMsiRecord);
    DWORD SetCsvFile(TMsiDatabase * pMsiDb);

    DWORD LoadCsvFileData(LPDWORD PtrFileSize);
    DWORD LoadFileSize();
    DWORD LoadFileData();

    void MakeItemNameFileSafe(std::tstring & strItemName);

    const MSI_BLOB & FileData();
    DWORD FileSize();
    LPCTSTR Name();

    protected:

    friend struct TMsiDatabase;

    LIST_ENTRY m_Entry;                     // Link to other files
    TMsiTable * m_pMsiTable;                // Pointer to the database table
    TMsiFile * m_pRefFile;                  // Reference to another file
    std::tstring m_strName;                 // File name
    MSIHANDLE m_hMsiRecord;                 // Handle to the MSI record (if binary file)
    MSI_BLOB m_Data;
    DWORD m_dwFileSize;                     // Size of the file
    DWORD m_dwRefs;
};

// Our structure describing open archive
struct TMsiDatabase
{
    TMsiDatabase(MSIHANDLE hMsiDb, FILETIME & ft);

    static TMsiDatabase * FromHandle(HANDLE hHandle);

    DWORD AddRef();
    DWORD Release();
    void  CloseAllFiles();
    void  UnlockAndRelease();
    void  AssertRefCount(DWORD dwRefs);

    TMsiFile * GetNextFile();
    TMsiFile * ReleaseLastFile(TMsiFile * pMsiFile = NULL);
    TMsiFile * FindReferencedFile(TMsiTable * pMsiTable, LPCTSTR szStreamName, LPTSTR szFileName, size_t ccFileName);

    DWORD LoadTableNames();
    DWORD LoadTables();
    DWORD LoadFiles();
    DWORD LoadMultipleStreamFiles(TMsiTable * pMsiTable);
    DWORD LoadSimpleCsvFile(TMsiTable * pMsiTable);
    
    TMsiFile * IsFilePresent(LPCTSTR szFileName);
    TMsiFile * LastFile();
    const FILETIME & FileTime()         { return m_FileTime; }

    protected:

    ~TMsiDatabase();
    
    template <typename LIST_ITEM>
    void DeleteLinkedList(LIST_ENTRY & Links)
    {
        PLIST_ENTRY pHeadEntry = &Links;
        PLIST_ENTRY pListEntry;
        LIST_ITEM * pMsiItem;

        for(pListEntry = pHeadEntry->Flink; pListEntry != pHeadEntry; )
        {
            // Get the TMsiTable out of it and move the pointer
            pMsiItem = CONTAINING_RECORD(pListEntry, LIST_ITEM, m_Entry);
            pListEntry = pListEntry->Flink;

            // Unlin and free the class
            RemoveEntryList(&pMsiItem->m_Entry);
            pMsiItem->Release();
        }
    }

    CRITICAL_SECTION m_Lock;
    MSI_STRING_LIST m_TableNames;
    PLIST_ENTRY m_pFileEntry;
    TMsiFile * m_pLastFile;                 // The last file found by ReadHeaders
    LIST_ENTRY m_Tables;                    // List of tables
    LIST_ENTRY m_Files;                     // List of files
    ULONGLONG m_MagicSignature;             // MSI_MAGIC_SIGNATURE
    MSIHANDLE m_hMsiDb;
    FILETIME m_FileTime;                    // File time of the MSI archive
    DWORD m_dwTables;                       // Number of tables
    DWORD m_dwFiles;                        // Number of files
    DWORD m_dwRefs;
};

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
// MSI helper functions

bool MsiRecordGetInteger(MSIHANDLE hMsiRecord, UINT nColumn, std::tstring & strValue);
bool MsiRecordGetString(MSIHANDLE hMsiRecord, UINT nColumn, std::tstring & strValue);
bool MsiRecordGetBinary(MSIHANDLE hMsiRecord, UINT nColumn, MSI_BLOB & binValue);

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
