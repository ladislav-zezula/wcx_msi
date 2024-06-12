/*****************************************************************************/
/* TMsi.h                                 Copyright (c) Ladislav Zezula 2023 */
/*---------------------------------------------------------------------------*/
/* Common header file for MSI classes                                        */
/*---------------------------------------------------------------------------*/
/*   Date    Ver   Who  Comment                                              */
/* --------  ----  ---  -------                                              */
/* 27.07.23  1.00  Lad  Created                                              */
/*****************************************************************************/

#ifndef __TMSI_H__
#define __TMSI_H__

#include <msiquery.h>                           // MSI.dll functions

//-----------------------------------------------------------------------------
// Defines

#define MSI_MAGIC_SIGNATURE  0x434947414D49534D // "MSIMAGIC"

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

        // Make sure there's at least one byte allocated
        cbSize = max(cbSize, 1);

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

    DWORD SetSummaryFile(TMsiDatabase * pMsiDb, MSIHANDLE hMsiSummary);
    DWORD SetBinaryFile(TMsiDatabase * pMsiDb, MSIHANDLE hMsiRecord);
    DWORD SetCsvFile(TMsiDatabase * pMsiDb);

    DWORD LoadSummaryFile(LPDWORD PtrFileSize);
    DWORD LoadBinaryFile(LPDWORD PtrFileSize);
    DWORD LoadCsvFile(LPDWORD PtrFileSize);
    
    DWORD LoadFileInternal(LPDWORD PtrFileSize);
    DWORD LoadFileData();

    void MakeItemNameFileSafe(std::tstring & strItemName);

    const MSI_BLOB & FileData();
    DWORD FileSize();
    LPCTSTR Name();

    enum MSI_FT
    {
        MsiFileNone = 0,            // Unknown / not specified
        MsiFileSummary,             // A summary file
        MsiFileBinary,              // A binary file
        MsiFileTable                // A MSI table file
    };

    protected:

    DWORD SetUniqueFileName(TMsiDatabase * pMsiDb, LPCTSTR szFolderName, LPCTSTR szBaseName, LPCTSTR szExtension);

    friend struct TMsiDatabase;

    LIST_ENTRY m_Entry;                     // Link to other files
    TMsiTable * m_pMsiTable;                // Pointer to the database table
    TMsiFile * m_pRefFile;                  // Reference to another file
    std::tstring m_strName;                 // File name
    MSIHANDLE m_hMsiHandle;                 // Handle to the MSI record (if binary file) or MSI summary (if summary file)
    MSI_BLOB m_Data;
    MSI_FT m_FileType;
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

    TMsiFile * GetNextFile();
    TMsiFile * ReleaseLastFile(TMsiFile * pMsiFile = NULL);
    TMsiFile * FindReferencedFile(TMsiTable * pMsiTable, LPCTSTR szStreamName, LPTSTR szFileName, size_t ccFileName);

    DWORD LoadTableNameIfExists(LPCTSTR szTableName);
    DWORD LoadTableNames();
    DWORD LoadTables();
    DWORD LoadFiles();
    DWORD LoadMultipleStreamFiles(TMsiTable * pMsiTable);
    DWORD LoadSimpleCsvFile(TMsiTable * pMsiTable);
    DWORD LoadSummaryFile(MSIHANDLE hMsiSummary);

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
// MSI handle diagnostics

#ifdef _DEBUG
#define MSI_LOG_OPEN_HANDLE(hMsiHandle) MSI_LOG_OPEN_HANDLE_EX(hMsiHandle, __FILE__, __LINE__)
void MSI_LOG_OPEN_HANDLE_EX(MSIHANDLE hMsiHandle, LPCSTR szFile, int nLine);
UINT MSI_CLOSE_HANDLE(MSIHANDLE hMsiHandle);
UINT MSI_DUMP_HANDLES();
#else
#define MSI_LOG_OPEN_HANDLE(hMsiHandle)  0
#define MSI_DUMP_HANDLES()               0
#define MSI_CLOSE_HANDLE                 MsiCloseHandle
#endif

//-----------------------------------------------------------------------------
// MSI helper functions

bool MsiRecordGetInteger(MSIHANDLE hMsiRecord, UINT nColumn, std::tstring & strValue);
bool MsiRecordGetString(MSIHANDLE hMsiRecord, UINT nColumn, std::tstring & strValue);
bool MsiRecordGetBinary(MSIHANDLE hMsiRecord, UINT nColumn, MSI_BLOB & binValue);

#endif // __TMSI_H__
