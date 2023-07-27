/*****************************************************************************/
/* TMsiDatabase.cpp                       Copyright (c) Ladislav Zezula 2023 */
/*---------------------------------------------------------------------------*/
/* Implementation of the TMsiDatabase class methods                          */
/*---------------------------------------------------------------------------*/
/*   Date    Ver   Who  Comment                                              */
/* --------  ----  ---  -------                                              */
/* 23.07.23  1.00  Lad  Created                                              */
/*****************************************************************************/

#include "wcx_msi.h"

//-----------------------------------------------------------------------------
// Local (non-class) functions

static bool FindStringInList(MSI_STRING_LIST & MsiTablesList, const std::tstring & strTableName)
{
    for(size_t i = 0; i < MsiTablesList.size(); i++)
    {
        if(!_tcsicmp(MsiTablesList[i].c_str(), strTableName.c_str()))
        {
            return true;
        }
    }
    return false;
}

//-----------------------------------------------------------------------------
// Constructor and destructor

TMsiDatabase::TMsiDatabase(MSIHANDLE hMsiDb, FILETIME & ft)
{
    InitializeCriticalSection(&m_Lock);
    m_MagicSignature = MSI_MAGIC_SIGNATURE;
    m_pFileEntry = NULL;
    m_pLastFile = NULL;
    m_FileTime = ft;
    m_hMsiDb = hMsiDb;
    m_dwTables = 0;
    m_dwFiles = 0;
    m_dwRefs = 1;

    // The list head is empty
    InitializeListHead(&m_Tables);
    InitializeListHead(&m_Files);
}

TMsiDatabase::~TMsiDatabase()
{
#ifdef _DEBUG
    UINT nHandleCount;

    // Only one handle should be open now
    if((nHandleCount = MSI_DUMP_HANDLES()) > 1)
    {
        Dbg(_T("Handle leak detected (%u handles)\n"), nHandleCount);
        __debugbreak();
    }

    // Ask MSI.dll about unclosed handles
    if((nHandleCount = MsiCloseAllHandles()) != 1)
    {
        Dbg(_T("Handle leak detected (%u handles)\n"), nHandleCount);
        __debugbreak();
    }

    // Reference count check
    if(m_dwRefs > 0)
    {
        Dbg(_T("Non-zero reference count detected (%u)\n"), m_dwRefs);
        __debugbreak();
    }
#endif

    // Free the MSI handle
    if(m_hMsiDb != NULL)
        MSI_CLOSE_HANDLE(m_hMsiDb);
    m_hMsiDb = NULL;

    // Delete the lock
    DeleteCriticalSection(&m_Lock);
}

//-----------------------------------------------------------------------------
// Public functions

TMsiDatabase * TMsiDatabase::FromHandle(HANDLE hHandle)
{
    TMsiDatabase * pMsiDb;

    if(hHandle != NULL && hHandle != INVALID_HANDLE_VALUE)
    {
        pMsiDb = static_cast<TMsiDatabase *>(hHandle);
        if(pMsiDb->m_MagicSignature == MSI_MAGIC_SIGNATURE)
        {
            EnterCriticalSection(&pMsiDb->m_Lock);
            pMsiDb->AddRef();
            return pMsiDb;
        }
    }
    return NULL;

}

DWORD TMsiDatabase::AddRef()
{
    return InterlockedIncrement((LONG *)(&m_dwRefs));
}

DWORD TMsiDatabase::Release()
{
    if(InterlockedDecrement((LONG *)(&m_dwRefs)) == 0)
    {
        delete this;
        return 0;
    }
    return m_dwRefs;
}

void TMsiDatabase::CloseAllFiles()
{
    // Free the last file, if any
    ReleaseLastFile();

    // Free the list of files
    DeleteLinkedList<TMsiFile>(m_Files);
    m_dwFiles = 0;

    // Free list of tables
    DeleteLinkedList<TMsiTable>(m_Tables);
    m_dwTables = 0;
}

TMsiFile * TMsiDatabase::ReleaseLastFile(TMsiFile * pMsiFile)
{
    // Release the old last file
    if(m_pLastFile != NULL)
        m_pLastFile->Release();
    m_pLastFile = NULL;

    // Set the new last file
    if(pMsiFile != NULL)
        pMsiFile->AddRef();
    m_pLastFile = pMsiFile;

    // Return the addded file
    return pMsiFile;
}

TMsiFile * TMsiDatabase::FindReferencedFile(TMsiTable * pMsiTable, LPCTSTR szStreamName, LPTSTR szFileName, size_t ccFileName)
{
    TMsiFile * pRefFile = NULL;
    LPCTSTR szDot;
    TCHAR szRefFile[MAX_PATH];

    // References to other files are only in the "_Streams" table
    if(pMsiTable->m_bIsStreamsTable)
    {
        // There must be dot in the name, like "Binary.bannrbmp"
        if((szDot = _tcschr(szStreamName, _T('.'))) != NULL)
        {
            // Create the referenced file name
            StringCchCopy(szRefFile, _countof(szRefFile), szStreamName);
            szRefFile[szDot - szStreamName] = _T('\\');

            // Is it in the database?
            if((pRefFile = IsFilePresent(szRefFile)) != NULL)
            {
                StringCchPrintf(szFileName, ccFileName, _T("%s\\%s"), pMsiTable->Name(), szDot + 1);
            }
        }
    }
    return pRefFile;
}

void TMsiDatabase::UnlockAndRelease()
{
    LeaveCriticalSection(&m_Lock);
    Release();
}

TMsiFile * TMsiDatabase::GetNextFile()
{
    PLIST_ENTRY pHeadEntry = &m_Files;
    TMsiFile * pMsiFile;
    DWORD dwErrCode = ERROR_SUCCESS;

    // Files are not loaded yet
    if(m_pFileEntry == NULL)
    {
        // Shall we load the tables?
        if(dwErrCode == ERROR_SUCCESS && m_TableNames.size() == 0)
            dwErrCode = LoadTableNames();

        // Shall we load all tables?
        if(dwErrCode == ERROR_SUCCESS && m_TableNames.size() && IsListEmpty(&m_Tables))
            dwErrCode = LoadTables();

        // Shall we load all files?
        if(dwErrCode == ERROR_SUCCESS && !IsListEmpty(&m_Tables) && IsListEmpty(&m_Files))
            dwErrCode = LoadFiles();

        // Setup the file iteration
        m_pFileEntry = m_Files.Flink;
    }

    // Do we have some files?
    if(m_pFileEntry != pHeadEntry)
    {
        // Get the next file
        pMsiFile = CONTAINING_RECORD(m_pFileEntry, TMsiFile, m_Entry);
        m_pFileEntry = m_pFileEntry->Flink;

        // Make sure that we have the file size
        pMsiFile->LoadFileSize();
        pMsiFile->AddRef();

        // Set the new last file
        return ReleaseLastFile(pMsiFile);
    }
    return NULL;
}

DWORD TMsiDatabase::LoadTableNames()
{
    std::tstring strTableName;
    MSIHANDLE hMsiRecord = NULL;
    MSIHANDLE hMsiView = NULL;
    DWORD dwErrCode;

    // The "_Validation" table contain list of all tables in the MSI
    if((dwErrCode = MsiDatabaseOpenView(m_hMsiDb, _T("SELECT * from _Validation"), &hMsiView)) == ERROR_SUCCESS)
    {
        // Log the handle for diagnostics
        MSI_LOG_OPEN_HANDLE(hMsiView);

        // Execute the query
        if((dwErrCode = MsiViewExecute(hMsiView, NULL)) == ERROR_SUCCESS)
        {
            // Dump all records
            while(MsiViewFetch(hMsiView, &hMsiRecord) == ERROR_SUCCESS)
            {
                // Log the handle for diagnostics
                MSI_LOG_OPEN_HANDLE(hMsiRecord);

                // Retrieve the table name
                if(MsiRecordGetString(hMsiRecord, 0, strTableName))
                {
                    if(!FindStringInList(m_TableNames, strTableName))
                    {
                        m_TableNames.push_back(strTableName);
                    }
                }
                MSI_CLOSE_HANDLE(hMsiRecord);
            }
            
            // Finalize the view
            MsiViewClose(hMsiView);
        }
        MSI_CLOSE_HANDLE(hMsiView);
    }

    // Check whether there is a "_Streams" table
    if((dwErrCode = MsiDatabaseOpenView(m_hMsiDb, _T("SELECT * from _Streams"), &hMsiView)) == ERROR_SUCCESS)
    {
        // Log the handle for diagnostics
        MSI_LOG_OPEN_HANDLE(hMsiView);

        // Insert the table to the table list
        if(!FindStringInList(m_TableNames, _T("_Streams")))
            m_TableNames.push_back(_T("_Streams"));
        MSI_CLOSE_HANDLE(hMsiView);
    }

    return m_TableNames.size() ? ERROR_SUCCESS : ERROR_NO_MORE_ITEMS;
}

DWORD TMsiDatabase::LoadTables()
{
    TMsiTable * pMsiTable;
    DWORD dwErrCode = ERROR_SUCCESS;

    // Enumerate table names
    for(size_t i = 0; i < m_TableNames.size(); i++)
    {
        const std::tstring & strTableName = m_TableNames[i];
        MSIHANDLE hMsiView = NULL;
        TCHAR szQuery[256];

        // Load the list of columns of the table
        StringCchPrintf(szQuery, _countof(szQuery), _T("SELECT * FROM %s"), strTableName.c_str());
        if(MsiDatabaseOpenView(m_hMsiDb, szQuery, &hMsiView) == ERROR_SUCCESS)
        {
            // Log the handle for diagnostics
            MSI_LOG_OPEN_HANDLE(hMsiView);

            // Create the TMsiTable object
            if((pMsiTable = new TMsiTable(this, strTableName, hMsiView)) != NULL)
            {
                if(pMsiTable->Load() == ERROR_SUCCESS)
                {
                    InsertTailList(&m_Tables, &pMsiTable->m_Entry);
                    InterlockedIncrement((LONG *)(&m_dwTables));
                    pMsiTable = NULL;
                }
                else
                {
                    pMsiTable->Release();
                    pMsiTable = NULL;
                }
            }
            else
            {
                dwErrCode = ERROR_NOT_ENOUGH_MEMORY;
                MSI_CLOSE_HANDLE(hMsiView);
                break;
            }
        }
    }
    return dwErrCode;
}

DWORD TMsiDatabase::LoadFiles()
{
    PLIST_ENTRY pHeadEntry = &m_Tables;
    PLIST_ENTRY pListEntry;

    for(pListEntry = pHeadEntry->Flink; pListEntry != pHeadEntry; pListEntry = pListEntry->Flink)
    {
        TMsiTable * pMsiTable = CONTAINING_RECORD(pListEntry, TMsiTable, m_Entry);

        // Is it a database table with stream field?
        if(pMsiTable->m_nStreamColumn != INVALID_SIZE_T && pMsiTable->m_nNameColumn != INVALID_SIZE_T)
        {
            LoadMultipleStreamFiles(pMsiTable);
        }

        // Simple database table - we simulate it as simple CSV file
        else
        {
            LoadSimpleCsvFile(pMsiTable);
        }
    }
    return ERROR_SUCCESS;
}

DWORD TMsiDatabase::LoadMultipleStreamFiles(TMsiTable * pMsiTable)
{
    TMsiFile * pMsiFile;
    MSIHANDLE hMsiRecord;
    MSIHANDLE hMsiView = pMsiTable->m_hMsiView;
    DWORD dwErrCode;

    // Execute the query
    if((dwErrCode = MsiViewExecute(hMsiView, NULL)) == ERROR_SUCCESS)
    {
        // Dump all records
        while(MsiViewFetch(hMsiView, &hMsiRecord) == ERROR_SUCCESS)
        {
            // Log the handle for diagnostics
            MSI_LOG_OPEN_HANDLE(hMsiRecord);

            // Create the TMsiFile object
            if((pMsiFile = new TMsiFile(pMsiTable)) != NULL)
            {
                if((dwErrCode = pMsiFile->SetBinaryFile(this, hMsiRecord)) == ERROR_SUCCESS)
                {
                    InsertTailList(&m_Files, &pMsiFile->m_Entry);
                    InterlockedIncrement((LONG *)(&m_dwFiles));
                }
                else
                {
                    MSI_CLOSE_HANDLE(hMsiRecord);
                    pMsiFile->Release();
                }
            }
        }

        // Finalize the executed view
        MsiViewClose(hMsiView);
    }
    return dwErrCode;
}

DWORD TMsiDatabase::LoadSimpleCsvFile(TMsiTable * pMsiTable)
{
    TMsiFile * pMsiFile;
    DWORD dwErrCode = ERROR_NOT_ENOUGH_MEMORY;

    if((pMsiFile = new TMsiFile(pMsiTable)) != NULL)
    {
        if((dwErrCode = pMsiFile->SetCsvFile(this)) == ERROR_SUCCESS)
        {
            InsertTailList(&m_Files, &pMsiFile->m_Entry);
            InterlockedIncrement((LONG *)(&m_dwFiles));
        }
        else
        {
            pMsiFile->Release();
        }
    }
    return dwErrCode;
}

TMsiFile * TMsiDatabase::IsFilePresent(LPCTSTR szFileName)
{
    PLIST_ENTRY pHeadEntry = &m_Files;
    PLIST_ENTRY pListEntry;

    for(pListEntry = pHeadEntry->Flink; pListEntry != pHeadEntry; pListEntry = pListEntry->Flink)
    {
        TMsiFile * pMsiFile = CONTAINING_RECORD(pListEntry, TMsiFile, m_Entry);

        if(!_tcsicmp(pMsiFile->Name(), szFileName))
        {
            return pMsiFile;
        }
    }
    return NULL;
}

TMsiFile * TMsiDatabase::LastFile()
{
    if(m_pLastFile != NULL)
        m_pLastFile->AddRef();
    return m_pLastFile;
}
