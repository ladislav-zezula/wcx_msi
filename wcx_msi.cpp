/*****************************************************************************/
/* wcx_msi.cpp                            Copyright (c) Ladislav Zezula 2023 */
/*---------------------------------------------------------------------------*/
/* Main file for Total Commander MSI plugin                                  */
/*---------------------------------------------------------------------------*/
/*   Date    Ver   Who  Comment                                              */
/* --------  ----  ---  -------                                              */
/* 22.07.23  1.00  Lad  Created                                              */
/*****************************************************************************/

#include "wcx_msi.h"                    // Our functions
#include "resource.h"                   // Resource symbols

#include "TStringConvert.inl"            // Conversion functions

#pragma comment(lib, "Msi.lib")

//-----------------------------------------------------------------------------
// Local variables

PFN_PROCESS_DATAA PfnProcessDataA;      // Process data procedure (ANSI)
PFN_PROCESS_DATAW PfnProcessDataW;      // Process data procedure (UNICODE)
PFN_CHANGE_VOLUMEA PfnChangeVolA;       // Change volume procedure (ANSI)
PFN_CHANGE_VOLUMEW PfnChangeVolW;       // Change volume procedure (UNICODE)

//-----------------------------------------------------------------------------
// CanYouHandleThisFile(W) allows the plugin to handle files with different
// extensions than the one defined in Total Commander
// https://www.ghisler.ch/wiki/index.php?title=CanYouHandleThisFile

BOOL WINAPI CanYouHandleThisFileW(LPCWSTR szFileName)
{
    MSIHANDLE hMsiDb = NULL;

    // Just try to open the database. If it succeeds,
    // then we can handle this file
    if(MsiOpenDatabase(szFileName, MSIDBOPEN_READONLY, &hMsiDb) == ERROR_SUCCESS)
    {
        // Log the handle for diagnostics
        MSI_LOG_OPEN_HANDLE(hMsiDb);

        // Close the handle
        MSI_CLOSE_HANDLE(hMsiDb);
    }
    return (hMsiDb != NULL) ? TRUE : FALSE;
}

BOOL WINAPI CanYouHandleThisFile(LPCSTR szFileName)
{
    return CanYouHandleThisFileW(TAnsiToWide(szFileName));
}

//-----------------------------------------------------------------------------
// OpenArchive(W) should perform all necessary operations when an archive is
// to be opened
// https://www.ghisler.ch/wiki/index.php?title=OpenArchive

static void FileTimeToDosFTime(DOS_FTIME & DosTime, const FILETIME & ft)
{
    SYSTEMTIME stUTC;                   // Local file time
    SYSTEMTIME st;                      // Local file time

    FileTimeToSystemTime(&ft, &stUTC);
    SystemTimeToTzSpecificLocalTime(NULL, &stUTC, &st);

    DosTime.ft_year = st.wYear - 60;
    DosTime.ft_month = st.wMonth;
    DosTime.ft_day = st.wDay;
    DosTime.ft_hour = st.wHour;
    DosTime.ft_min = st.wMinute;
    DosTime.ft_tsec = (st.wSecond / 2);
}

static HANDLE OpenArchiveAW(TOpenArchiveData * pArchiveData, LPCWSTR szArchiveName)
{
    WIN32_FIND_DATA wf;
    TMsiDatabase * pMsiDB = NULL;
    MSIHANDLE hMsiDb = NULL;
    HANDLE hFind;

    // Set the default error code
    pArchiveData->OpenResult = E_UNKNOWN_FORMAT;

    // Check the valid parameters
    if(pArchiveData && szArchiveName && szArchiveName[0])
    {
        // Check the valid archive access
        if(pArchiveData->OpenMode == PK_OM_LIST || pArchiveData->OpenMode == PK_OM_EXTRACT)
        {
            // Perform file search to in order to get filetime
            if((hFind = FindFirstFile(szArchiveName, &wf)) != INVALID_HANDLE_VALUE)
            {
                // Attempt to open the MSI
                if(MsiOpenDatabase(szArchiveName, MSIDBOPEN_READONLY, &hMsiDb) == ERROR_SUCCESS)
                {
                    // Log the handle for diagnostics
                    MSI_LOG_OPEN_HANDLE(hMsiDb);

                    // Create the TMsiDatabase object
                    if((pMsiDB = new TMsiDatabase(hMsiDb, wf.ftLastWriteTime)) != NULL)
                    {
                        pArchiveData->OpenResult = 0;
                        return (HANDLE)(pMsiDB);
                    }
                    else
                    {
                        pArchiveData->OpenResult = E_NO_MEMORY;
                        MSI_CLOSE_HANDLE(hMsiDb);
                    }
                }
                FindClose(hFind);
            }
        }
    }
    return NULL;
}

HANDLE WINAPI OpenArchiveW(TOpenArchiveData * pArchiveData)
{
    return OpenArchiveAW(pArchiveData, pArchiveData->szArchiveNameW);
}

HANDLE WINAPI OpenArchive(TOpenArchiveData * pArchiveData)
{
    return OpenArchiveAW(pArchiveData, TAnsiToWide(pArchiveData->szArchiveNameA));
}

//-----------------------------------------------------------------------------
// CloseArchive should perform all necessary operations when an archive
// is about to be closed.
// https://www.ghisler.ch/wiki/index.php?title=CloseArchive

int WINAPI CloseArchive(HANDLE hArchive)
{
    TMsiDatabase * pMsiDb;

    if((pMsiDb = TMsiDatabase::FromHandle(hArchive)) != NULL)
    {
        // Force-close all loaded files
        pMsiDb->CloseAllFiles();
        pMsiDb->UnlockAndRelease();

        // Finally release the database
        pMsiDb->Release();
    }
    return (pMsiDb != NULL) ? ERROR_SUCCESS : E_NOT_SUPPORTED;
}

//-----------------------------------------------------------------------------
// GetPackerCaps tells Totalcmd what features your packer plugin supports.
// https://www.ghisler.ch/wiki/index.php?title=GetPackerCaps

// GetPackerCaps tells Total Commander what features we support
int WINAPI GetPackerCaps()
{
    return(PK_CAPS_MULTIPLE |          // Archive can contain multiple files
           //PK_CAPS_OPTIONS |         // Has options dialog
           PK_CAPS_BY_CONTENT |        // Detect archive type by content
           PK_CAPS_SEARCHTEXT          // Allow searching for text in archives created with this plugin
          );
}

//-----------------------------------------------------------------------------
// ProcessFile should unpack the specified file or test the integrity of the archive.
// https://www.ghisler.ch/wiki/index.php?title=ProcessFile

static void MergePath(LPWSTR szFullPath, size_t ccFullPath, LPCWSTR szPath, LPCWSTR szName)
{
    // Always the buffer with zero
    if(ccFullPath != 0)
    {
        szFullPath[0] = 0;

        // Append destination path, if exists
        if(szPath && szPath[0])
        {
            StringCchCopy(szFullPath, ccFullPath, szPath);
        }

        // Append the name, if any
        if(szName && szName[0])
        {
            // Append backslash
            AddBackslash(szFullPath, ccFullPath);
            StringCchCat(szFullPath, ccFullPath, szName);
        }
    }
}

static int CallProcessDataProc(LPCWSTR szFullPath, int nSize)
{
    // Prioritize UNICODE version of the callback, if exists.
    // This leads to nicer progress dialog shown by Total Commander
    if(PfnProcessDataW != NULL)
        return PfnProcessDataW(szFullPath, nSize);

    // Call ANSI version of callback, if needed
    if(PfnProcessDataA != NULL)
        return PfnProcessDataA(TWideToAnsi(szFullPath), nSize);

    // Continue the operation
    return TRUE;
}

int WINAPI ProcessFileW(HANDLE hArchive, PROCESS_FILE_OPERATION nOperation, LPCWSTR szDestPath, LPCWSTR szDestName)
{
    TMsiDatabase * pMsiDb;
    TMsiFile * pMsiFile;
    HANDLE hLocFile = INVALID_HANDLE_VALUE;
    DWORD dwBytesWritten;
    DWORD dwFileOffset = 0;
    WCHAR szFullPath[MAX_PATH];
    int nResult = E_NOT_SUPPORTED;              // Result reported to Total Commander

    // Check whether it's the valid archive handle
    if((pMsiDb = TMsiDatabase::FromHandle(hArchive)) != NULL)
    {
        // If verify or skip the file, do nothing
        if(nOperation == PK_TEST || nOperation == PK_SKIP)
        {
            pMsiDb->UnlockAndRelease();
            return 0;
        }

        // Do we have to extract the file? If yes, the file must be saved in TMsiDatabase
        if(nOperation == PK_EXTRACT)
        {
            if((pMsiFile = pMsiDb->LastFile()) != NULL)
            {
                // Construct the full path name
                MergePath(szFullPath, _countof(szFullPath), szDestPath, szDestName);

                // Create the local file
                hLocFile = CreateFile(szFullPath, GENERIC_WRITE, FILE_SHARE_READ, NULL, CREATE_ALWAYS, 0, NULL);
                if(hLocFile != INVALID_HANDLE_VALUE)
                {
                    // Populate the cach with complete file data
                    if(pMsiFile->LoadFileData() == ERROR_SUCCESS)
                    {
                        const MSI_BLOB & FileData = pMsiFile->FileData();

                        // Tell Total Commader what we are doing
                        while(CallProcessDataProc(szFullPath, dwFileOffset))
                        {
                            DWORD dwBytesToWrite = 0x1000;

                            // Get the size of remaining data
                            if((dwFileOffset + dwBytesToWrite) > FileData.cbData)
                                dwBytesToWrite = FileData.cbData - dwFileOffset;

                            // Are we done?
                            if(dwBytesToWrite == 0)
                            {
                                nResult = 0;
                                break;
                            }

                            // Write the target file
                            if(!WriteFile(hLocFile, FileData.pbData + dwFileOffset, dwBytesToWrite, &dwBytesWritten, NULL))
                            {
                                nResult = E_EWRITE;
                                break;
                            }

                            // Increment the total bytes
                            dwFileOffset += dwBytesWritten;
                        }
                    }

                    // Close the local file
                    SetEndOfFile(hLocFile);
                    CloseHandle(hLocFile);
                }
                else
                {
                    nResult = E_ECREATE;
                }

                // Dereference the file
                pMsiFile->Release();
            }
        }

        // Unlock and release the archive
        pMsiDb->UnlockAndRelease();
    }
    return nResult;
}

int WINAPI ProcessFile(HANDLE hArchive, PROCESS_FILE_OPERATION nOperation, LPCSTR szDestPath, LPCSTR szDestName)
{
    return ProcessFileW(hArchive, nOperation, TAnsiToWide(szDestPath), TUTF8ToWide(szDestName));
}

//-----------------------------------------------------------------------------
// Totalcmd calls ReadHeader to find out what files are in the archive.
// https://www.ghisler.ch/wiki/index.php?title=ReadHeader

template <typename HDR>
static void StoreFoundFile(TMsiFile * pMsiFile, HDR * pHeaderData, DOS_FTIME & fileTime)
{
    // Store the file name
    StringCchCopyX(pHeaderData->FileName, _countof(pHeaderData->FileName), pMsiFile->Name());

    // Store the file time
    pHeaderData->FileTime = fileTime;

    // TODO: Store the file sizes
    pHeaderData->PackSize = pMsiFile->FileSize();
    pHeaderData->UnpSize = pMsiFile->FileSize();

    // Store file attributes
    pHeaderData->FileAttr = FILE_ATTRIBUTE_ARCHIVE;
}

template <typename HDR>
int ReadHeaderTemplate(HANDLE hArchive, HDR * pHeaderData)
{
    TMsiDatabase * pMsiDb;
    TMsiFile * pMsiFile;
    DOS_FTIME fileTime;
    DWORD dwErrCode = E_NOT_SUPPORTED;

    // Check the proper parameters
    if((pMsiDb = TMsiDatabase::FromHandle(hArchive)) != NULL)
    {
        // Convert the Windows filetime to DOS filetime
        FileTimeToDosFTime(fileTime, pMsiDb->FileTime());

        // Split the action
        if((pMsiFile = pMsiDb->GetNextFile()) != NULL)
        {
            StoreFoundFile(pMsiFile, pHeaderData, fileTime);
            pMsiFile->Release();
            dwErrCode = 0;
        }
        pMsiDb->UnlockAndRelease();
    }
    return dwErrCode;
}

int WINAPI ReadHeader(HANDLE hArchive, THeaderData * pHeaderData)
{
    // Use the common template function
    return ReadHeaderTemplate(hArchive, pHeaderData);
}

int WINAPI ReadHeaderEx(HANDLE hArchive, THeaderDataEx * pHeaderData)
{
    // Use the common template function
    return ReadHeaderTemplate(hArchive, pHeaderData);
}

int WINAPI ReadHeaderExW(HANDLE hArchive, THeaderDataExW * pHeaderData)
{
    // Use the common template function
    return ReadHeaderTemplate(hArchive, pHeaderData);
}

//-----------------------------------------------------------------------------
// SetChangeVolProc(W) allows you to notify user about changing a volume when packing files
// https://www.ghisler.ch/wiki/index.php?title=SetChangeVolProc

// This function allows you to notify user
// about changing a volume when packing files
void WINAPI SetChangeVolProc(HANDLE /* hArchive */, PFN_CHANGE_VOLUMEA PfnChangeVol)
{
    PfnChangeVolA = PfnChangeVol;
}

void WINAPI SetChangeVolProcW(HANDLE /* hArchive */, PFN_CHANGE_VOLUMEW PfnChangeVol)
{
    PfnChangeVolW = PfnChangeVol;
}

//-----------------------------------------------------------------------------
// SetProcessDataProc(W) allows you to notify user about the progress when you un/pack files
// Note that Total Commander may use INVALID_HANDLE_VALUE for the hArchive parameter
// https://www.ghisler.ch/wiki/index.php?title=SetProcessDataProc

void WINAPI SetProcessDataProc(HANDLE /* hArchive */, PFN_PROCESS_DATAA PfnProcessData)
{
    PfnProcessDataA = PfnProcessData;
}

void WINAPI SetProcessDataProcW(HANDLE /* hArchive */, PFN_PROCESS_DATAW PfnProcessData)
{
    PfnProcessDataW = PfnProcessData;
}

//-----------------------------------------------------------------------------
// PackFiles(W) specifies what should happen when a user creates, or adds files to the archive
// https://www.ghisler.ch/wiki/index.php?title=PackFiles

int WINAPI PackFilesW(LPCWSTR szPackedFile, LPCWSTR szSubPath, LPCWSTR szSrcPath, LPCWSTR szAddList, int nFlags)
{
    UNREFERENCED_PARAMETER(szPackedFile);
    UNREFERENCED_PARAMETER(szSubPath);
    UNREFERENCED_PARAMETER(szSrcPath);
    UNREFERENCED_PARAMETER(szAddList);
    UNREFERENCED_PARAMETER(nFlags);
    return E_NOT_SUPPORTED;
}

// PackFiles adds file(s) to an archive
int WINAPI PackFiles(LPCSTR szPackedFile, LPCSTR szSubPath, LPCSTR szSrcPath, LPCSTR szAddList, int nFlags)
{
    UNREFERENCED_PARAMETER(szPackedFile);
    UNREFERENCED_PARAMETER(szSubPath);
    UNREFERENCED_PARAMETER(szSrcPath);
    UNREFERENCED_PARAMETER(szAddList);
    UNREFERENCED_PARAMETER(nFlags);
    return E_NOT_SUPPORTED;
}

//-----------------------------------------------------------------------------
// DeleteFiles(W) should delete the specified files from the archive
// https://www.ghisler.ch/wiki/index.php?title=DeleteFiles

int WINAPI DeleteFilesW(LPCWSTR szPackedFile, LPCWSTR szDeleteList)
{
    UNREFERENCED_PARAMETER(szPackedFile);
    UNREFERENCED_PARAMETER(szDeleteList);
    return E_NOT_SUPPORTED;
}

int WINAPI DeleteFiles(LPCSTR szPackedFile, LPCSTR szDeleteList)
{
    UNREFERENCED_PARAMETER(szPackedFile);
    UNREFERENCED_PARAMETER(szDeleteList);
    return E_NOT_SUPPORTED;
}

//-----------------------------------------------------------------------------
// ConfigurePacker gets called when the user clicks the Configure button
// from within "Pack files..." dialog box in Totalcmd.
// https://www.ghisler.ch/wiki/index.php?title=ConfigurePacker

void WINAPI ConfigurePacker(HWND hParent, HINSTANCE hDllInstance)
{
    UNREFERENCED_PARAMETER(hDllInstance);
    UNREFERENCED_PARAMETER(hParent);
}

//-----------------------------------------------------------------------------
// PackSetDefaultParams is called immediately after loading the DLL, before
// any other function. This function is new in version 2.1. It requires Total
// Commander >=5.51, but is ignored by older versions.
// https://www.ghisler.ch/wiki/index.php?title=PackSetDefaultParams

void WINAPI PackSetDefaultParams(TPackDefaultParamStruct * dps)
{
    // Set default configuration.
    SetDefaultConfiguration();
    g_szIniFile[0] = 0;

    // If INI file, load it from it too.
    if(dps != NULL && dps->DefaultIniName[0])
    {
        StringCchCopyX(g_szIniFile, _countof(g_szIniFile), dps->DefaultIniName);
        LoadConfiguration();
    }
}
