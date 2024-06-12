/*****************************************************************************/
/* TMsiFile.cpp                           Copyright (c) Ladislav Zezula 2023 */
/*---------------------------------------------------------------------------*/
/* Implementation of the TMsiFile class methods                              */
/*---------------------------------------------------------------------------*/
/*   Date    Ver   Who  Comment                                              */
/* --------  ----  ---  -------                                              */
/* 24.07.23  1.00  Lad  Created                                              */
/*****************************************************************************/

#include "wcx_msi.h"

//-----------------------------------------------------------------------------
// MSI Summary info data types

struct MSI_PROPERTY
{
    VARENUM vType;
    LPCTSTR szName;
};

static const MSI_PROPERTY MsiPropertyList[] = 
{
    {VT_EMPTY},
    {VT_I2,       _T("Codepage")},
    {VT_LPSTR,    _T("Title")},
    {VT_LPSTR,    _T("Subject")},
    {VT_LPSTR,    _T("Author")},
    {VT_LPSTR,    _T("Keywords")},
    {VT_LPSTR,    _T("Comments")},
    {VT_LPSTR,    _T("Template")},
    {VT_LPSTR,    _T("Last Saved By")},
    {VT_LPSTR,    _T("Revision Number")},
    {VT_EMPTY},
    {VT_FILETIME, _T("Last Printed")},
    {VT_FILETIME, _T("Create Time / Date")},
    {VT_FILETIME, _T("Last Save Time / Date")},
    {VT_I4,       _T("Page Count")},
    {VT_I4,       _T("Word Count")},
    {VT_I4,       _T("Character Count")},
    {VT_EMPTY},
    {VT_LPSTR,    _T("Creating Application")},
    {VT_I4,       _T("Security")}
};

static LPCTSTR szCsvExtension = _T(".csv");

//-----------------------------------------------------------------------------
// Non-class members

static LPBYTE AppendNewLine(LPBYTE pbBufferPtr, LPBYTE pbBufferEnd)
{
    if((pbBufferPtr + 2) <= pbBufferEnd)
    {
        pbBufferPtr[0] = '\r';
        pbBufferPtr[1] = '\n';
    }
    return pbBufferPtr + 2;
}

static LPBYTE AppendUtf8Marker(LPBYTE pbBufferPtr, LPBYTE pbBufferEnd)
{
    if((pbBufferPtr + 3) <= pbBufferEnd)
    {
        pbBufferPtr[0] = 0xEF;
        pbBufferPtr[1] = 0xBB;
        pbBufferPtr[2] = 0xBF;
    }
    return pbBufferPtr + 3;
}

static LPBYTE AppendFieldString(LPBYTE pbBufferPtr, LPBYTE pbBufferEnd, const std::tstring & strValue, size_t nIndex)
{
    size_t nLength;

    // Append length of the UTF8 string plus 2x quotation marks
    nLength = ((nIndex > 0) ? 1 : 0)            // Comma at the beginning
            + 1                                 // Opening quotation mark
            + WideCharToMultiByte(CP_UTF8, 0, strValue.c_str(), (int)(strValue.size()), NULL, 0, NULL, NULL)
            + 1;                                // Closing quotation mark

    // Is this "dry run" (calculating the size)?
    if(pbBufferEnd == NULL)
        return pbBufferPtr + nLength;

    // Append comma
    if(nIndex > 0)
        *pbBufferPtr++ = ',';

    // Append opening quotation mark
    *pbBufferPtr++ = '\"';

    // Append the UTF-8 string
    pbBufferPtr += WideCharToMultiByte(CP_UTF8, 0, strValue.c_str(),
                                             (int)(strValue.size()),
                                           (LPSTR)(pbBufferPtr),
                                             (int)(pbBufferEnd - pbBufferPtr), NULL, NULL);
    // Append closing quotation mark
    *pbBufferPtr++ = '\"';

    return pbBufferPtr;
}

HRESULT StringCchPrintfFT(LPTSTR szBuffer, size_t ccBuffer, const FILETIME & ft)
{
    SYSTEMTIME st;
    LPTSTR szBufferEnd = szBuffer + ccBuffer;
    LPTSTR szBufferPtr = szBuffer;
    int nLength;

    // If the filetime is not present, do nothing
    if(ft.dwHighDateTime && ft.dwHighDateTime)
    {
        // Convert the file time to SYSTEMTIME
        FileTimeToSystemTime(&ft, &st);

        // Create date
        nLength = GetDateFormat(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &st, NULL, szBufferPtr, (int)(szBufferEnd - szBufferPtr));
        if(nLength != 0)
        {
            szBufferPtr += (nLength - 1);
            *szBufferPtr++ = _T(' ');
        }

        // Create time
        GetTimeFormat(LOCALE_USER_DEFAULT, TIME_FORCE24HOURFORMAT, &st, NULL, szBufferPtr, (int)(szBufferEnd - szBufferPtr));
    }
    else
    {
        StringCchCopy(szBuffer, ccBuffer, _T("N/A"));
    }
    return S_OK;
}

//-----------------------------------------------------------------------------
// TMsiFile functions

TMsiFile::TMsiFile(TMsiTable * pMsiTable)
{
    InitializeListHead(&m_Entry);
    m_hMsiHandle = NULL;
    m_dwFileSize = 0;
    m_FileType = MsiFileNone;
    m_dwRefs = 1;

    // Reset the referenced file
    m_pRefFile = NULL;

    // Reference the MSI table
    if((m_pMsiTable = pMsiTable) != NULL)
    {
        m_pMsiTable->AddRef();
    }
}

TMsiFile::~TMsiFile()
{
    // Sanity check
    assert(m_dwRefs == 0);
    
    // Dereference the referenced file
    if(m_pRefFile != NULL)
        m_pRefFile->Release();
    m_pRefFile = NULL;

    // Dereference the database table
    if(m_pMsiTable != NULL)
        m_pMsiTable->Release();
    m_pMsiTable = NULL;

    // Close the MSI handle, if any
    if(m_hMsiHandle != NULL)
        MSI_CLOSE_HANDLE(m_hMsiHandle);
    m_hMsiHandle = NULL;
}

//-----------------------------------------------------------------------------
// TMsiFile methods

DWORD TMsiFile::AddRef()
{
    return InterlockedIncrement((LONG *)(&m_dwRefs));
}

DWORD TMsiFile::Release()
{
    if(InterlockedDecrement((LONG *)(&m_dwRefs)) == 0)
    {
        delete this;
        return 0;
    }
    return m_dwRefs;
}

DWORD TMsiFile::SetSummaryFile(TMsiDatabase * pMsiDb, MSIHANDLE hMsiSummary)
{
    // Remember the summary info
    m_FileType = MsiFileSummary;
    m_hMsiHandle = hMsiSummary;

    // Ensure that we have an unique file name
    return SetUniqueFileName(pMsiDb, NULL, _T("_SummaryInformation"), szCsvExtension);
}

DWORD TMsiFile::SetBinaryFile(TMsiDatabase * pMsiDb, MSIHANDLE hMsiRecord)
{
    TMsiFile * pRefFile;
    std::tstring strItemName;
    LPTSTR szExtension;
    TCHAR szFileName[MAX_PATH];
    TCHAR szBaseName[MAX_PATH];
    TCHAR szFileExt[MAX_PATH] = {0};

    // Retrieve the name of the item
    if(MsiRecordGetString(hMsiRecord, (UINT)(m_pMsiTable->m_nNameColumn), strItemName))
    {
        // Is this just a reference to another file?
        pRefFile = pMsiDb->FindReferencedFile(m_pMsiTable, strItemName.c_str(), szFileName, _countof(szFileName));
        if((m_pRefFile = pRefFile) == NULL)
        {
            // Fix the item name to be file-safe
            MakeItemNameFileSafe(strItemName);

            // Split the base name and extension
            StringCchPrintf(szBaseName, _countof(szBaseName), strItemName.c_str());
            if((szExtension = GetFileExtension(szBaseName)) > szBaseName)
            {
                StringCchCopy(szFileExt, _countof(szFileExt), szExtension);
                szExtension[0] = 0;
            }

            // Setup the unique file name
            SetUniqueFileName(pMsiDb, m_pMsiTable->Name(), szBaseName, szFileExt);
        }
        else
        {
            m_strName.assign(szFileName);
            m_pRefFile->AddRef();
        }

        // Assign the file name and record handle
        m_FileType = MsiFileBinary;
        m_hMsiHandle = hMsiRecord;
        return ERROR_SUCCESS;
    }
    return ERROR_NOT_SUPPORTED;
}

DWORD TMsiFile::SetCsvFile(TMsiDatabase * pMsiDb)
{
    // Setup the handle
    m_FileType = MsiFileTable;
    m_hMsiHandle = NULL;

    // Generate unique file name
    return SetUniqueFileName(pMsiDb, NULL, m_pMsiTable->Name(), szCsvExtension);
}

DWORD TMsiFile::LoadSummaryFile(LPDWORD PtrFileSize)
{
    std::tstring strValue;
    LPBYTE pbBufferBegin = m_Data.pbData;
    LPBYTE pbBufferPtr = m_Data.pbData;
    LPBYTE pbBufferEnd = m_Data.pbData + m_Data.cbData;
    DWORD dwErrCode = ERROR_SUCCESS;
    UINT nPropertyCount = 0;

    // Append the UTF-8 marker and table header
    pbBufferPtr = AppendUtf8Marker(pbBufferPtr, pbBufferEnd);
    pbBufferPtr = AppendFieldString(pbBufferPtr, pbBufferEnd, _T("Name"), 0);
    pbBufferPtr = AppendFieldString(pbBufferPtr, pbBufferEnd, _T("Value"), 1);
    pbBufferPtr = AppendNewLine(pbBufferPtr, pbBufferEnd);

    // Read each field and store it tot he CSV file
    if(MsiSummaryInfoGetPropertyCount(m_hMsiHandle, &nPropertyCount) == ERROR_SUCCESS)
    {
        for(UINT i = 0; i < _countof(MsiPropertyList); i++)
        {
            FILETIME ft = {0};
            UINT uDataType = 0;
            INT iValue = 0;
            TCHAR szValue[1024] = {0};
            DWORD ccValue = _countof(szValue);

            if(MsiSummaryInfoGetProperty(m_hMsiHandle, i, &uDataType, &iValue, &ft, szValue, &ccValue) == ERROR_SUCCESS)
            {
                VARENUM vType = (VARENUM)(uDataType);
                bool bUnknownFormat = false;

                if(vType != VT_EMPTY)
                {
                    // Format the data type
                    switch(vType)
                    {
                        case VT_I2:
                        case VT_I4:
                            StringCchPrintf(szValue, _countof(szValue), _T("%i"), iValue);
                            break;

                        case VT_FILETIME:
                            StringCchPrintfFT(szValue, _countof(szValue), ft);
                            break;

                        case VT_LPSTR:
                            // Already there
                            break;

                        default:
                            bUnknownFormat = true;
                            assert(false);
                            break;
                    }

                    // Did we retrieve the format?
                    if(bUnknownFormat == false)
                    {
                        // Append name and value
                        pbBufferPtr = AppendFieldString(pbBufferPtr, pbBufferEnd, MsiPropertyList[i].szName, 0);
                        pbBufferPtr = AppendFieldString(pbBufferPtr, pbBufferEnd, szValue, 1);
                        pbBufferPtr = AppendNewLine(pbBufferPtr, pbBufferEnd);
                    }
                }
            }
        }
    }

    // Give the file size to the caller
    if(dwErrCode == ERROR_SUCCESS)
        PtrFileSize[0] = (DWORD)(pbBufferPtr - pbBufferBegin);
    return dwErrCode;
}

DWORD TMsiFile::LoadBinaryFile(LPDWORD PtrFileSize)
{
    DWORD dwFileSize = 0;
    DWORD dwErrCode;

    // "Load file data" mode?
    if(m_Data.pbData != NULL)
    {
        dwFileSize = m_dwFileSize;
        dwErrCode = MsiRecordReadStream(m_hMsiHandle, (UINT)(m_pMsiTable->m_nStreamColumn + 1), (char *)(m_Data.pbData), &dwFileSize);
    }
    else
    {
        dwFileSize = MsiRecordDataSize(m_hMsiHandle, (UINT)(m_pMsiTable->m_nStreamColumn + 1));
        dwErrCode = ERROR_SUCCESS;
    }

    // Give the file size to the caller
    if(dwErrCode == ERROR_SUCCESS)
        PtrFileSize[0] = dwFileSize;
    return dwErrCode;
}

DWORD TMsiFile::LoadCsvFile(LPDWORD PtrFileSize)
{
    const std::vector<TMsiColumn> & Columns = m_pMsiTable->Columns();
    std::tstring strValue;
    MSIHANDLE hMsiView = m_pMsiTable->MsiView();
    MSIHANDLE hMsiRecord;
    LPBYTE pbBufferBegin = m_Data.pbData;
    LPBYTE pbBufferPtr = m_Data.pbData;
    LPBYTE pbBufferEnd = m_Data.pbData + m_Data.cbData;
    DWORD dwErrCode = ERROR_SUCCESS;

    // Append the UTF-8 marker
    pbBufferPtr = AppendUtf8Marker(pbBufferPtr, pbBufferEnd);

    // Append the header columns
    for(size_t i = 0; i < Columns.size(); i++)
    {
        pbBufferPtr = AppendFieldString(pbBufferPtr, pbBufferEnd, Columns[i].m_strName, i);
    }

    // Append the end-of-line
    pbBufferPtr = AppendNewLine(pbBufferPtr, pbBufferEnd);

    // Execute the query on top of the view
    if(MsiViewExecute(hMsiView, NULL) == ERROR_SUCCESS)
    {
        // Fetch all records
        while(MsiViewFetch(hMsiView, &hMsiRecord) == ERROR_SUCCESS)
        {
            // Log the handle for diagnostics
            MSI_LOG_OPEN_HANDLE(hMsiRecord);

            // Dump all columns
            for(size_t i = 0; i < Columns.size(); i++)
            {
                // Retrieve the buffer data
                switch(Columns[i].m_Type)
                {
                    case MsiTypeInteger:
                        MsiRecordGetInteger(hMsiRecord, (UINT)(i), strValue);
                        pbBufferPtr = AppendFieldString(pbBufferPtr, pbBufferEnd, strValue, i);
                        break;

                    case MsiTypeString:
                        MsiRecordGetString(hMsiRecord, (UINT)(i), strValue);
                        pbBufferPtr = AppendFieldString(pbBufferPtr, pbBufferEnd, strValue, i);
                        break;

                    default:
                        dwErrCode = ERROR_NOT_SUPPORTED;
                        assert(false);
                        break;
                }
            }

            // Append a newline
            pbBufferPtr = AppendNewLine(pbBufferPtr, pbBufferEnd);

            // Close the MSI record
            MSI_CLOSE_HANDLE(hMsiRecord);
        }

        // Finalize the executed view
        MsiViewClose(hMsiView);
    }

    // Give the file size to the caller
    if(dwErrCode == ERROR_SUCCESS)
        PtrFileSize[0] = (DWORD)(pbBufferPtr - pbBufferBegin);
    return dwErrCode;
}

DWORD TMsiFile::LoadFileInternal(LPDWORD PtrFileSize)
{
    DWORD dwFileSize = 0;
    DWORD dwErrCode = ERROR_NOT_SUPPORTED;

    // Is there a referenced file?
    if(m_pRefFile != NULL)
        return m_pRefFile->LoadFileInternal(PtrFileSize);

    // File-type-specific
    switch(m_FileType)
    {
        case MsiFileSummary:
            dwErrCode = LoadSummaryFile(&dwFileSize);
            break;

        case MsiFileBinary:
            dwErrCode = LoadBinaryFile(&dwFileSize);
            break;

        case MsiFileTable:
            dwErrCode = LoadCsvFile(&dwFileSize);
            break;

        default:
            dwErrCode = ERROR_NOT_SUPPORTED;
            assert(false);
            break;
    }

    // Give the file size to the caller
    if(dwErrCode == ERROR_SUCCESS)
    {
        if(PtrFileSize != NULL)
            PtrFileSize[0] = dwFileSize;
        else
            m_dwFileSize = dwFileSize;
    }
    return dwErrCode;
}

DWORD TMsiFile::LoadFileData()
{
    DWORD dwErrCode = ERROR_SUCCESS;

    // Is there a referenced file?
    if(m_pRefFile != NULL)
        return m_pRefFile->LoadFileData();

    // Are the data already there?
    if(m_Data.cbData < m_dwFileSize)
    {
        if((dwErrCode = m_Data.Reserve(m_dwFileSize)) == ERROR_SUCCESS)
        {
            dwErrCode = LoadFileInternal(&m_Data.cbData);
        }
    }
    return dwErrCode;
}

void TMsiFile::MakeItemNameFileSafe(std::tstring & strItemName)
{
    for(size_t i = 0; i < strItemName.size(); i++)
    {
        if(0 <= strItemName[i] && strItemName[i] < 0x20)
        {
            strItemName[i] = _T('_');
        }
    }
}

const MSI_BLOB & TMsiFile::FileData()
{
    return m_pRefFile ? m_pRefFile->FileData() : m_Data;
}

DWORD TMsiFile::FileSize()
{
    return m_pRefFile ? m_pRefFile->FileSize() : m_dwFileSize;
}

LPCTSTR TMsiFile::Name()
{
    return m_strName.c_str();
}

DWORD TMsiFile::SetUniqueFileName(TMsiDatabase * pMsiDb, LPCTSTR szFolderName, LPCTSTR szBaseName, LPCTSTR szExtension)
{
    LPTSTR szFileNameEnd;
    LPTSTR szFileNamePtr;
    TCHAR szFileName[MAX_PATH];
    DWORD dwNameIndex = 1;

    // Setup the range
    szFileNameEnd = szFileName + _countof(szFileName);
    szFileNamePtr = szFileName;

    // If we have directory, append it first
    if(szFolderName && szFolderName[0])
    {
        StringCchPrintfEx(szFileNamePtr, (szFileNameEnd - szFileNamePtr), &szFileNamePtr, NULL, 0, _T("%s\\"), szFolderName);
    }

    // Construct the file name without numeric prefix
    StringCchPrintf(szFileNamePtr, (szFileNameEnd - szFileNamePtr), _T("%s%s"), szBaseName, szExtension);
    while(pMsiDb->IsFilePresent(szFileName))
        StringCchPrintf(szFileName, _countof(szFileName), _T("%s_%03u%s"), szBaseName, dwNameIndex++, szExtension);

    // Assign the file name to the string
    m_strName.assign(szFileName);
    return ERROR_SUCCESS;
}
