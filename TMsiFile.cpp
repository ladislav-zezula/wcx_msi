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

//-----------------------------------------------------------------------------
// TMsiFile functions

TMsiFile::TMsiFile(TMsiTable * pMsiTable)
{
    InitializeListHead(&m_Entry);
    m_hMsiRecord = NULL;
    m_dwFileSize = 0;
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

    // Close the MSI record handle
    if(m_hMsiRecord != NULL)
        MSI_CLOSE_HANDLE(m_hMsiRecord);
    m_hMsiRecord = NULL;
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

DWORD TMsiFile::SetBinaryFile(TMsiDatabase * pMsiDb, MSIHANDLE hMsiRecord)
{
    TMsiFile * pRefFile;
    std::tstring strItemName;
    LPTSTR szExtension;
    DWORD dwNameIndex = 1;
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

            // Construct the unique file name
            StringCchPrintf(szFileName, _countof(szFileName), _T("%s\\%s%s"), m_pMsiTable->Name(), szBaseName, szFileExt);
            while(pMsiDb->IsFilePresent(szFileName))
            {
                StringCchPrintf(szFileName, _countof(szFileName), _T("%s\\%s_%03u%s"), m_pMsiTable->Name(), szBaseName, dwNameIndex++, szFileExt);
            }
        }
        else
        {
            // Reference the file only
            m_pRefFile->AddRef();
        }

        // Assign the file name and record handle
        m_strName.assign(szFileName);
        m_hMsiRecord = hMsiRecord;
        return ERROR_SUCCESS;
    }

    // Store the MSI record handle
    return ERROR_NOT_SUPPORTED;
}

DWORD TMsiFile::SetCsvFile(TMsiDatabase * pMsiDb)
{
    DWORD dwNameIndex = 1; 
    TCHAR szFileName[MAX_PATH];

    // Construct the default file name
    StringCchPrintf(szFileName, _countof(szFileName), _T("%s.csv"), m_pMsiTable->Name());

    // We need to verify whether the file is in the database already
    while(pMsiDb->IsFilePresent(szFileName))
        StringCchPrintf(szFileName, _countof(szFileName), _T("%s_%03u.csv"), m_pMsiTable->Name(), dwNameIndex++);
    m_strName.assign(szFileName);
    return ERROR_SUCCESS;
}

DWORD TMsiFile::LoadCsvFileData(LPDWORD PtrFileSize)
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
                        MsiRecordGetInteger(hMsiRecord, (UINT)(i + 1), strValue);
                        pbBufferPtr = AppendFieldString(pbBufferPtr, pbBufferEnd, strValue, i);
                        break;

                    case MsiTypeString:
                        MsiRecordGetString(hMsiRecord, (UINT)(i + 1), strValue);
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

    // Give the file size to the called
    if(PtrFileSize != NULL)
        PtrFileSize[0] = (DWORD)(pbBufferPtr - pbBufferBegin);
    return dwErrCode;
}

DWORD TMsiFile::LoadFileSize()
{
    // Is there a referenced file?
    if(m_pRefFile != NULL)
        return m_pRefFile->LoadFileSize();

    // Is there a MSI record with the stream? 
    if(m_hMsiRecord != NULL)
    {
        m_dwFileSize = MsiRecordDataSize(m_hMsiRecord, (UINT)(m_pMsiTable->m_nStreamColumn + 1));
        return ERROR_SUCCESS;
    }
    else
    {
        return LoadCsvFileData(&m_dwFileSize);
    }
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
            // 1) Stream-based files: Read the stream data all at once
            if(m_hMsiRecord != NULL)
            {
                dwErrCode = MsiRecordReadStream(m_hMsiRecord, (UINT)(m_pMsiTable->m_nStreamColumn + 1), (char *)(m_Data.pbData), &m_Data.cbData);
            }
            else
            {
                dwErrCode = LoadCsvFileData(&m_Data.cbData);
            }
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
