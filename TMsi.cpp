/*****************************************************************************/
/* TMsi.cpp                               Copyright (c) Ladislav Zezula 2023 */
/*---------------------------------------------------------------------------*/
/* Global (non-class) MSI functions                                          */
/*---------------------------------------------------------------------------*/
/*   Date    Ver   Who  Comment                                              */
/* --------  ----  ---  -------                                              */
/* 27.07.23  1.00  Lad  Created                                              */
/*****************************************************************************/

#include "wcx_msi.h"

//-----------------------------------------------------------------------------
// Diagnostic functions

#ifdef _DEBUG

struct MSI_HANDLE_INFO
{
    MSI_HANDLE_INFO(MSIHANDLE hHandle, LPCSTR szFile, int nLine)
    {
        m_hHandle = hHandle;
        m_szFile = szFile;
        m_nLine = nLine;
    }

    MSI_HANDLE_INFO()
    {
        m_hHandle = NULL;
        m_szFile = NULL;
        m_nLine = 0;
    }

    MSIHANDLE m_hHandle;
    LPCSTR m_szFile;
    int m_nLine;
};

static std::vector<MSI_HANDLE_INFO> MsiHandleList;

void MSI_LOG_OPEN_HANDLE_EX(MSIHANDLE hMsiHandle, LPCSTR szFile, int nLine)
{
    // Turn the handle into index
    size_t nIndex = (size_t)(hMsiHandle);

    // Make sure we have enough entries in the list
    MsiHandleList.resize(nIndex + 1);
    MsiHandleList[nIndex] = MSI_HANDLE_INFO(hMsiHandle, szFile, nLine);

    // Log the opening of the handle
    //Dbg(_T("[*] MSI Handle opened: %p\n"), hMsiHandle);
}

UINT MSI_CLOSE_HANDLE(MSIHANDLE hMsiHandle)
{
    // Turn the handle into index
    size_t nIndex = (size_t)(hMsiHandle);

    // Set the entry to false
    if(MsiHandleList.size() > nIndex)
        MsiHandleList[nIndex].m_hHandle = NULL;
    else
        Dbg(_T("[x] Unknown MSI handle: %p\n"), hMsiHandle);

    // Log closing the handle and close it
    //Dbg(_T("[*] MSI Handle closed: %p\n"), hMsiHandle);
    return MsiCloseHandle(hMsiHandle);
}

UINT MSI_DUMP_HANDLES()
{
    UINT nOpenHandles = 0;

    for(size_t i = 0; i < MsiHandleList.size(); i++)
    {
        if(MsiHandleList[i].m_hHandle != NULL)
        {
            const MSI_HANDLE_INFO & MsiHandleInfo = MsiHandleList[i];

            Dbg(_T("[*] MSI handle open: %p, created in %hs(%u)\n"),
                MsiHandleInfo.m_hHandle,
                MsiHandleInfo.m_szFile,
                MsiHandleInfo.m_nLine);
            nOpenHandles++;
        }
    }
    return nOpenHandles;
}
#endif

//-----------------------------------------------------------------------------
// Local (non-class) functions

bool MsiRecordGetInteger(MSIHANDLE hMsiRecord, UINT nColumn, std::tstring & strValue)
{
    LPTSTR szEndString;
    TCHAR szIntValue[32];
    int nIntValue = MsiRecordGetInteger(hMsiRecord, nColumn + 1);

    if((nIntValue = MsiRecordGetInteger(hMsiRecord, nColumn + 1)) != MSI_NULL_INTEGER)
    {
        StringCchPrintfEx(szIntValue, _countof(szIntValue), &szEndString, NULL, 0, _T("%i"), nIntValue);
        strValue.assign(szIntValue, (size_t)(szEndString - szIntValue));
    }
    else
    {
        strValue.assign(_T("(null)"), 6);
    }
    return true;
}

bool MsiRecordGetString(MSIHANDLE hMsiRecord, UINT nColumn, std::tstring & strValue)
{
    #define STATIC_BUFFER_SIZE  0x80
    DWORD ccValue = 0;

    // Retrieve the length of the string without the ending EOS
    if(MsiRecordGetString(hMsiRecord, nColumn + 1, NULL, &ccValue) == ERROR_SUCCESS)
    {
        // Include EOS in the buffer
        ccValue++;

        // Can we retrieve the value without allocating buffer?
        if(ccValue > STATIC_BUFFER_SIZE)
        {
            LPTSTR szBuffer;

            if((szBuffer = new TCHAR[ccValue]) != NULL)
            {
                MsiRecordGetString(hMsiRecord, nColumn + 1, szBuffer, &ccValue);
                strValue.assign(szBuffer, ccValue);
                delete[] szBuffer;
                return true;
            }
        }
        else
        {
            TCHAR szBuffer[STATIC_BUFFER_SIZE];

            MsiRecordGetString(hMsiRecord, nColumn + 1, szBuffer, &ccValue);
            strValue.assign(szBuffer, ccValue);
            return true;
        }
    }
    return false;
}

bool MsiRecordGetBinary(MSIHANDLE hMsiRecord, UINT nColumn, MSI_BLOB & binValue)
{
    LPBYTE pbValue = NULL;
    DWORD cbValue = 0;

    // Retrieve the length of the string
    if((cbValue = MsiRecordDataSize(hMsiRecord, nColumn + 1)) > 0)
    {
        if((pbValue = new BYTE[cbValue]) != NULL)
        {
            MsiRecordReadStream(hMsiRecord, nColumn + 1, (char *)(pbValue), &cbValue);
        }
    }

    // Give the result
    binValue.pbData = pbValue;
    binValue.cbData = cbValue;
    return (binValue.pbData != NULL);
}
