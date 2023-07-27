/*****************************************************************************/
/* TMsiTable.cpp                          Copyright (c) Ladislav Zezula 2023 */
/*---------------------------------------------------------------------------*/
/* Implementation of the TMsiDatabase class methods                          */
/*---------------------------------------------------------------------------*/
/*   Date    Ver   Who  Comment                                              */
/* --------  ----  ---  -------                                              */
/* 24.07.23  1.00  Lad  Created                                              */
/*****************************************************************************/

#include "wcx_msi.h"

//-----------------------------------------------------------------------------
// TMsiColumn constructor

TMsiColumn::TMsiColumn(LPCTSTR szName, LPCTSTR szType)
{
    LPTSTR szEndPtr = NULL;
    TCHAR chLowerChar = (TCHAR)(DWORD_PTR)(CharLower((LPTSTR)(szType[0])));

    // Setup type and size
    m_Type = MsiTypeUnknown;
    m_Size = 0;

    // Assign name and type
    m_strName.assign(szName);
    m_strType.assign(szType);

    // Parse the type info
    switch(chLowerChar)
    {
        case 'i':
            m_Type = MsiTypeInteger;
            m_Size = 4;
            break;

        case 's':
        case 'l':
            m_Type = MsiTypeString;
            m_Size = 0;
            break;

        case 'v':
            m_Type = MsiTypeStream;
            m_Type = MsiTypeStream;
            m_Size = 0;
            break;

        default:            // Should never happen
            assert(false);
            break;
    }

    // If there is a number in the type info, extract the number
    if(isdigit(szType[1]))
    {
        m_Size = _tcstoul(szType + 1, &szEndPtr, 10);
    }
}

//-----------------------------------------------------------------------------
// TMsiTable functions

TMsiTable::TMsiTable(TMsiDatabase * pMsiDb, const std::tstring & strName, MSIHANDLE hMsiView)
{
    InitializeListHead(&m_Entry);
    m_strName = strName;
    m_hMsiView = hMsiView;
    m_nStreamColumn = INVALID_SIZE_T;
    m_nNameColumn = INVALID_SIZE_T;
    m_bIsStreamsTable = FALSE;
    m_dwRefs = 1;

    // Check for the "_Streams" table
    if(strName == _T("_Streams"))
        m_bIsStreamsTable = TRUE;

    // Reference the MSI database
    if((m_pMsiDb = pMsiDb) != NULL)
    {
        m_pMsiDb->AddRef();
    }
}

TMsiTable::~TMsiTable()
{
    // Sanity check
    assert(m_dwRefs == 0);

    // Release the database
    if(m_pMsiDb != NULL)
        m_pMsiDb->Release();
    m_pMsiDb = NULL;

    // Close the view handle
    if(m_hMsiView != NULL)
        MSI_CLOSE_HANDLE(m_hMsiView);
    m_hMsiView = NULL;
}

//-----------------------------------------------------------------------------
// TMsiTable methods

DWORD TMsiTable::AddRef()
{
    return InterlockedIncrement((LONG *)(&m_dwRefs));
}

DWORD TMsiTable::Release()
{
    if(InterlockedDecrement((LONG *)(&m_dwRefs)) == 0)
    {
        delete this;
        return 0;
    }
    return m_dwRefs;
}

DWORD TMsiTable::Load()
{
    DWORD dwErrCode;

    // Load the columns
    if((dwErrCode = LoadColumns()) == ERROR_SUCCESS)
    {
        // Check if we have a stream column
        for(size_t i = 0; i < m_Columns.size(); i++)
        {
            if(m_Columns[i].m_Type == MsiTypeStream)
            {
                m_nStreamColumn = i;
                break;
            }
        }

        // If we have a stream column, take the first string as name
        if(m_nStreamColumn != INVALID_SIZE_T)
        {
            for(size_t i = 0; i < m_Columns.size(); i++)
            {
                if(m_Columns[i].m_Type == MsiTypeString)
                {
                    m_nNameColumn = i;
                    break;
                }
            }
        }
    }
    return dwErrCode;
}

DWORD TMsiTable::LoadColumns()
{
    MSIHANDLE hMsiColTypes = NULL;
    MSIHANDLE hMsiColNames = NULL;
    DWORD dwErrCode;
    UINT uColumns1 = 0;
    UINT uColumns2 = 0;

    // Retrieve types
    if((dwErrCode = MsiViewGetColumnInfo(m_hMsiView, MSICOLINFO_TYPES, &hMsiColTypes)) == ERROR_SUCCESS)
    {
        // Log the handle for diagnostics
        MSI_LOG_OPEN_HANDLE(hMsiColTypes);

        // Retrieve names
        if((dwErrCode = MsiViewGetColumnInfo(m_hMsiView, MSICOLINFO_NAMES, &hMsiColNames)) == ERROR_SUCCESS)
        {
            // Log the handle for diagnostics
            MSI_LOG_OPEN_HANDLE(hMsiColNames);

            // Retrieve column counts
            uColumns1 = MsiRecordGetFieldCount(hMsiColTypes);
            uColumns2 = MsiRecordGetFieldCount(hMsiColNames);
            if(uColumns1 == uColumns2)
            {
                // Insert all columns
                for(UINT i = 0; i < uColumns1; i++)
                {
                    TCHAR szColumnName[128];
                    TCHAR szColumnType[128];
                    DWORD ccColumnName = _countof(szColumnName);
                    DWORD ccColumnType = _countof(szColumnType);

                    MsiRecordGetString(hMsiColNames, i + 1, szColumnName, &ccColumnName);
                    MsiRecordGetString(hMsiColTypes, i + 1, szColumnType, &ccColumnType);

                    if(ccColumnName && ccColumnType)
                    {
                        m_Columns.push_back(TMsiColumn(szColumnName, szColumnType));
                    }
                }
            }
            MSI_CLOSE_HANDLE(hMsiColNames);
        }
        MSI_CLOSE_HANDLE(hMsiColTypes);
    }
    return dwErrCode;
}
