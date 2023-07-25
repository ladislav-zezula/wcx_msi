/*****************************************************************************/
/* Config.cpp                             Copyright (c) Ladislav Zezula 2023 */
/*---------------------------------------------------------------------------*/
/* Functions for loading, editing and saving configuration of wcx_msi        */
/*---------------------------------------------------------------------------*/
/*   Date    Ver   Who  Comment                                              */
/* --------  ----  ---  -------                                              */
/* 23.07.23  1.00  Lad  Created                                              */
/*****************************************************************************/

#include "wcx_msi.h"
#include "resource.h"

//-----------------------------------------------------------------------------
// Global variables

TConfiguration g_cfg;
TCHAR g_szIniFile[MAX_PATH] = _T("");

//-----------------------------------------------------------------------------
// Local variables

static LPCTSTR szCfgSection = _T("wcx_msi");

//-----------------------------------------------------------------------------
// Public functions

//static BOOL WritePrivateProfileInt(LPCTSTR szSection, LPCTSTR szKeyName, UINT uValue, LPCTSTR szIniFile)
//{
//    TCHAR szIntValue[32];
//
//    StringCchPrintf(szIntValue, _countof(szIntValue), _T("%u"), uValue);
//    return WritePrivateProfileString(szSection, szKeyName, szIntValue, szIniFile);
//}

int SetDefaultConfiguration()
{
    // By default, all is zero or false
    ZeroMemory(&g_cfg, sizeof(TConfiguration));
    return ERROR_SUCCESS;
}

int LoadConfiguration()
{
    return ERROR_SUCCESS;
}

int SaveConfiguration()
{
    return ERROR_SUCCESS;
}
