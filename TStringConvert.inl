/*****************************************************************************/
/* TStringConvert.inl                     Copyright (c) Ladislav Zezula 2023 */
/*---------------------------------------------------------------------------*/
/* Unicode layer for Total Commander plugin                                  */
/*---------------------------------------------------------------------------*/
/*   Date    Ver   Who  Comment                                              */
/* --------  ----  ---  -------                                              */
/* 04.07.23  1.00  Lad  Created                                              */
/*****************************************************************************/

//-----------------------------------------------------------------------------
// Class template for conversions UNICODE <--> UTF-8

template <typename SRCCHAR, typename TRGCHAR, UINT CodePage>
class TConvertString
{
    public:

    TConvertString(const SRCCHAR * szSrcStr, const SRCCHAR * szSrcEnd = NULL)
    {
        // By default, set to NULL string
        m_szStr = NULL;
        m_nLen = 0;

        if(szSrcStr != NULL)
        {
            // Set to the pointer to the internal buffer
            m_szStr = m_StaticBuffer;
            m_StaticBuffer[0] = 0;

            // Retrieve the length of the UTF-8 string
            if((m_nLen = SrcStrToTrgStr(szSrcStr, szSrcEnd, NULL, 0)) != 0)
            {
                // Allocate buffer if too long
                if((m_nLen + 1) > _countof(m_StaticBuffer))
                {
                    if((m_szStr = new TRGCHAR[m_nLen + 1]) == NULL)
                    {
                        assert(false);
                        return;
                    }
                }

                // Convert the string
                SrcStrToTrgStr(szSrcStr, szSrcEnd, m_szStr, m_nLen + 1);
            }
        }
    }

    ~TConvertString()
    {
        if(m_szStr != NULL && m_szStr != m_StaticBuffer)
            delete [] m_szStr;
        m_szStr = NULL;
        m_nLen = 0;
    }

    size_t SrcStrToTrgStr(LPCWSTR szSrcStr, LPCWSTR szSrcEnd, LPSTR szTrgStr, size_t nTrgStr)
    {
        size_t nLength = (szSrcEnd > szSrcStr) ? (szSrcEnd - szSrcStr) : INVALID_SIZE_T;
        size_t nResult;

        // Perform the conversion
        nResult = WideCharToMultiByte(CodePage, 0, szSrcStr, (int)(nLength), szTrgStr, (int)(nTrgStr), NULL, NULL);

        // Terminate with zero, if needed
        if(szSrcEnd && szSrcEnd[0] && szTrgStr)
            szTrgStr[nResult] = 0;
        return nResult;
    }

    size_t SrcStrToTrgStr(LPCSTR szSrcStr, LPCSTR szSrcEnd, LPWSTR szTrgStr, size_t nTrgStr)
    {
        size_t nLength = (szSrcEnd > szSrcStr) ? (szSrcEnd - szSrcStr) : INVALID_SIZE_T;
        size_t nResult;

        // Perform the conversion
        nResult = MultiByteToWideChar(CodePage, 0, szSrcStr, (int)(nLength), szTrgStr, (int)(nTrgStr));

        // Terminate with zero, if needed
        if(szSrcEnd && szSrcEnd[0] && szTrgStr)
            szTrgStr[nResult] = 0;
        return nResult;
    }

    TRGCHAR * Buffer()
    {
        return m_szStr;
    }

    size_t Length()
    {
        return m_nLen;
    }

    operator const TRGCHAR *()
    {
        return m_szStr;
    }

    protected:

    TRGCHAR * m_szStr;
    size_t m_nLen;
    TRGCHAR m_StaticBuffer[0x80];
};

//-----------------------------------------------------------------------------
// Classes for conversions UNICODE <--> UTF-8

typedef TConvertString<WCHAR, char,  CP_UTF8> TWideToUTF8;
typedef TConvertString<char,  WCHAR, CP_UTF8> TUTF8ToWide;
typedef TConvertString<char,  WCHAR, CP_ACP>  TAnsiToWide;
typedef TConvertString<WCHAR, char,  CP_ACP>  TWideToAnsi;
