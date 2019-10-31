#ifndef _http_h_
#define _http_h_

// buffer length.
#define HTTP_BUFSIZE    4096

// headers max.
#define HTTP_HEADMAX    32



struct http{
    char        m_a_[4];                                                                // for safety of recv process.
    char        m_aBuf[HTTP_BUFSIZE];
    char        m_a__[2];                                                               // for safety of parse process.
    
    union{
    const char* m_pMethod;
    const char* m_pResVer;
    };
    
    union{
    const char* m_pURI;
    const char* m_pCode;
    };
    
    union{
    const char* m_pVersion;
    const char* m_pReason;
    };
    
    const char* m_aHeaderN[HTTP_HEADMAX];
    const char* m_aHeaderV[HTTP_HEADMAX];
    int         m_nHeaders;
    
};



struct http*    http_reset(struct http* p);
int             http_parse(struct http* p, int len);
const char*     http_header(struct http* p, const char* pName);



#endif //_http_h_
