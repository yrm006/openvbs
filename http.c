#include "http.h"
#include <string.h>

#ifdef _WIN32
    #define strcasecmp _stricmp
#else
#endif



struct http* http_reset(struct http* p)
{
    p->m_a_[0] = 0;
    p->m_a_[1] = 0;
    p->m_a_[2] = 0;
    p->m_a_[3] = 0;
    p->m_a__[0] = 0;
    p->m_a__[1] = 0;

    p->m_pMethod = NULL;
    p->m_pURI = NULL;
    p->m_pVersion = NULL;
    p->m_nHeaders = 0;
    
    return p;
}



int http_parse(struct http* p, int len){
    int i=0;
    
    
    
    p->m_pMethod = &p->m_aBuf[i];
    while(i<len && p->m_aBuf[i]!=' ') ++i;
                                                                                        if(len<=i) goto emergency;
    p->m_aBuf[i] = '\0';
    i += 1;
    
    p->m_pURI = &p->m_aBuf[i];
    while(i<len && p->m_aBuf[i]!=' ') ++i;
                                                                                        if(len<=i) goto emergency;
    p->m_aBuf[i] = '\0';
    i += 1;
    
    p->m_pVersion = &p->m_aBuf[i];
    while(i<len && p->m_aBuf[i]!='\r') ++i;
                                                                                        if(len<=i) goto emergency;
    p->m_aBuf[i] = '\0';
    i += 2;
    
    p->m_nHeaders = 0;
    while(i<len && p->m_aBuf[i]!='\r'){
                                                                                        // NOT SUPPORTED folded header.
                                                                                        if(p->m_aBuf[i]==' ' || p->m_aBuf[i]=='\t') goto emergency;
        
        p->m_aHeaderN[p->m_nHeaders] = &p->m_aBuf[i];
        while(i<len && p->m_aBuf[i]!=':') ++i;
                                                                                        if(len<=i) goto emergency;
        p->m_aBuf[i] = '\0';
        i += 1;
        
        while(i<len && p->m_aBuf[i]==' ') ++i;
        
        p->m_aHeaderV[p->m_nHeaders] = &p->m_aBuf[i];
        while(i<len && p->m_aBuf[i]!='\r') ++i;
                                                                                        if(len<=i) goto emergency;
        p->m_aBuf[i] = '\0';
        i += 2;
        
        ++p->m_nHeaders;
                                                                                        if(HTTP_HEADMAX <= p->m_nHeaders) goto emergency;
    }
                                                                                        if(len<=i) goto emergency;
    p->m_aBuf[i] = '\0';
    i += 2;                                                                             // This is safe, because m_aBuf is larger than HTTP_BUFSIZE 2bytes.
    
    
    
    return i;
    
emergency:
    p->m_pMethod = NULL;
    p->m_pURI = NULL;
    p->m_pVersion = NULL;
    p->m_nHeaders = 0;
    
    return -i;
}



const char* http_header(struct http* p, const char* pName){
    int i;
    for(i=0; i<p->m_nHeaders; ++i){
        if(strcasecmp(p->m_aHeaderN[i], pName) == 0){
            return p->m_aHeaderV[i];                                                    // NOT SUPPORTED same named header.
        }
    }
    
    return "";
}



