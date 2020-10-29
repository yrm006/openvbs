#include <stdio.h>
#include <stdint.h>
#include <string>
#include <vector>
#include <locale>

#ifdef _WIN32
 #include <winsock2.h>
 
 #include "websocket.h"
 #include "cocoserve.h"
#else
 #include <arpa/inet.h>
 #include <unistd.h>
 #include <pthread.h>

 extern "C"{
 #include "websocket.h"
 #include "cocoserve.h"
 }
#endif

#include "jujube.h"

#define LOG(m,...)        fprintf(stdout,m"\n",##__VA_ARGS__)
#define LO_(...)          
#define CHECK(e,m,...)    if(e)fprintf(stderr,m"\n",##__VA_ARGS__),exit(errno)



#ifdef _WIN32
    #define struct       
    #define usleep(m)    Sleep(m/1000)
    
    void WSA_STARTUP(){
        WORD wVersionRequested = MAKEWORD(2, 2);
        WSADATA wsaData;
        int n = WSAStartup(wVersionRequested, &wsaData);
                                                                                    CHECK(n != 0, "!WSAStartup");
    }
    void WSA_CLEANUP(){
        WSACleanup();
    }

    #define pthread_mutex_t          CRITICAL_SECTION
    #define pthread_mutex_init(a,b)  InitializeCriticalSection(a)
    #define pthread_mutex_destroy(a) DeleteCriticalSection(a)
    #define pthread_mutex_lock(a)    EnterCriticalSection(a)
    #define pthread_mutex_unlock(a)  LeaveCriticalSection(a)
#else
    #define WSA_STARTUP(...)
    #define WSA_CLEANUP(...)

    #define strcmpi strcasecmp
    #define _snwprintf swprintf
#endif



IClassFactory* CFVBScript  = nullptr;









//---------------------------
//  IchigoPage
//---------------------------

class JPO : public IDispatch{
    ULONG       m_refc   = 1;

    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){ return E_NOTIMPL; }
public:
    JPO(){/*wprintf(L"%s\n", __func__);*/}
    virtual ~JPO(){}
};

class JPO_LINK : public JPO{
    _bstr_t m_href;

    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        std::wstring buf;

        buf += L"<a href='";
        buf += m_href;
        buf += L"'>";
        if(pv->vt == VT_BSTR){
            buf += pv->bstrVal;
        }else
        if(pv->vt == VT_DISPATCH){
            DISPPARAMS param = { nullptr, nullptr, 0, 0 };
            _variant_t res;
            pv->pdispVal->Invoke(DISPID_VALUE, IID_NULL, 0, DISPATCH_METHOD, &param, &res, nullptr, nullptr);
            if(res.vt == VT_BSTR){
                buf += res.bstrVal;
            }
        }
        buf += L"</a>";

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = ::SysAllocString(buf.c_str());

        return S_OK;
    }

public:
    JPO_LINK(const wchar_t* href)
        : m_href(href)
    {}
    virtual ~JPO_LINK(){}
};

class JPO_IMG : public JPO{
    _bstr_t m_src;

    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        std::wstring buf;

        buf += L"<img src='";
        buf += m_src;
        buf += L"'>";

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = ::SysAllocString(buf.c_str());

        return S_OK;
    }

public:
    JPO_IMG(const wchar_t* src)
        : m_src(src)
    {}
    virtual ~JPO_IMG(){}
};

enum COLOR{
    C_RED,
    C_GREEN,
    C_BLUE,
};

const wchar_t* g_colors[] = {
    L"red",
    L"green",
    L"blue",
};

template<COLOR color>
class JPO_COLOR : public JPO{
    _variant_t m_contents;

    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        std::wstring buf;

        buf += L"<span style='color:" + std::wstring(g_colors[color]) + L"; '>";
        _variant_t v;
        if(::VariantChangeType(&v, &m_contents, 0, VT_BSTR) == S_OK){
            buf += v.bstrVal;
        }else
        if(m_contents.vt == VT_DISPATCH){
            DISPPARAMS param = { nullptr, nullptr, 0, 0 };
            _variant_t res;
            m_contents.pdispVal->Invoke(DISPID_VALUE, IID_NULL, 0, DISPATCH_METHOD, &param, &res, nullptr, nullptr);
            if(res.vt == VT_BSTR){
                buf += res.bstrVal;
            }
        }
        buf += L"</span>";

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = ::SysAllocString(buf.c_str());

        return S_OK;
    }

public:
    JPO_COLOR(_variant_t&& contents)
        : m_contents(contents)
    {}
    virtual ~JPO_COLOR(){}
};

enum FSIZE{
    SZ_SMALL,
    SZ_LARGE,
};

const wchar_t* g_sizes[] = {
    L"xx-small",
    L"xx-large",
};

template<FSIZE size>
class JPO_SIZE : public JPO{
    _variant_t m_contents;

    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        std::wstring buf;

        buf += L"<span style='font-size:" + std::wstring(g_sizes[size]) + L"; '>";
        _variant_t v;
        if(::VariantChangeType(&v, &m_contents, 0, VT_BSTR) == S_OK){
            buf += v.bstrVal;
        }else
        if(m_contents.vt == VT_DISPATCH){
            DISPPARAMS param = { nullptr, nullptr, 0, 0 };
            _variant_t res;
            m_contents.pdispVal->Invoke(DISPID_VALUE, IID_NULL, 0, DISPATCH_METHOD, &param, &res, nullptr, nullptr);
            if(res.vt == VT_BSTR){
                buf += res.bstrVal;
            }
        }
        buf += L"</span>";

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = ::SysAllocString(buf.c_str());

        return S_OK;
    }

public:
    JPO_SIZE(_variant_t&& contents)
        : m_contents(contents)
    {}
    virtual ~JPO_SIZE(){}
};

enum TALIGN{
    AL_LEFT,
    AL_CENTER,
    AL_RIGHT,
};

const wchar_t* g_aligns[] = {
    L"left",
    L"center",
    L"right",
};

template<TALIGN align>
class JPO_ALIGN : public JPO{
    _variant_t m_contents;

    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        std::wstring buf;

        buf += L"<pre style='text-align:" + std::wstring(g_aligns[align]) + L"; '>";
        _variant_t v;
        if(::VariantChangeType(&v, &m_contents, 0, VT_BSTR) == S_OK){
            buf += v.bstrVal;
        }else
        if(m_contents.vt == VT_DISPATCH){
            DISPPARAMS param = { nullptr, nullptr, 0, 0 };
            _variant_t res;
            m_contents.pdispVal->Invoke(DISPID_VALUE, IID_NULL, 0, DISPATCH_METHOD, &param, &res, nullptr, nullptr);
            if(res.vt == VT_BSTR){
                buf += res.bstrVal;
            }
        }
        buf += L"</pre>";

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = ::SysAllocString(buf.c_str());

        return S_OK;
    }

public:
    JPO_ALIGN(_variant_t&& contents)
        : m_contents(contents)
    {}
    virtual ~JPO_ALIGN(){}
};



class IchigoPage : public IDispatch{
private:
    enum PMODE{ PM_NONE, PM_CRLF };

public:
    std::wstring               resbuf;

public:
    IchigoPage(){}
    virtual ~IchigoPage(){}

private:
    ULONG       m_refc   = 1;
    PMODE       m_pmode  = PM_CRLF;

    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(
            _wcsicmp(*rgszNames, L"print") == 0 ||
            _wcsicmp(*rgszNames, L"?") == 0     ||
        false){
            *rgDispId = 1;
        }else
        if(_wcsicmp(*rgszNames, L";") == 0){
            *rgDispId = 2;
        }else
        if(_wcsicmp(*rgszNames, L"LINK") == 0){
            *rgDispId = 3;
        }else
        if(_wcsicmp(*rgszNames, L"IMG") == 0){
            *rgDispId = 4;
        }else
        if(_wcsicmp(*rgszNames, L"RED") == 0){
            *rgDispId = 101;
        }else
        if(_wcsicmp(*rgszNames, L"GREEN") == 0){
            *rgDispId = 102;
        }else
        if(_wcsicmp(*rgszNames, L"BLUE") == 0){
            *rgDispId = 103;
        }else
        if(_wcsicmp(*rgszNames, L"SMALL") == 0){
            *rgDispId = 201;
        }else
        if(_wcsicmp(*rgszNames, L"LARGE") == 0){
            *rgDispId = 202;
        }else
        if(_wcsicmp(*rgszNames, L"LEFT") == 0){
            *rgDispId = 301;
        }else
        if(_wcsicmp(*rgszNames, L"CENTER") == 0){
            *rgDispId = 302;
        }else
        if(_wcsicmp(*rgszNames, L"RIGHT") == 0){
            *rgDispId = 303;
        }else
        {
            return E_FAIL;
        }

        return S_OK;
    }
    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(dispIdMember == 1){//print,?
            int i = pDispParams->cArgs;
            while(0 <= --i){
                VARIANT* pv = pDispParams->rgvarg + i;
                if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

                _variant_t v;
                if(pv->vt == VT_BSTR && pv->bstrVal[0] == L'\b'){
                    m_pmode = PM_NONE;
                }else
                if(::VariantChangeType(&v, pv, 0, VT_BSTR) == S_OK){
                    resbuf += v.bstrVal;
                }else{
                    wchar_t buf[0x100];
                    swprintf(buf, 0x100, L"VARIANT:0x%x[0x%llx]", pv->vt, pv->llVal);
                    resbuf += buf;
                }
            }

            if(m_pmode == PM_CRLF){
                resbuf += L"\r\n";
            }else
            {
                m_pmode = PM_CRLF;
            }

            return S_OK;
        }else
        if(dispIdMember == 2){//;
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            pVarResult->vt = VT_BSTR;
            pVarResult->bstrVal = ::SysAllocString(L"\b");

            return S_OK;
        }else
        if(dispIdMember == 3){//link
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = new JPO_LINK(pv->bstrVal);

            return S_OK;
        }else
        if(dispIdMember == 4){//img
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = new JPO_IMG(pv->bstrVal);

            return S_OK;
        }else
        if(dispIdMember == 101){//red
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = new JPO_COLOR<C_RED>(*pv);

            return S_OK;
        }else
        if(dispIdMember == 102){//green
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = new JPO_COLOR<C_GREEN>(*pv);

            return S_OK;
        }else
        if(dispIdMember == 103){//blue
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = new JPO_COLOR<C_BLUE>(*pv);

            return S_OK;
        }else
        if(dispIdMember == 201){//small
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = new JPO_SIZE<SZ_SMALL>(*pv);

            return S_OK;
        }else
        if(dispIdMember == 202){//large
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = new JPO_SIZE<SZ_LARGE>(*pv);

            return S_OK;
        }else
        if(dispIdMember == 301){//left
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = new JPO_ALIGN<AL_LEFT>(*pv);

            return S_OK;
        }else
        if(dispIdMember == 302){//center
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = new JPO_ALIGN<AL_CENTER>(*pv);

            return S_OK;
        }else
        if(dispIdMember == 303){//right
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = new JPO_ALIGN<AL_RIGHT>(*pv);

            return S_OK;
        }else
        {
            return E_FAIL;
        }
    }
};



//---------------------------
//  OpenASP
//---------------------------

class Contents : public IDispatch{
private:
    ULONG       m_refc   = 1;

protected:
    std::map<_variant_t, _variant_t> m_map;
    pthread_mutex_t                  m_mtx;

public:
    Contents(){/*wprintf(L"%s\n", __func__);*/
        pthread_mutex_init(&m_mtx, NULL);
    }
    virtual ~Contents(){/*wprintf(L"%s\n", __func__);*/
        pthread_mutex_destroy(&m_mtx);
    }

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(_wcsicmp(*rgszNames, L"Count") == 0){
            *rgDispId = 1;
        }else
        {
wprintf(L"###%s: Implement here '%s' line %d. (%ls)\n", __func__, __FILE__, __LINE__, *rgszNames);
            return DISP_E_MEMBERNOTFOUND;
        }

        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(dispIdMember == 0){
            if(wFlags == DISPATCH_PROPERTYPUT){
                VARIANT* pv0 = &pDispParams->rgvarg[0];
                if(pv0->vt == (VT_BYREF|VT_VARIANT)) pv0 = pv0->pvarVal;
                VARIANT* pv1 = &pDispParams->rgvarg[1];
                if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
                
                pthread_mutex_lock(&m_mtx);
                    m_map[*pv1] = *pv0;
                pthread_mutex_unlock(&m_mtx);
            }else{
                pthread_mutex_lock(&m_mtx);
                    VariantCopy(pVarResult, &m_map[pDispParams->rgvarg[0]]);
                pthread_mutex_unlock(&m_mtx);
            }
        }else
        if(dispIdMember == 1){
            pthread_mutex_lock(&m_mtx);
                pVarResult->vt = VT_I8;
                pVarResult->llVal = m_map.size();
            pthread_mutex_unlock(&m_mtx);
        }else
        {
            return DISP_E_MEMBERNOTFOUND;
        }

        return S_OK;
    }
};

class Server : public IDispatch{
public:
    Server(){}
    virtual ~Server(){}

private:
    ULONG       m_refc   = 1;

    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
wprintf(L"###%s: Implement here '%s' line %d. (%ls)\n", __func__, __FILE__, __LINE__, *rgszNames);
        return DISP_E_MEMBERNOTFOUND;
    }
    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        return DISP_E_MEMBERNOTFOUND;
    }
};

class Application : public IDispatch{
public:
    Application(){
        m_pCtnts = new Contents();
    }
    virtual ~Application(){
        m_pCtnts->Release();
    }

private:
    ULONG       m_refc   = 1;

    Contents*   m_pCtnts;

    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(_wcsicmp(*rgszNames, L"Contents") == 0){
            *rgDispId = 1;
        }else
        {
wprintf(L"###%s: Implement here '%s' line %d. (%ls)\n", __func__, __FILE__, __LINE__, *rgszNames);
            return DISP_E_MEMBERNOTFOUND;
        }

        return S_OK;
    }
    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(dispIdMember == 0){
            return m_pCtnts->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
        }else
        if(dispIdMember == 1){
            pVarResult->vt = VT_DISPATCH;
            (pVarResult->pdispVal = m_pCtnts)->AddRef();
            return S_OK;
        }

        return DISP_E_MEMBERNOTFOUND;
    }
};

class Session : public IDispatch{
public:
    Session(){}
    virtual ~Session(){}

private:
    ULONG       m_refc   = 1;

    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
wprintf(L"###%s: Implement here '%s' line %d. (%ls)\n", __func__, __FILE__, __LINE__, *rgszNames);
        return DISP_E_MEMBERNOTFOUND;
    }
    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        return DISP_E_MEMBERNOTFOUND;
    }
};

class Request : public IDispatch{
public:
    Request(){}
    virtual ~Request(){}

private:
    ULONG       m_refc   = 1;

    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
wprintf(L"###%s: Implement here '%s' line %d. (%ls)\n", __func__, __FILE__, __LINE__, *rgszNames);
        return DISP_E_MEMBERNOTFOUND;
    }
    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        return DISP_E_MEMBERNOTFOUND;
    }
};

class Response : public IDispatch{
public:
    std::wstring               resbuf;
    std::vector<std::wstring>& soto;

public:
    Response(std::vector<std::wstring>& rsoto)
        : soto(rsoto)
    {}
    virtual ~Response(){}

private:
    ULONG       m_refc   = 1;

    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(_wcsicmp(*rgszNames, L"WriteBlock") == 0){
            *rgDispId = 1;
        }else
        if(_wcsicmp(*rgszNames, L"Write") == 0){
            *rgDispId = 2;
        }else
        {
wprintf(L"###%s: Implement here '%s' line %d. (%ls)\n", __func__, __FILE__, __LINE__, *rgszNames);
            return DISP_E_MEMBERNOTFOUND;
        }

        return S_OK;
    }
    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(dispIdMember == 1){
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            resbuf += soto[pv->lVal];

            return S_OK;
        }else
        if(dispIdMember == 2){
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            _variant_t v;
            if(::VariantChangeType(&v, pv, 0, VT_BSTR) == S_OK){
                resbuf += v.bstrVal;
            }else{
                wchar_t buf[0x100];
                swprintf(buf, 0x100, L"VARIANT:0x%x[0x%llx]", pv->vt, pv->llVal);
                resbuf += buf;
            }

            return S_OK;
        }

        return E_FAIL;
    }
};

Server      g_oServer;
Application g_oApplication;



//---------------------------
//  WebSock
//---------------------------

class Client : public IDispatch{
public:
    SOCKET    m_sock;

    Client(SOCKET sock)
        : m_sock(sock)
    {}
    virtual ~Client(){}

private:
    ULONG       m_refc   = 1;

    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(_wcsicmp(*rgszNames, L"Send") == 0){
            *rgDispId = 1;
        }else
        {
            return E_FAIL;
        }

        return S_OK;
    }
    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(dispIdMember == 1){
            VARIANT* pv = pDispParams->rgvarg;
            if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

            _variant_t v;
            if(::VariantChangeType(&v, pv, 0, VT_BSTR) == S_OK){
                size_t n = wchar_utf8(nullptr, 0, v.bstrVal);
                char* s = (char*)malloc(n);
                    wchar_utf8(s, n, v.bstrVal);
                    ws_send(m_sock, s, n-1, 0, NULL, WS_OP_TEXT);
                free(s);
            }else{
                wchar_t buf[0x100];
                swprintf(buf, 0x100, L"VARIANT:0x%x[0x%llx]", pv->vt, pv->llVal);

                size_t n = wchar_utf8(nullptr, 0, buf);
                char* s = (char*)malloc(n);
                    wchar_utf8(s, n, buf);
                    ws_send(m_sock, s, n-1, 0, NULL, WS_OP_TEXT);
                free(s);
            }

            return S_OK;
        }

        return E_FAIL;
    }
};



class Unit{
protected:
    class CVBScript{
    protected:
        IDispatch*  m_pVBScript;
    public:
        CVBScript() : m_pVBScript(nullptr){
            CFVBScript->CreateInstance(nullptr, IID_IDispatch, (void**)&m_pVBScript);
        }
        ~CVBScript(){
            if(m_pVBScript) m_pVBScript->Release();
        }
        IDispatch* operator &(){
            return m_pVBScript;
        }
    };

public:
    Client      m_oClient;
    CVBScript   m_oVBScript;
    CExtension  m_oExt;
    CProgram    m_oProgram;
    CProcessor  m_oProcessor;

    Unit(SOCKET sock, const wchar_t* pSource)
        : m_oClient(sock)
        , m_oVBScript()
        , m_oExt{
            { L"Server"     , (IDispatch*)&g_oServer },
            { L"Application", (IDispatch*)&g_oApplication },
            { L"Client"     , (IDispatch*)&m_oClient },
        }
        , m_oProgram(pSource)
        , m_oProcessor(&m_oProgram, &m_oVBScript, &m_oExt)
    {}

    ~Unit(){}
};









void on_start(struct sockaddr_in* bindto){
                                                                                    LO_("on_start");
    ccs_start(bindto);
}

void on_stop(){
                                                                                    LO_("on_stop");
    ccs_stop();
}



int  on_ws_handshake(SOCKET sock, context* ctx){
                                                                                    LOG("on_ws_handshake");
    int r = 0;

    // check URI
    int i = 0;
    while(
        ctx->hdr.m_pURI[i]                  &&
        !(
            ctx->hdr.m_pURI[i+0] == '.' &&
            ctx->hdr.m_pURI[i+1] == '.' &&
        1)                                  &&
    1) ++i;

    char aPath[256];
    sprintf(aPath, "wwwroot/%s", (ctx->hdr.m_pURI[i]) ? "" : &ctx->hdr.m_pURI[1]);

	FILE* f = fopen(aPath, "r");
    if(f){
        fclose(f);

        r = 1;
    }else{
        char aBuf[256];
        aBuf[sizeof(aBuf)-1] = '\0';                                                // For safety.

		int n = 0;
		n += snprintf(aBuf+n, sizeof(aBuf)-n,
			"HTTP/1.0 404 Not Found\r\n"
			"\r\n"
		);
		send(sock, aBuf, n, 0);
    }

    return r;
}

int  on_ws_connect(SOCKET sock, context* ctx){
                                                                                    {
                                                                                    uint32_t adr = ctx->from.sin_addr.s_addr;
                                                                                    char aBuf[3+1+3+1+3+1+3+1+5 +1];
                                                                                    sprintf(aBuf, "%d.%d.%d.%d:%d" 
                                                                                        , adr>> 0&0xff
                                                                                        , adr>> 8&0xff
                                                                                        , adr>>16&0xff
                                                                                        , adr>>24&0xff
                                                                                        , ctx->from.sin_port
                                                                                    );
                                                                                    LOG("on_ws_connect: %s %s", aBuf, ctx->hdr.m_pURI);
                                                                                    }
    int r = 0;

    std::wstring source;
    {
        // check URI will already ended at on_ws_handshake.
        char aPath[256];
        sprintf(aPath, "wwwroot/%s", &ctx->hdr.m_pURI[1]);

#ifdef _WIN32
        FILE* f = fopen(aPath, "rt,ccs=UTF-8");

        wchar_t readbuf[256];
        int readn;
        while( 0 < (readn = fread(readbuf, sizeof(readbuf[0]), 256, f)) ){
            source.append(readbuf, readbuf+readn);
        }
#else
        FILE* f = fopen(aPath, "r");

        std::string utf8;
        char readbuf[256];
        int readn;
        while( 0 < (readn = fread(readbuf, sizeof(readbuf[0]), 256, f)) ){
            utf8.append(readbuf, readbuf+readn);
        }

        size_t n = utf8_wchar(nullptr, 0, utf8.c_str());
        wchar_t* buf = (wchar_t*)malloc( sizeof(wchar_t)*n );
            utf8_wchar(buf, n, utf8.c_str());
            source += buf;
        free(buf);
#endif

        fclose(f);
    }



    // run
    Unit* pu = new Unit(sock, source.c_str());

    HRESULT hr = S_OK;
    CoInitializeEx(0, COINIT_MULTITHREADED);
    {
        hr = pu->m_oProcessor();

        if(FAILED(hr)){
            fwprintf(stderr, L"![0x%x]%ls in line:%d\n", hr, pu->m_oProcessor.m_err->bstrDescription, pu->m_oProcessor.m_err->dwHelpContext);
        }
    }
    CoUninitialize();

    if(FAILED(hr)){
                                                                            // char a[0x10];
                                                                            // strncpy(a, e->code, e->len)[e->len] = '\0';
                                                                            // printf("[ERROR!] %s('%s')\n", e->reason, a);
        delete pu;
    }else{
        ctx->nest = pu;

        r = 1;
    }

    return r;
}

int  on_ws_data(SOCKET sock, context* ctx, char* buf, size_t len, char opcode){
                                                                                    LOG("on_ws_data");
    Unit* pu = (Unit*)ctx->nest;

    size_t n = utf8_wchar(nullptr, 0, buf);
    BSTR s = SysAllocStringLen(NULL, n-1);
    utf8_wchar(s, n, buf);

    _variant_t v;{
        v.vt = VT_BSTR;
        v.bstrVal = s;
    }

    int r = 0;
    CoInitializeEx(0, COINIT_MULTITHREADED);
    {
        DISPID did;
        wchar_t name[] = L"on_data";
        LPOLESTR names[] = { name };
        if(pu->m_oProcessor.GetIDsOfNames(IID_NULL, names, 1, 0, &did) == S_OK){
            DISPPARAMS param = { &v, nullptr, 1, 0 };
            _variant_t res;
            HRESULT hr = pu->m_oProcessor.Invoke(did, IID_NULL, 0, DISPATCH_METHOD, &param, &res, nullptr, nullptr);

            if(FAILED(hr)){
                fwprintf(stderr, L"![0x%x]%ls in line:%d\n", hr, pu->m_oProcessor.m_err->bstrDescription, pu->m_oProcessor.m_err->dwHelpContext);
            }

            r = (hr == S_OK);
        }
    }
    CoUninitialize();

    return r;
}

void on_ws_close(SOCKET sock, context* ctx, char* buf, size_t len, char opcode){
                                                                                    LOG("on_ws_close");
    Unit* pu = (Unit*)ctx->nest;

    delete pu;
}

int  on_http_request_isp(SOCKET sock, context* ctx, FILE* f);
int  on_http_request_osp(SOCKET sock, context* ctx, FILE* f);

int  on_http_request(SOCKET sock, context* ctx){
                                                                                    {
                                                                                    uint32_t adr = ctx->from.sin_addr.s_addr;
                                                                                    char aBuf[3+1+3+1+3+1+3+1+5 +1];
                                                                                    sprintf(aBuf, "%d.%d.%d.%d:%d" 
                                                                                        , adr>> 0&0xff
                                                                                        , adr>> 8&0xff
                                                                                        , adr>>16&0xff
                                                                                        , adr>>24&0xff
                                                                                        , ctx->from.sin_port
                                                                                    );
                                                                                    LOG("%s: %s %s", __FUNCTION__, aBuf, ctx->hdr.m_pURI);
                                                                                    }
    // check URI
    int i = -1;
    while(
        ctx->hdr.m_pURI[++i]                &&
        !(
            ctx->hdr.m_pURI[i+0] == '.' &&
            ctx->hdr.m_pURI[i+1] == '.' &&
        1)                                  &&
    1);

    char aPath[256];
    sprintf(aPath, "wwwroot/%s", (ctx->hdr.m_pURI[i]) ? "" : &ctx->hdr.m_pURI[1]);

#ifdef _WIN32
	FILE* f = fopen(aPath, "rb");
#else
	FILE* f = fopen(aPath, "r");
#endif
	if(!f){
        char aBuf[256];
        aBuf[sizeof(aBuf)-1] = '\0';                                                // For safety.

		int n = 0;
		n += snprintf(aBuf+n, sizeof(aBuf)-n,
			"HTTP/1.0 404 Not Found\r\n"
			"\r\n"
		);
		send(sock, aBuf, n, 0);
	}else{
        const char* p = aPath + strlen(aPath);
        while(! (*--p=='/' || *p=='.') );

        if(*p=='.' && strcmpi(p+1, "osp")==0){
            fclose(f);
#ifdef _WIN32
            f = fopen(aPath, "rt,ccs=UTF-8");
#else
            f = fopen(aPath, "r");
#endif
            return on_http_request_osp(sock, ctx, f);
        }else
        if(*p=='.' && strcmpi(p+1, "isp")==0){
            fclose(f);
#ifdef _WIN32
            f = fopen(aPath, "rt,ccs=UTF-8");
#else
            f = fopen(aPath, "r");
#endif
            return on_http_request_isp(sock, ctx, f);
        }else
        {
            char aBuf[256];
            aBuf[sizeof(aBuf)-1] = '\0';                                        // For safety.

            int n = 0;
            n += snprintf(aBuf+n, sizeof(aBuf)-n,
                "HTTP/1.0 200 OK\r\n"
                "\r\n"
            );
            send(sock, aBuf, n, 0);

            char readbuf[1024];
            int readn;
            while( 0 < (readn = fread(readbuf, 1, sizeof(readbuf), f)) ){
                send(sock, readbuf, readn, 0);
            }

            fclose(f);
        }
	}

	return 0;
}

int  on_http_request_osp(SOCKET sock, context* ctx, FILE* f){
    std::wstring source;{
#ifdef _WIN32
        wchar_t readbuf[256];
        int readn;
        while( 0 < (readn = fread(readbuf, sizeof(readbuf[0]), 256, f)) ){
            source.append(readbuf, readbuf+readn);
        }
#else
        std::string utf8;
        char readbuf[256];
        int readn;
        while( 0 < (readn = fread(readbuf, sizeof(readbuf[0]), 256, f)) ){
            utf8.append(readbuf, readbuf+readn);
        }

        size_t n = utf8_wchar(nullptr, 0, utf8.c_str());
        wchar_t* buf = (wchar_t*)malloc( sizeof(wchar_t)*n );
            utf8_wchar(buf, n, utf8.c_str());
            source += buf;
        free(buf);
#endif
    }

    fclose(f);



    std::vector<std::wstring> soto;
    std::vector<std::wstring> naka;

    size_t start = 0;
    size_t found = 0;
    while(1){
        found = source.find(L"<%", start);
                                            if(found==std::wstring::npos) break;
        soto.push_back( source.substr(start, found-start) );
        start = found+2;
        found = source.find(L"%>", start);
        naka.push_back( source.substr(start, found-start) );
        start = found+2;
    }
    soto.push_back( source.substr(start) );

    std::wstring source2;
    int i=0;
    while(i < soto.size()){
        wchar_t buf2[128];
        swprintf(buf2, 128, L"Response.WriteBlock(%d):", i);

        source2 += buf2;
        if(i < naka.size()){
            if(naka[i][0] == L'='){
                source2 += L"Response.Write(" + naka[i].substr(1) + L"):";
            }else{
                source2 += naka[i] + L":";
            }
        }
        ++i;
    }



    {
        // run
        IDispatch* pVBScript = nullptr;
        CFVBScript->CreateInstance(nullptr, IID_IDispatch, (void**)&pVBScript);

        Session     oSession;
        Request     oRequest;
        Response    oResponse(soto);

        CExtension oExt = {
            { L"Server"     , (IDispatch*)&g_oServer },
            { L"Application", (IDispatch*)&g_oApplication },
            { L"Session"    , (IDispatch*)&oSession },
            { L"Request"    , (IDispatch*)&oRequest },
            { L"Response"   , (IDispatch*)&oResponse },
        };

        HRESULT hr = S_OK;
        CoInitializeEx(0, COINIT_APARTMENTTHREADED);
        {
            const wchar_t* pSource2 = source2.c_str();
            CProgram oProgram(pSource2);

            CProcessor oProcessor(&oProgram, pVBScript, &oExt);
                
            hr = oProcessor();

            if(FAILED(hr)){
                fwprintf(stderr, L"![0x%x]%ls in line:%d\n", hr, oProcessor.m_err->bstrDescription, oProcessor.m_err->dwHelpContext);
            }
        }
        CoUninitialize();

        if(FAILED(hr)){
            char aBuf[256];
            aBuf[sizeof(aBuf)-1] = '\0';                                        // For safety.

            int n = 0;
            n += snprintf(aBuf+n, sizeof(aBuf)-n,
                "HTTP/1.0 500 Error\r\n"
                "\r\n"
            );
            send(sock, aBuf, n, 0);

                                                                                // char a[0x10];
                                                                                // strncpy(a, e->code, e->len)[e->len] = '\0';
                                                                                // printf("[ERROR!] %s('%s')\n", e->reason, a);
        }else{
            char aBuf[256];
            aBuf[sizeof(aBuf)-1] = '\0';                                        // For safety.

            int n = 0;
            n += snprintf(aBuf+n, sizeof(aBuf)-n,
                "HTTP/1.0 200 OK\r\n"
                "content-type: text/html\r\n"
                "\r\n"
            );
            send(sock, aBuf, n, 0);

            send(sock, "\xEF\xBB\xBF", 3, 0);
            {
                size_t bufn = wchar_utf8(nullptr, 0, oResponse.resbuf.c_str());
                char* buf = (char*)malloc(bufn);
                    wchar_utf8(buf, bufn, oResponse.resbuf.c_str());
                    send(sock, buf, bufn-1, 0);
                free(buf);
            }
        }
    }

    return 0;
}

int  on_http_request_isp(SOCKET sock, context* ctx, FILE* f){
    std::wstring source;{
#ifdef _WIN32
        wchar_t readbuf[256];
        int readn;
        while( 0 < (readn = fread(readbuf, sizeof(readbuf[0]), 256, f)) ){
            source.append(readbuf, readbuf+readn);
        }
#else
        std::string utf8;
        char readbuf[256];
        int readn;
        while( 0 < (readn = fread(readbuf, sizeof(readbuf[0]), 256, f)) ){
            utf8.append(readbuf, readbuf+readn);
        }

        size_t n = utf8_wchar(nullptr, 0, utf8.c_str());
        wchar_t* buf = (wchar_t*)malloc( sizeof(wchar_t)*n );
            utf8_wchar(buf, n, utf8.c_str());
            source += buf;
        free(buf);
#endif
    }

    fclose(f);



    {
        // run
        IchigoPage oScript;

        CExtension oExt = {
        };

        HRESULT hr = S_OK;
        CoInitializeEx(0, COINIT_APARTMENTTHREADED);
        {
            const wchar_t* pSource = source.c_str();
            CProgram oProgram(pSource);

            CProcessor oProcessor(&oProgram, &oScript, &oExt);
            oProcessor.m_cc = 10000;
            
            hr = oProcessor();

            if(oProcessor.m_err->scode == 0x205){
               oScript.resbuf += L"[!Script Timeout]";
                hr = S_OK;
            }

            if(FAILED(hr)){
                fwprintf(stderr, L"![0x%x]%ls in line:%d\n", hr, oProcessor.m_err->bstrDescription, oProcessor.m_err->dwHelpContext);
            }
        }
        CoUninitialize();

        if(FAILED(hr)){
            char aBuf[256];
            aBuf[sizeof(aBuf)-1] = '\0';                                        // For safety.

            int n = 0;
            n += snprintf(aBuf+n, sizeof(aBuf)-n,
                "HTTP/1.0 500 Error\r\n"
                "\r\n"
            );
            send(sock, aBuf, n, 0);

                                                                                // char a[0x10];
                                                                                // strncpy(a, e->code, e->len)[e->len] = '\0';
                                                                                // printf("[ERROR!] %s('%s')\n", e->reason, a);
        }else{
            char aBuf[256];
            aBuf[sizeof(aBuf)-1] = '\0';                                        // For safety.

            int n = 0;
            n += snprintf(aBuf+n, sizeof(aBuf)-n,
                "HTTP/1.0 200 OK\r\n"
                "content-type: text/html\r\n"
                "\r\n"
            );
            send(sock, aBuf, n, 0);

            send(sock, "\xEF\xBB\xBF", 3, 0);
            {
                wchar_t aBuf2[1024];
                int n2 = 0;
                n2 += _snwprintf(aBuf2+n2, sizeof(aBuf2)/sizeof(wchar_t)-n2,
                    L"<!DOCTYPE html>\r\n"
                    L"<html>\r\n"
                    L"<head>\r\n"
                    L"<title>%ls</title>\r\n"
                    L"</head>\r\n"
                    L"<body>\r\n"
                    L"<pre>"
                    , L"IchigoPage"
                );
                {
                    size_t bufn = wchar_utf8(nullptr, 0, aBuf2);
                    char* buf = (char*)malloc(bufn);
                        wchar_utf8(buf, bufn, aBuf2);
                        send(sock, buf, bufn-1, 0);
                    free(buf);
                }
            }
            {
                size_t bufn = wchar_utf8(nullptr, 0, oScript.resbuf.c_str());
                char* buf = (char*)malloc(bufn);
                    wchar_utf8(buf, bufn, oScript.resbuf.c_str());
                    send(sock, buf, bufn-1, 0);
                free(buf);
            }
            {
                wchar_t aBuf2[1024];
                int n2 = 0;
                n2 += _snwprintf(aBuf2+n2, sizeof(aBuf2)/sizeof(wchar_t)-n2,
                    L"</pre>\r\n"
                    L"</body>\r\n"
                    L"</html>"
                );
                {
                    size_t bufn = wchar_utf8(nullptr, 0, aBuf2);
                    char* buf = (char*)malloc(bufn);
                        wchar_utf8(buf, bufn, aBuf2);
                        send(sock, buf, bufn-1, 0);
                    free(buf);
                }
            }
        }
    }

    return 0;
}









// int main(int argc, char* argv[]){
//     setlocale(LC_ALL, "");  //#?# too slow for printf...
//     WSA_STARTUP();
    
//     struct sockaddr_in bindto = {};
//     {
//         bindto.sin_family      = PF_INET;
//         bindto.sin_addr.s_addr = (1<argc) ? inet_addr(argv[1])   : INADDR_ANY;
//         bindto.sin_port        = (2<argc) ? htons(atoi(argv[2])) : htons(8000);
//     }
    
//     on_start(&bindto);
    
//     usleep(10000);
//     fprintf(stdout, "bindto: %s:%d.\n", inet_ntoa(bindto.sin_addr), (int)ntohs(bindto.sin_port));
//     fprintf(stdout, "Press ENTER to stop server.\n");
//     fgetc(stdin);                                                                   // waiting for DON'T exit this thread.
//     fprintf(stdout, "stopping server...\n");
    
//     on_stop();
    
//     WSA_CLEANUP();
//     return 0;
// }









class OpenASP : public IDispatch{
private:
    ULONG       m_refc   = 1;

public:
    OpenASP(){/*wprintf(L"%s\n", __func__);*/}
    virtual ~OpenASP(){/*wprintf(L"%s\n", __func__);*/}

    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        return E_NOTIMPL;
    }
    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        WSA_STARTUP();
        
        struct sockaddr_in bindto = {};
        {
            bindto.sin_family      = PF_INET;
            // bindto.sin_addr.s_addr = (1<argc) ? inet_addr(argv[1])   : INADDR_ANY;
            // bindto.sin_port        = (2<argc) ? htons(atoi(argv[2])) : htons(8000);
            bindto.sin_addr.s_addr = INADDR_ANY;
            bindto.sin_port        = htons(8000);
        }
        
        on_start(&bindto);
        
        usleep(10000);
        fprintf(stdout, "bindto: %s:%d.\n", inet_ntoa(bindto.sin_addr), (int)ntohs(bindto.sin_port));
        fprintf(stdout, "Press ENTER to stop server.\n");
        fgetc(stdin);                                                                   // waiting for DON'T exit this thread.
        fprintf(stdout, "stopping server...\n");
        
        on_stop();
        
        WSA_CLEANUP();

        return S_OK;
    }
};



class CFactory : public IClassFactory{
private:
    ULONG       m_refc   = 1;

public:
    CFactory(){}
    virtual ~CFactory(){}

public:
    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject){
        *ppvObject = new OpenASP;
        return S_OK;
    }
    
    HRESULT LockServer(BOOL fLock){
        m_refc += fLock ? 1 : -1;
        return S_OK;
    }
} g_oFactory;



extern "C"
HRESULT DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv){
    g_oFactory.AddRef();
    *ppv = (IClassFactory*)&g_oFactory;

    return S_OK;
}

extern "C"
HRESULT DllCanUnloadNow(){
    return S_FALSE;
}

static class Dll{
public:
    Dll(){
        HMODULE hm = GetModuleHandleW(nullptr);
        LPFNGETCLASSOBJECT pfn = (LPFNGETCLASSOBJECT)GetProcAddress(hm, "DllGetClassObject");
        if(pfn){
            {
                IClassFactory* cf = (IClassFactory*)L"VBScript";
                if( SUCCEEDED(pfn(CLSID_NULL, IID_IClassFactory, (void**)&cf)) ){
                    CFVBScript = cf;
                }
            }
        }
    }
    ~Dll(){
        if(CFVBScript)  CFVBScript->Release();
    }
} _;


