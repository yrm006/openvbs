#include "npole.h"

#include <stdio.h>
#include <string.h>
#include <vector>
#include <string>
#include <map>



#define DISPID_GETIMMEDIATELY DISPID_UNKNOWN
#define NAME L"Jujube"

enum VARENUMX{
    VTX_NONE            = 0x00,
    VTX_GROUND          = 0x01,
    VTX_INST            = 0x02,
    VTX_PROGRAM         = 0x03,
    VTX_JSONNEW         = 0x04,
    VTX_CLOSURE         = 0x05,
    VTX_CLASS           = 0x06,
    VTX_SKIPTOELSE      = 0x10,
    VTX_SKIPTOENDIF     = 0x11,
    VTX_SKIPTOCASE      = 0x12,
    VTX_SKIPTOENDSELECT = 0x13,
    VTX_SKIPTOLOOP      = 0x14,
    VTX_SKIPTONEXT      = 0x15,
    VTX_PCFORLOOPWHILE  = 0x20,
    VTX_PCFORLOOPUNTIL  = 0x21,
    VTX_POFORRETURN     = 0x22,
    VTX_PCFORRETURN     = 0x23,
    VTX_PCFORNEXT       = 0x24,
    VTX_RETURN          = 0x25,
};

enum DIMSCOPE{
    DSC_PUBLIC  = 0,
    DSC_PRIVATE = 1,
};

struct ichar_traits : public std::char_traits<wchar_t>{
	static int compare(const wchar_t *s1, const wchar_t *s2, size_t n){
		return _wcsnicmp(s1, s2, n);
	}
};

template <typename T>
class OnceAllocator{
public:
    typedef T value_type;

    T* allocate(size_t n){
        if(m_already){
            throw this;
            return nullptr;
        }else{
            m_already = true;
            return (T*)::operator new(sizeof(T)*n);
        }
    }

    void deallocate(T* p, size_t n){
        ::operator delete(p);
    }

private:
    bool m_already = false;
};

typedef std::basic_string<wchar_t, ichar_traits> istring;
typedef std::map<istring, _variant_t> CExtension;
typedef _com_ptr_t<_com_IIID<IDispatch, &IID_IDispatch> > _disp_ptr_t;
class CProgram;
typedef _com_ptr_t<_com_IIID<CProgram, &IID_NULL> > _prog_ptr_t;
class Error;
typedef _com_ptr_t<_com_IIID<Error, &IID_NULL> > _err_ptr_t;

struct word_t{ istring s; void* p; _variant_t v; size_t l; };



class JSONObjectEnum : public IEnumVARIANT{
private:
    ULONG       m_refc   = 1;

    const std::map<std::wstring, _variant_t>& m_r;
    std::map<std::wstring, _variant_t>::const_iterator m_i;

public:
    JSONObjectEnum(const std::map<std::wstring, _variant_t>& r)
        : m_r(r)
    {
        m_i = m_r.begin();
    }
    virtual ~JSONObjectEnum(){/*wprintf(L"%s\n", __func__);*/}

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){
        if( memcmp(&riid, &IID_IEnumVARIANT, sizeof(GUID)) == 0 ){
            AddRef();
            *ppvObject = (IEnumVARIANT*)this;
        }else{
            return E_NOINTERFACE;
        }

        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE Next(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched){
        ULONG i = 0;
        while(i < celt && m_i != m_r.end()){
            (rgVar+i)->vt = VT_BSTR;
            (rgVar+i)->bstrVal = SysAllocString(m_i->first.c_str());
            ++m_i;
            ++i;
        }

        if(pCeltFetched) *pCeltFetched = i;

        return (i == celt) ? S_OK : S_FALSE;
    }
    HRESULT STDMETHODCALLTYPE Skip(ULONG celt){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE Reset(){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT **ppEnum){ return E_NOTIMPL; }
};

class JSONObject : public IDispatch{
private:
    ULONG       m_refc   = 1;
    _variant_t* m_pvLast = nullptr;

public:
    std::map<std::wstring, _variant_t> m_map;

public:
    JSONObject(){/*wprintf(L"%s\n", __func__);*/}
    JSONObject(const wchar_t* json){
        from(json);
    }
    virtual ~JSONObject(){};

public:
    bool from(const wchar_t* c);

    std::wstring to(){
        std::wstring json;
        json += L"{";{
            auto i = m_map.begin();
            while(i != m_map.end()){
                json += L"\"";
                json += i->first;
                json += L"\"";
                json += L":";

                _variant_t v;
                if(i->second.vt == VT_BOOL){
                    json += (i->second.boolVal == VARIANT_TRUE) ? L"true" : L"false";
                }else
                if(i->second.vt == VT_NULL){
                    json += L"null";
                }else
                if( SUCCEEDED(VariantChangeType(&v, &i->second, 0, VT_BSTR)) ){
                    if(
                        (i->second.vt == VT_I8)       ||
                        (i->second.vt == VT_R8)       ||
                        (i->second.vt == VT_DISPATCH) ||
                    false){
                        json += v.bstrVal;
                    }else{
                        json += L"\"";
                        json += v.bstrVal;
                        json += L"\"";
                    }
                }else{
                    json += L"\"";
                    json += L"[Unknown Value]";
                    json += L"\"";
                }
                json += L",";

                ++i;
            }
        }
        json += L"}";

        return json;
    }

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        m_pvLast = &m_map[rgszNames[0]];
        rgDispId[0] = DISPID_GETIMMEDIATELY;

        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(dispIdMember == DISPID_VALUE){
            if(wFlags & DISPATCH_PROPERTYPUT){
                if(1 < pDispParams->cArgs){
                    VARIANT* pv = &pDispParams->rgvarg[0];
                    if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;
                    VARIANT* pi = &pDispParams->rgvarg[1];
                    if(pi->vt == (VT_BYREF|VT_VARIANT)) pi = pi->pvarVal;

                    if(pi->vt == VT_BSTR){
                        VariantCopy(&m_map[pi->bstrVal], pv);
                    }else{
                        return E_INVALIDARG;
                    }
                }else{
                    VARIANT* pv = &pDispParams->rgvarg[0];
                    if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

                    if(pv->vt == VT_BSTR && from(pv->bstrVal)){
                        pVarResult->vt = VT_DISPATCH;
                        pVarResult->pdispVal = this; this->AddRef();
                    }else{
                        return E_INVALIDARG;
                    }
                }
            }else{
                if(pDispParams->cArgs){
                    VARIANT* pv = &pDispParams->rgvarg[0];
                    if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

                    if(pv->vt == VT_BSTR){
                        pVarResult->vt = (VT_BYREF|VT_VARIANT);
                        pVarResult->pvarVal = &m_map[pv->bstrVal];
                    }else{
                        return E_INVALIDARG;
                    }
                }else{
                    pVarResult->vt = VT_BSTR;
                    pVarResult->bstrVal = SysAllocString(to().c_str());
                }
            }
        }else
        if(dispIdMember == DISPID_GETIMMEDIATELY){
            if(wFlags & DISPATCH_PROPERTYPUT){
                //#!# may not come here because DISPID_GETIMMEDIATELY
                *m_pvLast = pDispParams->rgvarg[0];
            }else{
                pVarResult->vt = (VT_BYREF|VT_VARIANT);
                pVarResult->pvarVal = m_pvLast;
            }
        }else
        if(dispIdMember == DISPID_NEWENUM){
            pVarResult->vt = VT_UNKNOWN;
            pVarResult->punkVal = new JSONObjectEnum(m_map);
        }else
        {
            return DISP_E_MEMBERNOTFOUND;
        }

        return S_OK;
    }
};



class JSONArrayEnum : public IEnumVARIANT{
private:
    ULONG       m_refc   = 1;

    const std::vector<_variant_t>& m_r;
    std::vector<_variant_t>::const_iterator m_i;

public:
    JSONArrayEnum(const std::vector<_variant_t>& r)
        : m_r(r)
    {
        m_i = m_r.begin();
    }
    virtual ~JSONArrayEnum(){/*wprintf(L"%s\n", __func__);*/}

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){
        if( memcmp(&riid, &IID_IEnumVARIANT, sizeof(GUID)) == 0 ){
            AddRef();
            *ppvObject = (IEnumVARIANT*)this;
        }else{
            return E_NOINTERFACE;
        }

        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE Next(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched){
        ULONG i = 0;
        while(i < celt && m_i != m_r.end()){
            VariantCopy(rgVar+i, &*m_i);
            ++m_i;
            ++i;
        }

        if(pCeltFetched) *pCeltFetched = i;

        return (i == celt) ? S_OK : S_FALSE;
    }
    HRESULT STDMETHODCALLTYPE Skip(ULONG celt){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE Reset(){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT **ppEnum){ return E_NOTIMPL; }
};

class JSONArray : public IDispatch{
private:
    ULONG       m_refc   = 1;

public:
    std::vector<_variant_t> m_vec;

public:
    JSONArray(){/*wprintf(L"%s\n", __func__);*/}
    JSONArray(const wchar_t* json){
        from(json);
    }
    virtual ~JSONArray(){};

public:
    bool from(const wchar_t* c);

    std::wstring to(){
        std::wstring json;
        json += L"[";{
            auto i = m_vec.begin();
            while(i != m_vec.end()){
                _variant_t v;
                if(i->vt == VT_BOOL){
                    json += (i->boolVal == VARIANT_TRUE) ? L"true" : L"false";
                }else
                if(i->vt == VT_NULL){
                    json += L"null";
                }else
                if( SUCCEEDED(VariantChangeType(&v, &*i, 0, VT_BSTR)) ){
                    if(
                        (i->vt == VT_I8)       ||
                        (i->vt == VT_R8)       ||
                        (i->vt == VT_DISPATCH) ||
                    false){
                        json += v.bstrVal;
                    }else{
                        json += L"\"";
                        json += v.bstrVal;
                        json += L"\"";
                    }
                }else{
                    json += L"\"";
                    json += L"[Unknown Value]";
                    json += L"\"";
                }
                json += L",";

                ++i;
            }
        }
        json += L"]";

        return json;
    }

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(_wcsicmp(*rgszNames, L"length") == 0){
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
        if(dispIdMember == DISPID_VALUE){
            if(wFlags & DISPATCH_PROPERTYPUT){
                if(1 < pDispParams->cArgs){
                    VARIANT* pv = &pDispParams->rgvarg[0];
                    if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;
                    VARIANT* pi = &pDispParams->rgvarg[1];
                    if(pi->vt == (VT_BYREF|VT_VARIANT)) pi = pi->pvarVal;

                    if(pi->vt == VT_I8){
                        VariantCopy(&m_vec[pi->llVal], pv);
                    }else{
                        return E_INVALIDARG;
                    }
                }else{
                    VARIANT* pv = &pDispParams->rgvarg[0];
                    if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

                    if(pv->vt == VT_BSTR && from(pv->bstrVal)){
                        pVarResult->vt = VT_DISPATCH;
                        pVarResult->pdispVal = this; this->AddRef();
                    }else{
                        return E_INVALIDARG;
                    }
                }
            }else{
                if(pDispParams->cArgs){
                    VARIANT* pv = &pDispParams->rgvarg[0];
                    if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

                    if(pv->vt == VT_I8 && (0 <= pv->llVal && pv->llVal < m_vec.size())){
                        VariantCopy(pVarResult, &m_vec[pv->llVal]);
                    }else{
                        return E_INVALIDARG;
                    }
                }else{
                    pVarResult->vt = VT_BSTR;
                    pVarResult->bstrVal = SysAllocString(to().c_str());
                }
            }
        }else
        if(dispIdMember == DISPID_NEWENUM){
            pVarResult->vt = VT_UNKNOWN;
            pVarResult->punkVal = new JSONArrayEnum(m_vec);
        }else
        if(dispIdMember == 1){
            pVarResult->vt = VT_I8;
            pVarResult->llVal = m_vec.size();
        }else{
            return DISP_E_MEMBERNOTFOUND;
        }

        return S_OK;
    }
};




class WeakDispatch : public IDispatch{
private:
    ULONG       m_refc   = 1;
    IDispatch*  m_pdisp;

public:
    WeakDispatch(IDispatch* pdisp)
        : m_pdisp(pdisp)
    {
        // wprintf(L"%s\n", __func__);
    }
    virtual ~WeakDispatch(){};

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return m_pdisp->GetTypeInfoCount(pctinfo); }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return m_pdisp->GetTypeInfo(iTInfo, lcid, ppTInfo); }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        return m_pdisp->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        return m_pdisp->Invoke(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
    }
};



class Error : public EXCEPINFO, public IDispatch{
private:
    ULONG   m_refc   = 1;

public:
    HRESULT m_hr = E_FAIL;
    UINT    m_an = 0;

    Error(){
        wCode = 0;
        wReserved = 0;
        bstrSource = nullptr;
        bstrDescription = nullptr;
        bstrHelpFile = nullptr;
        dwHelpContext = 0;
        pvReserved = nullptr;
        pfnDeferredFillIn = nullptr;
        scode = 0;
    }

    virtual ~Error(){
        SysFreeString(bstrSource);
        SysFreeString(bstrDescription);
        SysFreeString(bstrHelpFile);
    }

    void Clear(){
        wCode = 0;
        wReserved = 0;
        SysFreeString(bstrSource); bstrSource = nullptr;
        SysFreeString(bstrDescription); bstrDescription = nullptr;
        SysFreeString(bstrHelpFile); bstrHelpFile = nullptr;
        dwHelpContext = 0;
        pvReserved = nullptr;
        pfnDeferredFillIn = nullptr;
        scode = 0;
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(_wcsicmp(*rgszNames, L"number") == 0){
            *rgDispId = 1;
        }else
        if(_wcsicmp(*rgszNames, L"source") == 0){
            *rgDispId = 2;
        }else
        if(_wcsicmp(*rgszNames, L"description") == 0){
            *rgDispId = 3;
        }else
        if(_wcsicmp(*rgszNames, L"clear") == 0){
            *rgDispId = 10;
        }else
        if(_wcsicmp(*rgszNames, L"raise") == 0){
            *rgDispId = 11;
        }else
        {
            return DISP_E_MEMBERNOTFOUND;
        }

        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(dispIdMember == 1){
            if(wFlags & DISPATCH_PROPERTYPUT){
                this->wCode = pDispParams->rgvarg->llVal;
            }else{
                *((_variant_t*)pVarResult) = (long long)this->wCode;
            }

            return S_OK;
        }else
        if(dispIdMember == 2){
            if(wFlags & DISPATCH_PROPERTYPUT){
                this->bstrSource = SysAllocString(pDispParams->rgvarg->bstrVal);
            }else{
                *((_variant_t*)pVarResult) = this->bstrSource;
            }

            return S_OK;
        }else
        if(dispIdMember == 3){
            if(wFlags & DISPATCH_PROPERTYPUT){
                this->bstrDescription = SysAllocString(pDispParams->rgvarg->bstrVal);
            }else{
                *((_variant_t*)pVarResult) = this->bstrDescription;
            }

            return S_OK;
        }else
        if(dispIdMember == 10){
            Clear();

            return S_OK;
        }else
        if(dispIdMember == 11){
            if(1 <= pDispParams->cArgs){
                pExcepInfo->wCode = pDispParams->rgvarg[pDispParams->cArgs-1].uiVal;
            }//NOTELSE
            if(2 <= pDispParams->cArgs){
                SysFreeString(pExcepInfo->bstrSource);
                pExcepInfo->bstrSource = SysAllocString(pDispParams->rgvarg[pDispParams->cArgs-2].bstrVal);
            }//NOTELSE
            if(3 <= pDispParams->cArgs){
                SysFreeString(pExcepInfo->bstrDescription);
                pExcepInfo->bstrDescription = SysAllocString(pDispParams->rgvarg[pDispParams->cArgs-3].bstrVal);
            }//NOTELSE

            return E_FAIL;
        }else
        {
            return DISP_E_MEMBERNOTFOUND;
        }
    }
};



class Null : public IDispatch{
public:
    Null(){}
    virtual ~Null(){}

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return 1; }
    ULONG STDMETHODCALLTYPE Release(){ return 1; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        *rgDispId = DISPID_VALUE;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        pVarResult->vt = VT_EMPTY;
        return S_OK;
    }
};



class CProgram : public IDispatch{
private:
    static void* map_word(const istring& s);

private:
    typedef const wchar_t* (CProgram::*mode_t)(const wchar_t*, size_t);
    struct dim_names_t{ DIMSCOPE s; size_t i; };
    struct funcs_t{ DIMSCOPE s; _prog_ptr_t p; };

    mode_t m_parse_mode;

    const wchar_t* parse_option(const wchar_t* c, size_t len){
        if(len == 8 && _wcsnicmp(L"explicit", c, len) == 0){
            m_code.pop_back();
            m_option_explicit = true;
        }else
        {
            //#!# maybe syntax error
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        m_parse_mode = &CProgram::parse_;
        return c+len;
    }

    const wchar_t* parse_onerrorresume(const wchar_t* c, size_t len){
        if(len == 4 && _wcsnicmp(L"next", c, len) == 0){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }else
        {
            //#!# maybe syntax error
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        m_parse_mode = &CProgram::parse_;
        return c+len;
    }

    const wchar_t* parse_onerrorgoto(const wchar_t* c, size_t len){
        if(len == 1 && _wcsnicmp(L"0", c, len) == 0){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }else
        if(len == 5 && _wcsnicmp(L"catch", c, len) == 0){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }else
        {
            //#!# maybe syntax error
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        m_parse_mode = &CProgram::parse_;
        return c+len;
    }

    const wchar_t* parse_onerror(const wchar_t* c, size_t len){
        if(len == 6 && _wcsnicmp(L"resume", c, len) == 0){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_onerrorresume;
        }else
        if(len == 4 && _wcsnicmp(L"goto", c, len) == 0){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_onerrorgoto;
        }else
        if(len == 5 && _wcsnicmp(L"catch", c, len) == 0){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_;
        }else
        {
            //#!# maybe syntax error
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        return c+len;
    }

    const wchar_t* parse_on(const wchar_t* c, size_t len){
        if(len == 5 && _wcsnicmp(L"error", c, len) == 0){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_onerror;
        }else
        {
            //#!# maybe syntax error
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        return c+len;
    }

    istring parse_dim_name;
    std::vector<SAFEARRAYBOUND> parse_dim_array_sab;

    const wchar_t* parse_dim_array(const wchar_t* c, size_t len){
        if(len == 1 && *c == L')'){
            _variant_t& v = m_dim_defs.back();{
                v.vt = (VT_ARRAY|VT_VARIANT);
                if( parse_dim_array_sab.size() ){
                    v.parray = SafeArrayCreate(VT_VARIANT, parse_dim_array_sab.size(), &parse_dim_array_sab[0]);
                }else{
                    SAFEARRAYBOUND sab = { 0, 0 };
                    v.parray = SafeArrayCreate(VT_VARIANT, 1, &sab);
                }
            }

            m_parse_mode = &CProgram::parse_dim;
        }else{
            istring s(c, len);
            parse_dim_array_sab.push_back({ (ULONG)wcstoll(s.c_str(), nullptr, 10)+1, 0 });
        }

        return c+len;
    }

    const wchar_t* parse_dim(const wchar_t* c, size_t len){
        if(len == 1 && *c == L','){
            // none
        }else
        if(
            (len == 1 && *c == L'=')                    ||
            (len == 2 && (*c==L':' && *(c+1)==L'=') )   ||
        false){
            m_code.push_back( { parse_dim_name, map_word(parse_dim_name), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }else
        if(len == 1 && *c == L'('){
            m_parse_mode = &CProgram::parse_dim_array;
            parse_dim_array_sab.clear();
        }else
        if(len == 1 && (*c == L'\n' || *c == L':')){
            m_parse_mode = &CProgram::parse_;
            if(*c == L'\n') ++m_lines;
        }else
        {
            parse_dim_name = istring(c, len);

            auto n = m_dim_names.size();
            auto i = m_dim_names.insert({parse_dim_name, {DSC_PUBLIC, n}});
            if(i.second){
                m_dim_defs.push_back( _variant_t() );
            }else{
                //#!# re-definition
            }
        }

        return c+len;
    }

    const wchar_t* parse_public(const wchar_t* c, size_t len){
        if(len == 7 && _wcsnicmp(L"default", c, len) == 0){
            m_code.push_back( { L"default", map_word(istring(L"@")), _variant_t(), m_lines } );
        }else
        if(len == 8 && _wcsnicmp(L"function", c, len) == 0){
            m_code.push_back( { L"public", map_word(istring(L"@")), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_function;
        }else
        if(len == 8 && _wcsnicmp(L"property", c, len) == 0){
            m_code.push_back( { L"public", map_word(istring(L"@")), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_property;
        }else
        if(len == 1 && *c == L','){
            // none
        }else
        if(
            (len == 1 && *c == L'=')                    ||
            (len == 2 && (*c==L':' && *(c+1)==L'=') )   ||
        false){
            m_code.push_back( { parse_dim_name, map_word(parse_dim_name), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }else
        if(len == 1 && (*c == L'\n' || *c == L':')){
            m_parse_mode = &CProgram::parse_;
            if(*c == L'\n') ++m_lines;
        }else
        {
            parse_dim_name = istring(c, len);

            auto n = m_dim_names.size();
            auto i = m_dim_names.insert({parse_dim_name, {DSC_PUBLIC, n}});
            if(i.second){
                m_dim_defs.push_back( _variant_t() );
            }else{
                //#!# re-definition
            }
        }

        return c+len;
    }

    const wchar_t* parse_private(const wchar_t* c, size_t len){
        if(len == 8 && _wcsnicmp(L"function", c, len) == 0){
            m_code.push_back( { L"private", map_word(istring(L"@")), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_function;
        }else
        if(len == 8 && _wcsnicmp(L"property", c, len) == 0){
            m_code.push_back( { L"private", map_word(istring(L"@")), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_property;
        }else
        if(len == 1 && *c == L','){
            // none
        }else
        if(
            (len == 1 && *c == L'=')                    ||
            (len == 2 && (*c==L':' && *(c+1)==L'=') )   ||
        false){
            m_code.push_back( { parse_dim_name, map_word(parse_dim_name), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }else
        if(len == 1 && (*c == L'\n' || *c == L':')){
            m_parse_mode = &CProgram::parse_;
            if(*c == L'\n') ++m_lines;
        }else
        {
            parse_dim_name = istring(c, len);

            auto n = m_dim_names.size();
            auto i = m_dim_names.insert({parse_dim_name, {DSC_PRIVATE, n}});
            if(i.second){
                m_dim_defs.push_back( _variant_t() );
            }else{
                //#!# re-definition
            }
        }

        return c+len;
    }

    istring parse_const_name;
    bool    parse_const_positive;

    const wchar_t* parse_const_literal(const wchar_t* c, size_t len){
        if(
            (len == 1 && *c == L'=')                    ||
            (len == 2 && (*c==L':' && *(c+1)==L'=') )   ||
        false){
            // none
        }else
        if(len == 1 && *c == L'+'){
            parse_const_positive = true;
        }else
        if(len == 1 && *c == L'-'){
            parse_const_positive = false;
        }else
        if(len == 1 && (*c == L'\n' || *c == L':')){
            m_parse_mode = &CProgram::parse_;
            if(*c == L'\n') ++m_lines;
        }else
        {
            _variant_t v;

            if(L'0'<=*c && *c<=L'9'){
                if(*(c+len) == L'.'){
                    ++len;
                    while( (L'0'<=*(c+len) && *(c+len)<=L'9') ) ++len;

                    istring s(c, len);
                    {
                        v.vt = VT_R8;
                        v.dblVal = wcstod(s.c_str(), nullptr);
                        if(!parse_const_positive) v.dblVal *= -1;
                    }
                }else{
                    istring s(c, len);
                    {
                        v.vt = VT_I8;
                        v.llVal = wcstoll(s.c_str(), nullptr, 10);
                        if(!parse_const_positive) v.llVal *= -1;
                    }
                }
            }else
            if(*c==L'.' && (L'0'<=*(c+len) && *(c+len)<=L'9')){
                while( (L'0'<=*(c+len) && *(c+len)<=L'9') ) ++len;

                istring s(c, len);
                {
                    v.vt = VT_R8;
                    v.dblVal = wcstod(s.c_str(), nullptr);
                    if(!parse_const_positive) v.dblVal *= -1;
                }
            }else
            if(2 <= len && *c==L'&' && (*(c+1)==L'H' || *(c+1)==L'h')){
                istring s(c, len);
                {
                    v.vt = VT_I8;
                    v.llVal = wcstoll(s.c_str()+2, nullptr, 16);
                }
            }else
            if(2 <= len && *c==L'&' && (*(c+1)==L'B' || *(c+1)==L'b')){
                istring s(c, len);
                {
                    v.vt = VT_I8;
                    v.llVal = wcstoll(s.c_str()+2, nullptr, 2);
                }
            }else
            if(*c==L'"' && *(c+len-1)==*c){
                const wchar_t* lc = c + 1;
                std::wstring ls;
                while(1){
                    while(!( *lc==L'"' )) ls += *(lc++);
                    if(*++lc!=L'"') break;
                    ls += *(lc++);
                }

                istring s(ls.c_str());
                {
                    v.vt = VT_BSTR;
                    v.bstrVal = SysAllocString(ls.c_str());
                }
            }else
            if(*c==L'#' && *(c+len-1)==*c){
                istring s(c+1, len-2);
                DATE date = 0.0;
                VarDateFromStr(s.c_str(), 0, 0, &date);

                {
                    v.vt = VT_DATE;
                    v.date = date;
                }
            }else
            {
                // unknown literal
            }

            auto i = m_consts.insert({parse_const_name, v});
            if(i.second){
                // OK
            }else{
                //#!# re-definition
            }
        }

        return c+len;
    }

    const wchar_t* parse_const(const wchar_t* c, size_t len){
        parse_const_name = istring(c, len);
        parse_const_positive = true;

        m_parse_mode = &CProgram::parse_const_literal;

        return c+len;
    }

    const wchar_t* parse_redim(const wchar_t* c, size_t len){
        if(len == 8 && _wcsnicmp(L"preserve", c, len) == 0){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }else{
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        m_parse_mode = &CProgram::parse_;
        return c+len;
    }

    const wchar_t* parse_end(const wchar_t* c, size_t len){
        if(
            (len == 2 && _wcsnicmp(L"if", c, len) == 0)     ||
            (len == 6 && _wcsnicmp(L"select", c, len) == 0) ||
            (len == 4 && _wcsnicmp(L"with", c, len) == 0)   ||
        false){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }else
        if(
            (len == 8 && _wcsnicmp(L"function", c, len) == 0)   ||
            (len == 3 && _wcsnicmp(L"sub", c, len) == 0)        ||
            (len == 8 && _wcsnicmp(L"property", c, len) == 0)   ||
            (len == 5 && _wcsnicmp(L"class", c, len) == 0)      ||
        false){
            m_code.pop_back();
            m_parse_mode = nullptr;
            return c+len;
        }else
        {
            //#!# maybe syntax error
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        m_parse_mode = &CProgram::parse_;
        return c+len;
    }

    const wchar_t* parse_do(const wchar_t* c, size_t len){
        if(
            (len == 5 && _wcsnicmp(L"while", c, len) == 0) ||
            (len == 5 && _wcsnicmp(L"until", c, len) == 0) ||
        false){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_;
        }else
        {
            // do ~ loop while/until
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        return c+len;
    }

    const wchar_t* parse_loop(const wchar_t* c, size_t len){
        if(
            (len == 5 && _wcsnicmp(L"while", c, len) == 0) ||
            (len == 5 && _wcsnicmp(L"until", c, len) == 0) ||
        false){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_;
        }else
        {
            // do while/until ~ loop
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        return c+len;
    }

    const wchar_t* parse_for(const wchar_t* c, size_t len){
        if(len == 4 && _wcsnicmp(L"each", c, len) == 0){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }else
        {
            // for ~ to/in ~
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        return c+len;
    }

    const wchar_t* parse_select(const wchar_t* c, size_t len){
        if(len == 4 && _wcsnicmp(L"case", c, len) == 0){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }else
        {
            //#!# maybe syntax error
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        m_parse_mode = &CProgram::parse_;
        return c+len;
    }

    const wchar_t* parse_case(const wchar_t* c, size_t len){
        if(len == 4 && _wcsnicmp(L"else", c, len) == 0){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }else
        {
            // case XXX
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        return c+len;
    }

    bool parse_function_paran_default_positive;

    const wchar_t* parse_function_param_default(const wchar_t* c, size_t len){
        if(len == 1 && *c == L'+'){
            parse_function_paran_default_positive = true;
        }else
        if(len == 1 && *c == L'-'){
            parse_function_paran_default_positive = false;
        }else
        if(L'0'<=*c && *c<=L'9'){
            if(*(c+len) == L'.'){
                ++len;
                while( (L'0'<=*(c+len) && *(c+len)<=L'9') ) ++len;

                istring s(c, len);
                _variant_t& v = m_dim_defs.back();{
                    v.vt = VT_R8;
                    v.dblVal = wcstod(s.c_str(), nullptr);
                    if(!parse_function_paran_default_positive) v.dblVal *= -1;
                }
            }else{
                istring s(c, len);
                _variant_t& v = m_dim_defs.back();{
                    v.vt = VT_I8;
                    v.llVal = wcstoll(s.c_str(), nullptr, 10);
                    if(!parse_function_paran_default_positive) v.llVal *= -1;
                }
            }

            m_parse_mode = &CProgram::parse_function_param;
        }else
        if(*c==L'.' && (L'0'<=*(c+len) && *(c+len)<=L'9')){
            while( (L'0'<=*(c+len) && *(c+len)<=L'9') ) ++len;

            istring s(c, len);
            _variant_t& v = m_dim_defs.back();{
                v.vt = VT_R8;
                v.dblVal = wcstod(s.c_str(), nullptr);
                if(!parse_function_paran_default_positive) v.dblVal *= -1;
            }

            m_parse_mode = &CProgram::parse_function_param;
        }else
        if(2 <= len && *c==L'&' && (*(c+1)==L'H' || *(c+1)==L'h')){
            istring s(c, len);
            _variant_t& v = m_dim_defs.back();{
                v.vt = VT_I8;
                v.llVal = wcstoll(s.c_str()+2, nullptr, 16);
            }
        }else
        if(2 <= len && *c==L'&' && (*(c+1)==L'B' || *(c+1)==L'b')){
            istring s(c, len);
            _variant_t& v = m_dim_defs.back();{
                v.vt = VT_I8;
                v.llVal = wcstoll(s.c_str()+2, nullptr, 2);
            }
        }else
        if(*c==L'"' && *(c+len-1)==*c){
            const wchar_t* lc = c + 1;
            std::wstring ls;
            while(1){
                while(!( *lc==L'"' )) ls += *(lc++);
                if(*++lc!=L'"') break;
                ls += *(lc++);
            }

            istring s(ls.c_str());
            _variant_t& v = m_dim_defs.back();{
                v.vt = VT_BSTR;
                v.bstrVal = SysAllocString(ls.c_str());
            }

            m_parse_mode = &CProgram::parse_function_param;
        }else
        if(*c==L'#' && *(c+len-1)==*c){
            istring s(c+1, len-2);
            DATE date = 0.0;
            VarDateFromStr(s.c_str(), 0, 0, &date);

            _variant_t& v = m_dim_defs.back();{
                v.vt = VT_DATE;
                v.date = date;
            }

            m_parse_mode = &CProgram::parse_function_param;
        }else
        {
            istring s(c, len);
            auto found = m_consts.find(s);
            if(found != m_consts.end()){
                m_dim_defs.back() = found->second;
            }

            m_parse_mode = &CProgram::parse_function_param;
        }

        return c+len;
    }

    const wchar_t* parse_function_param(const wchar_t* c, size_t len){
        if(len == 1 && _wcsnicmp(L"(", c, len) == 0){
            // none
        }else
        if(len == 1 && _wcsnicmp(L",", c, len) == 0){
            // none
        }else
        if(len == 1 && _wcsnicmp(L"=", c, len) == 0){
            parse_function_paran_default_positive = true;
            m_parse_mode = &CProgram::parse_function_param_default;
        }else
        if(len == 1 && _wcsnicmp(L")", c, len) == 0){
            m_parse_mode = &CProgram::parse_;
        }else
        if(len == 1 && _wcsnicmp(L"\n", c, len) == 0){
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }else
        {
            m_params.push_back(istring(c, len));
            auto n = m_dim_names.size();
            m_dim_names[m_params.back()] = { DSC_PUBLIC, n };
            m_dim_defs.push_back( _variant_t{ {{{VT_ERROR,0,0,0,{0}}}} } );
        }

        return c+len;
    }

    const wchar_t* parse_function_name(const wchar_t* c, size_t len){
        m_name = istring(c, len);

        auto n = m_dim_names.size();
        m_dim_names[m_name] = { DSC_PUBLIC, n };
        m_dim_defs.push_back( _variant_t() );

        m_parse_mode = &CProgram::parse_function_param;
        return c+len;
    }

    const wchar_t* parse_function_noname(const wchar_t* c, size_t len){
        m_name = istring(L"");

        auto n = m_dim_names.size();
        m_dim_names[m_name] = { DSC_PUBLIC, n };
        m_dim_defs.push_back( _variant_t() );

        m_parse_mode = &CProgram::parse_function_param;
        return c;
    }

    const wchar_t* parse_function(const wchar_t* c, size_t len){
        if(*c == L'('){
            _prog_ptr_t prog(new CProgram(c, &CProgram::parse_function_noname, this), false);

            m_closures.push_back( prog );

            _variant_t v;{
                v.wReserved1 = VTX_CLOSURE;
                v.byref = prog;
            }
            m_code.push_back( { L"closure", map_word(istring(L"@closure")), v, m_lines } );

            m_lines += (prog->m_lines-m_lines);
        }else
        {
            ULONG type = DISPATCH_METHOD;
            if(m_code.back().s == L"let"){
                type = DISPATCH_PROPERTYPUT;
                m_code.pop_back();
            }else
            if(m_code.back().s == L"get"){
                type = DISPATCH_PROPERTYGET;
                m_code.pop_back();
            }

            DIMSCOPE ds = DSC_PUBLIC;
            bool def = false;
            if(m_code.back().s == L"public"){
                ds = DSC_PUBLIC;
                m_code.pop_back();
                if(m_code.back().s == L"default"){
                    def = true;
                    m_code.pop_back();
                }
            }else
            if(m_code.back().s == L"private"){
                ds = DSC_PRIVATE;
                m_code.pop_back();
            }

            istring name(c, len);
            _prog_ptr_t prog(new CProgram(c, &CProgram::parse_function_name, this), false);

            if(type == DISPATCH_PROPERTYPUT){
                m_plets.insert({name, {ds, prog}});
            }else{
                m_funcs.insert({name, {ds, prog}});
            }

            if(def){
                m_default = prog;
            }

            m_lines += (prog->m_lines-m_lines);
        }

        m_parse_mode = &CProgram::parse_;
        return c;
    }

    const wchar_t* parse_property(const wchar_t* c, size_t len){
        if(len == 3 && _wcsnicmp(L"let", c, len) == 0){
            m_code.push_back( { L"let", map_word(istring(L"@")), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_function;
        }else
        if(len == 3 && _wcsnicmp(L"get", c, len) == 0){
            m_code.push_back( { L"get", map_word(istring(L"@")), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_function;
        }else{
            m_parse_mode = &CProgram::parse_function;
            return parse_function(c, len);
        }
        
        return c+len;
    }

    const wchar_t* parse_exit(const wchar_t* c, size_t len){
        if(
            (len == 2 && _wcsnicmp(L"do", c, len) == 0)       ||
            (len == 3 && _wcsnicmp(L"for", c, len) == 0)      ||
            (len == 8 && _wcsnicmp(L"function", c, len) == 0) ||
            (len == 3 && _wcsnicmp(L"sub", c, len) == 0)      ||
            (len == 8 && _wcsnicmp(L"property", c, len) == 0) ||
        false){
            istring s(c, len);
            s = m_code.back().s + L" " + s;
            m_code.pop_back();
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }else
        {
            //#!# maybe syntax error
            m_parse_mode = &CProgram::parse_;
            return parse_(c, len);
        }

        m_parse_mode = &CProgram::parse_;
        return c+len;
    }

    const wchar_t* parse_class_name(const wchar_t* c, size_t len){
        m_name = istring(c, len);

        m_parse_mode = &CProgram::parse_;
        return c+len;
    }

    const wchar_t* parse_class(const wchar_t* c, size_t len){
        istring name(c, len);
        _prog_ptr_t prog(new CProgram(c, &CProgram::parse_class_name, this), false);

        m_classes.insert( decltype(m_classes)::value_type(name, prog) );

        m_lines += (prog->m_lines-m_lines);

        m_parse_mode = &CProgram::parse_;
        return c;
    }

    const wchar_t* parse_(const wchar_t* c, size_t len){
        if(len == 6 && _wcsnicmp(L"option", c, len) == 0){
            istring s(c, len);
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_option;
        }else
        if(len == 2 && _wcsnicmp(L"on", c, len) == 0){
            istring s(c, len);
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_on;
        }else
        if(len == 3 && _wcsnicmp(L"dim", c, len) == 0){
            m_parse_mode = &CProgram::parse_dim;
        }else
        if(len == 6 && _wcsnicmp(L"public", c, len) == 0){
            m_parse_mode = &CProgram::parse_public;
        }else
        if(len == 7 && _wcsnicmp(L"private", c, len) == 0){
            m_parse_mode = &CProgram::parse_private;
        }else
        if(len == 5 && _wcsnicmp(L"const", c, len) == 0){
            m_parse_mode = &CProgram::parse_const;
        }else
        if(len == 5 && _wcsnicmp(L"redim", c, len) == 0){
            istring s(c, len);
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_redim;
        }else
        if(len == 3 && _wcsnicmp(L"end", c, len) == 0){
            istring s(c, len);
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_end;
        }else
        if(len == 2 && _wcsnicmp(L"do", c, len) == 0){
            istring s(c, len);
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_do;
        }else
        if(len == 4 && _wcsnicmp(L"loop", c, len) == 0){
            istring s(c, len);
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_loop;
        }else
        if(len == 3 && _wcsnicmp(L"for", c, len) == 0){
            istring s(c, len);
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_for;
        }else
        if(len == 6 && _wcsnicmp(L"select", c, len) == 0){
            istring s(c, len);
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_select;
        }else
        if(len == 4 && _wcsnicmp(L"case", c, len) == 0){
            istring s(c, len);
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_case;
        }else
        if(len == 8 && _wcsnicmp(L"function", c, len) == 0){
            m_parse_mode = &CProgram::parse_function;
        }else
        if(len == 3 && _wcsnicmp(L"sub", c, len) == 0){
            m_parse_mode = &CProgram::parse_function;
        }else
        if(len == 8 && _wcsnicmp(L"property", c, len) == 0){
            m_parse_mode = &CProgram::parse_property;
        }else
        if(len == 4 && _wcsnicmp(L"exit", c, len) == 0){
            istring s(c, len);
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            m_parse_mode = &CProgram::parse_exit;
        }else
        if(len == 5 && _wcsnicmp(L"class", c, len) == 0){
            m_parse_mode = &CProgram::parse_class;
        }else
        if(len == 1 && (*c == L'\n' || *c == L':')){
            istring s(c, len);
            if(m_code.back().p != map_word(s)){
                m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
            }
            if(*c == L'\n') ++m_lines;
        }else
        if(len == 3 && _wcsnicmp(L"set", c, len) == 0){
            // none to remove
        }else
        if(len == 4 && _wcsnicmp(L"call", c, len) == 0){
            // none to remove
        }else
        if(len == 3 && _wcsnicmp(L"rem", c, len) == 0){
            while(!( *(c+len)==L'\n' || *(c+len)==L'\0' )) ++len;
            ++m_lines;
        }else
        if(*c==L'@' && (m_code.back().p == map_word(istring(L":")) || m_code.back().p == map_word(istring(L"\n")))){
            istring s(c, len);
            m_labels[s] = m_code.size()-1;
        }else
        if(L'0'<=*c && *c<=L'9'){
            if(m_code.back().p == map_word(istring(L":")) || m_code.back().p == map_word(istring(L"\n"))){
                istring s(c, len);
                m_labels[s] = m_code.size()-1;
            }else
            if(*(c+len) == L'.'){
                ++len;
                while( (L'0'<=*(c+len) && *(c+len)<=L'9') ) ++len;

                istring s(c, len);
                _variant_t v;{
                    v.vt = VT_R8;
                    v.dblVal = wcstod(s.c_str(), nullptr);
                }
                m_code.push_back( { s, map_word(istring(L"@literal")), v, m_lines } );
            }else{
                istring s(c, len);
                _variant_t v;{
                    v.vt = VT_I8;
                    v.llVal = wcstoll(s.c_str(), nullptr, 10);
                }
                m_code.push_back( { s, map_word(istring(L"@literal")), v, m_lines } );
            }
        }else
        if(*c==L'.' && (L'0'<=*(c+len) && *(c+len)<=L'9')){
            while( (L'0'<=*(c+len) && *(c+len)<=L'9') ) ++len;

            istring s(c, len);
            _variant_t v;{
                v.vt = VT_R8;
                v.dblVal = wcstod(s.c_str(), nullptr);
            }
            m_code.push_back( { s, map_word(istring(L"@literal")), v, m_lines } );
        }else
        if(2 <= len && *c==L'&' && (*(c+1)==L'H' || *(c+1)==L'h')){
            istring s(c, len);
            _variant_t v;{
                v.vt = VT_I8;
                v.llVal = wcstoll(s.c_str()+2, nullptr, 16);
            }
            m_code.push_back( { s, map_word(istring(L"@literal")), v, m_lines } );
        }else
        if(2 <= len && *c==L'&' && (*(c+1)==L'B' || *(c+1)==L'b')){
            istring s(c, len);
            _variant_t v;{
                v.vt = VT_I8;
                v.llVal = wcstoll(s.c_str()+2, nullptr, 2);
            }
            m_code.push_back( { s, map_word(istring(L"@literal")), v, m_lines } );
        }else
        if(*c==L'"' && *(c+len-1)==*c){
            const wchar_t* lc = c + 1;
            std::wstring ls;
            while(1){
                while(!( *lc==L'"' )) ls += *(lc++);
                if(*++lc!=L'"') break;
                ls += *(lc++);
            }

            istring s(ls.c_str());
            _variant_t v;{
                v.vt = VT_BSTR;
                v.bstrVal = SysAllocString(ls.c_str());
            }
            m_code.push_back( { s, map_word(istring(L"@literal")), v, m_lines } );
        }else
        if(*c==L'#' && *(c+len-1)==*c){
            istring s(c+1, len-2);
            DATE date = 0.0;
            VarDateFromStr(s.c_str(), 0, 0, &date);

            _variant_t v;{
                v.vt = VT_DATE;
                v.date = date;
            }
            m_code.push_back( { s, map_word(istring(L"@literal")), v, m_lines } );
        }else
        if(*c==L'{' && *(c+len-1)==L'}'){
            istring s(c+1, len-2);
            m_code.push_back( { s, map_word(istring(L"@json_obj")), _variant_t(), m_lines } );
        }else
        if(*c==L'[' && *(c+len-1)==L']'){
            istring s(c+1, len-2);
            m_code.push_back( { s, map_word(istring(L"@json_ary")), _variant_t(), m_lines } );
        }else
        {
            istring s(c, len);
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }

        return c+len;
    }

    void bind(){
        auto i = m_code.begin();
        while(i != m_code.end()){
            decltype(m_consts)::const_iterator foundC;
            decltype(m_dim_names)::const_iterator foundD;
            decltype(m_labels)::const_iterator foundL;

            if(i->p != map_word(istring(L"@"))){
                // none: already binded
            }else
            if((foundC = m_consts.find(i->s)) != m_consts.end()){
                i->p = map_word(istring(L"@literal"));
                i->v = foundC->second;
            }else
            if(m_parent && (foundC = m_parent->m_consts.find(i->s)) != m_parent->m_consts.end()){
                i->p = map_word(istring(L"@literal"));
                i->v = foundC->second;
            }else
            if(m_granpa && (foundC = m_granpa->m_consts.find(i->s)) != m_granpa->m_consts.end()){
                i->p = map_word(istring(L"@literal"));
                i->v = foundC->second;
            }else
            if((foundD = m_dim_names.find(i->s)) != m_dim_names.end()){
                if( i->s == m_name && (i+1)->s == L"(" ){
                    // none: recurse
                }else{
                    i->p = map_word(istring(L"@dim"));
                    i->v = foundD->second.i;
                }
            }else
            if(m_parent && (foundD = m_parent->m_dim_names.find(i->s)) != m_parent->m_dim_names.end()){
                if(m_name.length()){
                    i->p = map_word(istring(L"@pdim"));
                    i->v = foundD->second.i;
                }else{
                    auto n = m_clo_names.size();
                    auto j = m_clo_names.insert({foundD->first, n});
                    if(j.second){
                        m_clo_ids.push_back( foundD->second.i );
                        m_clo_defs.push_back( _variant_t() );
                    }

                    i->p = map_word(istring(L"@cdim"));
                    i->v = m_clo_names[foundD->first];
                }
            }else
            if(m_granpa && (foundD = m_granpa->m_dim_names.find(i->s)) != m_granpa->m_dim_names.end()){
                i->p = map_word(istring(L"@gdim"));
                i->v = foundD->second.i;
            }else
            if((foundL = m_labels.find(i->s)) != m_labels.end()){
                i->p = map_word(istring(L"@literal"));
                i->v.vt = VT_UI8;
                i->v.ullVal = foundL->second;
            }else
            {
                // none: bind the rests on run time.
            }

            ++i;
        }

        {
            auto j = m_funcs.begin();
            while(j != m_funcs.end()){
                j->second.p->m_option_explicit = m_option_explicit;
                j->second.p->bind();
                ++j;
            }
        }

        {
            auto j = m_plets.begin();
            while(j != m_plets.end()){
                j->second.p->m_option_explicit = m_option_explicit;
                j->second.p->bind();
                ++j;
            }
        }

        {
            auto j = m_closures.begin();
            while(j != m_closures.end()){
                (*j)->m_option_explicit = m_option_explicit;
                (*j)->bind();
                ++j;
            }
        }

        {
            auto j = m_classes.begin();
            while(j != m_classes.end()){
                j->second->m_option_explicit = m_option_explicit;
                j->second->bind();
                ++j;
            }
        }
    }

public:
    bool                            m_option_explicit;
    istring                         m_name;
    std::vector<istring>            m_params;
    std::map<istring, dim_names_t>  m_dim_names;
    std::vector<_variant_t>         m_dim_defs;
    std::map<istring, funcs_t>      m_funcs;
    std::map<istring, funcs_t>      m_plets;
    _prog_ptr_t                     m_default;
    std::map<istring, _prog_ptr_t>  m_classes;
    std::vector<_prog_ptr_t>        m_closures;
    std::map<istring, size_t>       m_clo_names;
    std::vector<_variant_t>         m_clo_defs;
    std::vector<size_t>             m_clo_ids;
    std::map<istring, _variant_t>   m_consts;
    std::map<istring, size_t>       m_labels;
    std::vector<word_t>             m_code;
    CProgram*                       m_parent;        //!!! weak
    CProgram*                       m_granpa;        //!!! weak

    size_t                          m_lines;

    //### this copy-constructor has error, check it.
    CProgram(const CProgram& r)
        : m_option_explicit(r.m_option_explicit)
        , m_name(r.m_name)
        , m_params(r.m_params)
        , m_dim_names(r.m_dim_names)
        , m_dim_defs(r.m_dim_defs)
        , m_clo_names(r.m_clo_names)
        , m_clo_defs(r.m_clo_defs)
        , m_clo_ids(r.m_clo_ids)
        , m_consts(r.m_consts)
        , m_labels(r.m_labels)
        , m_code(r.m_code)
        , m_parent(r.m_parent)
        , m_granpa(r.m_granpa)
        , m_lines(r.m_lines)
    {
        {
            auto i = r.m_funcs.begin();
            while(i != r.m_funcs.end()){
                _prog_ptr_t ptr(new CProgram(*i->second.p), false);
                ptr->m_parent = this;
                m_funcs[i->first] = {i->second.s, ptr};
                ++i;
            }
        }
        {
            auto i = r.m_plets.begin();
            while(i != r.m_plets.end()){
                _prog_ptr_t ptr(new CProgram(*i->second.p), false);
                ptr->m_parent = this;
                m_plets[i->first] = {i->second.s, ptr};
                ++i;
            }
        }
        {
            auto i = r.m_classes.begin();
            while(i != r.m_classes.end()){
                _prog_ptr_t ptr(new CProgram(*i->second), false);
                ptr->m_parent = this;
                ++i;
            }
        }
        {
            auto i = r.m_closures.begin();
            while(i != r.m_closures.end()){
                _prog_ptr_t ptr(new CProgram(**i), false);
                ptr->m_parent = this;
                m_closures.push_back( ptr );
                ++i;
            }
        }
// wprintf(L"$$$%s-copy: %ls\n", __func__, m_name.c_str());
    }

    CProgram(const wchar_t*& c, mode_t mode=&CProgram::parse_, CProgram* parent=nullptr)
        : m_parse_mode(mode)
        , m_option_explicit(false)
        , m_consts{
            {L"true"   , _variant_t{ {{{VT_BOOL,0,0,0,{VARIANT_TRUE}}}} } },
            {L"false"  , _variant_t{ {{{VT_BOOL,0,0,0,{VARIANT_FALSE}}}} }},
            {L"empty"  , _variant_t{ {{{VT_EMPTY,0,0,0,{0}}}} }           },
            {L"null"   , _variant_t{ {{{VT_NULL,0,0,0,{0}}}} }            },
            {L"nothing", _variant_t{ {{{VT_DISPATCH,0,0,0,{0}}}} }        },
        }
        , m_parent(parent)
        , m_granpa(parent ? parent->m_granpa : this)
        , m_lines(parent ? parent->m_lines : 0)
    {
        {
            istring s(L":");
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }

        while(*c && m_parse_mode){
            while(*c==L' ' || *c==L'\t') ++c;
            if(!*c) break;

            const wchar_t* c0 = c++;
            if(*c0==L'_' && *c==L'\n'){
                ++c;
                ++m_lines;
                continue;
            }else
            if(*c0==L'\''){
                while(!( *c==L'\n' || *c==L'\0' )) ++c;
                continue;
            }else
            if(L'0'<=*c0 && *c0<=L'9'){
                while( (L'0'<=*c && *c<=L'9') ) ++c;
            }else
            if(*c0==L'&' && (*c==L'H' || *c==L'h')){
                ++c;
                while( (L'0'<=*c && *c<=L'9') || (L'A'<=*c && *c<=L'F') || (L'a'<=*c && *c<=L'f') ) ++c;
            }else
            if(*c0==L'&' && (*c==L'B' || *c==L'b')){
                ++c;
                while( (L'0'<=*c && *c<=L'1') ) ++c;
            }else
            if(*c0==L'"'){
                while(1){
                    while(!( *c==L'"' || *c==L'\0' )) ++c;
                    if(*c==L'\0' || *++c!=L'"') break;
                    ++c;
                }
            }else
            if(*c0==L'#'){
                while(!( *c==L'#' || *c==L'\0' )) ++c;
                if(*c==L'#') ++c;
            }else
            if(*c0==L'{'){
                int nest = 1;
                while(nest){
                    if(*c==L'{') ++nest;
                    else
                    if(*c==L'}') --nest;
                    else
                    if(*c==L'\n') ++m_lines;

                    ++c;
                }
            }else
            if(*c0==L'['){
                int nest = 1;
                while(nest){
                    if(*c==L'[') ++nest;
                    else
                    if(*c==L']') --nest;
                    else
                    if(*c==L'\n') ++m_lines;

                    ++c;
                }
            }else
            if(
                (*c0==L'=' && *c==L'=' && *(c+1)==L'=') ||
                (*c0==L'!' && *c==L'=' && *(c+1)==L'=') ||
            false){
                ++c;++c;
            }else
            if(
                (*c0==L'<' && *c==L'=') ||
                (*c0==L'>' && *c==L'=') ||
                (*c0==L'<' && *c==L'>') ||
                (*c0==L'=' && *c==L'=') ||
                (*c0==L'+' && *c==L'+') ||
                (*c0==L'-' && *c==L'-') ||
                (*c0==L':' && *c==L'=') ||
                (*c0==L'?' && *c==L'.') ||
                (*c0==L'?' && *c==L'?') ||
            false){
                ++c;
            }else
            if(
                *c0==L'=' || *c0==L'+' || *c0==L'-' || *c0==L':' ||
                *c0==L'.' || *c0==L',' || *c0==L'(' || *c0==L')' ||
                *c0==L'\n'|| *c0==L'&' ||
                *c0==L'^' || *c0==L'*' || *c0==L'/' || *c0==L'<' || *c0==L'>' ||
                *c0==L'?' ||
            false){
                // none
            }else
            {
                while(!( 
                    *c==L'=' || *c==L'+' || *c==L'-' || *c==L':' ||
                    *c==L'.' || *c==L',' || *c==L'(' || *c==L')' ||
                    *c==L'\n'|| *c==L'&' ||
                    *c==L'^' || *c==L'*' || *c==L'/' || *c==L'<' || *c==L'>' ||
                    *c==L'?' ||
                    *c==L' ' || *c==L'\''|| *c==L'"' || *c==L'#' ||
                    *c==L'{' || *c==L'}' || *c==L'[' || *c==L']' ||
                    *c==L'\0'||
                false)) ++c;
            }

            c = (this->*m_parse_mode)(c0, c-c0);

            if(*c0==L',' && *c==L'\n'){
                ++c;
                ++m_lines;
            }
        }

        if(m_code.back().s != L":"){
            istring s(L":");
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }

        if(mode == &CProgram::parse_class_name){
            m_code.clear();
        }

        {
            istring s(L"");
            m_code.push_back( { s, map_word(s), _variant_t(), m_lines } );
        }



        if(!m_parent) bind();
// wprintf(L"$$$%s: %ls\n", __func__, m_name.c_str());
    }

    virtual ~CProgram(){
// wprintf(L"$$$%s: %ls\n", __func__, m_name.c_str());
    }

    void dump(){
        wprintf(L"***%s name:%ls\n", __func__, m_name.c_str());
        {
            auto i = m_params.begin();
            while(i != m_params.end()){
                wprintf(L"***%s params:%ls\n", __func__, i->c_str());
                ++i;
            }
        }
        {
            auto i = m_dim_names.begin();
            while(i != m_dim_names.end()){
                wprintf(L"***%s dims:%ls[%d,%zd]\n", __func__, i->first.c_str(), i->second.s, i->second.i);
                ++i;
            }
        }
        {
            auto i = m_code.begin();
            while(i != m_code.end()){
                const wchar_t* s = i->s.c_str();
                wprintf(L"***%s code:%ls(%p)[%x, %x, 0x%llx]\n", __func__, (*s==L'\n')?L"\\n":s, i->p, i->v.vt, i->v.wReserved1, i->v.llVal);
                ++i;
            }
        }
        {
            auto i = m_funcs.begin();
            while(i != m_funcs.end()){
                wprintf(L"***%s funcs:%ls(0x%p)\n", __func__, i->first.c_str(), &i->second.p);
                i->second.p->dump();
                ++i;
            }
        }
        {
            auto i = m_plets.begin();
            while(i != m_plets.end()){
                wprintf(L"***%s plets:%ls(0x%p)\n", __func__, i->first.c_str(), &i->second.p);
                i->second.p->dump();
                ++i;
            }
        }
        wprintf(L"***%s //\n", __func__);
    }



//IDispatch
private:
    ULONG       m_refc   = 1;

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){/*wprintf(L"%ls:+[%d]\n", m_name.c_str(), m_refc+1);*/ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){/*wprintf(L"%ls:-[%d]\n", m_name.c_str(), m_refc-1);*/ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        return E_NOTIMPL;
    }
};



class CProcessor : public IDispatch{
public:
    static Null s_oNull;
    typedef bool (CProcessor::*word_m)(word_t&);
    static const std::map<istring, word_m> s_words;
    static void* map_word(const istring& s);
    typedef bool (CProcessor::*inst_t)(VARIANT*);
    static const inst_t s_insts[];
    enum{
        INST_op_copy,
        INST_op_plus1,
        INST_op_plus,
        INST_op_minus1,
        INST_op_minus,
        INST_op_hat,
        INST_op_aster,
        INST_op_slash,
        INST_op_mod,
        INST_op_amp,
        INST_op_incR,
        INST_op_incL,
        INST_op_decR,
        INST_op_decL,
        INST_op_lthan,
        INST_op_gthan,
        INST_op_lethan,
        INST_op_gethan,
        INST_op_neq,
        INST_op_equal,
        INST_op_triequal,
        INST_op_trineq,
        INST_op_is,
        INST_op_not,
        INST_op_and,
        INST_op_or,
        INST_op_xor,
        INST_op_step,
        INST_op_parenL,
        INST_op_invoke,
        INST_op_invoke_put,
        INST_op_call,
        INST_op_array,
        INST_op_array_put,
        INST_op_getref,
        INST_op_exit,
        INST_stmt_if,
        INST_stmt_selectcase,
        INST_stmt_case,
        INST_stmt_caseelse,
        INST_stmt_do,
        INST_stmt_foreach,
        INST_stmt_in,
        INST_stmt_for,
        INST_stmt_next,
        INST_stmt_goto,
        INST_stmt_return,
        INST_stmt_with,
        INST_stmt_new,
        INST_stmt_redim,
        INST_stmt_redimpreserve,
    };

private:
    typedef bool (CProcessor::*mode_t)(word_t&);
    typedef bool (CProcessor::*onerr_t)();

    CProcessor*  m_parent;  //!!! weak
    CExtension*  m_ext;
    _disp_ptr_t  m_env;
    _prog_ptr_t  m_p0;
    _prog_ptr_t  m_pp;
    size_t       m_pc;
    std::vector<std::vector<_variant_t>, OnceAllocator<std::vector<_variant_t> > >              m_scope;
    std::vector<std::map<istring, _variant_t>, OnceAllocator<std::map<istring, _variant_t> > >  m_autodim;
    std::vector<std::vector<_variant_t>, OnceAllocator<std::vector<_variant_t> > >              m_with;
    std::vector<onerr_t, OnceAllocator<onerr_t> >                                               m_onerr;
    std::vector<_variant_t, OnceAllocator<_variant_t> >                                         m_s;

    std::vector<_variant_t>         m_cdims;

    std::vector<_variant_t>*        m_pgdims;
    std::map<istring, _variant_t>*  m_pgadims;

    mode_t m_mode;

    bool op_copy(VARIANT* p){
        //---OLD
        // VARIANT* pvL = p-1;
        // VARIANT* pvR = p+1;

        // if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        // if(pvL->vt == (VT_BYREF|VT_VARIANT)){
        //     VariantCopy(pvL->pvarVal, pvR);
        // }else{
        //     m_mode = &CProcessor::clock_throw_invalidcopy;
        //     --m_pc;
        //     return false;
        // }

        // while(p-1 < &m_s.back()) m_s.pop_back();//keep L for return
        //---//

        // multi-copy experiment
        int i = 0;
        while(p < &m_s.back()){
            VARIANT* pvL = p-(++i);
            VARIANT* pvR = &m_s.back();

            if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

            if(pvL->vt == (VT_BYREF|VT_VARIANT)){
                VariantCopy(pvL->pvarVal, pvR);
            }else{
                m_mode = &CProcessor::clock_throw_invalidcopy;
                --m_pc;
                return false;
            }

            m_s.pop_back();
        }

        m_s.pop_back();

        return true;
    }

    bool op_plus1(VARIANT* p){
        VARIANT* pvR = p+1;

        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v0(0LL);
        _variant_t v;
        VarAdd(&v0, pvR, &v);

        while(p-1 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_plus(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v;
        VarAdd(pvL, pvR, &v);

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_minus1(VARIANT* p){
        VARIANT* pvR = p+1;

        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v0(0LL);
        _variant_t v;
        VarSub(&v0, pvR, &v);

        while(p-1 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_minus(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v;
        VarSub(pvL, pvR, &v);

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_hat(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v;
        VarPow(pvL, pvR, &v);

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_aster(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v;
        VarMul(pvL, pvR, &v);

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_slash(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v;
        if(FAILED( m_err->m_hr = VarDiv(pvL, pvR, &v) )){
            m_mode = &CProcessor::clock_throw_;
            --m_pc;
            return false;
        }

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_mod(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v;
        VarMod(pvL, pvR, &v);

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_amp(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v;
        VarCat(pvL, pvR, &v);

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_incR(VARIANT* p){
        VARIANT* pvR = p+1;

        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v1(1LL);
        VarAdd(pvR, &v1, pvR);

        while(p-1 < &m_s.back()) m_s.pop_back();

        m_s.push_back(*pvR);

        return true;
    }

    bool op_incL(VARIANT* p){
        VARIANT* pvL = p-1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;

        _variant_t v(*pvL);

        _variant_t v1(1LL);
        VarAdd(pvL, &v1, pvL);

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_decR(VARIANT* p){
        VARIANT* pvR = p+1;

        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v1(1LL);
        VarSub(pvR, &v1, pvR);

        while(p-1 < &m_s.back()) m_s.pop_back();

        m_s.push_back(*pvR);

        return true;
    }

    bool op_decL(VARIANT* p){
        VARIANT* pvL = p-1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;

        _variant_t v(*pvL);

        _variant_t v1(1LL);
        VarSub(pvL, &v1, pvL);

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_lthan(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v( VarCmp(pvL, pvR, 0, 0) == VARCMP_LT );

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_gthan(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v( VarCmp(pvL, pvR, 0, 0) == VARCMP_GT );

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_lethan(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v( VarCmp(pvL, pvR, 0, 0) == VARCMP_LT || VarCmp(pvL, pvR, 0, 0) == VARCMP_EQ );

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_gethan(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v( VarCmp(pvL, pvR, 0, 0) == VARCMP_GT || VarCmp(pvL, pvR, 0, 0) == VARCMP_EQ );

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_neq(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v( VarCmp(pvL, pvR, 0, 0) != VARCMP_EQ );

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_equal(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v( VarCmp(pvL, pvR, 0, 0) == VARCMP_EQ );

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_triequal(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        bool b = (
            pvL->vt == pvR->vt &&
            (
                (pvL->vt == VT_EMPTY)                                      ||
                (pvL->vt == VT_NULL)                                       ||
                (pvL->vt == VT_ERROR)                                      ||
                (pvL->vt == VT_DISPATCH           && pvL->pdispVal == pvR->pdispVal) ||
                (pvL->vt == (VT_ARRAY|VT_VARIANT) && pvL->parray == pvR->parray)     ||
                // (pvL->vt == VT_I1   && pvL-> == pvR->) ||
                // (pvL->vt == VT_I2   && pvL-> == pvR->) ||
                (pvL->vt == VT_I4   && pvL->lVal  == pvR->lVal)      ||
                (pvL->vt == VT_I8   && pvL->llVal == pvR->llVal)     ||
                // (pvL->vt == VT_UI1  && pvL-> == pvR->) ||
                (pvL->vt == VT_UI2  && pvL->uiVal == pvR->uiVal)     ||
                // (pvL->vt == VT_UI4  && pvL-> == pvR->) ||
                (pvL->vt == VT_UI8  && pvL->ullVal == pvR->ullVal)   ||
                // (pvL->vt == VT_R4   && pvL-> == pvR->) ||
                (pvL->vt == VT_R8   && pvL->dblVal == pvR->dblVal)   ||
                // (pvL->vt == VT_CY   && pvL-> == pvR->) ||
                (pvL->vt == VT_DATE && pvL->date == pvR->date)       ||
                (pvL->vt == VT_BOOL && pvL->boolVal == pvR->boolVal) ||
                (pvL->vt == VT_BSTR && std::wcscmp(pvL->bstrVal, pvR->bstrVal) == 0) ||
            false)             &&
        true);
       _variant_t v( b );

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_trineq(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        bool b = (
            pvL->vt == pvR->vt &&
            (
                (pvL->vt == VT_EMPTY)                                      ||
                (pvL->vt == VT_NULL)                                       ||
                (pvL->vt == VT_ERROR)                                      ||
                (pvL->vt == VT_DISPATCH           && pvL->pdispVal == pvR->pdispVal) ||
                (pvL->vt == (VT_ARRAY|VT_VARIANT) && pvL->parray == pvR->parray)     ||
                // (pvL->vt == VT_I1   && pvL-> == pvR->) ||
                // (pvL->vt == VT_I2   && pvL-> == pvR->) ||
                (pvL->vt == VT_I4   && pvL->lVal  == pvR->lVal)      ||
                (pvL->vt == VT_I8   && pvL->llVal == pvR->llVal)     ||
                // (pvL->vt == VT_UI1  && pvL-> == pvR->) ||
                (pvL->vt == VT_UI2  && pvL->uiVal == pvR->uiVal)     ||
                // (pvL->vt == VT_UI4  && pvL-> == pvR->) ||
                (pvL->vt == VT_UI8  && pvL->ullVal == pvR->ullVal)   ||
                // (pvL->vt == VT_R4   && pvL-> == pvR->) ||
                (pvL->vt == VT_R8   && pvL->dblVal == pvR->dblVal)   ||
                // (pvL->vt == VT_CY   && pvL-> == pvR->) ||
                (pvL->vt == VT_DATE && pvL->date == pvR->date)       ||
                (pvL->vt == VT_BOOL && pvL->boolVal == pvR->boolVal) ||
                (pvL->vt == VT_BSTR && std::wcscmp(pvL->bstrVal, pvR->bstrVal) == 0) ||
            false)             &&
        true);
       _variant_t v( !b );

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_is(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v( pvL->vt==VT_DISPATCH && pvR->vt==VT_DISPATCH && pvL->pdispVal==pvR->pdispVal );

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_not(VARIANT* p){
        VARIANT* pvR = p+1;

        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v;
        VarNot(pvR, &v);

        while(p-1 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_and(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v;
        VarAnd(pvL, pvR, &v);

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_or(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v;
        VarOr(pvL, pvR, &v);

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_xor(VARIANT* p){
        VARIANT* pvL = p-1;
        VARIANT* pvR = p+1;

        if(pvL->vt == (VT_BYREF|VT_VARIANT)) pvL = pvL->pvarVal;
        if(pvR->vt == (VT_BYREF|VT_VARIANT)) pvR = pvR->pvarVal;

        _variant_t v;
        VarXor(pvL, pvR, &v);

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_step(VARIANT* p){
        VARIANT* pvI = p-3;
        VARIANT* pvS = p-1;

        if(pvI->vt == (VT_BYREF|VT_VARIANT)) pvI = pvI->pvarVal;
        if(pvS->vt == (VT_BYREF|VT_VARIANT)) pvS = pvS->pvarVal;

        VarAdd(pvI, pvS, pvI);

        m_s.pop_back();//VTX_PCFORNEXT
        m_s.pop_back();//me

        ++m_pc;

        return true;
    }

    bool op_parenL(VARIANT* p){
        int i = 0;
        while(p+i < &m_s.back()){
            *(p+i) = *(p+i+1);
            (p+i+1)->vt = VT_EMPTY;
            ++i;
        }

        m_s.pop_back();

        return false;
    }

    bool op_invoke(VARIANT* p){
        // '// PROP-----------------------------
        //  x = o.comparemode:         '| GET
        //      o.comparemode = 0      '| PUT
        // 'x = o.comparemode()        runtime: invalid invoke
        //      o.comparemode() = 0    '| PUT?
        // 'call o.comparemode:        runtime: invalid invoke
        // 'call o.comparemode()       runtime: invalid invoke
        // 'x = o.item "a":            compile: invalid end
        // '    o.item "a" = 6         runtime: incorrect type "a"
        //  x = o.item("a")            '| GET|CALL
        //      o.item("a") = 6        '| PUT
        // 'call o.item "a":           compile: invalid end
        // 'call o.item("a")           runtime: invalid invoke
        // '// METHOD---------------------------
        // 'x = o.exists "a":          compile: invalid end
        //      o.exists "a"  = true   '| ?
        //  x = o.exists("a")          '| CALL|GET
        // '    o.exists("a") = true   runtime: invalid invoke
        // 'x = o.add "b", "c":        compile: invalid end
        // '    o.add "b", "c" = 0     runtime: incorrect type "c"
        //  x = o.add("b", "c")        '| CALL|GET
        // '    o.add("b", "c") = 0    runtime: invalid invoke
        // 'call o.add "b", "c":       compile: invalid end
        // call o.add("d", "e")        '| CALL
        VARIANT* pvO = p-2;
        VARIANT* pvD = p-1;

        if(pvO->vt == (VT_BYREF|VT_VARIANT)) pvO = pvO->pvarVal;
        if(pvD->vt == (VT_BYREF|VT_VARIANT)) pvD = pvD->pvarVal;

        _variant_t args[32];
        UINT n = (VARIANT*)&m_s.back() - p;{
            if(sizeof(args)/sizeof(args[0]) < n) throw _com_error(DISP_E_BADPARAMCOUNT);

            int i = n;
            while(0 < i){
                args[n-i].Attach( *(p+i) );
                --i;
            }
        }

        DISPID dids[1] = { DISPID_PROPERTYPUT };

        DISPPARAMS param = {
            args,
            dids,
            n,
            0,
        };

        _variant_t vO(*pvO);
        _variant_t vD(*pvD);

        while(p-3 < &m_s.back()) m_s.pop_back();

        _variant_t res;
        if(FAILED( m_err->m_hr = vO.pdispVal->Invoke(vD.lVal, IID_NULL, 0, DISPATCH_METHOD|DISPATCH_PROPERTYGET, &param, &res, m_err, &m_err->m_an) )){
            m_mode = &CProcessor::clock_throw_object;
            --m_pc;
            return false;
        }

        m_s.push_back(res);

        return true;
    }

    //!!! ALMOST same as op_invoke
    bool op_invoke_put(VARIANT* p){
        VARIANT* pvO = p-2;
        VARIANT* pvD = p-1;

        if(pvO->vt == (VT_BYREF|VT_VARIANT)) pvO = pvO->pvarVal;
        if(pvD->vt == (VT_BYREF|VT_VARIANT)) pvD = pvD->pvarVal;

        _variant_t args[32];
        UINT n = (VARIANT*)&m_s.back() - p;{
            if(sizeof(args)/sizeof(args[0]) < n) throw _com_error(DISP_E_BADPARAMCOUNT);

            int i = n;
            while(0 < i){
                args[n-i].Attach( *(p+i) );
                --i;
            }
        }

        DISPID dids[1] = { DISPID_PROPERTYPUT };

        DISPPARAMS param = {
            args,
            dids,
            n,
            1,//!!! DIFF
        };

        _variant_t vO(*pvO);
        _variant_t vD(*pvD);

        while(p-3 < &m_s.back()) m_s.pop_back();

        _variant_t res;
        if(FAILED( m_err->m_hr = vO.pdispVal->Invoke(vD.lVal, IID_NULL, 0, DISPATCH_PROPERTYPUT, &param, &res, m_err, &m_err->m_an) )){//!!! DIFF
            m_mode = &CProcessor::clock_throw_object;
            --m_pc;
            return false;
        }

        m_s.push_back(res);

        return true;
    }

    bool op_call(VARIANT* p){
        VARIANT* pv = p-1;

        CProgram* prog = (CProgram*)pv->byref;
        m_scope.push_back( prog->m_dim_defs );
        std::vector<_variant_t>& newscope = m_scope.back();
        m_autodim.push_back( {} );
        m_with.push_back( { {{{{VT_EMPTY,0,0,0,{}}}}} } );
        m_onerr.push_back(&CProcessor::onerr_goto0);

        int i = 0;
        while( (p+1)+i <= &m_s.back() ){
            if(i < prog->m_params.size()){
                if(((p+1)+i)->vt == VT_ERROR){
                    // none
                }else{
                    newscope[i+1].Attach( *((p+1)+i) );
                }
            }else{
                m_mode = &CProcessor::clock_throw_toomanyarg;
                --m_pc;
                return false;
            }
            ++i;
        }

        while(p-2 < &m_s.back()) m_s.pop_back();

        {
            m_s.push_back(_variant_t());

            _variant_t& ret = newscope[0];
            ret.vt = (VT_BYREF|VT_VARIANT);
            ret.pvarVal = &m_s.back();
        }
        {
            _variant_t v;
            v.wReserved1 = VTX_POFORRETURN;
            v.byref = m_pp;
            m_s.push_back(v);
        }
        {
            _variant_t v;
            v.wReserved1 = VTX_PCFORRETURN;
            v.llVal = m_pc;
            m_s.push_back(v);
        }
        {
            _variant_t v;
            v.wReserved1 = VTX_GROUND;
            m_s.push_back(v);
        }

        m_pp = prog;
        m_pc = 0;

        return true;
    }

    bool op_array(VARIANT* p){
        VARIANT* pa = p-1;

        if(pa->vt == (VT_BYREF|VT_VARIANT)) pa = pa->pvarVal;

        int n = 0;{
            int i = 0;
            int m = 1;
            while( (p+1)+i <= &m_s.back()-1 ){
                VARIANT* pn = (p+1)+i;
                if(pn->vt == (VT_BYREF|VT_VARIANT)) pn = pn->pvarVal;

                if(pa->parray->rgsabound[pa->parray->cDims-(i+1)].cElements <= pn->llVal){
                    m_mode = &CProcessor::clock_throw_invalidindex;
                    --m_pc;
                    return false;
                }

                m *= pa->parray->rgsabound[pa->parray->cDims-(i+1)].cElements;
                n += m * pn->llVal;

                ++i;
            }

            {
                VARIANT* pn = (p+1)+i;
                if(pn->vt == (VT_BYREF|VT_VARIANT)) pn = pn->pvarVal;

                if(pa->parray->rgsabound[pa->parray->cDims-(i+1)].cElements <= pn->llVal){
                    m_mode = &CProcessor::clock_throw_invalidindex;
                    --m_pc;
                    return false;
                }

                n += pn->llVal;
            }
        }

        _variant_t v;{
            VARIANT* pav;
            SafeArrayAccessData(pa->parray, (void**)&pav);
            {
                VariantCopy(&v, &pav[n]);
            }
            SafeArrayUnaccessData(pa->parray);
        }

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_array_put(VARIANT* p){
        VARIANT* pa = p-1;
        VARIANT* pv = &m_s.back();

        if(pa->vt == (VT_BYREF|VT_VARIANT)) pa = pa->pvarVal;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        int n = 0;{
            int i = 0;
            int m = 1;
            while( (p+1)+i <= &m_s.back()-2 ){
                VARIANT* pn = (p+1)+i;
                if(pn->vt == (VT_BYREF|VT_VARIANT)) pn = pn->pvarVal;

                if(pa->parray->rgsabound[pa->parray->cDims-(i+1)].cElements <= pn->llVal){
                    m_mode = &CProcessor::clock_throw_invalidindex;
                    --m_pc;
                    return false;
                }

                m *= pa->parray->rgsabound[pa->parray->cDims-(i+1)].cElements;
                n += m * pn->llVal;

                ++i;
            }
            {
                VARIANT* pn = (p+1)+i;
                if(pn->vt == (VT_BYREF|VT_VARIANT)) pn = pn->pvarVal;

                if(pa->parray->rgsabound[pa->parray->cDims-(i+1)].cElements <= pn->llVal){
                    m_mode = &CProcessor::clock_throw_invalidindex;
                    --m_pc;
                    return false;
                }

                n += pn->llVal;
            }
        }

        _variant_t v;{
            VARIANT* pav = nullptr;
            SafeArrayAccessData(pa->parray, (void**)&pav);
            {
                VariantCopy(&pav[n], pv);
            }
            SafeArrayUnaccessData(pa->parray);
        }

        while(p-2 < &m_s.back()) m_s.pop_back();

        m_s.push_back(v);

        return true;
    }

    bool op_getref(VARIANT* p){
        VARIANT* pv = p+1;

        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        auto& funcs = m_pp->m_parent ? m_pp->m_parent->m_funcs : m_pp->m_funcs;
        auto found = funcs.find(pv->bstrVal);
        if(found != funcs.end()){
            while(p-1 < &m_s.back()) m_s.pop_back();

            CProcessor* pros( new CProcessor(*this, found->second.p) );
            pros->m_pc = pros->m_pp->m_code.size()-1;   // to don't-run by default

            _variant_t v;
            v.vt = VT_DISPATCH;
            v.pdispVal = pros;
            m_s.push_back(v);
        }else{
            m_mode = &CProcessor::clock_throw_funcnotfound;
            --m_pc;
            return false;
        }

        return true;
    }

    bool op_exit(VARIANT* p){
        while(p-1 < &m_s.back()) m_s.pop_back();

        m_mode = &CProcessor::clock_exit;
        --m_pc;
        return false;
    }

    bool stmt_if(VARIANT* p){
        VARIANT* pv = p+1;

        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        if( *(_variant_t*)pv ){
            while(p-1 < &m_s.back()) m_s.pop_back();
        }else{
            while(p-1 < &m_s.back()) m_s.pop_back();

            _variant_t v;
            v.wReserved1 = VTX_SKIPTOELSE;
            v.llVal = 1;
            m_s.push_back(v);

            m_mode = &CProcessor::clock_skiptoelse;
        }

        return false;
    }

    bool stmt_selectcase(VARIANT* p){
        // none

        return false;
    }

    bool stmt_case(VARIANT* p){
        VARIANT* pvS = p-1;
        if(pvS->vt == (VT_BYREF|VT_VARIANT)) pvS = pvS->pvarVal;

        int i=1;
        while(p+i <= &m_s.back()){
            VARIANT* pvC = p+i;
            if(pvC->vt == (VT_BYREF|VT_VARIANT)) pvC = pvC->pvarVal;

            if(::VarCmp(pvS, pvC, 0, 0) == VARCMP_EQ){
                while(p-3 < &m_s.back()) m_s.pop_back();

                _variant_t v;
                v.wReserved1 = VTX_GROUND;
                m_s.push_back(v);

                return false;
            }

            ++i;
        }

        {
            while(p-1 < &m_s.back()) m_s.pop_back();

            _variant_t v;
            v.wReserved1 = VTX_SKIPTOCASE;
            v.llVal = 1;
            m_s.push_back(v);

            m_mode = &CProcessor::clock_skiptocase;
        }

        return false;
    }

    bool stmt_caseelse(VARIANT* p){
        while(p-3 < &m_s.back()) m_s.pop_back();

        _variant_t v;
        v.wReserved1 = VTX_GROUND;
        m_s.push_back(v);

        return false;
    }

    bool stmt_do(VARIANT* p){
        VARIANT* pve = p+1;
        VARIANT* pvc = p-1;

        if(pve->vt == (VT_BYREF|VT_VARIANT)) pve = pve->pvarVal;

        bool b = (pvc->wReserved1==VTX_PCFORLOOPWHILE) ? (bool)*(_variant_t*)pve : !(bool)*(_variant_t*)pve;

        if( b ){
            while(p-0 < &m_s.back()) m_s.pop_back();// keep stmt_do, VTX_PCFORLOOPxxxxx

            _variant_t v;
            v.wReserved1 = VTX_GROUND;
            m_s.push_back(v);
        }else{
            while(p-2 < &m_s.back()) m_s.pop_back();

            _variant_t v;
            v.wReserved1 = VTX_SKIPTOLOOP;
            v.llVal = 1;
            m_s.push_back(v);

            m_mode = &CProcessor::clock_skiptoloop;
        }

        return false;
    }

    bool stmt_foreach(VARIANT* p){
        VARIANT* pvI = p+1;
        VARIANT* pvO = p+2;

        if(pvI->vt == (VT_BYREF|VT_VARIANT)) pvI = pvI->pvarVal;
        if(pvO->vt == (VT_BYREF|VT_VARIANT)) pvO = pvO->pvarVal;

        _com_ptr_t<_com_IIID<IEnumVARIANT, &IID_IEnumVARIANT> > pe(*pvO);
        ULONG n;
        _com_util::CheckError( pe->Next(1, pvI, &n) );
        if(n){
            {
                _variant_t v;
                v.wReserved1 = VTX_GROUND;
                m_s.push_back(v);
            }
        }else{
            while(p-1 < &m_s.back()) m_s.pop_back();

            _variant_t v;
            v.wReserved1 = VTX_SKIPTONEXT;
            v.llVal = 1;
            m_s.push_back(v);

            m_mode = &CProcessor::clock_skiptonext;
        }

        return false;
    }

    bool stmt_in(VARIANT* p){
        VARIANT* pv = p+1;

        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        DISPPARAMS param = {
            nullptr,
            nullptr,
            0,
            0,
        };
        _variant_t res;
        _com_util::CheckError( pv->pdispVal->Invoke(DISPID_NEWENUM, IID_NULL, 0, DISPATCH_METHOD, &param, &res, nullptr, nullptr) );

        while(p-1 < &m_s.back()) m_s.pop_back();

        m_s.push_back(res);
        {
            _variant_t v;
            v.wReserved1 = VTX_PCFORNEXT;
            v.llVal = m_pc-1;
            m_s.push_back( v );
        }

        return true;
    }

    bool stmt_for(VARIANT* p){
        if(((VARIANT*)&m_s.back()-p) == 2){
            _variant_t v;{
                v.vt = VT_I8;
                v.llVal = 1;
            }
            m_s.push_back( v );
        }

        VARIANT* pvI = p+1;
        VARIANT* pvT = p+2;
        VARIANT* pvS = p+3;

        if(pvI->vt == (VT_BYREF|VT_VARIANT)) pvI = pvI->pvarVal;
        if(pvT->vt == (VT_BYREF|VT_VARIANT)) pvT = pvT->pvarVal;
        if(pvS->vt == (VT_BYREF|VT_VARIANT)) pvS = pvS->pvarVal;

        bool b = (0 < pvS->llVal) ? pvI->llVal<=pvT->llVal : pvT->llVal<=pvI->llVal;

        if(b){
            {
                _variant_t v;
                v.wReserved1 = VTX_INST;
                v.byref = (void*)&s_insts[INST_op_step];
                m_s.push_back( v );
            }
            {
                _variant_t v;
                v.wReserved1 = VTX_PCFORNEXT;
                v.llVal = m_pc-1;
                m_s.push_back( v );
            }
            {
                _variant_t v;
                v.wReserved1 = VTX_GROUND;
                m_s.push_back(v);
            }
        }else
        {
            while(p-1 < &m_s.back()) m_s.pop_back();

            _variant_t v;
            v.wReserved1 = VTX_SKIPTONEXT;
            v.llVal = 1;
            m_s.push_back(v);

            m_mode = &CProcessor::clock_skiptonext;
        }

        return false;
    }

    bool stmt_next(VARIANT* p){
        while(p-1 < &m_s.back()) m_s.pop_back();

        m_s.pop_back();     //VTX_GROUND
        m_pc = m_s.back().llVal;

        return true;
    }

    bool stmt_goto(VARIANT* p){
        size_t pc = SIZE_MAX;{
            VARIANT* pv = p+1;

            if(pv->vt == VT_UI8){
                pc = pv->ullVal;
            }else{
                // goto by line-head number
                if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

                if(pv->vt == VT_I8){
                    wchar_t buf[19+1];
                    swprintf(buf, 19+1, L"%lld", pv->llVal);
                    auto found = m_pp->m_labels.find(istring(buf));
                    if(found != m_pp->m_labels.end()){
                        pc = found->second;
                    }
                }
            }

            if(pc == SIZE_MAX){
                m_mode = &CProcessor::clock_throw_cantgotothere;
                --m_pc;
                return false;
            }
        }

        while(p-1 < &m_s.back()) m_s.pop_back();

        int ngnd = 0;
        if(pc < m_pc){
            while(pc < --m_pc){
                if(
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_dowhile    ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_dountil    ||
                false){
                    if(ngnd == 0){
                        while(! (m_s.back().wReserved1==VTX_INST && *((inst_t*)m_s.back().byref)==&CProcessor::stmt_do) ) m_s.pop_back();
                        m_s.pop_back();//stmt_do
                        m_s.pop_back();//VTX_PCFORLOOP
                    }else{
                        --ngnd;
                    }
                }else
                if( *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_for ){
                    if(ngnd == 0){
                        while(! (m_s.back().wReserved1==VTX_INST && *((inst_t*)m_s.back().byref)==&CProcessor::stmt_for) ) m_s.pop_back();
                        m_s.pop_back();//stmt_for||stmt_foreach
                    }else{
                        --ngnd;
                    }
                }else
                if( *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_foreach ){
                    if(ngnd == 0){
                        while(! (m_s.back().wReserved1==VTX_INST && *((inst_t*)m_s.back().byref)==&CProcessor::stmt_foreach) ) m_s.pop_back();
                        m_s.pop_back();//stmt_for||stmt_foreach
                    }else{
                        --ngnd;
                    }
                }else
                if( *((word_m*)&m_pp->m_code[m_pc].p) == &CProcessor::word_selectcase ){
                    if(ngnd == 0){
                        while(! (m_s.back().wReserved1==VTX_GROUND) ) m_s.pop_back();
                        m_s.pop_back();//VTX_GROUND
                    }else{
                        --ngnd;
                    }
                }else
                if( *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_with ){
                    if(ngnd == 0){
                        m_with.back().pop_back();
                    }else{
                        --ngnd;
                    }
                }else
                if(
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_loop       ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_next       ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_endselect  ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_endwith    ||
                false){
                    ++ngnd;
                }
            }
        }else{
            --m_pc;
            while(++m_pc < pc){
                if(
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_dowhile    ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_dountil    ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_for        ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_foreach    ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_selectcase ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_with       ||
                false){
                    ++ngnd;
                }else
                if( *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_loop ){
                    if(ngnd == 0){
                        while(! (m_s.back().wReserved1==VTX_INST && *((inst_t*)m_s.back().byref)==&CProcessor::stmt_do) ) m_s.pop_back();
                        m_s.pop_back();//stmt_do
                        m_s.pop_back();//VTX_PCFORLOOP
                    }else{
                        --ngnd;
                    }
                }else
                if( *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_next ){
                    if(ngnd == 0){
                        while(!( 
                            (m_s.back().wReserved1==VTX_INST && *((inst_t*)m_s.back().byref)==&CProcessor::stmt_for)     ||
                            (m_s.back().wReserved1==VTX_INST && *((inst_t*)m_s.back().byref)==&CProcessor::stmt_foreach) ||
                        false)) m_s.pop_back();
                        m_s.pop_back();//stmt_for||stmt_foreach
                    }else{
                        --ngnd;
                    }
                }else
                if( *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_endselect ){
                    if(ngnd == 0){
                        while(! (m_s.back().wReserved1==VTX_GROUND) ) m_s.pop_back();
                        m_s.pop_back();//VTX_GROUND
                    }else{
                        --ngnd;
                    }
                }else
                if( *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_endwith ){
                    if(ngnd == 0){
                        m_with.back().pop_back();
                    }else{
                        --ngnd;
                    }
                }
            }
        }

        if(ngnd){
            m_mode = &CProcessor::clock_throw_cantgotothere;
            --m_pc;
            return false;
        }

        return true;
    }

    bool stmt_return(VARIANT* p){
        p->wReserved1 = VTX_RETURN;

        m_pc = m_pp->m_code.size() - 1;

        return false;
    }

    bool stmt_with(VARIANT* p){
        VARIANT* pv = p+1;

        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        m_with.back().push_back(*pv);

        while(p-1 < &m_s.back()) m_s.pop_back();

        return false;
    }

    bool stmt_new(VARIANT* p){
        VARIANT* pv = p+1;

        _disp_ptr_t newo;
        if(pv->vt == VT_DISPATCH){
            newo = pv->pdispVal;
        }else{
            //#?# keep this ptr as parent-child
            CProcessor* pros = new CProcessor(*this, (CProgram*)pv->byref);
            pros->m_pc = pros->m_pp->m_code.size()-1;   // to don't-run by default
            newo.Attach(pros);
        }

        while(p-1 < &m_s.back()) m_s.pop_back();

        {
            _variant_t v;
            v.vt = VT_DISPATCH;
            v.pdispVal = newo.Detach();
            m_s.push_back(v);
        }

        return true;
    }

    bool stmt_redim(VARIANT* p){
        VARIANT* pa = p+1;

        if(pa->vt == (VT_BYREF|VT_VARIANT)) pa = pa->pvarVal;

        int n = 0;
        SAFEARRAYBOUND sab[64] = {};{
            int i = 0;
            while( (p+2)+i <= &m_s.back() ){
                if(i < sizeof(sab)/sizeof(sab[0])){
                    VARIANT* pn = (p+2)+i;
                    if(pn->vt == (VT_BYREF|VT_VARIANT)) pn = pn->pvarVal;

                    sab[i] = { (ULONG)pn->llVal+1, 0 };
                }else{
                    m_mode = &CProcessor::clock_throw_toomanyarg;
                    --m_pc;
                    return false;
                }
                ++i;
            }
            n = i;
        }

        VariantClear(pa);
        pa->vt = (VT_ARRAY|VT_VARIANT);
        pa->parray = SafeArrayCreate(VT_VARIANT, n, sab);

        while(p-1 < &m_s.back()) m_s.pop_back();

        m_s.push_back( _variant_t() );

        return false;
    }

    bool stmt_redimpreserve(VARIANT* p){
        VARIANT* pa = p+1;

        if(pa->vt == (VT_BYREF|VT_VARIANT)) pa = pa->pvarVal;

        int n = 0;
        SAFEARRAYBOUND sab[64] = {};{
            int i = 0;
            while( (p+2)+i <= &m_s.back() ){
                if(i < sizeof(sab)/sizeof(sab[0])){
                    VARIANT* pn = (p+2)+i;
                    if(pn->vt == (VT_BYREF|VT_VARIANT)) pn = pn->pvarVal;

                    sab[i] = { (ULONG)pn->llVal+1, 0 };
                }else{
                    m_mode = &CProcessor::clock_throw_toomanyarg;
                    --m_pc;
                    return false;
                }
                ++i;
            }
            n = i;
        }

        {
            SAFEARRAY* newa = SafeArrayCreate(VT_VARIANT, n, sab);

            int nElements = 1;{
                int i=0;
                while(i < pa->parray->cDims){
                    nElements *= pa->parray->rgsabound[i++].cElements;
                }
            }

            VARIANT* po;
            VARIANT* pn;
            SafeArrayAccessData(pa->parray, (void**)&po);
            SafeArrayAccessData(newa, (void**)&pn);
            {
                int i=0;
                while(i < nElements){
                    VariantCopy(&pn[i], &po[i]);
                    ++i;
                }
            }
            SafeArrayUnaccessData(newa);
            SafeArrayUnaccessData(pa->parray);

            VariantClear(pa);
            pa->vt = (VT_ARRAY|VT_VARIANT);
            pa->parray = newa;
        }

        while(p-1 < &m_s.back()) m_s.pop_back();

        m_s.push_back( _variant_t() );

        return false;
    }

private:
    bool clock_skiptoelse(word_t& pc){
        VARIANT* pv = &m_s.back();

        if(*((word_m*)pc.p) == &CProcessor::word_exit){
            return false;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_if){
            ++pv->llVal;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_endif){
            if(--pv->llVal == 0){
                m_s.pop_back();
                m_mode = &CProcessor::clock_;
            }
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_else && pv->llVal == 1){
            m_s.pop_back();
            m_mode = &CProcessor::clock_;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_elseif && pv->llVal == 1){
            m_s.pop_back();
            m_mode = &CProcessor::clock_;

            return word_if(pc);
        }else
        {
            // none
        }

        return true;
    }

    bool clock_skiptoendif(word_t& pc){
        VARIANT* pv = &m_s.back();

        if(*((word_m*)pc.p) == &CProcessor::word_exit){
            return false;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_if){
            ++pv->llVal;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_endif){
            if(--pv->llVal == 0){
                m_s.pop_back();
                m_mode = &CProcessor::clock_;
            }
        }else
        {
            // none
        }

        return true;
    }

    bool clock_skiptocase(word_t& pc){
        VARIANT* pv = &m_s.back();

        if(*((word_m*)pc.p) == &CProcessor::word_exit){
            return false;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_selectcase){
            ++pv->llVal;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_endselect){
            if(--pv->llVal == 0){
                m_s.pop_back();
                m_mode = &CProcessor::clock_;
                --m_pc;
            }
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_case && pv->llVal == 1){
            m_s.pop_back();
            m_mode = &CProcessor::clock_;
            --m_pc;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_caseelse && pv->llVal == 1){
            m_s.pop_back();
            m_mode = &CProcessor::clock_;
            --m_pc;
        }else
        {
            // none
        }

        return true;
    }

    bool clock_skiptoendselect(word_t& pc){
        VARIANT* pv = &m_s.back();

        if(*((word_m*)pc.p) == &CProcessor::word_exit){
            return false;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_selectcase){
            ++pv->llVal;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_endselect){
            if(--pv->llVal == 0){
                m_s.pop_back();
                m_mode = &CProcessor::clock_;
                --m_pc;
            }
        }else
        {
            // none
        }

        return true;
    }

    bool clock_skiptoloop(word_t& pc){
        VARIANT* pv = &m_s.back();

        if(*((word_m*)pc.p) == &CProcessor::word_exit){
            return false;
        }else
        if(
            *((word_m*)pc.p) == &CProcessor::word_dowhile  ||
            *((word_m*)pc.p) == &CProcessor::word_dountil  ||
        false){
            ++pv->llVal;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_loop){
            if(--pv->llVal == 0){
                m_s.pop_back();
                m_mode = &CProcessor::clock_;
            }
        }else
        {
            // none
        }

        return true;
    }

    bool clock_skiptonext(word_t& pc){
        VARIANT* pv = &m_s.back();

        if(*((word_m*)pc.p) == &CProcessor::word_exit){
            return false;
        }else
        if(
            *((word_m*)pc.p) == &CProcessor::word_for     ||
            *((word_m*)pc.p) == &CProcessor::word_foreach ||
        false){
            ++pv->llVal;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_next){
            if(--pv->llVal == 0){
                m_s.pop_back();
                m_mode = &CProcessor::clock_;

                if(*((word_m*)m_pp->m_code[m_pc].p) != &CProcessor::word_colon) ++m_pc;
            }
        }else
        {
            // none
        }

        return true;
    }

    bool clock_skiptocatch(word_t& pc){
        // VARIANT* pv = &m_s.back();

        if(*((word_m*)pc.p) == &CProcessor::word_exit){
            return false;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_onerrorcatch){
            m_mode = &CProcessor::clock_;
        }else
        {
            // none
        }

        return true;
    }

    bool clock_skiptogoto0(word_t& pc){
        // VARIANT* pv = &m_s.back();

        if(*((word_m*)pc.p) == &CProcessor::word_exit){
            return false;
        }else
        if(*((word_m*)pc.p) == &CProcessor::word_onerrorgoto0){
            m_mode = &CProcessor::clock_;
        }else
        {
            // none
        }

        return true;
    }

    bool clock_shortcircuit(word_t& pc){
        --m_pc;

        int nest = 1;
        while(nest){
            if(*((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_parenL){
                ++nest;
            }else
            if(*((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_parenR){
                --nest;
            }else
            if(
                nest == 1  &&
                (
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_comma     ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_colon     ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_else      ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_elseif    ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_endif     ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_in        ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_to        ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_step      ||
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_exit      ||
                false)      &&
            true){
                --nest;
            }else
            if(*((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_exit){
                return false;
            }else
            {}

            if(nest) ++m_pc;
        }

        m_mode = &CProcessor::clock_;
        return true;
    }

    bool clock_dispatch(word_t& pc){
        VARIANT* pv = &m_s.back();
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        DISPID id;
        LPOLESTR name = (LPOLESTR)pc.s.c_str();
        if(pv->vt == VT_DISPATCH && pv->pdispVal){
            if(SUCCEEDED(m_err->m_hr = pv->pdispVal->GetIDsOfNames(IID_NULL, &name, 1, 0, &id))){
                if(id == DISPID_GETIMMEDIATELY){
                    DISPPARAMS param = { nullptr, nullptr, 0, 0 };
                    _variant_t res;
                    _com_util::CheckError( pv->pdispVal->Invoke(id, IID_NULL, 0, DISPATCH_PROPERTYGET, &param, &res, nullptr, nullptr) );
                    m_s.pop_back();
                    m_s.push_back( res );
                }else{
                    bool put = ( *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_coleq );
                    if(put) ++m_pc;

                    {
                        _variant_t v;
                        v.vt = VT_I4;
                        v.lVal = id;
                        m_s.push_back( v );
                    }
                    {
                        _variant_t v;
                        v.wReserved1 = VTX_INST;
                        v.byref = put ? (void*)&s_insts[INST_op_invoke_put] : (void*)&s_insts[INST_op_invoke];
                        m_s.push_back( v );
                    }

                    if( *((word_m*)m_pp->m_code[m_pc].p) != &CProcessor::word_parenL ){
                        int ni = 0;{
                            auto i = m_s.rbegin();
                            while(!( i->wReserved1 == VTX_GROUND )){
                                if(i->wReserved1 == VTX_INST) ++ni;
                                ++i;
                            }
                        }

                        if(1 < ni){
                            do_left_invoke();
                        }
                    }
                }

                m_mode = &CProcessor::clock_;
                return true;
            }else{
                m_mode = &CProcessor::clock_throw_nomember;
                --m_pc;
                return true;//#?# throw
            }
        }else{
            m_mode = &CProcessor::clock_throw_noobject;
            --m_pc;
            return true;//#?# throw
        }

        return false;
    }

    bool clock_throw_invalidcopy(word_t& pc){
        volatile int e=1;(void)e;// dummy now
        return false;
    }

    bool clock_throw_toomanyarg(word_t& pc){
        volatile int e=2;(void)e;// dummy now
        return false;
    }

    bool clock_throw_json(word_t& pc){
        volatile int e=3;(void)e;// dummy now
        return false;
    }

    bool clock_throw_funcnotfound(word_t& pc){
        volatile int e=4;(void)e;// dummy now
        return false;
    }

    bool clock_throw_timeout(word_t& pc){
        volatile int e=5;(void)e;// dummy now
        return false;
    }

    bool clock_throw_cantgotothere(word_t& pc){
        volatile int e=6;(void)e;// dummy now
        return false;
    }

    bool clock_throw_invalidindex(word_t& pc){
        volatile int e=7;(void)e;// dummy now
        return false;
    }

    bool clock_throw_unexceptedend(word_t& pc){
        volatile int e=8;(void)e;// dummy now
        return false;
    }

    bool clock_throw_object(word_t& pc){
        if(m_err->m_hr == DISP_E_BADPARAMCOUNT){
            m_err->wCode = 450;
            SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
            SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Bad param count");
        }else
        if(m_err->m_hr == CO_E_CLASSSTRING){
            m_err->wCode = 429;
            SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
            SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Bad class name");
        }
        return (this->*m_onerr.back())();
    }

    bool clock_throw_unknownname(word_t& pc){
        m_err->wCode = 500;
        SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
        SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Unknown name");
        return (this->*m_onerr.back())();
    }

    bool clock_throw_noobject(word_t& pc){
        m_err->wCode = 424;
        SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
        SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"No object");
        return (this->*m_onerr.back())();
    }

    bool clock_throw_nomember(word_t& pc){
        m_err->wCode = 438;
        SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
        SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"No member");
        return (this->*m_onerr.back())();
    }

    bool clock_throw_(word_t& pc){
        if(m_err->m_hr == DISP_E_DIVBYZERO){
            m_err->wCode = 11;
            SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
            SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Div by 0");
        }
        return (this->*m_onerr.back())();
    }

    bool clock_exit(word_t& pc){
        return false;
    }

    bool clock_(word_t& pc){
       return (this->*(*(word_m*)pc.p))(pc);
    }

private:
    bool onerr_resumenext(){
        while(!( *((word_m*)m_pp->m_code[m_pc-1].p) == &CProcessor::word_colon )) ++m_pc;
        while(!( m_s.back().wReserved1 == VTX_GROUND )) m_s.pop_back();

        m_mode = &CProcessor::clock_;

        return true;
    }

    bool onerr_gotocatch(){
        while(!( m_s.back().wReserved1 == VTX_GROUND )) m_s.pop_back();

        m_mode = &CProcessor::clock_skiptocatch;

        m_onerr.back() = &CProcessor::onerr_goto0;

        return true;
    }

    bool onerr_goto0(){
        int n = m_onerr.size();
        while( 0 <= --n && m_onerr[n] == &CProcessor::onerr_goto0 );

        if(0 <= n){
            while(n < m_onerr.size()-1){
                while(!( m_s[m_s.size()-2].wReserved1 == VTX_PCFORRETURN && m_s[m_s.size()-1].wReserved1 == VTX_GROUND )) m_s.pop_back();

                m_s.pop_back();
                m_pc = m_s.back().llVal;
                m_s.pop_back();
                m_pp = (CProgram*)m_s.back().byref;
                m_s.pop_back();

                m_scope.pop_back();
                m_autodim.pop_back();
                m_with.pop_back();
                m_onerr.pop_back();
            }

            return (this->*m_onerr.back())();
        }

        return false;
    }

private:
    void do_left_invoke(){
        _variant_t& lv = m_s.back();
        if(lv.wReserved1==VTX_INST && *((inst_t*)lv.byref)==&CProcessor::op_invoke){
            (this->*(*((inst_t*)lv.byref)))(&lv);
        }
    }

    bool do_left_op(inst_t p){
        struct priority_t{ inst_t m; size_t p; };
        static priority_t priority[] = {
            { &CProcessor::op_plus1,    0 },
            { &CProcessor::op_minus1,   0 },
            { &CProcessor::op_incL,     1 },
            { &CProcessor::op_incR,     1 },
            { &CProcessor::op_decL,     1 },
            { &CProcessor::op_decR,     1 },
            { &CProcessor::op_hat,      2 },
            { &CProcessor::op_aster,    3 },
            { &CProcessor::op_slash,    3 },
            { &CProcessor::op_mod,      4 },
            { &CProcessor::op_plus,     5 },
            { &CProcessor::op_minus,    5 },
            { &CProcessor::op_amp,      6 },
            { &CProcessor::op_equal,   10 },
            { &CProcessor::op_neq,     10 },
            { &CProcessor::op_lthan,   10 },
            { &CProcessor::op_gthan,   10 },
            { &CProcessor::op_lethan,  10 },
            { &CProcessor::op_gethan,  10 },
            { &CProcessor::op_triequal,10 },
            { &CProcessor::op_trineq,  10 },
            { &CProcessor::op_is,      11 },
            { &CProcessor::op_not,     20 },
            { &CProcessor::op_and,     21 },
            { &CProcessor::op_or,      22 },
            { &CProcessor::op_xor,     23 },
            { nullptr,                 99 },
        };

        size_t lop = m_s.size()-1;
        while(!( m_s[lop].wReserved1==VTX_INST || m_s[lop].wReserved1==VTX_GROUND )) --lop;

        if(m_s[lop].wReserved1 == VTX_INST){
            inst_t l = *((inst_t*)m_s[lop].byref);
            int lp=0;
            while(!( priority[lp].m==l || !priority[lp].m )) ++lp;
            int pp=0;
            while(!( priority[pp].m==p || !priority[pp].m )) ++pp;

            if(priority[lp].p <= priority[pp].p){
                (this->*l)(&m_s[lop]);
                return do_left_op(p);
            }
        }

        return false;
    }

private:
    bool word_if(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_if];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_else(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_SKIPTOENDIF;
            v.llVal = 1;
            m_s.push_back( v );
        }
        m_mode = &CProcessor::clock_skiptoendif;

        return true;
    }

    bool word_elseif(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_SKIPTOENDIF;
            v.llVal = 1;
            m_s.push_back( v );
        }
        m_mode = &CProcessor::clock_skiptoendif;

        return true;
    }

    bool word_endif(word_t& pc){
        // none

        return true;
    }

    bool word_selectcase(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_selectcase];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_case(word_t& pc){
        if(m_s.back().wReserved1 == VTX_GROUND){
            {
                _variant_t v;
                v.wReserved1 = VTX_SKIPTOENDSELECT;
                v.llVal = 1;
                m_s.push_back(v);
            }
            m_mode = &CProcessor::clock_skiptoendselect;
        }else{
            {
                _variant_t v;
                v.wReserved1 = VTX_INST;
                v.byref = (void*)&s_insts[INST_stmt_case];
                m_s.push_back( v );
            }
        }

        return true;
    }

    bool word_caseelse(word_t& pc){
        if(m_s.back().wReserved1 == VTX_GROUND){
            {
                _variant_t v;
                v.wReserved1 = VTX_SKIPTOENDSELECT;
                v.llVal = 1;
                m_s.push_back(v);
            }
            m_mode = &CProcessor::clock_skiptoendselect;
        }else{
            {
                _variant_t v;
                v.wReserved1 = VTX_INST;
                v.byref = (void*)&s_insts[INST_stmt_caseelse];
                m_s.push_back( v );
            }
        }

        return true;
    }

    bool word_endselect(word_t& pc){
        if(m_s.back().wReserved1 == VTX_GROUND){
            m_s.pop_back();//VTX_GROUND
        }else{
            m_s.pop_back();//value
            m_s.pop_back();//stmt_selectcase
        }

        return true;
    }

    bool word_dowhile(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_PCFORLOOPWHILE;
            v.llVal = m_pc;
            m_s.push_back( v );
        }
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_do];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_dountil(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_PCFORLOOPUNTIL;
            v.llVal = m_pc;
            m_s.push_back( v );
        }
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_do];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_loop(word_t& pc){
        m_s.pop_back();     //VTX_GROUND
        m_pc = m_s[m_s.size()-2].llVal;

        return true;
    }

    bool word_foreach(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_foreach];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_in(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_in];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_for(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_for];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_to(word_t& pc){
        decltype(m_s)::reverse_iterator i = m_s.rbegin();
        while(! (i->wReserved1 == VTX_INST && *((inst_t*)i->byref) == &CProcessor::stmt_for) ){
            if(i->wReserved1 == VTX_INST){
                if(! (this->*(*(inst_t*)i->byref))(&*i) ) break;
                i = m_s.rbegin();
            }else{
                ++i;
            }
        }

        return true;
    }

    bool word_step(word_t& pc){
        decltype(m_s)::reverse_iterator i = m_s.rbegin();
        while(! (i->wReserved1 == VTX_INST && *((inst_t*)i->byref) == &CProcessor::stmt_for) ){
            if(i->wReserved1 == VTX_INST){
                if(! (this->*(*(inst_t*)i->byref))(&*i) ) break;
                i = m_s.rbegin();
            }else{
                ++i;
            }
        }

        return true;
    }

    bool word_next(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_next];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_goto(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_goto];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_return(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_return];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_exitfunction(word_t& pc){
        while(! (m_s[m_s.size()-1].wReserved1==VTX_GROUND && m_s[m_s.size()-2].wReserved1==VTX_PCFORRETURN) ) m_s.pop_back();

        m_pc = m_pp->m_code.size() - 1;

        return true;
    }

    bool word_exitdo(word_t& pc){
        while(! (m_s[m_s.size()-1].wReserved1==VTX_GROUND && m_s[m_s.size()-2].wReserved1==VTX_INST && *((inst_t*)m_s[m_s.size()-2].byref)==&CProcessor::stmt_do) ) m_s.pop_back();

        m_s.pop_back();//VTX_GROUND
        m_s.pop_back();//VTX_INST
        m_s.pop_back();//VTX_PCFORLOOPWHILE

        _variant_t v;
        v.wReserved1 = VTX_SKIPTOLOOP;
        v.llVal = 1;
        m_s.push_back(v);

        m_mode = &CProcessor::clock_skiptoloop;

        return true;
    }

    bool word_exitfor(word_t& pc){
        while(! (m_s[m_s.size()-1].wReserved1==VTX_GROUND && m_s[m_s.size()-2].wReserved1==VTX_PCFORNEXT) ) m_s.pop_back();

        if( m_s[m_s.size()-5].wReserved1==VTX_INST && *((inst_t*)m_s[m_s.size()-5].byref)==&CProcessor::stmt_foreach ){
            m_s.pop_back();//VTX_GROUND
            m_s.pop_back();//VTX_PCFORNEXT
            m_s.pop_back();//IEnumVARIANT
            m_s.pop_back();//VT_BYREF|VT_VARIANT
            m_s.pop_back();//stmt_foreach

            _variant_t v;
            v.wReserved1 = VTX_SKIPTONEXT;
            v.llVal = 1;
            m_s.push_back(v);

            m_mode = &CProcessor::clock_skiptonext;
        }else
        if( m_s[m_s.size()-7].wReserved1==VTX_INST && *((inst_t*)m_s[m_s.size()-7].byref)==&CProcessor::stmt_for ){
            m_s.pop_back();//VTX_GROUND
            m_s.pop_back();//VTX_PCFORNEXT
            m_s.pop_back();//op_step
            m_s.pop_back();//VT_I8
            m_s.pop_back();//VT_I8
            m_s.pop_back();//VT_BYREF|VT_VARIANT
            m_s.pop_back();//stmt_for

            _variant_t v;
            v.wReserved1 = VTX_SKIPTONEXT;
            v.llVal = 1;
            m_s.push_back(v);

            m_mode = &CProcessor::clock_skiptonext;
        }

        return true;
    }

    bool word_with(word_t& pc){
        do_left_op(&CProcessor::stmt_with);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_with];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_endwith(word_t& pc){
        m_with.back().pop_back();

        return true;
    }

    bool word_equal(word_t& pc){
        int ni = 0;{
            auto i = m_s.rbegin();
            while(!( i->wReserved1 == VTX_GROUND )){
                if(i->wReserved1 == VTX_INST) ++ni;
                ++i;
            }
        }

        auto i = m_s.rbegin();
        while(!( i->wReserved1 == VTX_GROUND || i->wReserved1 == VTX_INST )) ++i;

        if(i->wReserved1 == VTX_INST && *((inst_t*)i->byref) == &CProcessor::op_invoke && ni<2){
            i->byref = (void*)&s_insts[INST_op_invoke_put];//&CProcessor::op_invoke_put;
        }else
        if(i->wReserved1 == VTX_INST && *((inst_t*)i->byref) == &CProcessor::op_array && ni<2){
            i->byref = (void*)&s_insts[INST_op_array_put];//&CProcessor::op_array_put;
        }else
        {
            VARIANT* pv = &m_s.back();
            if(
                (i->wReserved1 == VTX_GROUND)                                                           ||
                ((pv-1)->wReserved1 == VTX_INST && *((inst_t*)(pv-1)->byref) == &CProcessor::stmt_for)  ||
            false){
                {
                    _variant_t v;
                    v.wReserved1 = VTX_INST;
                    v.byref = (void*)&s_insts[INST_op_copy];
                    m_s.push_back( v );
                }
            }else{
                do_left_invoke();

                do_left_op(&CProcessor::op_equal);
                {
                    _variant_t v;
                    v.wReserved1 = VTX_INST;
                    v.byref = (void*)&s_insts[INST_op_equal];
                    m_s.push_back( v );
                }
            }
        }
        
        return true;
    }

    bool word_coleq(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_copy];
            m_s.push_back( v );
        }
        
        return true;
    }

    bool word_plus(word_t& pc){
        int inst;
        if(m_s.back().wReserved1 == VTX_NONE){
            do_left_invoke();

            inst = INST_op_plus;
            do_left_op(&CProcessor::op_plus);
        }else{
            inst = INST_op_plus1;
            do_left_op(&CProcessor::op_plus1);
        }

        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[inst];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_minus(word_t& pc){
        int inst;
        if(
            m_s.back().wReserved1 == VTX_NONE                       &&
            *((word_m*)(&pc-1)->p) != &CProcessor::word_comma &&
            *((word_m*)(&pc-1)->p) != &CProcessor::word_to    &&
            *((word_m*)(&pc-1)->p) != &CProcessor::word_step  &&
        true){
            do_left_invoke();

            inst = INST_op_minus;
            do_left_op(&CProcessor::op_minus);
        }else{
            inst = INST_op_minus1;
            do_left_op(&CProcessor::op_minus1);
        }

        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[inst];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_hat(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_hat);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_hat];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_aster(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_aster);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_aster];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_slash(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_slash);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_slash];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_mod(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_mod);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_mod];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_amp(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_amp);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_amp];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_inc(word_t& pc){
        int inst;
        if(m_s.back().wReserved1 == VTX_NONE){
            inst = INST_op_incL;
            do_left_op(&CProcessor::op_incL);
        }else{
            inst = INST_op_incR;
            do_left_op(&CProcessor::op_incR);
        }

        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[inst];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_dec(word_t& pc){
        int inst;
        if(m_s.back().wReserved1 == VTX_NONE){
            inst = INST_op_decL;
            do_left_op(&CProcessor::op_decL);
        }else{
            inst = INST_op_decR;
            do_left_op(&CProcessor::op_decR);
        }

        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[inst];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_lthan(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_lthan);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_lthan];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_gthan(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_gthan);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_gthan];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_lethan(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_lethan);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_lethan];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_gethan(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_gethan);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_gethan];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_neq(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_neq);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_neq];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_eqeq(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_equal);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_equal];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_eqeqeq(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_triequal);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_triequal];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_exeqeq(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_trineq);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_trineq];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_is(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_is);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_is];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_not(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_not);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_not];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_and(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_and);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_and];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_or(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_or);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_or];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_xor(word_t& pc){
        do_left_invoke();

        do_left_op(&CProcessor::op_xor);
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_xor];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_dot(word_t& pc){
        if( *((word_m*)m_pp->m_code[m_pc-2].p) != &CProcessor::word_parenL ){
            do_left_invoke();
        }

        VARIANT* pv = &m_s.back();
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        if(pv->wReserved1 != VTX_NONE){
            m_s.push_back( m_with.back().back() );
        }

        m_mode = &CProcessor::clock_dispatch;
        return true;
    }

    bool word_quesdot(word_t& pc){
        if( *((word_m*)m_pp->m_code[m_pc-2].p) != &CProcessor::word_parenL ){
            do_left_invoke();
        }

        VARIANT* pv = &m_s.back();
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        if( (pv->vt == VT_DISPATCH && pv->pdispVal == nullptr) || pv->vt == VT_EMPTY ){
            VariantClear(pv);
            pv->vt = VT_DISPATCH;
            pv->pdispVal = &s_oNull;
        }

        m_mode = &CProcessor::clock_dispatch;
        return true;
    }

    bool word_quesques(word_t& pc){
        if( *((word_m*)m_pp->m_code[m_pc-2].p) != &CProcessor::word_parenL ){
            do_left_invoke();
        }

        VARIANT* pv = &m_s.back();
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        if( (pv->vt == VT_DISPATCH && pv->pdispVal == nullptr) || pv->vt == VT_EMPTY ){
            m_s.pop_back();
        }else{
            m_mode = &CProcessor::clock_shortcircuit;
        }

        return true;
    }

    bool word_comma(word_t& pc){
        int ni = 0;{
            auto i = m_s.rbegin();
            while(!( i->wReserved1 == VTX_GROUND )){
                if(i->wReserved1 == VTX_INST) ++ni;
                ++i;
            }
        }

        _variant_t& lv = m_s.back();
        if(lv.wReserved1 == VTX_INST && *((inst_t*)lv.byref) == &CProcessor::op_invoke && 1<ni){
            do_left_invoke();
        }else
        if( 
            *((word_m*)(&pc-1)->p) == &CProcessor::word_comma  ||
            *((word_m*)(&pc-1)->p) == &CProcessor::word_parenL ||
            (
                lv.wReserved1 == VTX_INST   &&
                (
                    *((inst_t*)lv.byref) == &CProcessor::op_invoke    ||
                    *((inst_t*)lv.byref) == &CProcessor::op_call      ||
                    *((inst_t*)lv.byref) == &CProcessor::op_copy      ||
                    *((inst_t*)lv.byref) == &CProcessor::stmt_return  ||
                false)                      &&
            true)                                                   ||
        false){
            m_s.push_back( _variant_t{ {{{VT_ERROR,0,0,0,{0}}}} } );
        }

        auto i = m_s.rbegin();
        while(!( i->wReserved1 == VTX_GROUND )){
            if(i->wReserved1 == VTX_INST){
                if(
                    *((inst_t*)i->byref) == &CProcessor::op_invoke            ||
                    *((inst_t*)i->byref) == &CProcessor::op_call              ||
                    *((inst_t*)i->byref) == &CProcessor::op_array             ||
                    *((inst_t*)i->byref) == &CProcessor::op_copy              ||
                    *((inst_t*)i->byref) == &CProcessor::stmt_case            ||
                    *((inst_t*)i->byref) == &CProcessor::stmt_redim           ||
                    *((inst_t*)i->byref) == &CProcessor::stmt_redimpreserve   ||
                    *((inst_t*)i->byref) == &CProcessor::stmt_return          ||
                false){
                    break;
                }else{
                    if( (this->*(*(inst_t*)i->byref))(&*i) ){
                        i = m_s.rbegin();
                    }else{
                        break;
                    }
                }
            }else{
                ++i;
            }
        }

        return true;
    }

    bool word_parenL(word_t& pc){
        //### think about do_left_invoke.

        VARIANT* pv = &m_s.back();
        VARIANT* pvp = pv-1;

        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        if(pv->vt == VT_DISPATCH){
            {
                _variant_t v;
                v.vt = VT_I4;
                v.lVal = DISPID_VALUE;
                m_s.push_back( v );
            }
            {
                _variant_t v;
                v.wReserved1 = VTX_INST;
                v.byref = (void*)&s_insts[INST_op_invoke];
                m_s.push_back( v );
            }
        }else
        if(
            pv->vt == (VT_ARRAY|VT_VARIANT) &&
            !(
                pvp->wReserved1==VTX_INST       &&
                (
                    *((inst_t*)pvp->byref) == &CProcessor::stmt_redim         ||
                    *((inst_t*)pvp->byref) == &CProcessor::stmt_redimpreserve ||
                false)                          &&
            true)                           &&
        true)
        {
            {
                _variant_t v;
                v.wReserved1 = VTX_INST;
                v.byref = (void*)&s_insts[INST_op_array];
                m_s.push_back( v );
            }
        }

        pv = &m_s.back();
        if(pv->wReserved1==VTX_INST){
            bool b = ( (&pc-1)->p == map_word(L"(") );//###
            if(!b && *((inst_t*)pv->byref) == &CProcessor::op_invoke){
                // none
            }else
            if(!b && *((inst_t*)pv->byref) == &CProcessor::op_call){
                // none
            }else
            if(!b && *((inst_t*)pv->byref) == &CProcessor::op_array){
                // none
            }else
            {
                {
                    _variant_t v;
                    v.wReserved1 = VTX_INST;
                    v.byref = (void*)&s_insts[INST_op_parenL];
                    m_s.push_back( v );
                }
            }
        }

        return true;
    }

    bool word_parenR(word_t& pc){
        auto i = m_s.rbegin();
        while(!( i->wReserved1 == VTX_GROUND )){
            if(i->wReserved1 == VTX_INST){
                if(
                    *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_equal &&
                    (
                        *((inst_t*)i->byref) == &CProcessor::op_invoke    ||
                        *((inst_t*)i->byref) == &CProcessor::op_array     ||
                    false)                                                      &&
                true)
                {
                    i = m_s.rend();
                    break;
                }else
                if( *((word_m*)m_pp->m_code[m_pc].p) == &CProcessor::word_coleq ){
                    if(*((inst_t*)i->byref) == &CProcessor::op_invoke){
                        i->byref = (void*)&s_insts[INST_op_invoke_put];
                        ++m_pc;
                        i = m_s.rend();
                        break;
                    }else
                    if(*((inst_t*)i->byref) == &CProcessor::op_array){
                        i->byref = (void*)&s_insts[INST_op_array_put];
                        ++m_pc;
                        i = m_s.rend();
                        break;
                    }else
                    {}
                }

                inst_t inst = *((inst_t*)i->byref);
                if(
                    !(this->*inst)(&*i)             ||
                    inst==&CProcessor::op_call      ||
                    inst==&CProcessor::op_invoke    ||
                    inst==&CProcessor::op_array     ||
                false){
                    i = m_s.rend();
                    break;
                }
                i = m_s.rbegin();
            }else{
                ++i;
            }
        }

        return (i == m_s.rend());
    }

    bool word_literal(word_t& pc){
        {
            m_s.push_back( pc.v );
        }

        return true;
    }

    bool word_jsonobj(word_t& pc){
        auto object = new JSONObject(pc.s.c_str());
        {
            _variant_t v;
            v.vt = VT_DISPATCH;
            v.pdispVal = object;
            m_s.push_back( v );
        }

        return true;
    }

    bool word_jsonary(word_t& pc){
        auto array = new JSONArray(pc.s.c_str());
        {
            _variant_t v;
            v.vt = VT_DISPATCH;
            v.pdispVal = array;
            m_s.push_back( v );
        }

        return true;
    }

    bool word_dim(word_t& pc){
        _variant_t& dim = m_scope.back()[pc.v.ullVal];
        if(dim.vt & VT_BYREF){
            m_s.push_back(_variant_t());
            *(VARIANT*)&m_s.back() = dim;
        }else
        {
            _variant_t v;
            v.vt = (VT_BYREF|VT_VARIANT);
            v.pvarVal = &dim;
            m_s.push_back( v );
        }

        return true;
    }

    bool word_pdim(word_t& pc){
        _variant_t& dim = m_scope.front()[pc.v.ullVal];
        if(dim.vt & VT_BYREF){
            m_s.push_back(_variant_t());
            *(VARIANT*)&m_s.back() = dim;
        }else
        {
            _variant_t v;
            v.vt = (VT_BYREF|VT_VARIANT);
            v.pvarVal = &dim;
            m_s.push_back( v );
        }

        return true;
    }

    bool word_gdim(word_t& pc){
        _variant_t& dim = m_parent->m_scope.front()[pc.v.ullVal];
        if(dim.vt & VT_BYREF){
            m_s.push_back(_variant_t());
            *(VARIANT*)&m_s.back() = dim;
        }else
        {
            _variant_t v;
            v.vt = (VT_BYREF|VT_VARIANT);
            v.pvarVal = &dim;
            m_s.push_back( v );
        }

        return true;
    }

    bool word_cdim(word_t& pc){
        _variant_t& dim = m_cdims[pc.v.ullVal];
        {
            _variant_t v;
            v.vt = (VT_BYREF|VT_VARIANT);
            v.pvarVal = &dim;
            m_s.push_back( v );
        }

        return true;
    }

    bool word_closure(word_t& pc){
        CProcessor* pros( new CProcessor(*this, (CProgram*)pc.v.byref) );
        pros->m_pc = pros->m_pp->m_code.size()-1;   // to don't-run by default

        size_t i = 0;
        while(i < pros->m_pp->m_clo_ids.size()){
            _variant_t& v = m_scope.back()[pros->m_pp->m_clo_ids[i]];
            if(v.vt == VT_DISPATCH){
                // for circulation reference.
                pros->m_cdims[i].vt = VT_DISPATCH;
                pros->m_cdims[i].pdispVal = new WeakDispatch(v.pdispVal);
            }else{
                pros->m_cdims[i] = v;
            }
            ++i;
        }

        _variant_t v;
        v.vt = VT_DISPATCH;
        v.pdispVal = pros;
        m_s.push_back(v);

        return true;
    }

    bool word_func(word_t& pc){
        _variant_t& lv = m_s.back();
        bool bNew = (lv.wReserved1 == VTX_INST && *((inst_t*)lv.byref) == &CProcessor::stmt_new);

        {
            _variant_t v;
            v.wReserved1 = VTX_PROGRAM;
            v.byref = pc.v.byref;
            m_s.push_back( v );
        }
        if(!bNew){
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_call];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_class(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_CLASS;
            v.byref = pc.v.byref;
            m_s.push_back( v );
        }

        return true;
    }

    bool word_exit(word_t& pc){
        std::vector<_variant_t> rets;
        {
            auto i = m_s.rbegin();
            while(!( i == m_s.rend() || i->wReserved1 == VTX_GROUND )){
                if(i->wReserved1 == VTX_RETURN){
                    while(!( &m_s.back() == &*i )){
                        rets.push_back( m_s.back() );
                        m_s.pop_back();
                    }
                    m_s.pop_back();
                    break;
                }
                ++i;
            }
        }

        if(m_s.back().wReserved1 == VTX_GROUND){
            m_s.pop_back();

            if(m_s.size()){
                if(m_s.back().wReserved1 != VTX_PCFORRETURN){
                    m_mode = &CProcessor::clock_throw_unexceptedend;
                    --m_pc;
                    return true;//#?# throw
                }

                m_pc = m_s.back().llVal;
                m_s.pop_back();
                m_pp = (CProgram*)m_s.back().byref;
                m_s.pop_back();

                if( rets.size() ){
                    m_s.pop_back();
                    while( rets.size() ){
                        VARIANT* pv = &rets.back();
                        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

                        m_s.push_back( _variant_t() );
                        *(VARIANT*)&m_s.back() = *pv;

                        pv->vt = VT_EMPTY;
                        rets.pop_back();
                    }
                }

                m_scope.pop_back();
                m_autodim.pop_back();
                m_with.pop_back();
                m_onerr.pop_back();

                return true;
            }
        }

        return false;
    }

    bool word_colon(word_t& pc){
        auto i = m_s.rbegin();
        while(!( i->wReserved1 == VTX_GROUND )){
            if(i->wReserved1 == VTX_INST){
                if(*((inst_t*)i->byref) == &CProcessor::op_call) --m_pc;

                if(! (this->*(*(inst_t*)i->byref))(&*i) ){
                    i = m_s.rend();
                    break;
                }
                i = m_s.rbegin();
            }else{
                ++i;
            }
        }

        if(i != m_s.rend()){// means (i->wReserved1 == VTX_GROUND)
            while(&*i < &m_s.back()) m_s.pop_back();
        }

        return true;
    }

    bool word_ext(word_t& pc){
        m_s.push_back(_variant_t());
        *(VARIANT*)&m_s.back() = pc.v;

        return true;
    }

    bool word_extdisp(word_t& pc){
        {
            m_s.push_back( pc.v );
        }
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_invoke];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_env(word_t& pc){
        {
            _variant_t v((IDispatch*)m_env);
            v.wReserved1 = VTX_NONE;
            m_s.push_back( v );
        }
        {
            m_s.push_back( pc.v );
        }
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_invoke];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_noop(word_t& pc){
        volatile int n=0;(void)n;// dummy now
        return true;
    }

    bool word_getref(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_getref];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_me(word_t& pc){
        {
            _variant_t v;
            v.vt = VT_DISPATCH;
            v.pdispVal = new CProcessor(*this);
            m_s.push_back( v );
        }

        return true;
    }

    bool word_err(word_t& pc){
        {
            _variant_t v;
            v.vt = VT_DISPATCH;
            (v.pdispVal = m_err)->AddRef();
            m_s.push_back( v );
        }

        return true;
    }

    bool word_new(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_new];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_redim(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_redim];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_redimpreserve(word_t& pc){
        {
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_stmt_redimpreserve];
            m_s.push_back( v );
        }

        return true;
    }

    bool word_onerrorresumenext(word_t& pc){
        m_onerr.back() = &CProcessor::onerr_resumenext;
        m_err->Clear();

        return true;
    }

    bool word_onerrorgotocatch(word_t& pc){
        m_onerr.back() = &CProcessor::onerr_gotocatch;
        m_err->Clear();

        return true;
    }

    bool word_onerrorcatch(word_t& pc){
        m_mode = &CProcessor::clock_skiptogoto0;

        return true;
    }

    bool word_onerrorgoto0(word_t& pc){
        m_onerr.back() = &CProcessor::onerr_goto0;
        m_err->Clear();

        return true;
    }

    bool word_DUMP(word_t& pc){
        dump();
        return true;
    }

    bool word_(word_t& pc){
        if(m_pp->m_option_explicit){
            m_mode = &CProcessor::clock_throw_unknownname;
            --m_pc;
            return true;//#?# throw
        }

        auto pb = &m_autodim.back();
        auto pf = &m_autodim.front();

        VARIANT* pv;
        auto found = pb->find(pc.s);
        if(found != pb->end()){
            pv = &(*pb)[pc.s];
        }else
        if(pb!=pf && (found=pf->find(pc.s))!=pf->end()){
            pv = &(*pf)[pc.s];
        }else{
            pv = &(*pb)[pc.s];
        }

        {
            _variant_t v;
            v.vt = (VT_BYREF|VT_VARIANT);
            v.pvarVal = pv;
            m_s.push_back( v );
        }

        return true;
    }

private:
    void bind(CProgram* pp){
        auto i = pp->m_code.begin();
        while(i != pp->m_code.end()){
            if(*((word_m*)i->p) == &CProcessor::word_){
                if(*((word_m*)(i-1)->p) == &CProcessor::word_dot){
                    if(*((word_m*)(i-2)->p) == &CProcessor::word_ext){
                        DISPID id;
                        LPOLESTR name = (LPOLESTR)i->s.c_str();
                        if((i-2)->v.pvarVal->pdispVal->GetIDsOfNames(IID_NULL, &name, 1, 0, &id) == S_OK){
                            {
                                i->p = map_word(L"@extdisp");

                                i->v.vt = VT_I4;
                                i->v.lVal = id;
                            }
                            {
                                (i-1)->p = map_word(L"@noop");
                            }
                        }
                    }
                }else{
                    auto pfuncs = pp->m_parent ? &pp->m_parent->m_funcs : &pp->m_funcs;
                    auto gfuncs = &pp->m_granpa->m_funcs;
                    decltype(CProgram::m_funcs)::const_iterator foundF;
                    auto pclasses = &pp->m_granpa->m_classes;
                    decltype(CProgram::m_classes)::const_iterator foundA;
                    auto ext = m_ext->find(i->s);
                    DISPID id;
                    LPOLESTR name = (LPOLESTR)i->s.c_str();
                    HRESULT hr;

                    if((foundF = pfuncs->find(i->s)) != pfuncs->end()){
                        {
                            i->p = map_word(L"@func");
                            i->v.vt = VT_EMPTY;
                            i->v.wReserved1 = VTX_PROGRAM;
                            i->v.byref = foundF->second.p;
                        }
                    }else
                    if(gfuncs && (foundF = gfuncs->find(i->s)) != gfuncs->end()){
                        {
                            i->p = map_word(L"@func");
                            i->v.vt = VT_EMPTY;
                            i->v.wReserved1 = VTX_PROGRAM;
                            i->v.byref = foundF->second.p;
                        }
                    }else
                    if((foundA = pclasses->find(i->s)) != pclasses->end()){
                        {
                            i->p = map_word(L"@class");
                            i->v.vt = VT_EMPTY;
                            i->v.wReserved1 = VTX_CLASS;
                            i->v.byref = foundA->second;
                        }
                    }else
                    if(ext != m_ext->end()){
                        i->p = map_word(L"@ext");

                        if(ext->second.vt & VT_BYREF){
                            *(VARIANT*)&i->v = ext->second;
                        }else
                        {
                            i->v.vt = (VT_BYREF|VT_VARIANT);
                            i->v.pvarVal = &ext->second;
                        }
                    }else
                    if(m_env && SUCCEEDED(hr = m_env->GetIDsOfNames(IID_NULL, &name, 1, 0, &id))){
                        if(hr == S_OK){
                            i->p = map_word(L"@env");

                            i->v.vt = VT_I4;
                            i->v.lVal = id;
                        }else{
                            i->p = map_word(L"@literal");
                            m_env->Invoke(id, IID_NULL, 0, 0, nullptr, &i->v, nullptr, nullptr);
                        }
                    }else{
                        // none: unknown word
                    }
                }
            }

            ++i;
        }

        {
            auto j = pp->m_funcs.begin();
            while(j != pp->m_funcs.end()){
                bind(j->second.p);
                ++j;
            }
        }

        {
            auto j = pp->m_plets.begin();
            while(j != pp->m_plets.end()){
                bind(j->second.p);
                ++j;
            }
        }

        {
            auto j = pp->m_closures.begin();
            while(j != pp->m_closures.end()){
                bind(*j);
                ++j;
            }
        }

        {
            auto j = pp->m_classes.begin();
            while(j != pp->m_classes.end()){
                bind(j->second);
                ++j;
            }
        }
    }

public:
    size_t  m_cc;
    _err_ptr_t   m_err;

    CProcessor(CProgram* pp, IDispatch* env, CExtension* ext, size_t size=0x100)
        : m_parent(nullptr)
        , m_ext(ext)
        , m_env(env)
        , m_p0(pp)
        , m_pp(pp)
        , m_pc(0)
        //---
        , m_cdims(pp->m_clo_defs)
        //, m_pgdims()
        //, m_pgadims()
        , m_mode(&CProcessor::clock_)
        //---
        , m_cc(0)
        , m_err(new Error(), false)
    {
        m_scope.reserve(size);
        m_scope.push_back( m_pp->m_dim_defs );
        m_pgdims = &m_scope.back();

        m_autodim.reserve(size);
        m_autodim.push_back( {} );
        m_pgadims = &m_autodim.back();

        m_with.reserve(size);
        m_with.push_back( { {{{{VT_EMPTY,0,0,0,{}}}}} } );

        m_onerr.reserve(size);
        m_onerr.push_back(&CProcessor::onerr_goto0);

        m_s.reserve(size*8);

        bind(m_pp);
// wprintf(L"$$$%s\n", __func__);
    }

    CProcessor(CProcessor& parent, CProgram* pp)
        : m_parent(&parent)
        , m_ext(parent.m_ext)
        , m_env(parent.m_env)
        , m_p0(pp)
        , m_pp(pp)
        , m_pc(0)
        //---
        , m_cdims(pp->m_clo_defs)
        //, m_pgdims()
        //, m_pgadims()
        , m_mode(&CProcessor::clock_)
        //---
        , m_cc(0)//### to ref
        , m_err(new Error(), false)
    {
        m_scope.reserve(parent.m_scope.capacity());
        m_scope.push_back( m_pp->m_dim_defs );
        m_pgdims = parent.m_pgdims;

        m_autodim.reserve(parent.m_autodim.capacity());
        m_autodim.push_back( {} );
        m_pgadims = parent.m_pgadims;

        m_with.reserve(parent.m_with.capacity());
        m_with.push_back( { {{{{VT_EMPTY,0,0,0,{}}}}} } );

        m_onerr.reserve(parent.m_onerr.capacity());
        m_onerr.push_back(&CProcessor::onerr_goto0);

        m_s.reserve(parent.m_s.capacity());

        DISPID did;
        LPOLESTR init = (LPOLESTR)L"class_initialize";
        if(this->GetIDsOfNames(IID_NULL, &init, 1, 0, &did) == S_OK){
            DISPPARAMS param = { nullptr, nullptr, 0, 0 };
            _variant_t res;
            _com_util::CheckError( this->Invoke(did, IID_NULL, 0, DISPATCH_METHOD, &param, &res, nullptr, nullptr) );
        }
// wprintf(L"$$$%s-child\n", __func__);
    }

    CProcessor(CProcessor& source)
        : m_parent(source.m_parent)
        , m_ext(source.m_ext)
        , m_env(source.m_env)
        , m_p0(source.m_p0)
        , m_pp(source.m_p0)
        , m_pc(source.m_p0->m_code.size()-1)
        //---
        , m_cdims(source.m_p0->m_clo_defs)
        //, m_pgdims()
        //, m_pgadims()
        , m_mode(&CProcessor::clock_)
        //---
        , m_cc(0)//### to ref
        , m_err(new Error(), false)
    {
        m_scope.reserve(source.m_scope.capacity());
        m_scope.push_back( {} );
        m_pgdims = source.m_pgdims;

        m_autodim.reserve(source.m_autodim.capacity());
        m_autodim.push_back( {} );
        m_pgdims = source.m_pgdims;

        m_with.reserve(source.m_with.capacity());
        m_with.push_back( { {{{{VT_EMPTY,0,0,0,{}}}}} } );

        m_onerr.reserve(source.m_onerr.capacity());
        m_onerr.push_back(&CProcessor::onerr_goto0);

        m_s.reserve(source.m_s.capacity());
// wprintf(L"$$$%s-copy\n", __func__);
    }

    virtual ~CProcessor(){
        DISPID did;
        LPOLESTR init = (LPOLESTR)L"class_terminate";
        if(this->GetIDsOfNames(IID_NULL, &init, 1, 0, &did) == S_OK){
            DISPPARAMS param = { nullptr, nullptr, 0, 0 };
            _variant_t res;
            _com_util::CheckError( this->Invoke(did, IID_NULL, 0, DISPATCH_METHOD, &param, &res, nullptr, nullptr) );
        }
// wprintf(L"$$$%s%ls\n", __func__, m_parent?L"-child":L"");
    }

    void dump(){
        {
            auto i = m_scope.begin();
            while(i != m_scope.end()){
                auto j = i->begin();
                while(j != i->end()){
                    wprintf(L"***%s-scope (vt:0x%x, [0x%llx p:%p])\n", __func__, j->vt, j->llVal, &*j);
                    ++j;
                }
                wprintf(L"***%s-scope ---\n", __func__);
                ++i;
            }
        }
        {
            auto i = m_autodim.begin();
            while(i != m_autodim.end()){
                auto j = i->begin();
                while(j != i->end()){
                    wprintf(L"***%s-atdim %ls(vt:0x%x, [0x%llx])\n", __func__, j->first.c_str(), j->second.vt, j->second.llVal);
                    ++j;
                }
                wprintf(L"***%s-atdim ---\n", __func__);
                ++i;
            }
        }
        {
            auto i = m_with.begin();
            while(i != m_with.end()){
                auto j = i->begin();
                while(j != i->end()){
                    wprintf(L"***%s-with vt:0x%x, wR:0x%x, [0x%llx]\n", __func__, j->vt, j->wReserved1, j->llVal);
                    ++j;
                }
                wprintf(L"***%s-with ---\n", __func__);
                ++i;
            }
        }
        {
            auto i = m_s.begin();
            while(i != m_s.end()){
                wprintf(L"***%s-stack vt:0x%x, wR:0x%x, [0x%llx p:%p]\n", __func__, i->vt, i->wReserved1, i->llVal, &(*i));
                ++i;
            }
        }
        wprintf(L"***%s       //\n", __func__);
    }

    HRESULT operator()(CProgram& prog, VARIANT args[], int argn){
        m_scope.push_back( prog.m_dim_defs );
        std::vector<_variant_t>& newscope = m_scope.back();
        m_autodim.push_back( {} );
        m_with.push_back( { {{{{VT_EMPTY,0,0,0,{}}}}} } );
        m_onerr.push_back(&CProcessor::onerr_goto0);

        int i = 0;
        while(i<argn){
            if(i < prog.m_params.size()){
                newscope[i+1].vt = (VT_BYREF|VT_VARIANT);
                newscope[i+1].pvarVal = args+(argn-1-i);
            }else{
                m_mode = &CProcessor::clock_throw_toomanyarg;
                --m_pc;
                return false;
            }
            ++i;
        }
        
        //#?# no need this stack push?
        if(m_s.size()){
            _variant_t v;
            v.wReserved1 = VTX_INST;
            v.byref = (void*)&s_insts[INST_op_exit];
            m_s.push_back( v );
        }
        {
            m_s.push_back( _variant_t() );

            _variant_t& ret = newscope[0];
            ret.vt = (VT_BYREF|VT_VARIANT);
            ret.pvarVal = &m_s.back();
        }
        {
            _variant_t v;
            v.wReserved1 = VTX_POFORRETURN;
            v.byref = m_pp;
            m_s.push_back(v);
        }
        {
            _variant_t v;
            v.wReserved1 = VTX_PCFORRETURN;
            v.llVal = m_pc;
            m_s.push_back(v);
        }

        m_pp = &prog;
        m_pc = 0;

        return operator()();
    }

    HRESULT operator()(){
        HRESULT hr = S_OK;

        try{
            m_s.push_back(_variant_t());
            m_s.back().wReserved1 = VTX_GROUND;

            if(m_cc){
                while( (this->*m_mode)(m_pp->m_code[m_pc++]) ){
                    if(!--m_cc) m_mode = &CProcessor::clock_throw_timeout;
                }
            }else{
                while( (this->*m_mode)(m_pp->m_code[m_pc++]) );
            }
        }catch(_com_error& e){
            hr = e.Error();
        }catch(...){
            // none
        }

        --m_pc;

        word_t& last = m_pp->m_code[m_pc];



        if(hr != S_OK){
            m_err->scode = 0x003;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Object error");
            }
        }else
        if(m_mode == &CProcessor::clock_throw_invalidcopy){
            m_err->scode = 0x201;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Invalid copy");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_throw_toomanyarg){
            m_err->scode = 0x202;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Too many args");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_throw_json){
            m_err->scode = 0x203;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Invalid JSON");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_throw_funcnotfound){
            m_err->scode = 0x204;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Function not found");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_throw_timeout){
            m_err->scode = 0x205;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Timeout");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_throw_cantgotothere){
            m_err->scode = 0x206;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Can't go to there");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_throw_invalidindex){
            m_err->scode = 0x207;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Invalid array index");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_throw_unexceptedend){
            m_err->scode = 0x208;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Unexcepted end");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_throw_object){
            m_err->scode = 0x209;
            m_err->dwHelpContext = last.l+1;
            hr = m_err->m_hr;
        }else
        if(m_mode == &CProcessor::clock_skiptoelse){
            m_err->scode = 0x101;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Not found 'else'");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_skiptoendif){
            m_err->scode = 0x102;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Not found 'end if'");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_skiptocase){
            m_err->scode = 0x103;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Not found 'case'");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_skiptoendselect){
            m_err->scode = 0x104;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Not found 'end select'");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_skiptoloop){
            m_err->scode = 0x105;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Not found 'loop'");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_skiptonext){
            m_err->scode = 0x106;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Not found 'next'");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_skiptocatch){
            m_err->scode = 0x107;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Not found 'on error catch'");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_dispatch){
            m_err->scode = 0x108;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Unknown member");
            }
            hr = E_FAIL;
        }else
        if(m_mode == &CProcessor::clock_exit){
            m_mode = &CProcessor::clock_;   // no error & return to previous operator()
        }else
        if(*((word_m*)last.p) == &CProcessor::word_exit){
            // no error
        }else
        if(m_scope.size() == m_scope.capacity()){
            m_err->scode = 0x004;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Memory over");
            }
            hr = E_FAIL;
        }else
        if(m_s.size() == m_s.capacity()){
            m_err->scode = 0x005;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Stack over flow");
            }
            hr = E_FAIL;
        }else
        {
            m_err->scode = 0x001;
            m_err->dwHelpContext = last.l+1;
            if(!m_err->wCode){
                SysFreeString(m_err->bstrSource); m_err->bstrSource = SysAllocString(NAME);
                SysFreeString(m_err->bstrDescription); m_err->bstrDescription = SysAllocString(L"Runtime error");
            }
            hr = E_FAIL;
        }

        return hr;
    }



//IDispatch
private:
    ULONG       m_refc   = 1;

    _variant_t* m_pvDispatching;
    CProgram*   m_pfDispatching;
    CProgram*   m_ppDispatching;

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        m_pvDispatching = nullptr;
        m_pfDispatching = nullptr;
        m_ppDispatching = nullptr;

        decltype(m_p0->m_dim_names)::const_iterator foundD;
        if((foundD = m_p0->m_dim_names.find(*rgszNames)) != m_p0->m_dim_names.end() && foundD->second.s == DSC_PUBLIC){
            if(m_parent){
                m_pvDispatching = &m_scope.front()[foundD->second.i];
            }else{
                m_pvDispatching = &(*m_pgdims)[foundD->second.i];
            }
            *rgDispId = 1;
        }else{
            decltype(m_p0->m_funcs)::const_iterator foundF;
            if((foundF = m_p0->m_funcs.find(*rgszNames)) != m_p0->m_funcs.end() && foundF->second.s == DSC_PUBLIC){
                m_pfDispatching = foundF->second.p;
            }

            decltype(m_p0->m_plets)::const_iterator foundL;
            if((foundL = m_p0->m_plets.find(*rgszNames)) != m_p0->m_plets.end() && foundL->second.s == DSC_PUBLIC){
                m_ppDispatching = foundL->second.p;
            }

            if(m_pfDispatching || m_ppDispatching){
                *rgDispId = 2;
            }
        }

        return (m_pvDispatching || m_pfDispatching || m_ppDispatching) ? S_OK : E_FAIL;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        HRESULT hr = DISP_E_MEMBERNOTFOUND;

        if(dispIdMember == 0){
            //### if(m_pc == 0) (*this)(); else
            if(m_p0->m_default){
                if( SUCCEEDED(hr = (*this)(*m_p0->m_default, pDispParams->rgvarg, pDispParams->cArgs)) ){
                    *pVarResult = m_s.back().Detach();
                    m_s.pop_back();
                }
            }else{
                //### think about 'me()'
                if( SUCCEEDED(hr = (*this)(*m_p0, pDispParams->rgvarg, pDispParams->cArgs)) ){
                    *pVarResult = m_s.back().Detach();
                    m_s.pop_back();
                }
            }
        }else
        if(dispIdMember == 1){
            if(wFlags == DISPATCH_PROPERTYPUT){
                hr = VariantCopy(m_pvDispatching, pDispParams->rgvarg);
            }else{
                hr = VariantCopy(pVarResult, m_pvDispatching);
            }
        }else
        if(dispIdMember == 2){
            if(wFlags == DISPATCH_PROPERTYPUT){
                if(m_ppDispatching){
                    if(SUCCEEDED( hr = (*this)(*m_ppDispatching, pDispParams->rgvarg, pDispParams->cArgs) )){
                        *pVarResult = m_s.back().Detach();
                        m_s.pop_back();
                    }
                }
            }else{
                if(m_pfDispatching){
                    if(SUCCEEDED( hr = (*this)(*m_pfDispatching, pDispParams->rgvarg, pDispParams->cArgs) )){
                        *pVarResult = m_s.back().Detach();
                        m_s.pop_back();
                    }
                }
            }
        }

        if(m_s.size()){
            hr = E_INVALIDARG;
        }

        return hr;
    }
};


