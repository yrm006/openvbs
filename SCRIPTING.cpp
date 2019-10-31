#include "npole.h"
#include <string>
#include <map>



class DictionaryEnum : public IEnumVARIANT{
private:
    ULONG       m_refc   = 1;

    const std::map<_variant_t, _variant_t>& m_r;
    std::map<_variant_t, _variant_t>::const_iterator m_i;

public:
    DictionaryEnum(const std::map<_variant_t, _variant_t>& r)
        : m_r(r)
    {
        m_i = m_r.begin();
    }
    virtual ~DictionaryEnum(){/*wprintf(L"%s\n", __func__);*/}

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
            VariantCopy(rgVar+i, &m_i->first);
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



class Dictionary : public IDispatch{
private:
    ULONG       m_refc   = 1;

public:
    std::map<_variant_t, _variant_t> m_map;

public:
    Dictionary(){/*wprintf(L"%s\n", __func__);*/}
    virtual ~Dictionary(){/*wprintf(L"%s\n", __func__);*/}

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(_wcsicmp(*rgszNames, L"Item") == 0){
            *rgDispId = 0;
        }else
        if(_wcsicmp(*rgszNames, L"Add") == 0){
            *rgDispId = 1;
        }else
        if(_wcsicmp(*rgszNames, L"Count") == 0){
            *rgDispId = 2;
        }else
        if(_wcsicmp(*rgszNames, L"CompareMode") == 0){
            *rgDispId = 3;
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
                m_map[pDispParams->rgvarg[1]] = pDispParams->rgvarg[0];
            }else{
                VariantCopy(pVarResult, &m_map[pDispParams->rgvarg[0]]);
            }
        }else
        if(dispIdMember == 1){
            if(pDispParams->cArgs == 2){
                m_map[pDispParams->rgvarg[1]] = pDispParams->rgvarg[0];
            }else{
                return DISP_E_BADPARAMCOUNT;
            }
        }else
        if(dispIdMember == 2){
            pVarResult->vt = VT_I8;
            pVarResult->llVal = m_map.size();
        }else
        if(dispIdMember == 3){
            //###TODO
        }else
        if(dispIdMember == DISPID_NEWENUM){
            pVarResult->vt = VT_UNKNOWN;
            pVarResult->punkVal = new DictionaryEnum(m_map);
        }else
        {
            return DISP_E_MEMBERNOTFOUND;
        }

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
        *ppvObject = new Dictionary;
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



