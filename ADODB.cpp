#include "npole.h"



class Recordset : public IDispatch{
private:
    ULONG       m_refc   = 1;

public:
    Recordset(){/*wprintf(L"%s\n", __func__);*/}
    virtual ~Recordset(){/*wprintf(L"%s\n", __func__);*/}

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(_wcsicmp(*rgszNames, L"eof") == 0){
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
        if(dispIdMember == 1){
            //### eof
            pVarResult->vt = VT_BOOL;
            pVarResult->boolVal = VARIANT_TRUE;
        }else
        {
            return DISP_E_MEMBERNOTFOUND;
        }

        return S_OK;
    }
};



class Connection : public IDispatch{
private:
    ULONG       m_refc   = 1;

public:
    Connection(){/*wprintf(L"%s\n", __func__);*/}
    virtual ~Connection(){/*wprintf(L"%s\n", __func__);*/}

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(_wcsicmp(*rgszNames, L"open") == 0){
            *rgDispId = 1;
        }else
        if(_wcsicmp(*rgszNames, L"execute") == 0){
            *rgDispId = 2;
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
        if(dispIdMember == 1){
            //### open
        }else
        if(dispIdMember == 2){
            //### execute
            pVarResult->vt = VT_DISPATCH;
            pVarResult->pdispVal = new Recordset();
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
        *ppvObject = new Connection();
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



