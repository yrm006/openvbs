#include "npole.h"



class GUI : public IDispatch{
private:
    ULONG       m_refc   = 1;

public:
    // GUI(){wprintf(L"%s\n", __FUNCTIONW__);}
    // ~GUI(){wprintf(L"%s\n", __FUNCTIONW__);}

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(_wcsicmp(*rgszNames, L"MsgBox") == 0){
            *rgDispId = 1;
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
            pVarResult->vt = VT_I8;
            pVarResult->llVal = MessageBoxW(nullptr, pDispParams->rgvarg[0].bstrVal, L"", MB_OK);
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
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject){
        *ppvObject = new GUI;
        return S_OK;
    }
    
    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock){
        m_refc += fLock ? 1 : -1;
        return S_OK;
    }
} g_oFactory;



extern "C"
HRESULT CALLBACK DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv){
    g_oFactory.AddRef();
    *ppv = (IClassFactory*)&g_oFactory;

    return S_OK;
}

extern "C"
HRESULT CALLBACK DllCanUnloadNow(){
    return S_FALSE;
}



