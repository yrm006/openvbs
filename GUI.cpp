#include "npole.h"



class GUI : public IDispatch{
private:
    ULONG       m_refc   = 1;

public:
    // GUI(){wprintf(L"%s\n", __FUNCTIONW__);}
    // ~GUI(){wprintf(L"%s\n", __FUNCTIONW__);}

    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(_wcsicmp(*rgszNames, L"MsgBox") == 0){
            *rgDispId = 1;
        }else
        {
            return DISP_E_MEMBERNOTFOUND;
        }

        return S_OK;
    }
    HRESULT Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
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
    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject){
        *ppvObject = new GUI;
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



