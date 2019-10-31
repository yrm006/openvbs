#include "npole.h"



class CFactory : public IClassFactory{
private:
    ULONG       m_refc   = 1;

    CLSID       m_clsid = {0x3F4DACA4,0x160D,0x11D2,{0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F}}; // VBScript.RegExp

public:
    HRESULT QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG AddRef(){ return ++m_refc; }
    ULONG Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject){
        return CoCreateInstance(m_clsid, nullptr, CLSCTX_INPROC_SERVER, IID_IDispatch, ppvObject);
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



