#include "npole.h"



class CFactory : public IClassFactory{
private:
    volatile LONG       m_refc   = 1;

    CLSID       m_clsid = {0x3F4DACA4,0x160D,0x11D2,{0xA8,0xE9,0x00,0x10,0x4B,0x36,0x5C,0x9F}}; // VBScript.RegExp

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef() { return InterlockedIncrement(&m_refc); }
    ULONG STDMETHODCALLTYPE Release(){ LONG refc = InterlockedDecrement(&m_refc); if(!refc){ delete this; return 0; } return refc; }

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject){
        return CoCreateInstance(m_clsid, nullptr, CLSCTX_INPROC_SERVER, IID_IDispatch, ppvObject);
    }
    
    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock){
        if(fLock){
            InterlockedIncrement(&m_refc);
        }else{
            InterlockedDecrement(&m_refc);
        }
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



