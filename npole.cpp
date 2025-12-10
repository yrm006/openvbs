#include "npole.h"



#ifndef _WIN32
#include <time.h>
#include <wchar.h>
#include <ctype.h>
#include <stdlib.h>
#include <math.h>
#include <stdio.h>
#include <sys/time.h>

#define D18991230 693899.000000

const CLSID CLSID_NULL      = {0x00000000,0x0000,0x0000,{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00}};
const IID IID_NULL          = {0x00000000,0x0000,0x0000,{0x00,0x00,0x00,0x00,0x00,0x00,0x00,0x00}};
const IID IID_IClassFactory = {0x00000001,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
const IID IID_IDispatch     = {0x00020400,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
const IID IID_IEnumVARIANT  = {0x00020404,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};



// system
HMODULE LoadLibraryW(LPOLESTR lpszLibName){
    char name[256];
    if(wchar_utf8(name, 256, lpszLibName) < 256){
        int i = 0;
        while(!( name[i]=='\0' || name[i]=='.' )) ++i;

        char path[256];
        if(name[i]){
            snprintf(path, 256, "%s%s",  getenv("REGISTRY"), name);
        }else{
            snprintf(path, 256, "%s%s.so",  getenv("REGISTRY"), name);
        }

        return dlopen(path, RTLD_LAZY);
    }

    return nullptr;
}

HMODULE GetModuleHandleW(LPCWSTR lpModuleName){
    if(lpModuleName){
        //###TODO
        return nullptr;
    }else{
        return dlopen(nullptr, RTLD_LAZY);
    }
}

PROC GetProcAddress(HMODULE hModule, LPCSTR lpProcName){
    return (PROC)dlsym(hModule, lpProcName);
}

void GetLocalTime(LPSYSTEMTIME lpSystemTime){
    timeval tv;
    gettimeofday(&tv, nullptr);
    tm t;
    localtime_r(&tv.tv_sec, &t);
    lpSystemTime->wYear  = t.tm_year+1900;
    lpSystemTime->wMonth = t.tm_mon+1;
    lpSystemTime->wDay   = t.tm_mday;
    lpSystemTime->wDayOfWeek = t.tm_wday;
    lpSystemTime->wHour  = t.tm_hour;
    lpSystemTime->wMinute = t.tm_min;
    lpSystemTime->wSecond = t.tm_sec;
    lpSystemTime->wMilliseconds = tv.tv_usec/1000;
}

// ole
HRESULT CLSIDFromProgID(LPCOLESTR lpszProgID, LPCLSID lpclsid){
    char name[256];
    wchar_utf8(name, 256, lpszProgID);
    for(size_t i=0; name[i]; ++i) name[i] = toupper(name[i]);

    char path[256];
    snprintf(path, 256, "%s%s.clsid",  getenv("REGISTRY"), name);

    FILE* pf = fopen(path, "r");
    if(pf){
        // long, short, short, char[8] ex:{EE09B103-97E0-11CF-978F-00A02463E06F}
        int scan = fscanf(pf, "{%8lx-%4hx-%4hx-%2hhx%2hhx-%2hhx%2hhx%2hhx%2hhx%2hhx%2hhx}",
            &lpclsid->Data1, &lpclsid->Data2, &lpclsid->Data3,
            &lpclsid->Data4[0], &lpclsid->Data4[1], &lpclsid->Data4[2], &lpclsid->Data4[3],
            &lpclsid->Data4[4], &lpclsid->Data4[5], &lpclsid->Data4[6], &lpclsid->Data4[7]);
        fclose(pf);
        return (scan == 11) ? S_OK : E_FAIL;
    }

    return CO_E_CLASSSTRING;
}

HRESULT CLSIDFromString(LPCOLESTR lpsz, LPCLSID pclsid){
wprintf(L"###%s: Implement here '%s' line %d.\n", __func__, __FILE__, __LINE__);
    return E_NOTIMPL;
}

HRESULT CoCreateInstance(REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext, REFIID riid, LPVOID *ppv){
    wchar_t path[38+1];
    swprintf(path, 38+1, L"{%08X-%04hX-%04hX-%02hhX%02hhX-%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX}",
        rclsid.Data1, rclsid.Data2, rclsid.Data3,
        rclsid.Data4[0], rclsid.Data4[1], rclsid.Data4[2], rclsid.Data4[3], rclsid.Data4[4], rclsid.Data4[5], rclsid.Data4[6], rclsid.Data4[7]);

    HMODULE hm = LoadLibraryW(path);
    if(hm){
        LPFNGETCLASSOBJECT pfn =  (LPFNGETCLASSOBJECT)GetProcAddress(hm, "DllGetClassObject");
        if(pfn){
            IClassFactory* cf = (IClassFactory*)path;
            if( SUCCEEDED(pfn(CLSID_NULL, IID_IClassFactory, (void**)&cf)) ){
                HRESULT hr = cf->CreateInstance(pUnkOuter, riid, ppv);
                cf->Release();
                return hr;
            }
        }

        //### manage hm!
    }

    return E_FAIL;
}

HRESULT CoGetObject(LPCWSTR pszName, BIND_OPTS *pBindOptions, REFIID riid, void **ppv){
    if(_wcsnicmp(pszName, L"new:", 4) == 0){
        pszName += 4;

        CLSID clsid;
        HRESULT hr;
        if(
            SUCCEEDED( hr = (pszName[0] == L'{') ? CLSIDFromString(pszName, &clsid) : CLSIDFromProgID(pszName, &clsid) ) &&
            SUCCEEDED( hr = CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER, riid, ppv) )                          &&
        true){
            return hr;
        }
    }

    return E_FAIL;
}

HRESULT CoInitializeEx(LPVOID pvReserved, DWORD dwCoInit){
    // none
    return S_OK;
}

void CoUninitialize(){
    // none
}

// oleaut
HRESULT VarNot(LPVARIANT pvarIn, LPVARIANT pvarResult){
    if(pvarIn->vt == (VT_BYREF|VT_VARIANT)) pvarIn = pvarIn->pvarVal;
    
    if(pvarIn->vt == VT_BOOL){
        pvarResult->vt = VT_BOOL;
        pvarResult->boolVal = ~pvarIn->boolVal;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d)\n", __func__, __FILE__, __LINE__, pvarIn->vt);
        return E_NOTIMPL;
    }

    return S_OK;
}

HRESULT VarAnd(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult){
    if(pvarLeft->vt == (VT_BYREF|VT_VARIANT)) pvarLeft = pvarLeft->pvarVal;
    if(pvarRight->vt == (VT_BYREF|VT_VARIANT)) pvarRight = pvarRight->pvarVal;
    
    if(pvarLeft->vt == VT_BOOL && pvarRight->vt == VT_BOOL){
        pvarResult->vt = VT_BOOL;
        pvarResult->boolVal = pvarLeft->boolVal & pvarRight->boolVal;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarLeft->vt, pvarRight->vt);
        return E_NOTIMPL;
    }

    return S_OK;
}

HRESULT VarOr(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult){
    if(pvarLeft->vt == (VT_BYREF|VT_VARIANT)) pvarLeft = pvarLeft->pvarVal;
    if(pvarRight->vt == (VT_BYREF|VT_VARIANT)) pvarRight = pvarRight->pvarVal;
    
    if(pvarLeft->vt == VT_BOOL && pvarRight->vt == VT_BOOL){
        pvarResult->vt = VT_BOOL;
        pvarResult->boolVal = pvarLeft->boolVal | pvarRight->boolVal;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarLeft->vt, pvarRight->vt);
        return E_NOTIMPL;
    }

    return S_OK;
}

HRESULT VarXor(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult){
    if(pvarLeft->vt == (VT_BYREF|VT_VARIANT)) pvarLeft = pvarLeft->pvarVal;
    if(pvarRight->vt == (VT_BYREF|VT_VARIANT)) pvarRight = pvarRight->pvarVal;
    
    if(pvarLeft->vt == VT_BOOL && pvarRight->vt == VT_BOOL){
        pvarResult->vt = VT_BOOL;
        pvarResult->boolVal = pvarLeft->boolVal ^ pvarRight->boolVal;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarLeft->vt, pvarRight->vt);
        return E_NOTIMPL;
    }

    return S_OK;
}

HRESULT VarCmp(LPVARIANT pvarLeft, LPVARIANT pvarRight, LCID lcid, ULONG dwFlags){
    if(pvarLeft->vt == (VT_BYREF|VT_VARIANT)) pvarLeft = pvarLeft->pvarVal;
    if(pvarRight->vt == (VT_BYREF|VT_VARIANT)) pvarRight = pvarRight->pvarVal;
    
    _variant_t v;
    if(pvarLeft->vt == VT_BSTR && pvarRight->vt != VT_BSTR){
        VariantChangeType(&v, pvarRight, 0, VT_BSTR);
        pvarRight = &v;
    }else
    if(pvarRight->vt == VT_BSTR && pvarLeft->vt != VT_BSTR){
        VariantChangeType(&v, pvarLeft, 0, VT_BSTR);
        pvarLeft = &v;
    }else
    if(pvarLeft->vt == VT_R8 && pvarRight->vt != VT_R8){
        VariantChangeType(&v, pvarRight, 0, VT_R8);
        pvarRight = &v;
    }else
    if(pvarRight->vt == VT_R8 && pvarLeft->vt != VT_R8){
        VariantChangeType(&v, pvarLeft, 0, VT_R8);
        pvarLeft = &v;
    }else
    if(pvarLeft->vt == VT_I8 && pvarRight->vt != VT_I8){
        VariantChangeType(&v, pvarRight, 0, VT_I8);
        pvarRight = &v;
    }else
    if(pvarRight->vt == VT_I8 && pvarLeft->vt != VT_I8){
        VariantChangeType(&v, pvarLeft, 0, VT_I8);
        pvarLeft = &v;
    }else
    {}

    if(pvarLeft->vt == VT_I8 && pvarRight->vt == VT_I8){
        return  (pvarLeft->llVal < pvarRight->llVal) ? VARCMP_LT
              : (pvarLeft->llVal > pvarRight->llVal) ? VARCMP_GT
              : VARCMP_EQ;
    }else
    if(pvarLeft->vt == VT_I4 && pvarRight->vt == VT_I4){
        return  (pvarLeft->lVal < pvarRight->lVal) ? VARCMP_LT
              : (pvarLeft->lVal > pvarRight->lVal) ? VARCMP_GT
              : VARCMP_EQ;
    }else
    if(pvarLeft->vt == VT_R8 && pvarRight->vt == VT_R8){
        return  (pvarLeft->dblVal < pvarRight->dblVal) ? VARCMP_LT
              : (pvarLeft->dblVal > pvarRight->dblVal) ? VARCMP_GT
              : VARCMP_EQ;
    }else
    if(pvarLeft->vt == VT_BSTR && pvarRight->vt == VT_BSTR){
        int cmp = wcscmp(pvarLeft->bstrVal, pvarRight->bstrVal);
        return  (cmp < 0) ? VARCMP_LT
              : (0 < cmp) ? VARCMP_GT
              : VARCMP_EQ;
    }else
    if(pvarLeft->vt == VT_BOOL && pvarRight->vt == VT_BOOL){
        return  (pvarLeft->boolVal < pvarRight->boolVal) ? VARCMP_LT
              : (pvarLeft->boolVal > pvarRight->boolVal) ? VARCMP_GT
              : VARCMP_EQ;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarLeft->vt, pvarRight->vt);
        return E_NOTIMPL;
    }
}

HRESULT VarRound(LPVARIANT pvarIn, int cDecimals, LPVARIANT pvarResult){
    if(pvarIn->vt == VT_I8){
        pvarResult->vt = VT_I8;
        pvarResult->llVal = pvarIn->llVal;
    }else
    if(pvarIn->vt == VT_R8){
        double dbl = pvarIn->dblVal;

        int i=0;
        while(i++ < cDecimals) dbl *= 10;

        dbl -= 0.5;
        dbl = (uint64_t)dbl - (dbl < 0) + 1;

        i=0;
        while(i++ < cDecimals) dbl /= 10;

        pvarResult->vt = VT_R8;
        pvarResult->dblVal = dbl;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d)\n", __func__, __FILE__, __LINE__, pvarIn->vt);
        return E_INVALIDARG;
    }

    return S_OK;
}

HRESULT VarInt(LPVARIANT pvarIn, LPVARIANT pvarResult){
    // Do NOT change VT
    if(pvarIn->vt == VT_R8){
        pvarResult->vt = VT_R8;
        pvarResult->dblVal = (int64_t)pvarIn->dblVal - (pvarIn->dblVal < 0);
    }else
    if(pvarIn->vt == VT_I8){
        pvarResult->vt = VT_I8;
        pvarResult->llVal = pvarIn->llVal;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d)\n", __func__, __FILE__, __LINE__, pvarIn->vt);
        return E_INVALIDARG;
    }

    return S_OK;
}

HRESULT VarFix(LPVARIANT pvarIn, LPVARIANT pvarResult){
    if(pvarIn->vt == VT_R8){
        pvarResult->vt = VT_I8;
        pvarResult->llVal = pvarIn->dblVal;
    }else
    if(pvarIn->vt == VT_I8){
        pvarResult->vt = VT_I8;
        pvarResult->llVal = pvarIn->llVal;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d)\n", __func__, __FILE__, __LINE__, pvarIn->vt);
        return E_INVALIDARG;
    }

    return S_OK;
}

HRESULT VarAbs(LPVARIANT pvarIn, LPVARIANT pvarResult){
    if(pvarIn->vt == VT_R8){
        pvarResult->vt = VT_R8;
        pvarResult->dblVal = (pvarIn->dblVal < 0) ? -pvarIn->dblVal : pvarIn->dblVal;
    }else
    if(pvarIn->vt == VT_I8){
        pvarResult->vt = VT_I8;
        pvarResult->llVal = (pvarIn->llVal < 0) ? -pvarIn->llVal : pvarIn->llVal;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d)\n", __func__, __FILE__, __LINE__, pvarIn->vt);
        return E_INVALIDARG;
    }

    return S_OK;
}

HRESULT VarPow(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult){
    if(pvarLeft->vt == (VT_BYREF|VT_VARIANT)) pvarLeft = pvarLeft->pvarVal;
    if(pvarRight->vt == (VT_BYREF|VT_VARIANT)) pvarRight = pvarRight->pvarVal;
    
    if(pvarLeft->vt == VT_I8 && pvarRight->vt == VT_I8){
        pvarResult->vt = VT_I8;
        pvarResult->llVal = pow(pvarLeft->llVal, pvarRight->llVal);
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarLeft->vt, pvarRight->vt);
        return E_NOTIMPL;
    }

    return S_OK;
}

HRESULT VarMul(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult){
    if(pvarLeft->vt == (VT_BYREF|VT_VARIANT)) pvarLeft = pvarLeft->pvarVal;
    if(pvarRight->vt == (VT_BYREF|VT_VARIANT)) pvarRight = pvarRight->pvarVal;
    
    _variant_t vL;
    if(pvarRight->vt == VT_R8 && pvarLeft->vt != VT_R8){
        VariantChangeType(&vL, pvarLeft, 0, VT_R8);
        pvarLeft = &vL;
    }
    _variant_t vR;
    if(pvarLeft->vt == VT_R8 && pvarRight->vt != VT_R8){
        VariantChangeType(&vR, pvarRight, 0, VT_R8);
        pvarRight = &vR;
    }

    if(pvarLeft->vt == VT_I8 && pvarRight->vt == VT_I8){
        pvarResult->vt = VT_I8;
        pvarResult->llVal = pvarLeft->llVal * pvarRight->llVal;
    }else
    if(pvarLeft->vt == VT_R8 && pvarRight->vt == VT_R8){
        pvarResult->vt = VT_R8;
        pvarResult->dblVal = pvarLeft->dblVal * pvarRight->dblVal;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarLeft->vt, pvarRight->vt);
        return E_NOTIMPL;
    }

    return S_OK;
}

HRESULT VarMod(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult){
    if(pvarLeft->vt == (VT_BYREF|VT_VARIANT)) pvarLeft = pvarLeft->pvarVal;
    if(pvarRight->vt == (VT_BYREF|VT_VARIANT)) pvarRight = pvarRight->pvarVal;

    if(pvarLeft->vt == VT_I8 && pvarRight->vt == VT_I8){
        if(pvarRight->llVal == 0) return DISP_E_DIVBYZERO;

        pvarResult->vt = VT_I8;
        pvarResult->llVal = pvarLeft->llVal % pvarRight->llVal;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarLeft->vt, pvarRight->vt);
        return E_NOTIMPL;
    }

    return S_OK;
}

HRESULT VarDiv(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult){
    if(pvarLeft->vt == (VT_BYREF|VT_VARIANT)) pvarLeft = pvarLeft->pvarVal;
    if(pvarRight->vt == (VT_BYREF|VT_VARIANT)) pvarRight = pvarRight->pvarVal;
    
    _variant_t vL;
    if(pvarRight->vt == VT_R8 && pvarLeft->vt != VT_R8){
        VariantChangeType(&vL, pvarLeft, 0, VT_R8);
        pvarLeft = &vL;
    }
    _variant_t vR;
    if(pvarLeft->vt == VT_R8 && pvarRight->vt != VT_R8){
        VariantChangeType(&vR, pvarRight, 0, VT_R8);
        pvarRight = &vR;
    }

    if(pvarLeft->vt == VT_I8 && pvarRight->vt == VT_I8){
        if(pvarRight->llVal == 0) return DISP_E_DIVBYZERO;

        pvarResult->vt = VT_R8;
        pvarResult->dblVal = (double)pvarLeft->llVal / (double)pvarRight->llVal;
    }else
    if(pvarLeft->vt == VT_R8 && pvarRight->vt == VT_R8){
        if(pvarRight->dblVal == 0.0) return DISP_E_DIVBYZERO;

        pvarResult->vt = VT_R8;
        pvarResult->dblVal = pvarLeft->dblVal / pvarRight->dblVal;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarLeft->vt, pvarRight->vt);
        return E_NOTIMPL;
    }

    return S_OK;
}

HRESULT VarAdd(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult){
    if(pvarLeft->vt == (VT_BYREF|VT_VARIANT)) pvarLeft = pvarLeft->pvarVal;
    if(pvarRight->vt == (VT_BYREF|VT_VARIANT)) pvarRight = pvarRight->pvarVal;
    
    _variant_t vL;
    if(pvarRight->vt == VT_R8 && pvarLeft->vt != VT_R8){
        VariantChangeType(&vL, pvarLeft, 0, VT_R8);
        pvarLeft = &vL;
    }
    _variant_t vR;
    if(pvarLeft->vt == VT_R8 && pvarRight->vt != VT_R8){
        VariantChangeType(&vR, pvarRight, 0, VT_R8);
        pvarRight = &vR;
    }

    if(pvarLeft->vt == VT_I8 && pvarRight->vt == VT_I8){
        pvarResult->vt = VT_I8;
        pvarResult->llVal = pvarLeft->llVal + pvarRight->llVal;
    }else
    if(pvarLeft->vt == VT_R8 && pvarRight->vt == VT_R8){
        pvarResult->vt = VT_R8;
        pvarResult->dblVal = pvarLeft->dblVal + pvarRight->dblVal;
    }else
    if(pvarLeft->vt == VT_BSTR && pvarRight->vt == VT_BSTR){
        return VarCat(pvarLeft, pvarRight, pvarResult);
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarLeft->vt, pvarRight->vt);
        return E_NOTIMPL;
    }

    return S_OK;
}

HRESULT VarSub(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult){
    if(pvarLeft->vt == (VT_BYREF|VT_VARIANT)) pvarLeft = pvarLeft->pvarVal;
    if(pvarRight->vt == (VT_BYREF|VT_VARIANT)) pvarRight = pvarRight->pvarVal;
    
    _variant_t vL;
    if(pvarRight->vt == VT_R8 && pvarLeft->vt != VT_R8){
        VariantChangeType(&vL, pvarLeft, 0, VT_R8);
        pvarLeft = &vL;
    }
    _variant_t vR;
    if(pvarLeft->vt == VT_R8 && pvarRight->vt != VT_R8){
        VariantChangeType(&vR, pvarRight, 0, VT_R8);
        pvarRight = &vR;
    }

    if(pvarLeft->vt == VT_I8 && pvarRight->vt == VT_I8){
        pvarResult->vt = VT_I8;
        pvarResult->llVal = pvarLeft->llVal - pvarRight->llVal;
    }else
    if(pvarLeft->vt == VT_R8 && pvarRight->vt == VT_R8){
        pvarResult->vt = VT_R8;
        pvarResult->dblVal = pvarLeft->dblVal - pvarRight->dblVal;
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarLeft->vt, pvarRight->vt);
        return E_NOTIMPL;
    }

    return S_OK;
}

HRESULT VarCat(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult){
    if(pvarLeft->vt == (VT_BYREF|VT_VARIANT)) pvarLeft = pvarLeft->pvarVal;
    if(pvarRight->vt == (VT_BYREF|VT_VARIANT)) pvarRight = pvarRight->pvarVal;
    
    if(pvarLeft->vt == VT_NULL && pvarRight->vt == VT_NULL){
        pvarResult->vt = VT_NULL;
        return S_OK;
    }else
    if(pvarLeft->vt == VT_NULL){
        return VariantChangeType(pvarResult, pvarRight, 0, VT_BSTR);
    }else
    if(pvarRight->vt == VT_NULL){
        return VariantChangeType(pvarResult, pvarLeft, 0, VT_BSTR);
    }

    _variant_t vL;
    if(pvarLeft->vt != VT_BSTR){
        VariantChangeType(&vL, pvarLeft, 0, VT_BSTR);
        pvarLeft = &vL;
    }
    _variant_t vR;
    if(pvarRight->vt != VT_BSTR){
        VariantChangeType(&vR, pvarRight, 0, VT_BSTR);
        pvarRight = &vR;
    }

    if(pvarLeft->vt == VT_BSTR && pvarRight->vt == VT_BSTR){
        UINT nL = SysStringLen(pvarLeft->bstrVal);
        UINT nR = SysStringLen(pvarRight->bstrVal);
        BSTR bs = SysAllocStringLen(nullptr, nL+nR);
        memcpy(bs   , pvarLeft->bstrVal , sizeof(*bs)*nL);
        memcpy(bs+nL, pvarRight->bstrVal, sizeof(*bs)*nR);
        pvarResult->vt = VT_BSTR;
        pvarResult->bstrVal = bs;

        return S_OK;
    }

wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarLeft->vt, pvarRight->vt);
    return E_NOTIMPL;
}

HRESULT VarDateFromStr(LPCOLESTR strIn, LCID lcid, ULONG dwFlags, DATE *pdateOut){
    char s[256];
    wchar_utf8(s, 256, strIn);

    tm t = {};
    if(
        strptime(s, "%Y/%m/%d %H:%M:%S", &t) ||
        strptime(s, "%Y/%m/%d %H:%M", &t) ||
        strptime(s, "%Y/%m/%d", &t) ||
        strptime(s, "%Y-%m-%d %H:%M:%S", &t) ||
        strptime(s, "%Y-%m-%d %H:%M", &t) ||
        strptime(s, "%Y-%m-%d", &t) ||
    false){
        *pdateOut = tm_double(t);
    }else
    if(
        strptime(s, "%H:%M:%S", &t) ||
    false){
        t.tm_year = 1899 -1900;
        t.tm_mon  = 12 -1;
        t.tm_mday = 30;
        *pdateOut = tm_double(t);
    }else
    {
        return E_INVALIDARG;
    }

    return S_OK;
}

HRESULT VariantChangeType(VARIANT *pvargDest, const VARIANT *pvarSrc, USHORT wFlags, VARTYPE vt){
    if(pvarSrc->vt == (VT_BYREF|VT_VARIANT)) pvarSrc = pvarSrc->pvarVal;

    if(pvarSrc->vt == vt){
        return VariantCopy(pvargDest, pvarSrc);
    }else
    if(vt == VT_BSTR){
        if(pvarSrc->vt == VT_EMPTY){
            pvargDest->vt = VT_BSTR;
            pvargDest->bstrVal = SysAllocString(L"");
        }else
        if(pvarSrc->vt == VT_BSTR){
            pvargDest->vt = VT_BSTR;
            pvargDest->bstrVal = SysAllocString(pvarSrc->bstrVal);
        }else
        if(pvarSrc->vt == VT_I8){
            wchar_t buf[1+19 +1];
            swprintf(buf, 1+19+1, L"%lld", pvarSrc->llVal);
            pvargDest->vt = VT_BSTR;
            pvargDest->bstrVal = SysAllocString(buf);
        }else
        if(pvarSrc->vt == VT_UI8){
            wchar_t buf[1+19 +1];
            swprintf(buf, 1+19+1, L"%llu", pvarSrc->ullVal);
            pvargDest->vt = VT_BSTR;
            pvargDest->bstrVal = SysAllocString(buf);
        }else
        if(pvarSrc->vt == VT_I4){
            wchar_t buf[1+19 +1];
            swprintf(buf, 1+19+1, L"%lld", (long long)pvarSrc->lVal);
            pvargDest->vt = VT_BSTR;
            pvargDest->bstrVal = SysAllocString(buf);
        }else
        if(pvarSrc->vt == VT_UI1){
            wchar_t buf[1+19 +1];
            swprintf(buf, 1+19+1, L"%lld", (long long)pvarSrc->bVal);
            pvargDest->vt = VT_BSTR;
            pvargDest->bstrVal = SysAllocString(buf);
        }else
        if(pvarSrc->vt == VT_BOOL){
            wchar_t buf[1+19 +1];
            swprintf(buf, 1+19+1, L"%lld", (long long)pvarSrc->boolVal);
            pvargDest->vt = VT_BSTR;
            pvargDest->bstrVal = SysAllocString(buf);
        }else
        if(pvarSrc->vt == VT_R8){
            wchar_t buf[1+30 +1];
            swprintf(buf, 1+30+1, L"%g", pvarSrc->dblVal);
            pvargDest->vt = VT_BSTR;
            pvargDest->bstrVal = SysAllocString(buf);
        }else
        if(pvarSrc->vt == VT_DATE){
            tm t = double_tm(pvarSrc->date);
            wchar_t buf[4+1+2+1+2+1+2+1+2+1+2 +1];
            if(0.0 <= pvarSrc->date && pvarSrc->date < 1.0){
                swprintf(buf, 4+1+2+1+2+1+2+1+2+1+2+1, L"%02d:%02d:%02d", t.tm_hour, t.tm_min, t.tm_sec);
            }else
            if(t.tm_hour + t.tm_min + t.tm_sec){
                swprintf(buf, 4+1+2+1+2+1+2+1+2+1+2+1, L"%04d/%02d/%02d %02d:%02d:%02d", t.tm_year+1900, t.tm_mon+1, t.tm_mday, t.tm_hour, t.tm_min, t.tm_sec);
            }else{
                swprintf(buf, 4+1+2+1+2+1+2+1+2+1+2+1, L"%04d/%02d/%02d", t.tm_year+1900, t.tm_mon+1, t.tm_mday);
            }
            pvargDest->vt = VT_BSTR;
            pvargDest->bstrVal = SysAllocString(buf);
        }else
        if(pvarSrc->vt == VT_DISPATCH){
            DISPPARAMS param = { nullptr, nullptr, 0, 0 };
            _variant_t v;
            HRESULT hr = pvarSrc->pdispVal->Invoke(0, IID_NULL, 0, DISPATCH_METHOD, &param, &v, nullptr, nullptr);

            return SUCCEEDED(hr) ? VariantChangeType(pvargDest, &v, wFlags, vt) : hr;
        }else
        {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarSrc->vt, vt);
            return E_INVALIDARG;
        }
    }else
    if(vt == VT_BOOL){
        if(pvarSrc->vt == VT_EMPTY){
            pvargDest->vt = VT_BOOL;
            pvargDest->boolVal = VARIANT_FALSE;
        }else
        if(pvarSrc->vt == VT_I8){
            pvargDest->vt = VT_BOOL;
            pvargDest->boolVal = pvarSrc->llVal ? VARIANT_TRUE : VARIANT_FALSE;
        }else
        if(pvarSrc->vt == VT_BSTR){
            VARIANT_BOOL b;
            if( _wcsicmp(pvarSrc->bstrVal, L"true") == 0 ){
                b = VARIANT_TRUE;
            }else
            if( _wcsicmp(pvarSrc->bstrVal, L"false") == 0 ){
                b = VARIANT_FALSE;
            }else
            {
                return E_INVALIDARG;
            }

            pvargDest->vt = VT_BOOL;
            pvargDest->boolVal = b;
        }else
        {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarSrc->vt, vt);
            return E_INVALIDARG;
        }
    }else
    if(vt == VT_R8){
        if(pvarSrc->vt == VT_EMPTY){
            pvargDest->vt = VT_R8;
            pvargDest->dblVal = 0.0;
        }else
        if(pvarSrc->vt == VT_I8){
            pvargDest->vt = VT_R8;
            pvargDest->dblVal = pvarSrc->llVal;
        }else
        if(pvarSrc->vt == VT_BSTR){
            wchar_t* pend;
            double dbl = wcstod(pvarSrc->bstrVal, &pend);
            if(pvarSrc->bstrVal == pend) return E_INVALIDARG;

            pvargDest->vt = VT_R8;
            pvargDest->dblVal = dbl;
        }else
        {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarSrc->vt, vt);
            return E_INVALIDARG;
        }
    }else
    if(vt == VT_I8){
        if(pvarSrc->vt == VT_EMPTY){
            pvargDest->vt = VT_I8;
            pvargDest->llVal = 0;
        }else
        if(pvarSrc->vt == VT_BSTR){
            wchar_t* pend;
            long long ll = wcstoll(pvarSrc->bstrVal, &pend, 10);
            if(pvarSrc->bstrVal == pend) return E_INVALIDARG;

            pvargDest->vt = VT_I8;
            pvargDest->llVal = ll;
        }else
        if(pvarSrc->vt == VT_I4){
            pvargDest->vt = VT_I8;
            pvargDest->llVal = pvarSrc->lVal;
        }else
        if(pvarSrc->vt == VT_UI1){
            pvargDest->vt = VT_I8;
            pvargDest->llVal = pvarSrc->bVal;
        }else
        if(pvarSrc->vt == VT_R8){
            pvargDest->vt = VT_I8;
            pvargDest->llVal = pvarSrc->dblVal;
        }else
        {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarSrc->vt, vt);
            return E_INVALIDARG;
        }
    }else
    if(vt == VT_I4){
        if(pvarSrc->vt == VT_EMPTY){
            pvargDest->vt = VT_I4;
            pvargDest->lVal = 0;
        }else
        if(pvarSrc->vt == VT_I8){
            pvargDest->vt = VT_I4;
            pvargDest->lVal = pvarSrc->llVal;
        }else
        {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarSrc->vt, vt);
            return E_INVALIDARG;
        }
    }else
    if(vt == VT_UI1){
        if(pvarSrc->vt == VT_EMPTY){
            pvargDest->vt = VT_UI1;
            pvargDest->bVal = 0;
        }else
        if(pvarSrc->vt == VT_I8){
            pvargDest->vt = VT_UI1;
            pvargDest->bVal = pvarSrc->llVal;
        }else
        {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarSrc->vt, vt);
            return E_INVALIDARG;
        }
    }else
    if(vt == VT_DATE){
        if(pvarSrc->vt == VT_EMPTY){
            pvargDest->vt = VT_DATE;
            pvargDest->date = 0.0;
        }else
        if(pvarSrc->vt == VT_I8){
            pvargDest->vt = VT_DATE;
            pvargDest->date = pvarSrc->llVal;
        }else
        if(pvarSrc->vt == VT_BSTR){
            DATE date;
            HRESULT hr = VarDateFromStr(pvarSrc->bstrVal, 0, 0, &date);
            if(FAILED(hr)) return hr;

            pvargDest->vt = VT_DATE;
            pvargDest->date = date;
        }else
        {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarSrc->vt, vt);
            return E_INVALIDARG;
        }
    }else
    {
wprintf(L"###%s: Implement here '%s' line %d. (vt:%d->%d)\n", __func__, __FILE__, __LINE__, pvarSrc->vt, vt);
        return E_INVALIDARG;
    }

    return S_OK;
}

HRESULT VariantCopy(VARIANT *pvargDest, const VARIANT *pvargSrc){
    VariantClear(pvargDest);
    *pvargDest = *pvargSrc;

    if(pvargDest->vt == VT_DISPATCH){
        if(pvargDest->pdispVal)
            pvargDest->pdispVal->AddRef();
    }else
    if(pvargDest->vt == VT_BSTR){
        pvargDest->bstrVal = SysAllocString(pvargSrc->bstrVal);
    }else
    if(pvargDest->vt & VT_ARRAY){
        SafeArrayCopy(pvargSrc->parray, &pvargDest->parray);
    }

    return S_OK;
}

HRESULT VariantClear(VARIANT *pvarg){
    if(pvarg->vt == VT_DISPATCH){
        if(pvarg->pdispVal)
            pvarg->pdispVal->Release();
    }else
    if(pvarg->vt == VT_BSTR){
        SysFreeString(pvarg->bstrVal);
    }else
    if(pvarg->vt & VT_ARRAY){
        SafeArrayDestroy(pvarg->parray);
    }

    pvarg->vt = VT_EMPTY;

    return S_OK;
}

void VariantInit(VARIANT *pvarg){
    memset(pvarg, 0, sizeof(*pvarg));
}

BSTR SysAllocString(const OLECHAR *psz){
    BSTR ret = nullptr;
    if(psz){
        size_t len = wcslen(psz);
        BYTE* p = (BYTE*)malloc(sizeof(size_t) + sizeof(wchar_t)*(len+1));
        *(size_t*)p = len;
        ret = wcscpy( (BSTR)(p+sizeof(size_t)), psz );
    }

    return ret;
}

BSTR SysAllocStringLen(const OLECHAR *strIn, UINT ui){
    BYTE* p = (BYTE*)malloc(sizeof(size_t) + sizeof(wchar_t)*(ui+1));
    *(size_t*)p = ui;
    BSTR ret = (BSTR)(p+sizeof(size_t));
    if(strIn){
        memcpy( ret, strIn, sizeof(*strIn)*ui );
    }
    ret[ui] = L'\0';

    return ret;
}

UINT SysStringLen(BSTR pbstr){
    return pbstr ? *(size_t*)((BYTE*)pbstr-sizeof(size_t)) : 0;
}

void SysFreeString(BSTR bstrString){
    if(bstrString){
        free( (BYTE*)bstrString-sizeof(size_t) );
    }
}

SAFEARRAY* SafeArrayCreate(VARTYPE vt, UINT cDims, SAFEARRAYBOUND *rgsabound){
    if(vt != VT_VARIANT) return nullptr;//###NOTIMPL

    ULONG total = 0;{
        UINT i=0;
        while(i<cDims) total += rgsabound[i++].cElements;
    }

    VARIANT* data;{
        data = (VARIANT*)malloc(sizeof(VARIANT)*total);
        ULONG i=0;
        while(i<total) VariantInit(&data[i++]);
    }

    SAFEARRAY* psa = (SAFEARRAY*)malloc(sizeof(SAFEARRAY) + sizeof(SAFEARRAYBOUND)*cDims);
    {
        psa->cDims = cDims;
        psa->fFeatures = FADF_VARIANT;
        psa->cbElements = sizeof(VARIANT)*total;
        psa->cLocks = 0;
        psa->pvData = data;
        memcpy(psa->rgsabound, rgsabound, sizeof(SAFEARRAYBOUND)*cDims);
        psa->rgsabound[cDims].cElements = 0;
        psa->rgsabound[cDims].lLbound = 0;
    }

    return psa;
}

HRESULT SafeArrayCopy(SAFEARRAY *psa, SAFEARRAY **ppsaOut){
    if(!(psa->fFeatures&FADF_VARIANT)) return E_NOTIMPL;

    ULONG total = psa->cbElements/sizeof(VARIANT);

    VARIANT* data;{
        data = (VARIANT*)malloc(sizeof(VARIANT)*total);
        ULONG i=0;
        while(i<total) VariantInit(&data[i++]);
        VARIANT* psrc = (VARIANT*)psa->pvData;
        i=0;
        while(i<total){ VariantCopy(&data[i], &psrc[i]); ++i; }
    }

    *ppsaOut = (SAFEARRAY*)malloc(sizeof(SAFEARRAY) + sizeof(SAFEARRAYBOUND)*psa->cDims);
    {
        (*ppsaOut)->cDims = psa->cDims;
        (*ppsaOut)->fFeatures = psa->fFeatures;
        (*ppsaOut)->cbElements = psa->cbElements;
        (*ppsaOut)->cLocks = 0;
        (*ppsaOut)->pvData = data;
        memcpy((*ppsaOut)->rgsabound, psa->rgsabound, sizeof(SAFEARRAYBOUND)*psa->cDims);
        (*ppsaOut)->rgsabound[psa->cDims].cElements = 0;
        (*ppsaOut)->rgsabound[psa->cDims].lLbound = 0;
    }

    return S_OK;
}

HRESULT SafeArrayDestroy(SAFEARRAY *psa){
    if(!(psa->fFeatures&FADF_VARIANT)) return E_NOTIMPL;

    ULONG total = psa->cbElements/sizeof(VARIANT);
    VARIANT* data = (VARIANT*)psa->pvData;

    ULONG i=0;
    while(i<total) VariantClear(&data[i++]);

    free(psa->pvData);
    free(psa);

    return S_OK;
}

HRESULT SafeArrayGetUBound(SAFEARRAY *psa, UINT nDim, LONG *plUbound){
    *plUbound = psa->rgsabound[nDim-1].lLbound + psa->rgsabound[nDim-1].cElements - 1;

    return S_OK;
}

HRESULT SafeArrayGetLBound(SAFEARRAY *psa, UINT nDim, LONG *plLbound){
    *plLbound = psa->rgsabound[nDim-1].lLbound;

    return S_OK;
}

HRESULT SafeArrayAccessData(SAFEARRAY  *psa, void **ppvData){
    *ppvData = psa->pvData;

    return S_OK;
}

HRESULT SafeArrayUnaccessData(SAFEARRAY *psa){
    // none

    return S_OK;
}

INT SystemTimeToVariantTime(LPSYSTEMTIME lpSystemTime, DOUBLE *pvtime){
    tm t;{
        t.tm_year = lpSystemTime->wYear-1900;
        t.tm_mon  = lpSystemTime->wMonth-1;
        t.tm_mday = lpSystemTime->wDay;
        t.tm_hour = lpSystemTime->wHour;
        t.tm_min  = lpSystemTime->wMinute;
        t.tm_sec  = lpSystemTime->wSecond;
    }

    *pvtime = tm_double(t);

    return TRUE;
}

INT VariantTimeToSystemTime(DOUBLE vtime, LPSYSTEMTIME lpSystemTime){
    tm t = double_tm(vtime);

    lpSystemTime->wYear         = t.tm_year+1900;
    lpSystemTime->wMonth        = t.tm_mon+1;
    lpSystemTime->wDayOfWeek    = t.tm_wday;
    lpSystemTime->wDay          = t.tm_mday;
    lpSystemTime->wHour         = t.tm_hour;
    lpSystemTime->wMinute       = t.tm_min;
    lpSystemTime->wSecond       = t.tm_sec;
    lpSystemTime->wMilliseconds = 0;

    return TRUE;
}

// others
double tm_double(tm t){
    t.tm_year += 1900;
    t.tm_mon += 1;

    if(t.tm_mon < 3){
        --t.tm_year;
        t.tm_mon += 12;
    }

    uint64_t dy = 365 * t.tm_year;
    uint64_t du = t.tm_year/4 - t.tm_year/100 + t.tm_year/400;
    uint64_t dm = (153*(t.tm_mon-3)+2)/5;

    uint64_t days = dy + du + dm + t.tm_mday - 1;
    uint64_t secs = 60*60*t.tm_hour + 60*t.tm_min + t.tm_sec;

    return (double)days + (double)secs/(60*60*24) - D18991230;
}

tm double_tm(double d){
    const uint64_t D1   = 365;
    const uint64_t D4   = D1*4+1;   // 1461
    const uint64_t D100 = D4*25-1;  // 36524
    const uint64_t D400 = D100*4+1; // 146097

    d += D18991230;

    uint64_t dd = (uint64_t)d;
    uint64_t y400 = dd / D400;
    uint64_t y100 = (dd%D400) / D100;
    uint64_t y4   = ((dd%D400)%D100) / D4;
    uint64_t y1   = (((dd%D400)%D100)%D4) / D1;

    tm t = {};
    t.tm_year = y400*400 + y100*100 + y4*4 + y1;
    t.tm_mon  = (5*((((dd%D400)%D100)%D4)%D1) + 2) / 153 + 3;
    t.tm_mday = ((((dd%D400)%D100)%D4)%D1) - (153*(t.tm_mon-3)+2)/5 + 1;
    t.tm_wday = -1;//###TODO

    uint64_t ds = (d - dd) * (60*60*24);
    t.tm_hour = ds / (60*60);
    t.tm_min  = (ds%(60*60)) / 60;
    t.tm_sec  = (ds%(60*60)) % 60;

    if(12 < t.tm_mon){
        t.tm_mon -= 12;
        ++t.tm_year;
    }

    t.tm_year -= 1900;
    t.tm_mon -= 1;

    return t;
}
#endif



size_t utf8_wchar(wchar_t* out, size_t outc, const char* in){
    size_t i = 0;

    while(*in && i<outc-1){
        if(*in == '\r'){
            ++in;
        }else
        if((*(in+0)&0b10000000) == 0b00000000){
            if(outc) out[i] = *in;
            ++i;
            in += 1;
        }else
        if((*(in+0)&0b11100000) == 0b11000000 && (*(in+1)&0b11000000) == 0b10000000){
            if(outc) out[i] = (wchar_t)(*(in+0) & 0b00011111)<<6 | (wchar_t)(*(in+1) & 0b00111111);
            ++i;
            in += 2;
        }else
        if((*(in+0)&0b11110000) == 0b11100000 && (*(in+1)&0b11000000) == 0b10000000 && (*(in+2)&0b11000000) == 0b10000000){
            if(outc) out[i] = (wchar_t)(*(in+0) & 0b00001111)<<12 | (wchar_t)(*(in+1) & 0b00111111)<<6 | (wchar_t)(*(in+2) & 0b00111111);
            ++i;
            in += 3;
        }else
#if 0xffff<WCHAR_MAX
        if((*(in+0)&0b11111000) == 0b11110000 && (*(in+1)&0b11000000) == 0b10000000 && (*(in+2)&0b11000000) == 0b10000000 && (*(in+3)&0b11000000) == 0b10000000){
            if(outc) out[i] = (wchar_t)(*(in+0) & 0b00001111)<<18 | (wchar_t)(*(in+1) & 0b00111111)<<12 | (wchar_t)(*(in+2) & 0b00111111)<<6 | (wchar_t)(*(in+3) & 0b00111111);
            ++i;
            in += 4;
        }else
#endif
        {
            return 0;
        }
    }

    if(outc) out[i] = 0;
    ++i;

    return i;
}

size_t wchar_utf8(char* out, size_t outc, const wchar_t* in){
    size_t i = 0;

    while(*in && i<outc-1){
        if((*in&0b11111111111111111111111110000000) == 0){
            if(outc) out[i] = *in;
            i += 1;
            ++in;
        }else
        if((*in&0b11111111111111111111100000000000) == 0){
            if(outc){out[i+0] = 0b11000000 | (*in>>6);
                     out[i+1] = 0b10000000 | (*in&0b0000000000111111);}
            i += 2;
            ++in;
        }else
        if((*in&0b11111111111111110000000000000000) == 0){
            if(outc){out[i+0] = 0b11100000 | (*in>>12);
                     out[i+1] = 0b10000000 | ((*in>>6)&0b0000000000111111);
                     out[i+2] = 0b10000000 | (*in&0b0000000000111111);}
            i += 3;
            ++in;
        }else
#if 0xffff<WCHAR_MAX
        {
            if(outc){out[i+0] = 0b11110000 | (*in>>18);
                     out[i+1] = 0b10000000 | ((*in>>12)&0b0000000000111111);
                     out[i+2] = 0b10000000 | ((*in>>6)&0b0000000000111111);
                     out[i+3] = 0b10000000 | (*in&0b0000000000111111);}
            i += 4;
            ++in;
        }
#else
        {
            // none
        }
#endif
    }

    if(outc) out[i] = 0;
    ++i;

    return i;
}



bool operator<(const _variant_t& l, const _variant_t& r){
    return (VarCmp((LPVARIANT)&l, (LPVARIANT)&r, 0, 0) == VARCMP_LT);
}


