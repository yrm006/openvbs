#include "jujube.h"

#include <random>
#include <locale.h>
#include <stdlib.h>









class VBScript : public IDispatch{
private:
    ULONG       m_refc   = 1;

    typedef HRESULT (VBScript::*invoke_t)(DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*);
    union invoke_cast{ invoke_t i; void* v; };
    static int                       s_disps_n;
    static std::map<istring, DISPID> s_disps_ids;
    static std::vector<_variant_t>   s_disps;

    std::mt19937_64 m_rnd;
    double          m_rnd_last = (double)m_rnd() / UINT64_MAX;

    std::vector<_com_ptr_t<_com_IIID<IClassFactory, &IID_IClassFactory> > > m_cfs;

public:
    VBScript(){}
    virtual ~VBScript(){}

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        auto foundC = s_disps_ids.find(*rgszNames);
        if(foundC != s_disps_ids.end()){
            *rgDispId = foundC->second;
            return (s_disps[*rgDispId].vt == VT_EMPTY) ? S_OK : S_FALSE;
        }else{
            wchar_t path[256];{
                wcsncpy(path, *rgszNames, 256)[255] = L'\0';
                for(size_t i=0; path[i]; ++i) path[i] = toupper(path[i]);
            }

            HMODULE hm = LoadLibraryW(path);
            if(hm){
                LPFNGETCLASSOBJECT pfn =  (LPFNGETCLASSOBJECT)GetProcAddress(hm, "DllGetClassObject");
                if(pfn){
                    IClassFactory* cf = (IClassFactory*)*rgszNames;
                    if( SUCCEEDED(pfn(CLSID_NULL, IID_IClassFactory, (void**)&cf)) ){
                        m_cfs.push_back(nullptr);
                        m_cfs.back().Attach(cf);

                        *rgDispId = s_disps_n + (m_cfs.size()-1);
                        return S_OK;
                    }
                }
            }
        }

        return E_FAIL;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(dispIdMember < s_disps_n){
            _variant_t& v = s_disps[dispIdMember];
            if(v.vt == VT_EMPTY){
                return (this->*((invoke_cast*)&v.byref)->i)(dispIdMember, riid, lcid, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
            }else{
                return VariantCopy(pVarResult, &v);
            }
        }else{
            IDispatch* obj = (IDispatch*)pDispParams;
            if(SUCCEEDED(m_cfs[dispIdMember-s_disps_n]->CreateInstance(nullptr, IID_IDispatch, (void**)&obj))){
                pVarResult->vt = VT_DISPATCH;
                pVarResult->pdispVal = obj;
                return S_OK;
            }
        }

        return E_FAIL;
    }

    // DISPID_VALUE
    HRESULT _(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        return E_NOTIMPL;
    }

    // functions
    HRESULT vbCreateObject(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;

        CLSID clsid;
        HRESULT hr;
        if(
            SUCCEEDED( hr = (pv1->bstrVal[0] == L'{') ? CLSIDFromString(pv1->bstrVal, &clsid) : CLSIDFromProgID(pv1->bstrVal, &clsid) ) &&
            SUCCEEDED( hr = CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER|CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pVarResult->pdispVal) )   &&
        true){
            pVarResult->vt = VT_DISPATCH;
        }

        return hr;
    }

    HRESULT vbGetObject(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;

        HRESULT hr;
        if( SUCCEEDED( hr = CoGetObject(pv1->bstrVal, nullptr, IID_IDispatch, (void**)&pVarResult->pdispVal) ) ){
            pVarResult->vt = VT_DISPATCH;
        }

        return hr;
    }

    HRESULT vbTypeName(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        BSTR name = nullptr;

        VARTYPE mvt = (pv->vt & VT_TYPEMASK);
        if(mvt == VT_EMPTY){
            name = SysAllocString(L"Empty");
        }else
        if(mvt == VT_NULL){
            name = SysAllocString(L"Null");
        }else
        if(mvt == VT_I1 || mvt == VT_UI1){
            name = SysAllocString(L"Byte");
        }else
        if(mvt == VT_I2 || mvt == VT_UI2){
            name = SysAllocString(L"Integer");
        }else
        if(mvt == VT_I4 || mvt == VT_UI4){
            name = SysAllocString(L"Long");
        }else
        if(mvt == VT_I8 || mvt == VT_UI8){
            name = SysAllocString(L"Longlong");
        }else
        if(mvt == VT_R4){
            name = SysAllocString(L"Single");
        }else
        if(mvt == VT_R8){
            name = SysAllocString(L"Double");
        }else
        if(mvt == VT_CY){
            name = SysAllocString(L"Currency");
        }else
        if(mvt == VT_DATE){
            name = SysAllocString(L"Date");
        }else
        if(mvt == VT_BSTR){
            name = SysAllocString(L"String");
        }else
        if(mvt == VT_DISPATCH){
            if(pv->pdispVal){
                name = SysAllocString(L"Object");
            }else{
                name = SysAllocString(L"Nothing");
            }
        }else
        if(mvt == VT_UNKNOWN){
            name = SysAllocString(L"Unknown");
        }else
        if(mvt == VT_BOOL){
            name = SysAllocString(L"Boolean");
        }else
        if(mvt == VT_DECIMAL){
            name = SysAllocString(L"Decimal");
        }else
        if(mvt == VT_ERROR){
            name = SysAllocString(L"Error");
        }else
        if(mvt == VT_VARIANT){
            name = SysAllocString(L"Variant");
        }else
        {
            //none
        }

        if(name){
            if(pv->vt & VT_ARRAY){
                std::wstring s(name);
                SysFreeString(name);
                name = SysAllocString( (s+L"()").c_str() );
            }

            pVarResult->vt = VT_BSTR;
            pVarResult->bstrVal = name;
            return S_OK;
        }

        return E_FAIL;
    }

    HRESULT vbVarDump(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        std::wstring s;

        VARTYPE mvt = (pv->vt & VT_TYPEMASK);
        if(mvt == VT_EMPTY){
            s += (std::wstring)L"Empty";
        }else
        if(mvt == VT_NULL){
            s += (std::wstring)L"Null";
        }else
        if(mvt == VT_UI1){
            wchar_t buf[0xff];
            swprintf(buf, 0xff, L"%u", pv->bVal);
            s += (std::wstring)L"Byte(" + buf + L")";
        }else
        if(mvt == VT_UI2){
            wchar_t buf[0xff];
            swprintf(buf, 0xff, L"%hu", pv->uiVal);
            s += (std::wstring)L"Integer(" + buf + L")";
        }else
        if(mvt == VT_I4){
            wchar_t buf[0xff];
            swprintf(buf, 0xff, L"%d", pv->lVal);
            s += (std::wstring)L"Long(" + buf + L")";
        }else
        if(mvt == VT_I8 || mvt == VT_UI8){
            wchar_t buf[0xff];
            swprintf(buf, 0xff, L"%lld", pv->llVal);
            s += (std::wstring)L"Longlong(" + buf + L")";
        }else
        if(mvt == VT_R8){
            wchar_t buf[0xff];
            swprintf(buf, 0xff, L"%g", pv->dblVal);
            s += (std::wstring)L"Double(" + buf + L")";
        }else
        if(mvt == VT_BSTR){
            wchar_t buf[0xff];
            swprintf(buf, 0xff, L"%u", SysStringLen(pv->bstrVal));
            s += (std::wstring)L"String(" + buf + L") \"" + pv->bstrVal + L"\"";
        }else
        if(mvt == VT_DISPATCH){
            if(pv->pdispVal){
                s += (std::wstring)L"Object";
                //### show JSON strin when this is json-object.
            }else{
                s += (std::wstring)L"Nothing";
            }
        }else
        if(mvt == VT_BOOL){
            wchar_t buf[0xff];
            swprintf(buf, 0xff, L"%ls", pv->boolVal==VARIANT_TRUE ? L"true" : L"false");
            s += (std::wstring)L"Boolean(" + buf + L")";
        }else
        if(mvt == VT_VARIANT){
            s += (std::wstring)L"Variant";
        }else
        {
            s += (std::wstring)L"[Unknown]";
        }

        if(pv->vt & VT_ARRAY){
            wchar_t buf[0xff];
            swprintf(buf, 0xff, L"%u", pv->parray->rgsabound[0].cElements); //### many dimension
            s += (std::wstring)L" Array(" + buf + L")";
            //### show array elements
        }

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocString(s.c_str());

        return S_OK;
    }

    HRESULT vbVarType(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        pVarResult->vt = VT_UI8;
        pVarResult->llVal = pv->vt;

        return S_OK;
    }

    HRESULT vbIsArray(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        pVarResult->vt = VT_BOOL;
        pVarResult->boolVal = (pv->vt & VT_ARRAY) ? VARIANT_TRUE : VARIANT_FALSE;

        return S_OK;
    }

    HRESULT vbIsObject(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        pVarResult->vt = VT_BOOL;
        pVarResult->boolVal = (pv->vt == VT_DISPATCH) ? VARIANT_TRUE : VARIANT_FALSE;

        return S_OK;
    }

    HRESULT vbIsEmpty(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        pVarResult->vt = VT_BOOL;
        pVarResult->boolVal = (pv->vt == VT_EMPTY) ? VARIANT_TRUE : VARIANT_FALSE;

        return S_OK;
    }

    HRESULT vbIsNull(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        pVarResult->vt = VT_BOOL;
        pVarResult->boolVal = (pv->vt == VT_NULL) ? VARIANT_TRUE : VARIANT_FALSE;

        return S_OK;
    }

    HRESULT vbIsDate(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        _variant_t v;
        bool b = SUCCEEDED( VariantChangeType(&v, pv, 0, VT_DATE) );

        pVarResult->vt = VT_BOOL;
        pVarResult->boolVal = b ? VARIANT_TRUE : VARIANT_FALSE;

        return S_OK;
    }

    HRESULT vbIsNumeric(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        _variant_t v1, v2;
        bool b = SUCCEEDED( VariantChangeType(&v1, pv, 0, VT_I8) ) || SUCCEEDED( VariantChangeType(&v2, pv, 0, VT_R8) );

        pVarResult->vt = VT_BOOL;
        pVarResult->boolVal = b ? VARIANT_TRUE : VARIANT_FALSE;

        return S_OK;
    }

    HRESULT vbCBool(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv, 0, VT_BOOL) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbCByte(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv, 0, VT_UI1) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbCCur(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv, 0, VT_CY) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbCDate(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv, 0, VT_DATE) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbCDbl(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv, 0, VT_R8) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbCInt(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv, 0, VT_I2) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbCLng(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv, 0, VT_I4) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbCLnglng(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv, 0, VT_I8) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbCSng(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv, 0, VT_R4) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbCStr(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(!( pDispParams->cArgs == 1 )) return E_INVALIDARG;

        VARIANT* pv = pDispParams->rgvarg;
        if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv, 0, VT_BSTR) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbLBound(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_I8,0,0,0,{1}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;

        if(!( pv1->vt&VT_ARRAY )) return E_INVALIDARG;
        if(!( pv2->vt==VT_I8   )) return E_INVALIDARG;

        long res;
        HRESULT hr;
        if( SUCCEEDED( hr = SafeArrayGetLBound(pv1->parray, pv2->lVal, &res ) ) ){
            pVarResult->vt = VT_I8;
            pVarResult->llVal = res;
            return S_OK;
        }

        return hr;
    }

    HRESULT vbUBound(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_I8,0,0,0,{1}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;

        if(!( pv1->vt&VT_ARRAY )) return E_INVALIDARG;
        if(!( pv2->vt==VT_I8   )) return E_INVALIDARG;

        long res;
        HRESULT hr;
        if( SUCCEEDED( hr = SafeArrayGetUBound(pv1->parray, pv2->lVal, &res ) ) ){
            pVarResult->vt = VT_I8;
            pVarResult->llVal = res;
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbArray(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        SAFEARRAYBOUND sab{pDispParams->cArgs, 0};
        SAFEARRAY* sa = SafeArrayCreate(VT_VARIANT, 1, &sab);

        VARIANT* data;
        SafeArrayAccessData(sa, (void**)&data);
        {
            int i = pDispParams->cArgs;
            while(i--){
                VariantCopy(&data[pDispParams->cArgs-1-i], &pDispParams->rgvarg[i]);
            }
        }
        SafeArrayUnaccessData(sa);

        pVarResult->vt = (VT_ARRAY|VT_VARIANT);
        pVarResult->parray = sa;

        return S_OK;
    }

    HRESULT vbRandomize(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_ERROR || pv1->vt==VT_I8 )) return E_INVALIDARG;

        if(pv1->vt == VT_I8){
            m_rnd.seed(pv1->llVal);
        }else{
            std::random_device rd;
            m_rnd.seed(rd());
        }

        m_rnd_last = (double)m_rnd() / UINT64_MAX;
        
        return S_OK;
    }

    HRESULT vbRnd(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_I8,0,0,0,{1}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_I8 )) return E_INVALIDARG;

        if(pv1->llVal < 0){
            m_rnd.seed(pv1->llVal);
            m_rnd_last = (double)m_rnd() / UINT64_MAX;
        }else
        if(0 < pv1->llVal){
            m_rnd_last = (double)m_rnd() / UINT64_MAX;
        }else
        {
            // none
        }

        pVarResult->vt = VT_R8;
        pVarResult->dblVal = m_rnd_last;

        return S_OK;
    }

    HRESULT vbInt(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(   pv1->vt==VT_ERROR  ) return E_INVALIDARG;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VarInt(pv1, &v) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbFix(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(   pv1->vt==VT_ERROR  ) return E_INVALIDARG;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VarFix(pv1, &v) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbRound(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_I8,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;

        if(   pv1->vt==VT_ERROR  ) return E_INVALIDARG;
        if(!( pv2->vt==VT_I8    )) return E_INVALIDARG;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VarRound(pv1, pv2->llVal, &v) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbSgn(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(   pv1->vt==VT_ERROR  ) return E_INVALIDARG;

        _variant_t v(0LL);
        HRESULT hr;
        if( SUCCEEDED( hr = VarCmp(pv1, &v, 0, 0) ) ){
            if(hr == VARCMP_LT){
                pVarResult->vt = VT_I8;
                pVarResult->llVal = -1;
            }else
            if(hr == VARCMP_GT){
                pVarResult->vt = VT_I8;
                pVarResult->llVal = +1;
            }else
            {
                pVarResult->vt = VT_I8;
                pVarResult->llVal =  0;
            }
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbAbs(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(   pv1->vt==VT_ERROR  ) return E_INVALIDARG;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VarAbs(pv1, &v) ) ){
            *pVarResult = v.Detach();
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbSqr(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(   pv1->vt==VT_ERROR  ) return E_INVALIDARG;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv1, 0, VT_R8) ) ){
            if( v.dblVal < 0.0 ) return E_INVALIDARG;

            pVarResult->vt = VT_R8;
            pVarResult->dblVal = std::sqrt(v.dblVal);
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbSin(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(   pv1->vt==VT_ERROR  ) return E_INVALIDARG;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv1, 0, VT_R8) ) ){
            pVarResult->vt = VT_R8;
            pVarResult->dblVal = std::sin(v.dblVal);
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbCos(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(   pv1->vt==VT_ERROR  ) return E_INVALIDARG;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv1, 0, VT_R8) ) ){
            pVarResult->vt = VT_R8;
            pVarResult->dblVal = std::cos(v.dblVal);
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbTan(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(   pv1->vt==VT_ERROR  ) return E_INVALIDARG;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv1, 0, VT_R8) ) ){
            pVarResult->vt = VT_R8;
            pVarResult->dblVal = std::tan(v.dblVal);
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbAtn(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(   pv1->vt==VT_ERROR  ) return E_INVALIDARG;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv1, 0, VT_R8) ) ){
            pVarResult->vt = VT_R8;
            pVarResult->dblVal = std::atan(v.dblVal);
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbLog(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(   pv1->vt==VT_ERROR  ) return E_INVALIDARG;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv1, 0, VT_R8) ) ){
            pVarResult->vt = VT_R8;
            pVarResult->dblVal = std::log(v.dblVal);
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbExp(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(   pv1->vt==VT_ERROR  ) return E_INVALIDARG;

        _variant_t v;
        HRESULT hr;
        if( SUCCEEDED( hr = VariantChangeType(&v, pv1, 0, VT_R8) ) ){
            pVarResult->vt = VT_R8;
            pVarResult->dblVal = std::exp(v.dblVal);
            return S_OK;
        }
        
        return hr;
    }

    HRESULT vbLen(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;

        pVarResult->vt = VT_I8;
        pVarResult->llVal = SysStringLen(pv1->bstrVal);
        
        return S_OK;
    }

    HRESULT vbLeft(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv2->vt==VT_I8   )) return E_INVALIDARG;

        UINT pos = 0;
        if(SysStringLen(pv1->bstrVal) < pos) pos = 0;
        UINT len = pv2->llVal;
        if(SysStringLen(pv1->bstrVal)-pos < len) len = SysStringLen(pv1->bstrVal)-pos;

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocStringLen(pv1->bstrVal+pos, len);

        return S_OK;
    }

    HRESULT vbRight(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv2->vt==VT_I8   )) return E_INVALIDARG;

        UINT pos = SysStringLen(pv1->bstrVal) - pv2->llVal;
        if(SysStringLen(pv1->bstrVal) < pos) pos = 0;
        UINT len = pv2->llVal;
        if(SysStringLen(pv1->bstrVal)-pos < len) len = SysStringLen(pv1->bstrVal)-pos;

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocStringLen(pv1->bstrVal+pos, len);

        return S_OK;
    }

    HRESULT vbMid(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd3{ {{{VT_I8,0,0,0,{INT64_MAX}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;
        VARIANT* pv3 = (0 <= an-3) ? &pDispParams->rgvarg[an-3] : &vd3;
        if(pv3->vt == (VT_BYREF|VT_VARIANT)) pv3 = pv3->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv2->vt==VT_I8   )) return E_INVALIDARG;
        if(!( pv3->vt==VT_I8   )) return E_INVALIDARG;

        UINT pos = pv2->llVal-1;
        if(SysStringLen(pv1->bstrVal) < pos) pos = SysStringLen(pv1->bstrVal);
        UINT len = pv3->llVal;
        if(SysStringLen(pv1->bstrVal)-pos < len) len = SysStringLen(pv1->bstrVal)-pos;

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocStringLen(pv1->bstrVal+pos, len);

        return S_OK;
    }

    HRESULT vbLCase(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;

        UINT len = SysStringLen(pv1->bstrVal);

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocStringLen(nullptr, len);

        int i=0;
        while(i<len){
            pVarResult->bstrVal[i] = pv1->bstrVal[i];
            if(L'A'<=pVarResult->bstrVal[i] && pVarResult->bstrVal[i]<=L'Z'){
                pVarResult->bstrVal[i] += L'a' - L'A';
            }
            ++i;
        }

        return S_OK;
    }

    HRESULT vbUCase(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;

        UINT len = SysStringLen(pv1->bstrVal);

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocStringLen(nullptr, len);

        int i=0;
        while(i<len){
            pVarResult->bstrVal[i] = pv1->bstrVal[i];
            if(L'a'<=pVarResult->bstrVal[i] && pVarResult->bstrVal[i]<=L'z'){
                pVarResult->bstrVal[i] -= L'a' - L'A';
            }
            ++i;
        }

        return S_OK;
    }

    HRESULT vbLTrim(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;

        UINT len = SysStringLen(pv1->bstrVal);

        int s=0;
        while(s<len && pv1->bstrVal[s]==L' ') ++s;

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocStringLen(pv1->bstrVal+s, len-s);

        return S_OK;
    }

    HRESULT vbRTrim(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;

        UINT len = SysStringLen(pv1->bstrVal);

        int e=len-1;
        while(0<=e && pv1->bstrVal[e]==L' ') --e;

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocStringLen(pv1->bstrVal, e+1);

        return S_OK;
    }

    HRESULT vbTrim(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;

        UINT len = SysStringLen(pv1->bstrVal);

        int s=0;
        while(s<len && pv1->bstrVal[s]==L' ') ++s;
        int e=len-1;
        while(0<=e && pv1->bstrVal[e]==L' ') --e;

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocStringLen(pv1->bstrVal+s, (e+1)-s);

        return S_OK;
    }

    HRESULT vbInStr(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_I8,0,0,0,{1}}}} };
        _variant_t vd2{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd3{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd4{ {{{VT_I8,0,0,0,{0}}}} };

        VARIANT* pv1 = nullptr;
        VARIANT* pv2 = nullptr;
        VARIANT* pv3 = nullptr;
        VARIANT* pv4 = nullptr;

        int an = pDispParams->cArgs;
        if(2 < an){
            pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
            if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
            pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
            if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;
            pv3 = (0 <= an-3) ? &pDispParams->rgvarg[an-3] : &vd3;
            if(pv3->vt == (VT_BYREF|VT_VARIANT)) pv3 = pv3->pvarVal;
            pv4 = (0 <= an-4) ? &pDispParams->rgvarg[an-4] : &vd4;
            if(pv4->vt == (VT_BYREF|VT_VARIANT)) pv4 = pv4->pvarVal;
        }else{
            pv1 = &vd1;
            pv2 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd2;
            if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;
            pv3 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd3;
            if(pv3->vt == (VT_BYREF|VT_VARIANT)) pv3 = pv3->pvarVal;
            pv4 = &vd4;
        }

        if(!( pv1->vt==VT_I8   )) return E_INVALIDARG;
        if(!( pv2->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv3->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv4->vt==VT_I8   )) return E_INVALIDARG;

        int l = SysStringLen(pv3->bstrVal);
        int e = SysStringLen(pv2->bstrVal)-l;
        int i = pv1->llVal-1;
        while(i <= e){
            int cmp = pv4->llVal
                ? _wcsnicmp(pv2->bstrVal+i, pv3->bstrVal, l)
                : std::wcsncmp(pv2->bstrVal+i, pv3->bstrVal, l);
            if(cmp == 0) break;
            ++i;
        }

        pVarResult->vt = VT_I8;
        pVarResult->llVal = (i <= e) ? i+1 : 0;

        return S_OK;
    }

    HRESULT vbInStrRev(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd3{ {{{VT_I8,0,0,0,{-1}}}} };
        _variant_t vd4{ {{{VT_I8,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;
        VARIANT* pv3 = (0 <= an-3) ? &pDispParams->rgvarg[an-3] : &vd3;
        if(pv3->vt == (VT_BYREF|VT_VARIANT)) pv3 = pv3->pvarVal;
        VARIANT* pv4 = (0 <= an-4) ? &pDispParams->rgvarg[an-4] : &vd4;
        if(pv4->vt == (VT_BYREF|VT_VARIANT)) pv4 = pv4->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv2->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv3->vt==VT_I8   )) return E_INVALIDARG;
        if(!( pv4->vt==VT_I8   )) return E_INVALIDARG;

        int l = SysStringLen(pv2->bstrVal);
        int e = 0;
        int i = (0 < pv3->llVal) ? pv3->llVal-1 : SysStringLen(pv1->bstrVal)+pv3->llVal;
        while(e <= i){
            int cmp = pv4->llVal
                ? _wcsnicmp(pv1->bstrVal+i, pv2->bstrVal, l)
                : std::wcsncmp(pv1->bstrVal+i, pv2->bstrVal, l);
            if(cmp == 0) break;
            --i;
        }

        pVarResult->vt = VT_I8;
        pVarResult->llVal = (e <= i) ? i+1 : 0;

        return S_OK;
    }

    HRESULT vbStrComp(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd3{ {{{VT_I8,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;
        VARIANT* pv3 = (0 <= an-3) ? &pDispParams->rgvarg[an-3] : &vd3;
        if(pv3->vt == (VT_BYREF|VT_VARIANT)) pv3 = pv3->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv2->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv3->vt==VT_I8   )) return E_INVALIDARG;

        pVarResult->vt = VT_I8;
        pVarResult->llVal = (pv3->llVal)
            ? _wcsicmp(pv1->bstrVal, pv2->bstrVal)
            : std::wcscmp(pv1->bstrVal, pv2->bstrVal);

        if(pVarResult->llVal) pVarResult->llVal /= std::abs(pVarResult->llVal);

        return S_OK;
    }

    HRESULT vbReplace(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd3{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd4{ {{{VT_I8,0,0,0,{1}}}} };
        _variant_t vd5{ {{{VT_I8,0,0,0,{-1}}}} };
        _variant_t vd6{ {{{VT_I8,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;
        VARIANT* pv3 = (0 <= an-3) ? &pDispParams->rgvarg[an-3] : &vd3;
        if(pv3->vt == (VT_BYREF|VT_VARIANT)) pv3 = pv3->pvarVal;
        VARIANT* pv4 = (0 <= an-4) ? &pDispParams->rgvarg[an-4] : &vd4;
        if(pv4->vt == (VT_BYREF|VT_VARIANT)) pv4 = pv4->pvarVal;
        VARIANT* pv5 = (0 <= an-5) ? &pDispParams->rgvarg[an-5] : &vd5;
        if(pv5->vt == (VT_BYREF|VT_VARIANT)) pv5 = pv5->pvarVal;
        VARIANT* pv6 = (0 <= an-6) ? &pDispParams->rgvarg[an-6] : &vd6;
        if(pv6->vt == (VT_BYREF|VT_VARIANT)) pv6 = pv6->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv2->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv3->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv4->vt==VT_I8   )) return E_INVALIDARG;
        if(!( pv5->vt==VT_I8   )) return E_INVALIDARG;
        if(!( pv6->vt==VT_I8   )) return E_INVALIDARG;

        BSTR src = pv1->bstrVal;
        BSTR tgt = pv2->bstrVal;
        BSTR rpl = pv3->bstrVal;
        int s = pv4->llVal;
        int n = pv5->llVal;
        int (*cmp)(const wchar_t*, const wchar_t*, size_t) = (pv6->llVal) ? _wcsnicmp : std::wcsncmp;

        std::wstring rs;

        int lt = SysStringLen(tgt);
        int e = SysStringLen(src)-lt;
        int i = s-1;
        while( i <= e ){
            if(n && cmp(src+i, tgt, lt) == 0){
                rs += rpl;
                i += lt;
                n -= (n < 0) ? 0 : 1;
            }else{
                rs += src[i];
                i += 1;
            }
        }

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocString(rs.c_str());

        return S_OK;
    }

    HRESULT vbStrReverse(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;

        UINT len = SysStringLen(pv1->bstrVal);
        BSTR s = SysAllocStringLen(nullptr, len);

        int i = 0;
        while(i < len){
            s[i] = pv1->bstrVal[len-1-i];
            ++i;
        }

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = s;

        return S_OK;
    }

    HRESULT vbString(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;

        if(!( pv1->vt==VT_I8 )) return E_INVALIDARG;
        if(!( pv2->vt==VT_I8 || pv2->vt==VT_BSTR )) return E_INVALIDARG;

        UINT len = pv1->llVal;
        BSTR s = SysAllocStringLen(nullptr, len);
        wchar_t c = (pv2->vt == VT_BSTR) ? pv2->bstrVal[0] : pv2->llVal;

        int i = 0;
        while(i < len){
            s[i] = c;
            ++i;
        }

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = s;

        return S_OK;
    }

    HRESULT vbSpace(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_I8 )) return E_INVALIDARG;

        UINT len = pv1->llVal;
        BSTR s = SysAllocStringLen(nullptr, len);
        wchar_t c = L' ';

        int i = 0;
        while(i < len){
            s[i] = c;
            ++i;
        }

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = s;

        return S_OK;
    }

    HRESULT vbChrW(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_I8 )) return E_INVALIDARG;

        BSTR s = SysAllocStringLen(nullptr, 1);
        s[0] = pv1->llVal;

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = s;

        return S_OK;
    }

    HRESULT vbAscW(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;

        pVarResult->vt = VT_I8;
        pVarResult->llVal = pv1->bstrVal[0];

        return S_OK;
    }

    HRESULT vbOct(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_I8 )) return E_INVALIDARG;

        wchar_t buf[22+1];    // log8(2^64) + NULL
        swprintf(buf, sizeof(buf)/sizeof(*buf), L"%llo", pv1->llVal);

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocString(buf);

        return S_OK;
    }

    HRESULT vbHex(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_I8 )) return E_INVALIDARG;

        wchar_t buf[16+1];    // log16(2^64) + NULL
        swprintf(buf, sizeof(buf)/sizeof(*buf), L"%llX", pv1->llVal);

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocString(buf);

        return S_OK;
    }

    HRESULT vbJoin(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_BSTR,0,0,0,{(long long)SysAllocString(L" ")}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;

        if(!( pv1->vt==(VT_ARRAY|VT_VARIANT) && pv1->parray->cDims==1 )) return E_INVALIDARG;
        if(!( pv2->vt==VT_BSTR )) return E_INVALIDARG;

        std::wstring rs;
        HRESULT hr = S_OK;

        VARIANT* data;
        SafeArrayAccessData(pv1->parray, (void**)&data);
        {
            int n = pv1->parray->rgsabound[0].cElements;
            int i = 0;
            while(i < n){
                _variant_t v;
                VARIANT* pvs = &data[i];
                if(pvs->vt!=VT_BSTR){
                    if( FAILED(hr = VariantChangeType(&v, pvs, 0, VT_BSTR)) ) break;
                    pvs = &v;
                }

                if(0 < i) rs += pv2->bstrVal;
                rs += pvs->bstrVal;

                ++i;
            }
        }
        SafeArrayUnaccessData(pv1->parray);

        pVarResult->vt = VT_BSTR;
        pVarResult->bstrVal = SysAllocString(rs.c_str());

        return hr;
    }

    HRESULT vbSplit(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_BSTR,0,0,0,{(long long)SysAllocString(L" ")}}}} };
        _variant_t vd3{ {{{VT_I8,0,0,0,{-1}}}} };
        _variant_t vd4{ {{{VT_I8,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;
        VARIANT* pv3 = (0 <= an-3) ? &pDispParams->rgvarg[an-3] : &vd3;
        if(pv3->vt == (VT_BYREF|VT_VARIANT)) pv3 = pv3->pvarVal;
        VARIANT* pv4 = (0 <= an-4) ? &pDispParams->rgvarg[an-4] : &vd4;
        if(pv4->vt == (VT_BYREF|VT_VARIANT)) pv4 = pv4->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv2->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv3->vt==VT_I8   )) return E_INVALIDARG;
        if(!( pv4->vt==VT_I8   )) return E_INVALIDARG;

        BSTR src = pv1->bstrVal;
        BSTR tgt = pv2->bstrVal;
        int n = pv3->llVal;
        int (*cmp)(const wchar_t*, const wchar_t*, size_t) = (pv4->llVal) ? _wcsnicmp : std::wcsncmp;

        std::vector<BSTR> rv;
        if(n){
            n -= (n < 0) ? 0 : 1;

            int lt = SysStringLen(tgt);
            int e = SysStringLen(src)-lt;
            int s = 0;
            int i = s;
            while( i <= e ){
                if(n && cmp(src+i, tgt, lt) == 0){
                    rv.push_back( SysAllocStringLen(src+s, i-s) );
                    s = i + lt;
                    i = s;
                    n -= (n < 0) ? 0 : 1;
                }else{
                    i += 1;
                }
            }
            rv.push_back( SysAllocStringLen(src+s, i-s) );
        }

        // Make Array
        SAFEARRAYBOUND sab{(ULONG)rv.size(), 0};
        SAFEARRAY* sa = SafeArrayCreate(VT_VARIANT, 1, &sab);

        VARIANT* data;
        SafeArrayAccessData(sa, (void**)&data);
        {
            int i = 0;
            while(i < sab.cElements){
                data[i].vt = VT_BSTR;
                data[i].bstrVal = rv[i];
                rv[i] = nullptr;
                ++i;
            }
        }
        SafeArrayUnaccessData(sa);

        pVarResult->vt = (VT_ARRAY|VT_VARIANT);
        pVarResult->parray = sa;

        return S_OK;
    }

    HRESULT vbFilter(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd3{ {{{VT_BOOL,0,0,0,{VARIANT_TRUE}}}} };
        _variant_t vd4{ {{{VT_I8,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;
        VARIANT* pv3 = (0 <= an-3) ? &pDispParams->rgvarg[an-3] : &vd3;
        if(pv3->vt == (VT_BYREF|VT_VARIANT)) pv3 = pv3->pvarVal;
        VARIANT* pv4 = (0 <= an-4) ? &pDispParams->rgvarg[an-4] : &vd4;
        if(pv4->vt == (VT_BYREF|VT_VARIANT)) pv4 = pv4->pvarVal;

        if(!( pv1->vt==(VT_ARRAY|VT_VARIANT) && pv1->parray->cDims==1 )) return E_INVALIDARG;
        if(!( pv2->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv3->vt==VT_BOOL )) return E_INVALIDARG;
        if(!( pv4->vt==VT_I8   )) return E_INVALIDARG;

        BSTR tgt = pv2->bstrVal;
        int lt = SysStringLen(tgt);
        bool inc = (pv3->boolVal == VARIANT_TRUE);
        int (*cmp)(const wchar_t*, const wchar_t*, size_t) = (pv4->llVal) ? _wcsnicmp : std::wcsncmp;

        std::vector<BSTR> rv;
        HRESULT hr = S_OK;

        VARIANT* data;
        SafeArrayAccessData(pv1->parray, (void**)&data);
        {
            int n = pv1->parray->rgsabound[0].cElements;
            int i = 0;
            while(i < n){
                if(data[i].vt!=VT_BSTR){
                    hr = E_INVALIDARG;
                    break;
                }

                int e = SysStringLen(data[i].bstrVal)-lt;
                int j = 0;
                while(j <= e){
                    if(cmp(data[i].bstrVal+j, tgt, lt) == 0){
                        break;
                    }
                    ++j;
                }

                if( ((j <= e)&&inc) || (!(j <= e)&&!inc) ){
                    rv.push_back( SysAllocString(data[i].bstrVal) );
                }

                ++i;
            }
        }
        SafeArrayUnaccessData(pv1->parray);

        // Make Array
        SAFEARRAYBOUND sab{(ULONG)rv.size(), 0};
        SAFEARRAY* sa = SafeArrayCreate(VT_VARIANT, 1, &sab);

        VARIANT* rdata;
        SafeArrayAccessData(sa, (void**)&rdata);
        {
            int i = 0;
            while(i < sab.cElements){
                rdata[i].vt = VT_BSTR;
                rdata[i].bstrVal = rv[i];
                rv[i] = nullptr;
                ++i;
            }
        }
        SafeArrayUnaccessData(sa);

        pVarResult->vt = (VT_ARRAY|VT_VARIANT);
        pVarResult->parray = sa;

        return hr;
    }

    HRESULT vbNow(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        SYSTEMTIME st;
        GetLocalTime(&st);
        double vt;
        SystemTimeToVariantTime(&st, &vt);

        pVarResult->vt = VT_DATE;
        pVarResult->date = vt;

        return S_OK;
    }

    HRESULT vbDate(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        SYSTEMTIME st;
        GetLocalTime(&st);
        double vt;
        SystemTimeToVariantTime(&st, &vt);

        pVarResult->vt = VT_DATE;
        pVarResult->date = std::floor(vt);

        return S_OK;
    }

    HRESULT vbTime(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        SYSTEMTIME st;
        GetLocalTime(&st);
        double vt;
        SystemTimeToVariantTime(&st, &vt);

        pVarResult->vt = VT_DATE;
        pVarResult->date = vt - std::floor(vt);

        return S_OK;
    }

    HRESULT vbYear(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_DATE )) return E_INVALIDARG;

        SYSTEMTIME st;
        VariantTimeToSystemTime(pv1->date, &st);

        pVarResult->vt = VT_I8;
        pVarResult->llVal = st.wYear;

        return S_OK;
    }

    HRESULT vbMonth(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_DATE )) return E_INVALIDARG;

        SYSTEMTIME st;
        VariantTimeToSystemTime(pv1->date, &st);

        pVarResult->vt = VT_I8;
        pVarResult->llVal = st.wMonth;

        return S_OK;
    }

    HRESULT vbDay(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_DATE )) return E_INVALIDARG;

        SYSTEMTIME st;
        VariantTimeToSystemTime(pv1->date, &st);

        pVarResult->vt = VT_I8;
        pVarResult->llVal = st.wDay;

        return S_OK;
    }

    HRESULT vbWeekday(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_DATE )) return E_INVALIDARG;

        SYSTEMTIME st;
        VariantTimeToSystemTime(pv1->date, &st);

        pVarResult->vt = VT_I8;
        pVarResult->llVal = st.wDayOfWeek+1;

        return S_OK;
    }

    HRESULT vbHour(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_DATE )) return E_INVALIDARG;

        SYSTEMTIME st;
        VariantTimeToSystemTime(pv1->date, &st);

        pVarResult->vt = VT_I8;
        pVarResult->llVal = st.wHour;

        return S_OK;
    }

    HRESULT vbMinute(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_DATE )) return E_INVALIDARG;

        SYSTEMTIME st;
        VariantTimeToSystemTime(pv1->date, &st);

        pVarResult->vt = VT_I8;
        pVarResult->llVal = st.wMinute;

        return S_OK;
    }

    HRESULT vbSecond(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;

        if(!( pv1->vt==VT_DATE )) return E_INVALIDARG;

        SYSTEMTIME st;
        VariantTimeToSystemTime(pv1->date, &st);

        pVarResult->vt = VT_I8;
        pVarResult->llVal = st.wSecond;

        return S_OK;
    }

    HRESULT vbDateAdd(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        //###NEED CHECK
        _variant_t vd1{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd2{ {{{VT_ERROR,0,0,0,{0}}}} };
        _variant_t vd3{ {{{VT_ERROR,0,0,0,{0}}}} };

        int an = pDispParams->cArgs;
        VARIANT* pv1 = (0 <= an-1) ? &pDispParams->rgvarg[an-1] : &vd1;
        if(pv1->vt == (VT_BYREF|VT_VARIANT)) pv1 = pv1->pvarVal;
        VARIANT* pv2 = (0 <= an-2) ? &pDispParams->rgvarg[an-2] : &vd2;
        if(pv2->vt == (VT_BYREF|VT_VARIANT)) pv2 = pv2->pvarVal;
        VARIANT* pv3 = (0 <= an-3) ? &pDispParams->rgvarg[an-3] : &vd3;
        if(pv3->vt == (VT_BYREF|VT_VARIANT)) pv3 = pv3->pvarVal;

        if(!( pv1->vt==VT_BSTR )) return E_INVALIDARG;
        if(!( pv2->vt==VT_I8   )) return E_INVALIDARG;
        if(!( pv3->vt==VT_DATE )) return E_INVALIDARG;

        SYSTEMTIME st;
        VariantTimeToSystemTime(pv3->date, &st);

        istring s(pv1->bstrVal);
        if(s == L"yyyy"){
            st.wYear += pv2->llVal;
        }else
        if(s == L"q"){
            st.wMonth += 3 * pv2->llVal;
        }else
        if(s == L"m"){
            st.wMonth += pv2->llVal;
        }else
        if(s == L"y"){
            st.wDay += pv2->llVal;//#?#
        }else
        if(s == L"d"){
            st.wDay += pv2->llVal;
        }else
        if(s == L"w"){
            st.wDay += pv2->llVal;//#?#
        }else
        if(s == L"ww"){
            st.wDay += 7 * pv2->llVal;
        }else
        if(s == L"h"){
            st.wHour += pv2->llVal;
        }else
        if(s == L"n"){
            st.wMinute += pv2->llVal;
        }else
        if(s == L"s"){
            st.wSecond += pv2->llVal;
        }else
        {
            return E_INVALIDARG;
        }

        double vt;
        SystemTimeToVariantTime(&st, &vt);

        pVarResult->vt = VT_DATE;
        pVarResult->date = vt;

        return S_OK;
    }

    HRESULT vbDateDiff(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
wprintf(L"###%s: Implement here '%s' line %d.\n", __func__, __FILE__, __LINE__);
        return E_NOTIMPL;
    }

    HRESULT vbDatePart(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
wprintf(L"###%s: Implement here '%s' line %d.\n", __func__, __FILE__, __LINE__);
        return E_NOTIMPL;
    }

    HRESULT vbDateSerial(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
wprintf(L"###%s: Implement here '%s' line %d.\n", __func__, __FILE__, __LINE__);
        return E_NOTIMPL;
    }

    HRESULT vbDateValue(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
wprintf(L"###%s: Implement here '%s' line %d.\n", __func__, __FILE__, __LINE__);
        return E_NOTIMPL;
    }

    HRESULT vbTimeSerial(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
wprintf(L"###%s: Implement here '%s' line %d.\n", __func__, __FILE__, __LINE__);
        return E_NOTIMPL;
    }

    HRESULT vbTimeValue(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
wprintf(L"###%s: Implement here '%s' line %d.\n", __func__, __FILE__, __LINE__);
        return E_NOTIMPL;
    }

    HRESULT vbFormatCurrency(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
wprintf(L"###%s: Implement here '%s' line %d.\n", __func__, __FILE__, __LINE__);
        return E_NOTIMPL;
    }

    HRESULT vbFormatDateTime(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
wprintf(L"###%s: Implement here '%s' line %d.\n", __func__, __FILE__, __LINE__);
        return E_NOTIMPL;
    }

    HRESULT vbFormatNumber(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
wprintf(L"###%s: Implement here '%s' line %d.\n", __func__, __FILE__, __LINE__);
        return E_NOTIMPL;
    }

    HRESULT vbFormatPercent(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
wprintf(L"###%s: Implement here '%s' line %d.\n", __func__, __FILE__, __LINE__);
        return E_NOTIMPL;
    }

    HRESULT vbTimer(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        SYSTEMTIME st;
        GetLocalTime(&st);

        pVarResult->vt = VT_R8;
        pVarResult->dblVal = (double)((60*60)*st.wHour+(60)*st.wMinute+st.wSecond) + ((double)st.wMilliseconds/1000);

        return S_OK;
    }

};

int VBScript::s_disps_n = 0;
std::map<istring, DISPID> VBScript::s_disps_ids{
    {L"",                       s_disps_n++},
    // consts
    {L"vbBinaryCompare",        s_disps_n++},
    {L"vbTextCompare",          s_disps_n++},
    {L"vbSunday",               s_disps_n++},
    {L"vbMonday",               s_disps_n++},
    {L"vbTuesday",              s_disps_n++},
    {L"vbWednesday",            s_disps_n++},
    {L"vbThursday",             s_disps_n++},
    {L"vbFriday",               s_disps_n++},
    {L"vbSaturday",             s_disps_n++},
    {L"vbUseSystemDayOfWeek",   s_disps_n++},
    {L"vbFirstJan1",            s_disps_n++},
    {L"vbFirstFourDays",        s_disps_n++},
    {L"vbFirstFullWeek",        s_disps_n++},
    {L"vbGeneralDate",          s_disps_n++},
    {L"vbLongDate",             s_disps_n++},
    {L"vbShortDate",            s_disps_n++},
    {L"vbLongTime",             s_disps_n++},
    {L"vbShortTime",            s_disps_n++},
    {L"vbObjectError",          s_disps_n++},
    {L"vbUseDefault",           s_disps_n++},
    {L"vbTrue",                 s_disps_n++},
    {L"vbFalse",                s_disps_n++},
    {L"vbEmpty",                s_disps_n++},
    {L"vbNull",                 s_disps_n++},
    {L"vbInteger",              s_disps_n++},
    {L"vbLong",                 s_disps_n++},
    {L"vbSingle",               s_disps_n++},
    {L"vbDouble",               s_disps_n++},
    {L"vbCurrency",             s_disps_n++},
    {L"vbDate",                 s_disps_n++},
    {L"vbString",               s_disps_n++},
    {L"vbObject",               s_disps_n++},
    {L"vbError",                s_disps_n++},
    {L"vbBoolean",              s_disps_n++},
    {L"vbVariant",              s_disps_n++},
    {L"vbDataObject",           s_disps_n++},
    {L"vbDecimal",              s_disps_n++},
    {L"vbByte",                 s_disps_n++},
    {L"vbArray",                s_disps_n++},
    {L"vbCr",                   s_disps_n++},
    {L"VbCrLf",                 s_disps_n++},
    {L"vbFormFeed",             s_disps_n++},
    {L"vbLf",                   s_disps_n++},
    {L"vbNewLine",              s_disps_n++},
    {L"vbNullChar",             s_disps_n++},
    {L"vbNullString",           s_disps_n++},
    {L"vbTab",                  s_disps_n++},
    {L"vbVerticalTab",          s_disps_n++},
    // functions
    {L"CreateObject",           s_disps_n++},
    {L"GetObject",              s_disps_n++},
    {L"TypeName",               s_disps_n++},
    {L"VarDump",                s_disps_n++},
    {L"VarType",                s_disps_n++},
    {L"IsArray",                s_disps_n++},
    {L"IsObject",               s_disps_n++},
    {L"IsEmpty",                s_disps_n++},
    {L"IsNull",                 s_disps_n++},
    {L"IsDate",                 s_disps_n++},
    {L"IsNumeric",              s_disps_n++},
    {L"CBool",                  s_disps_n++},
    {L"CByte",                  s_disps_n++},
    {L"CCur",                   s_disps_n++},
    {L"CDate",                  s_disps_n++},
    {L"CDbl",                   s_disps_n++},
    {L"CInt",                   s_disps_n++},
    {L"CLng",                   s_disps_n++},
    {L"CLnglng",                s_disps_n++},
    {L"CSng",                   s_disps_n++},
    {L"CStr",                   s_disps_n++},
    {L"LBound",                 s_disps_n++},
    {L"UBound",                 s_disps_n++},
    {L"Array",                  s_disps_n++},
    {L"Randomize",              s_disps_n++},
    {L"Rnd",                    s_disps_n++},
    {L"Int",                    s_disps_n++},
    {L"Fix",                    s_disps_n++},
    {L"Round",                  s_disps_n++},
    {L"Sgn",                    s_disps_n++},
    {L"Abs",                    s_disps_n++},
    {L"Sqr",                    s_disps_n++},
    {L"Sin",                    s_disps_n++},
    {L"Cos",                    s_disps_n++},
    {L"Tan",                    s_disps_n++},
    {L"Atn",                    s_disps_n++},
    {L"Log",                    s_disps_n++},
    {L"Exp",                    s_disps_n++},
    {L"Len",                    s_disps_n++},
    {L"Left",                   s_disps_n++},
    {L"Right",                  s_disps_n++},
    {L"Mid",                    s_disps_n++},
    {L"LCase",                  s_disps_n++},
    {L"UCase",                  s_disps_n++},
    {L"LTrim",                  s_disps_n++},
    {L"RTrim",                  s_disps_n++},
    {L"Trim",                   s_disps_n++},
    {L"InStr",                  s_disps_n++},
    {L"InStrRev",               s_disps_n++},
    {L"StrComp",                s_disps_n++},
    {L"Replace",                s_disps_n++},
    {L"StrReverse",             s_disps_n++},
    {L"String",                 s_disps_n++},
    {L"Space",                  s_disps_n++},
    {L"ChrW",                   s_disps_n++},
    {L"AscW",                   s_disps_n++},
    {L"Oct",                    s_disps_n++},
    {L"Hex",                    s_disps_n++},
    {L"Join",                   s_disps_n++},
    {L"Split",                  s_disps_n++},
    {L"Filter",                 s_disps_n++},
    {L"Now",                    s_disps_n++},
    {L"Date",                   s_disps_n++},
    {L"Time",                   s_disps_n++},
    {L"Year",                   s_disps_n++},
    {L"Month",                  s_disps_n++},
    {L"Day",                    s_disps_n++},
    {L"Weekday",                s_disps_n++},
    {L"Hour",                   s_disps_n++},
    {L"Minute",                 s_disps_n++},
    {L"Second",                 s_disps_n++},
    {L"Timer",                  s_disps_n++},
    {L"DateAdd",                s_disps_n++},
    {L"DateDiff",               s_disps_n++},
    {L"DatePart",               s_disps_n++},
    {L"DateSerial",             s_disps_n++},
    {L"DateValue",              s_disps_n++},
    {L"TimeSerial",             s_disps_n++},
    {L"TimeValue",              s_disps_n++},
    {L"FormatCurrency",         s_disps_n++},
    {L"FormatDateTime",         s_disps_n++},
    {L"FormatNumber",           s_disps_n++},
    {L"FormatPercent",          s_disps_n++},
};

std::vector<_variant_t> VBScript::s_disps{
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::_}).v}}}} },
    // consts
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)0}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)1}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)1}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)2}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)3}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)4}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)5}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)6}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)7}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)0}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)1}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)2}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)3}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)0}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)1}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)2}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)3}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)4}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)0x80040000}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)-2}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)-1}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)0}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)0}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)1}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)2}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)3}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)4}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)5}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)6}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)7}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)8}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)9}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)10}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)11}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)12}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)13}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)14}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)17}}}} },
    _variant_t{ {{{VT_I8  ,0,0,0,{(long long)0x2000}}}} },
    _variant_t{ {{{VT_BSTR,0,0,0,{(long long)SysAllocString(L"\x0D")}}}} },
    _variant_t{ {{{VT_BSTR,0,0,0,{(long long)SysAllocString(L"\x0D\x0A")}}}} },
    _variant_t{ {{{VT_BSTR,0,0,0,{(long long)SysAllocString(L"\x0C")}}}} },
    _variant_t{ {{{VT_BSTR,0,0,0,{(long long)SysAllocString(L"\x0A")}}}} },
    _variant_t{ {{{VT_BSTR,0,0,0,{(long long)SysAllocString(L"\x0A")}}}} },
    _variant_t{ {{{VT_BSTR,0,0,0,{(long long)SysAllocString(L"\x00")}}}} },
    _variant_t{ {{{VT_BSTR,0,0,0,{(long long)SysAllocString(nullptr)}}}} },
    _variant_t{ {{{VT_BSTR,0,0,0,{(long long)SysAllocString(L"\x09")}}}} },
    _variant_t{ {{{VT_BSTR,0,0,0,{(long long)SysAllocString(L"\x0b")}}}} },
    // functions
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbCreateObject}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbGetObject}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbTypeName}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbVarDump}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbVarType}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbIsArray}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbIsObject}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbIsEmpty}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbIsNull}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbIsDate}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbIsNumeric}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbCBool}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbCByte}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbCCur}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbCDate}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbCDbl}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbCInt}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbCLng}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbCLnglng}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbCSng}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbCStr}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbLBound}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbUBound}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbArray}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbRandomize}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbRnd}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbInt}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbFix}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbRound}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbSgn}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbAbs}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbSqr}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbSin}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbCos}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbTan}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbAtn}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbLog}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbExp}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbLen}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbLeft}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbRight}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbMid}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbLCase}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbUCase}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbLTrim}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbRTrim}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbTrim}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbInStr}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbInStrRev}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbStrComp}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbReplace}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbStrReverse}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbString}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbSpace}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbChrW}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbAscW}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbOct}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbHex}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbJoin}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbSplit}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbFilter}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbNow}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbDate}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbTime}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbYear}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbMonth}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbDay}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbWeekday}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbHour}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbMinute}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbSecond}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbTimer}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbDateAdd}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbDateDiff}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbDatePart}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbDateSerial}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbDateValue}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbTimeSerial}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbTimeValue}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbFormatCurrency}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbFormatDateTime}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbFormatNumber}).v}}}} },
    _variant_t{ {{{VT_EMPTY,0,0,0,{(long long)(invoke_cast{&VBScript::vbFormatPercent}).v}}}} },
};



class WScript : public IDispatch{
private:
    ULONG       m_refc   = 1;

public:
    WScript(){}
    virtual ~WScript(){}

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo){ return E_NOTIMPL; }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, UINT cNames, LCID lcid, DISPID *rgDispId){
        if(_wcsicmp(*rgszNames, L"echo") == 0){
            *rgDispId = 1;
        }else{
wprintf(L"###%s: Implement here '%s' line %d. (%ls)\n", __func__, __FILE__, __LINE__, *rgszNames);
            return DISP_E_MEMBERNOTFOUND;
        }

        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, 
        DISPPARAMS *pDispParams, VARIANT *pVarResult, EXCEPINFO *pExcepInfo, UINT *puArgErr)
    {
        if(dispIdMember == 1){
            UINT i = pDispParams->cArgs;
            while(--i < pDispParams->cArgs){
                VARIANT* pv = pDispParams->rgvarg + i;
                if(pv->vt == (VT_BYREF|VT_VARIANT)) pv = pv->pvarVal;

                _variant_t v;
                if(::VariantChangeType(&v, pv, 0, VT_BSTR) == S_OK){
                    wprintf(L"%ls%ls", v.bstrVal, i ? L" " : L"\n");
                }else{
                    wprintf(L"VARIANT:0x%x[0x%llx]\n", v.vt, v.llVal);
                }
            }
        }else{
            return DISP_E_MEMBERNOTFOUND;
        }

        return S_OK;
    }
};









class CFVBScript : public IClassFactory{
private:
    ULONG       m_refc   = 1;

public:
    CFVBScript(){}
    virtual ~CFVBScript(){}

public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject){ return E_NOTIMPL; }
    ULONG STDMETHODCALLTYPE AddRef(){ return ++m_refc; }
    ULONG STDMETHODCALLTYPE Release(){ if(!--m_refc){ delete this; return 0; } return m_refc; }

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject){
        *ppvObject = new VBScript;
        return S_OK;
    }
    
    HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock){
        m_refc += fLock ? 1 : -1;
        return S_OK;
    }
} g_oCFVBScript;



extern "C"
HRESULT DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv){
    if(wcscmp((const wchar_t*)*ppv, L"VBScript") == 0){
        g_oCFVBScript.AddRef();
        *ppv = (IClassFactory*)&g_oCFVBScript;
    }else
    {
        return CO_E_CLASSSTRING;
    }

    return S_OK;
}









int main(int argn, const char* argc[]){
#ifdef _WIN32
    setlocale(LC_ALL, ".65001");
#else
    setlocale(LC_ALL, "");
    
    if(!getenv("REGISTRY")){
        char buf[256];
        strncpy(buf, argc[0], 256)[255] = '\0';
        int i = strlen(buf);
        while(0 <= i && buf[i] != '/') --i;
        buf[i+1] = '\0';
        setenv("REGISTRY", buf, 0);
    }
#endif

    #define MAX_SIZE 0x8000
    wchar_t aSource[MAX_SIZE+1];
    const wchar_t* pSource = nullptr;

    // load source
    if(1 < argn){
#ifdef _WIN32
        FILE* pf = fopen(argc[1], "rt,ccs=UTF-8");
        if(pf){
            {
                aSource[MAX_SIZE] = 0xffff;
                size_t read = fread(aSource, sizeof(aSource[0]), sizeof(aSource)/sizeof(aSource[0])-1, pf);
                aSource[read] = L'\0';
            }

            // check size
            if(!aSource[MAX_SIZE]){ fwprintf(stderr, L"!Max source size is %zd byte.\n", sizeof(aSource)); return -1; };

            pSource = aSource;
            fclose(pf);
        }
#else
        FILE* pf = fopen(argc[1], "r");
        if(pf){
            char aBuf[MAX_SIZE+1];{
                aBuf[MAX_SIZE] = 0xff;
                size_t read = fread(aBuf, sizeof(aBuf[0]), sizeof(aBuf)/sizeof(aBuf[0])-1, pf);
                aBuf[read] = 0x00;
            }

            // check size
            if(!aBuf[MAX_SIZE]){ fwprintf(stderr, L"!Max source size is %zd byte.\n", sizeof(aBuf)); return -1; };

            if(!utf8_wchar(aSource, sizeof(aSource)/sizeof(aSource[0]), aBuf)){ fwprintf(stderr, L"!Need UTF8 format.\n"); return -1; }

            pSource = aSource;
            fclose(pf);
        }
#endif
    }

    if(!pSource){
        char aBuf[MAX_SIZE+1];{
            aBuf[MAX_SIZE] = 0xff;
            size_t read = fread(aBuf, sizeof(aBuf[0]), sizeof(aBuf)/sizeof(aBuf[0])-1, stdin);
            aBuf[read] = 0x00;
        }

        // check size
        if(!aBuf[MAX_SIZE]){ fwprintf(stderr, L"!Max source size is %zd byte.\n", sizeof(aBuf)); return -1; };

        if(!utf8_wchar(aSource, sizeof(aSource)/sizeof(aSource[0]), aBuf)){ fwprintf(stderr, L"!Need UTF8 format.\n"); return -1; }

        pSource = aSource;
    }



    // run
    VBScript oVBScript;
    WScript  oWScript;

    CExtension oExt = {
        { L"wscript", (IDispatch*)&oWScript },
    };

    CProgram oProgram(pSource);

    HRESULT hr = S_OK;    
    CoInitializeEx(0, COINIT_APARTMENTTHREADED);
    {
        // CProgram copy(oProgram);//### error, check it.
        CProcessor oProcessor(&oProgram, &oVBScript, &oExt);

        // oProcessor.m_cc = 10000;
        hr = oProcessor();

        if(FAILED(hr)){
            fwprintf(stderr, L"![0x%x]%ls in line:%d\n", hr, oProcessor.m_err->bstrDescription, oProcessor.m_err->dwHelpContext);
        }
    }
    CoUninitialize();

    return SUCCEEDED(hr) ? 0 : 1;
}


