#ifdef _WIN32
    #include <comdef.h>
#else
    #include <stdint.h>
    #include <stddef.h>
    #include <memory.h>
    #include <wchar.h>
    #include <dlfcn.h>

    #define _wcsnicmp wcsncasecmp
    #define _wcsicmp wcscasecmp
    #define STDMETHODCALLTYPE 

    #define S_OK             0x00000000
    #define S_FALSE          0x00000001
    #define E_NOTIMPL        0x80004001
    #define E_FAIL           0x80004005
    #define E_INVALIDARG     0x80070057
    #define E_NOINTERFACE    0x80004002
    #define CO_E_CLASSSTRING 0x800401F3
    #define DISP_E_MEMBERNOTFOUND 0x80020003
    #define DISP_E_BADPARAMCOUNT  0x8002000E
    #define DISP_E_DIVBYZERO      0x80020012

    #define	DISPID_VALUE        ( 0)
    #define DISPID_UNKNOWN      (-1)
    #define	DISPID_PROPERTYPUT	(-3)
    #define	DISPID_NEWENUM      (-4)
    #define DISPATCH_METHOD         0x1
    #define DISPATCH_PROPERTYGET    0x2
    #define DISPATCH_PROPERTYPUT    0x4

    #define VARIANT_TRUE ((VARIANT_BOOL)-1)
    #define VARIANT_FALSE ((VARIANT_BOOL)0)

    #define SUCCEEDED(hr) (0 <= ((HRESULT)(hr)))
    #define FAILED(hr)    (((HRESULT)(hr)) <  0)

    #define VARCMP_LT   0
    #define VARCMP_EQ   1
    #define VARCMP_GT   2
    #define VARCMP_NULL 3

    #define TRUE  1
    #define FALSE 0

    #define CLSCTX_INPROC_SERVER 1
    #define CLSCTX_LOCAL_SERVER  4
    #define COINIT_MULTITHREADED     0x0
    #define COINIT_APARTMENTTHREADED 0x2

    struct GUID{
        unsigned long   Data1;
        unsigned short  Data2;
        unsigned short  Data3;
        unsigned char   Data4[8];
    };

    typedef int32_t     HRESULT;
    typedef uint32_t    ULONG;
    typedef uint32_t    UINT;
    typedef uint32_t    LCID;
    typedef wchar_t     OLECHAR;
    typedef wchar_t*    LPOLESTR;
    typedef const wchar_t* LPCOLESTR;
    typedef int32_t     DISPID;
    typedef uint16_t    WORD;
    typedef uint32_t    DWORD;
    typedef wchar_t*    BSTR;
    typedef int32_t     SCODE;
    typedef void*       PVOID;
    typedef GUID        IID;
    typedef GUID        CLSID;
    typedef CLSID*      LPCLSID;
    typedef const CLSID& REFCLSID;
    typedef const IID&   REFIID;
    typedef uint16_t    VARTYPE;
    typedef uint16_t    USHORT;
    typedef uint64_t    ULONGLONG;
    typedef uint8_t     BYTE;
    typedef int64_t     LONGLONG;
    typedef double      DOUBLE;
    typedef double      DATE;
    typedef long        LONG;
    typedef int         INT;
    typedef void*       LPVOID;
    typedef const wchar_t* LPCWSTR;
    typedef const char*    LPCSTR;
    typedef short       VARIANT_BOOL;
    typedef int         BOOL;
    typedef void*       HMODULE;
    typedef void*       (*PROC)();
    typedef HRESULT (*LPFNGETCLASSOBJECT)(REFCLSID, REFIID, LPVOID *);

    extern const IID IID_NULL;
    extern const CLSID CLSID_NULL;

    struct IUnknown{
        virtual HRESULT STDMETHODCALLTYPE QueryInterface(REFIID, void**) = 0;
        virtual ULONG STDMETHODCALLTYPE AddRef() = 0;
        virtual ULONG STDMETHODCALLTYPE Release() = 0;
    };
    typedef IUnknown* LPUNKNOWN;

    extern const IID IID_IClassFactory;
    struct IClassFactory : public IUnknown{
        virtual HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppvObject) = 0;
        virtual HRESULT STDMETHODCALLTYPE LockServer(BOOL fLock) = 0;
    };

    struct SAFEARRAYBOUND{
        ULONG cElements;
        LONG lLbound;
    };

    #define FADF_VARIANT        ( 0x800 )

    struct SAFEARRAY{
        USHORT cDims;
        USHORT fFeatures;
        ULONG cbElements;
        ULONG cLocks;
        PVOID pvData;
        SAFEARRAYBOUND rgsabound[1];
    };

    struct SYSTEMTIME{
        WORD wYear;
        WORD wMonth;
        WORD wDayOfWeek;
        WORD wDay;
        WORD wHour;
        WORD wMinute;
        WORD wSecond;
        WORD wMilliseconds;
    };
    typedef SYSTEMTIME* LPSYSTEMTIME;

    struct BIND_OPTS{
        DWORD cbStruct;
        DWORD grfFlags;
        DWORD grfMode;
        DWORD dwTickCountDeadline;
    };

    enum VARENUM{
        VT_EMPTY	= 0,
        VT_NULL	= 1,
        VT_I2	= 2,
        VT_I4	= 3,
        VT_R4	= 4,
        VT_R8	= 5,
        VT_CY	= 6,
        VT_DATE	= 7,
        VT_BSTR	= 8,
        VT_DISPATCH	= 9,
        VT_ERROR	= 10,
        VT_BOOL	= 11,
        VT_VARIANT	= 12,
        VT_UNKNOWN	= 13,
        VT_DECIMAL	= 14,
        VT_I1	= 16,
        VT_UI1	= 17,
        VT_UI2	= 18,
        VT_UI4	= 19,
        VT_I8	= 20,
        VT_UI8	= 21,
        // VT_INT	= 22,
        // VT_UINT	= 23,
        // VT_VOID	= 24,
        // VT_HRESULT	= 25,
        // VT_PTR	= 26,
        // VT_SAFEARRAY	= 27,
        // VT_CARRAY	= 28,
        // VT_USERDEFINED	= 29,
        // VT_LPSTR	= 30,
        // VT_LPWSTR	= 31,
        // VT_RECORD	= 36,
        // VT_INT_PTR	= 37,
        // VT_UINT_PTR	= 38,
        // VT_FILETIME	= 64,
        // VT_BLOB	= 65,
        // VT_STREAM	= 66,
        // VT_STORAGE	= 67,
        // VT_STREAMED_OBJECT	= 68,
        // VT_STORED_OBJECT	= 69,
        // VT_BLOB_OBJECT	= 70,
        // VT_CF	= 71,
        // VT_CLSID	= 72,
        // VT_VERSIONED_STREAM	= 73,
        // VT_BSTR_BLOB	= 0xfff,
        // VT_VECTOR	= 0x1000,
        VT_ARRAY	= 0x2000,
        VT_BYREF	= 0x4000,
        // VT_RESERVED	= 0x8000,
        // VT_ILLEGAL	= 0xffff,
        // VT_ILLEGALMASKED	= 0xfff,
        VT_TYPEMASK	= 0xfff
    };

    struct DECIMAL{
        USHORT wReserved;
        union {
            struct {
                BYTE scale;
                BYTE sign;
            } DUMMYSTRUCTNAME;
            USHORT signscale;
        } DUMMYUNIONNAME;
        ULONG Hi32;
        union {
            struct {
                ULONG Lo32;
                ULONG Mid32;
            } DUMMYSTRUCTNAME2;
            ULONGLONG Lo64;
        } DUMMYUNIONNAME2;
    };

    struct IDispatch;

    struct VARIANT{
        union{
            struct{
                VARTYPE vt;
                WORD wReserved1;
                WORD wReserved2;
                WORD wReserved3;
                union{
                    LONGLONG llVal;
                    LONG lVal;
                    BYTE bVal;
                    // SHORT iVal;
                    // FLOAT fltVal;
                    DOUBLE dblVal;
                    VARIANT_BOOL boolVal;
                    // VARIANT_BOOL __OBSOLETE__VARIANT_BOOL;
                    // SCODE scode;
                    // CY cyVal;
                    DATE date;
                    BSTR bstrVal;
                    IUnknown *punkVal;
                    IDispatch *pdispVal;
                    SAFEARRAY *parray;
                    // BYTE *pbVal;
                    // SHORT *piVal;
                    // LONG *plVal;
                    // LONGLONG *pllVal;
                    // FLOAT *pfltVal;
                    // DOUBLE *pdblVal;
                    // VARIANT_BOOL *pboolVal;
                    // VARIANT_BOOL *__OBSOLETE__VARIANT_PBOOL;
                    // SCODE *pscode;
                    // CY *pcyVal;
                    // DATE *pdate;
                    // BSTR *pbstrVal;
                    // IUnknown **ppunkVal;
                    // IDispatch **ppdispVal;
                    // SAFEARRAY **pparray;
                    VARIANT *pvarVal;
                    PVOID byref;
                    // CHAR cVal;
                    USHORT uiVal;
                    // ULONG ulVal;
                    ULONGLONG ullVal;
                    // INT intVal;
                    // UINT uintVal;
                    // DECIMAL *pdecVal;
                    // CHAR *pcVal;
                    // USHORT *puiVal;
                    // ULONG *pulVal;
                    // ULONGLONG *pullVal;
                    // INT *pintVal;
                    // UINT *puintVal;
                    // struct{
                    //     PVOID pvRecord;
                    //     IRecordInfo *pRecInfo;
                    // };
                };
            };
            DECIMAL decVal;
        };
    };
    typedef VARIANT* LPVARIANT;

    struct DISPPARAMS{
        VARIANT *rgvarg;
        DISPID  *rgdispidNamedArgs;
        UINT    cArgs;
        UINT    cNamedArgs;
    };

    struct EXCEPINFO{
        WORD  wCode;
        WORD  wReserved;
        BSTR  bstrSource;
        BSTR  bstrDescription;
        BSTR  bstrHelpFile;
        DWORD dwHelpContext;
        PVOID pvReserved;
        HRESULT (*pfnDeferredFillIn)(EXCEPINFO*);
        SCODE scode;
    };

    struct ITypeInfo;
    extern const IID IID_IDispatch;
    struct IDispatch : public IUnknown{
        virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT*) = 0;
        virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo**) = 0;
        virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) = 0;
        virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) = 0;
    };

    extern const IID IID_IEnumVARIANT;
    struct IEnumVARIANT : public IUnknown{
        virtual HRESULT STDMETHODCALLTYPE Next(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched) = 0;
        virtual HRESULT STDMETHODCALLTYPE Skip(ULONG celt) = 0;
        virtual HRESULT STDMETHODCALLTYPE Reset() = 0;
        virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT **ppEnum) = 0;
    };

    // system
    HMODULE LoadLibraryW(LPOLESTR lpszLibName);
    HMODULE GetModuleHandleW(LPCWSTR lpModuleName);
    PROC GetProcAddress(HMODULE hModule, LPCSTR lpProcName);
    void GetLocalTime(LPSYSTEMTIME lpSystemTime);

    // ole
    HRESULT CLSIDFromProgID(LPCOLESTR lpszProgID, LPCLSID lpclsid);
    HRESULT CLSIDFromString(LPCOLESTR lpsz, LPCLSID pclsid);
    HRESULT CoCreateInstance(REFCLSID rclsid, LPUNKNOWN pUnkOuter, DWORD dwClsContext, REFIID riid, LPVOID *ppv);
    HRESULT CoGetObject(LPCWSTR pszName, BIND_OPTS *pBindOptions, REFIID riid, void **ppv);
    HRESULT CoInitializeEx(LPVOID pvReserved, DWORD dwCoInit);
    void CoUninitialize();

    // oleaut
    HRESULT VarNot(LPVARIANT pvarIn, LPVARIANT pvarResult);
    HRESULT VarAnd(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
    HRESULT VarOr(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
    HRESULT VarXor(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
    HRESULT VarCmp(LPVARIANT pvarLeft, LPVARIANT pvarRight, LCID lcid, ULONG dwFlags);
    HRESULT VarRound(LPVARIANT pvarIn, int cDecimals, LPVARIANT pvarResult);
    HRESULT VarInt(LPVARIANT pvarIn, LPVARIANT pvarResult);
    HRESULT VarFix(LPVARIANT pvarIn, LPVARIANT pvarResult);
    HRESULT VarAbs(LPVARIANT pvarIn, LPVARIANT pvarResult);
    HRESULT VarPow(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
    HRESULT VarMul(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
    HRESULT VarMod(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
    HRESULT VarDiv(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
    HRESULT VarAdd(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
    HRESULT VarSub(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
    HRESULT VarCat(LPVARIANT pvarLeft, LPVARIANT pvarRight, LPVARIANT pvarResult);
    HRESULT VarDateFromStr(LPCOLESTR strIn, LCID lcid, ULONG dwFlags, DATE *pdateOut);
    HRESULT VariantChangeType(VARIANT *pvargDest, const VARIANT *pvarSrc, USHORT wFlags, VARTYPE vt);
    HRESULT VariantCopy(VARIANT *pvargDest, const VARIANT *pvargSrc);
    HRESULT VariantClear(VARIANT *pvarg);
    void VariantInit(VARIANT *pvarg);
    BSTR SysAllocString(const OLECHAR *psz);
    BSTR SysAllocStringLen(const OLECHAR *strIn, UINT ui);
    UINT SysStringLen(BSTR pbstr);
    void SysFreeString(BSTR bstrString);
    SAFEARRAY* SafeArrayCreate(VARTYPE vt, UINT cDims, SAFEARRAYBOUND *rgsabound);
    HRESULT SafeArrayCopy(SAFEARRAY *psa, SAFEARRAY **ppsaOut);
    HRESULT SafeArrayDestroy(SAFEARRAY *psa);
    HRESULT SafeArrayGetUBound(SAFEARRAY *psa, UINT nDim, LONG *plUbound);
    HRESULT SafeArrayGetLBound(SAFEARRAY *psa, UINT nDim, LONG *plLbound);
    HRESULT SafeArrayAccessData(SAFEARRAY  *psa, void **ppvData);
    HRESULT SafeArrayUnaccessData(SAFEARRAY *psa);
    INT SystemTimeToVariantTime(LPSYSTEMTIME lpSystemTime, DOUBLE *pvtime);
    INT VariantTimeToSystemTime(DOUBLE vtime, LPSYSTEMTIME lpSystemTime);

    // others
    double tm_double(tm t);
    tm double_tm(double d);

    // utility
    struct _com_error{
        HRESULT m_hr;
        _com_error(HRESULT hr) : m_hr(hr){}
        HRESULT Error(){ return m_hr; };
    };

    namespace _com_util{
        inline void CheckError(HRESULT hr){
            if(FAILED(hr)) throw _com_error(hr);
        }
    }

    struct _bstr_t{
        _bstr_t() : m_s(nullptr){
        }
        _bstr_t(const wchar_t* s){
            m_s = SysAllocString(s);
        }
        operator const wchar_t*(){
            return m_s;
        }
        ~_bstr_t(){
            if(m_s) SysFreeString(m_s);
        }
    protected:
        BSTR m_s;
    };

    struct _variant_t : public VARIANT{
        _variant_t(){
            VariantInit(this);
        }
        _variant_t(long long r){
            vt = VT_I8;
            llVal = r;
        }
        _variant_t(bool r){
            vt = VT_BOOL;
            boolVal = r ? VARIANT_TRUE : VARIANT_FALSE;
        }
        _variant_t(IDispatch* pdisp){
            vt = VT_DISPATCH;
            pdispVal = pdisp;
            if(pdispVal) pdispVal->AddRef();
        }
        _variant_t(const VARIANT& r){
            vt = VT_EMPTY;
            VariantCopy(this, &r);
        }
        _variant_t(VARIANT&& r){
            *(VARIANT*)this = r;
            r.vt = VT_EMPTY;
        }
        _variant_t(const _variant_t& r){
            vt = VT_EMPTY;
            VariantCopy(this, &r);
        }
        _variant_t(_variant_t&& r){
            *(VARIANT*)this = r;
            r.vt = VT_EMPTY;
        }
        _variant_t& operator=(const long long r){
            VariantClear(this);
            vt = VT_I8;
            llVal = r;
            return *this;
        }
        _variant_t& operator=(const BSTR r){
            VariantClear(this);
            vt = VT_BSTR;
            bstrVal = SysAllocString(r);
            return *this;
        }
        _variant_t& operator=(const VARIANT& r){
            VariantCopy(this, &r);
            return *this;
        }
        _variant_t& operator=(const _variant_t& r){
            VariantCopy(this, &r);
            return *this;
        }
        bool operator<(const _variant_t& r) const{
            return (VarCmp((LPVARIANT)this, (LPVARIANT)&r, 0, 0) == VARCMP_LT);
        }
        void Attach(VARIANT& r){
            VariantClear(this);
            *(VARIANT*)this = r;
            r.vt = VT_EMPTY;
        }
        VARIANT Detach(){
            VARIANT ret = *(VARIANT*)this;
            this->vt = VT_EMPTY;
            return ret;
        }
        operator bool(){
            if(vt == VT_BOOL){
                return (boolVal == VARIANT_TRUE);
            }else{
                _variant_t v;
                _com_util::CheckError( VariantChangeType(&v, this, 0, VT_BOOL) );
                return (v.boolVal == VARIANT_TRUE);
            }
        }
        ~_variant_t(){
            VariantClear(this);
        }
    };

    template<typename TI, const IID* pIID>
    struct _com_IIID{
        typedef TI I;
        static const IID* sIID(){ return pIID; }
    };

    template<typename TII>
    class _com_ptr_t{
    public:
        typedef typename TII::I I;
        _com_ptr_t() : m_p(nullptr){
            // none
        }
        _com_ptr_t(const _com_ptr_t<TII>& r) : m_p(r.m_p){
            if(m_p) m_p->AddRef();
        }
        _com_ptr_t(_com_ptr_t<TII>&& r) : m_p(r.m_p){
            r.m_p = nullptr;
        }
        _com_ptr_t(I* p, bool addref=true) : m_p(p){
            if(m_p && addref) m_p->AddRef();
        }
        _com_ptr_t(const _variant_t& r) : m_p(nullptr){
            if(r.vt == VT_UNKNOWN){
                r.pdispVal->QueryInterface(*TII::sIID(), (void**)&m_p);
            }else
            if(r.vt == VT_DISPATCH){
                if( memcmp(TII::sIID(), &IID_IDispatch, sizeof(GUID)) == 0 ){
                    m_p = (I*)r.pdispVal;
                    if(m_p) m_p->AddRef();
                }else{
                    r.pdispVal->QueryInterface(*TII::sIID(), (void**)&m_p);
                }
            }
        }
        operator I*() const { return m_p; }
        I* operator ->(){ return m_p; }
        I& operator *(){ return *m_p; }
        _com_ptr_t& operator =(I* p){
            if(m_p) m_p->Release();
            m_p = p;
            if(m_p) m_p->AddRef();
            return *this;
        }
        _com_ptr_t& operator =(const _com_ptr_t<TII>& r){
            if(m_p) m_p->Release();
            m_p = r.m_p;
            if(m_p) m_p->AddRef();
            return *this;
        }
        void Attach(I* p){
            if(m_p) m_p->Release();
            m_p = p;
        }
        I* Detach(){
            I* p = m_p;
            m_p = nullptr;
            return p;
        }
        ~_com_ptr_t(){
            if(m_p) m_p->Release();
        }
    protected:
        I* m_p;
    };



#endif



size_t utf8_wchar(wchar_t*, size_t, const char*);
size_t wchar_utf8(char*, size_t, const wchar_t*);


