/*
 * winshim.dll - in-process COM server that publishes WinsockVB6.dll under the
 * MSWINSCK.OCX identity, plus the OLE control layer a form-hosted control needs.
 *
 *   CLSID {248DD896-BB45-11CF-9ABC-0080C7E7B78D} -> this DLL
 *     CreateInstance:
 *       inner = WinsockVB6.dll!New_WinsockVB6()      (socket object)
 *       outer = the control below, which forwards IMSWinsockControl,
 *               IDispatch and IConnectionPointContainer to inner
 *
 * Implemented OLE interfaces: IOleObject, IViewObject/IViewObject2.
 * IQuickActivate is deliberately not offered so the host takes the slow path.
 * Persistence and in-place interfaces are not implemented yet; the trace tells
 * us which one VB6 asks for next (M0/M1).
 *
 * Every QI and every method call is appended to winshim.log next to the host.
 * WINSHIM_NO_LOG compiles the log out entirely (VS2019 Release); Debug and the
 * TCC build keep it.
 */

#include <windows.h>

/* Entry points: TCC needs __declspec(dllexport) for build.ps1 to find them.
 * The VS2019 project defines WINSHIM_NO_DLLEXPORT and exports the same four
 * names through vs2019\winshim.def, so only the plain names are exported. */
#ifdef WINSHIM_NO_DLLEXPORT
#define WINSHIM_API
#else
#define WINSHIM_API __declspec(dllexport)
#endif

/*------------------------------------------------------------------------------
 * Minimal COM surface
 *----------------------------------------------------------------------------*/
typedef struct IUnknown IUnknown;
typedef struct IUnknownVtbl
{
    HRESULT (__stdcall *QueryInterface)(IUnknown *, const GUID *, void **);
    ULONG   (__stdcall *AddRef)(IUnknown *);
    ULONG   (__stdcall *Release)(IUnknown *);
} IUnknownVtbl;
struct IUnknown { const IUnknownVtbl *lpVtbl; };

typedef struct IOleClientSite IOleClientSite;
typedef struct IAdviseSink IAdviseSink;

typedef struct IOleObject IOleObject;
typedef struct IOleObjectVtbl
{
    HRESULT (__stdcall *QueryInterface)(IOleObject *, const GUID *, void **);
    ULONG   (__stdcall *AddRef)(IOleObject *);
    ULONG   (__stdcall *Release)(IOleObject *);
    HRESULT (__stdcall *SetClientSite)(IOleObject *, IOleClientSite *);
    HRESULT (__stdcall *GetClientSite)(IOleObject *, IOleClientSite **);
    HRESULT (__stdcall *SetHostNames)(IOleObject *, const wchar_t *, const wchar_t *);
    HRESULT (__stdcall *Close)(IOleObject *, DWORD);
    HRESULT (__stdcall *SetMoniker)(IOleObject *, DWORD, void *);
    HRESULT (__stdcall *GetMoniker)(IOleObject *, DWORD, DWORD, void **);
    HRESULT (__stdcall *InitFromData)(IOleObject *, void *, BOOL, DWORD);
    HRESULT (__stdcall *GetClipboardData)(IOleObject *, DWORD, void **);
    HRESULT (__stdcall *DoVerb)(IOleObject *, LONG, void *, DWORD, HWND, const RECT *);
    HRESULT (__stdcall *EnumVerbs)(IOleObject *, void **);
    HRESULT (__stdcall *Update)(IOleObject *);
    HRESULT (__stdcall *IsUpToDate)(IOleObject *);
    HRESULT (__stdcall *GetUserClassID)(IOleObject *, GUID *);
    HRESULT (__stdcall *GetUserType)(IOleObject *, DWORD, wchar_t **);
    HRESULT (__stdcall *SetExtent)(IOleObject *, DWORD, SIZEL *);
    HRESULT (__stdcall *GetExtent)(IOleObject *, DWORD, SIZEL *);
    HRESULT (__stdcall *Advise)(IOleObject *, IAdviseSink *, DWORD *);
    HRESULT (__stdcall *Unadvise)(IOleObject *, DWORD);
    HRESULT (__stdcall *EnumAdvise)(IOleObject *, void **);
    HRESULT (__stdcall *GetMiscStatus)(IOleObject *, DWORD, DWORD *);
    HRESULT (__stdcall *SetColorScheme)(IOleObject *, void *);
} IOleObjectVtbl;
struct IOleObject { const IOleObjectVtbl *lpVtbl; };

typedef struct IViewObject IViewObject;
typedef struct IViewObjectVtbl
{
    HRESULT (__stdcall *QueryInterface)(IViewObject *, const GUID *, void **);
    ULONG   (__stdcall *AddRef)(IViewObject *);
    ULONG   (__stdcall *Release)(IViewObject *);
    HRESULT (__stdcall *Draw)(IViewObject *, DWORD, LONG, void *, void *, void *, HDC, RECT *, void *, DWORD);
    HRESULT (__stdcall *GetColorSet)(IViewObject *, DWORD, LONG, void *, void *, HDC, void **);
    HRESULT (__stdcall *Freeze)(IViewObject *, DWORD, LONG, void *, DWORD *);
    HRESULT (__stdcall *Unfreeze)(IViewObject *, DWORD);
    HRESULT (__stdcall *SetAdvise)(IViewObject *, DWORD, DWORD, IAdviseSink *);
    HRESULT (__stdcall *GetAdvise)(IViewObject *, DWORD *, DWORD *, IAdviseSink **);
    HRESULT (__stdcall *GetExtent)(IViewObject *, DWORD, LONG, void *, SIZEL *);
} IViewObjectVtbl;
struct IViewObject { const IViewObjectVtbl *lpVtbl; };

typedef struct IPersistStreamInit IPersistStreamInit;
typedef struct IPersistStreamInitVtbl
{
    HRESULT (__stdcall *QueryInterface)(IPersistStreamInit *, const GUID *, void **);
    ULONG   (__stdcall *AddRef)(IPersistStreamInit *);
    ULONG   (__stdcall *Release)(IPersistStreamInit *);
    HRESULT (__stdcall *GetClassID)(IPersistStreamInit *, GUID *);
    HRESULT (__stdcall *IsDirty)(IPersistStreamInit *);
    HRESULT (__stdcall *InitNew)(IPersistStreamInit *);
    HRESULT (__stdcall *Load)(IPersistStreamInit *, void *);
    HRESULT (__stdcall *Save)(IPersistStreamInit *, void *, BOOL);
    HRESULT (__stdcall *GetSizeMax)(IPersistStreamInit *, ULARGE_INTEGER *);
} IPersistStreamInitVtbl;
struct IPersistStreamInit { const IPersistStreamInitVtbl *lpVtbl; };

typedef struct IPersistPropertyBag IPersistPropertyBag;
typedef struct IPersistPropertyBagVtbl
{
    HRESULT (__stdcall *QueryInterface)(IPersistPropertyBag *, const GUID *, void **);
    ULONG   (__stdcall *AddRef)(IPersistPropertyBag *);
    ULONG   (__stdcall *Release)(IPersistPropertyBag *);
    HRESULT (__stdcall *GetClassID)(IPersistPropertyBag *, GUID *);
    HRESULT (__stdcall *InitNew)(IPersistPropertyBag *);
    HRESULT (__stdcall *Load)(IPersistPropertyBag *, void *, void *);
    HRESULT (__stdcall *Save)(IPersistPropertyBag *, void *, BOOL, BOOL);
} IPersistPropertyBagVtbl;
struct IPersistPropertyBag { const IPersistPropertyBagVtbl *lpVtbl; };

typedef struct IOleControl IOleControl;
typedef struct { ULONG cb; HACCEL hAccel; USHORT cProps; DWORD *pProps; } CONTROLINFO;
typedef LONG DISPID;
typedef struct IOleControlVtbl
{
    HRESULT (__stdcall *QueryInterface)(IOleControl *, const GUID *, void **);
    ULONG   (__stdcall *AddRef)(IOleControl *);
    ULONG   (__stdcall *Release)(IOleControl *);
    HRESULT (__stdcall *GetControlInfo)(IOleControl *, CONTROLINFO *);
    HRESULT (__stdcall *OnMnemonic)(IOleControl *, MSG *);
    HRESULT (__stdcall *OnAmbientPropertyChange)(IOleControl *, DISPID);
    HRESULT (__stdcall *FreezeEvents)(IOleControl *, BOOL);
} IOleControlVtbl;
struct IOleControl { const IOleControlVtbl *lpVtbl; };

/* IDispatch: argument types we never interpret are passed as void* so the
   header set stays minimal; layouts are pointer-compatible. */
typedef struct IDispatch IDispatch;
typedef struct IDispatchVtbl
{
    HRESULT (__stdcall *QueryInterface)(IDispatch *, const GUID *, void **);
    ULONG   (__stdcall *AddRef)(IDispatch *);
    ULONG   (__stdcall *Release)(IDispatch *);
    HRESULT (__stdcall *GetTypeInfoCount)(IDispatch *, UINT *);
    HRESULT (__stdcall *GetTypeInfo)(IDispatch *, UINT, LCID, void **);
    HRESULT (__stdcall *GetIDsOfNames)(IDispatch *, const GUID *, wchar_t **, UINT, LCID, DISPID *);
    HRESULT (__stdcall *Invoke)(IDispatch *, DISPID, const GUID *, LCID, WORD, void *,
                                void *, void *, UINT *);
} IDispatchVtbl;
struct IDispatch { const IDispatchVtbl *lpVtbl; };


typedef struct IClassFactory IClassFactory;
typedef struct IClassFactoryVtbl
{
    HRESULT (__stdcall *QueryInterface)(IClassFactory *, const GUID *, void **);
    ULONG   (__stdcall *AddRef)(IClassFactory *);
    ULONG   (__stdcall *Release)(IClassFactory *);
    HRESULT (__stdcall *CreateInstance)(IClassFactory *, IUnknown *, const GUID *, void **);
    HRESULT (__stdcall *LockServer)(IClassFactory *, BOOL);
} IClassFactoryVtbl;
struct IClassFactory { const IClassFactoryVtbl *lpVtbl; };

typedef struct IClassFactory2 IClassFactory2;
typedef struct IClassFactory2Vtbl
{
    HRESULT (__stdcall *QueryInterface)(IClassFactory2 *, const GUID *, void **);
    ULONG   (__stdcall *AddRef)(IClassFactory2 *);
    ULONG   (__stdcall *Release)(IClassFactory2 *);
    HRESULT (__stdcall *CreateInstance)(IClassFactory2 *, IUnknown *, const GUID *, void **);
    HRESULT (__stdcall *LockServer)(IClassFactory2 *, BOOL);
    HRESULT (__stdcall *GetLicInfo)(IClassFactory2 *, void *);
    HRESULT (__stdcall *RequestLicKey)(IClassFactory2 *, DWORD, wchar_t **);
    HRESULT (__stdcall *CreateInstanceLic)(IClassFactory2 *, IUnknown *, IUnknown *,
                                           const GUID *, wchar_t *, void **);
} IClassFactory2Vtbl;
struct IClassFactory2 { const IClassFactory2Vtbl *lpVtbl; };

#define S_OK_        ((HRESULT)0)
#define S_FALSE_     ((HRESULT)1)
#define E_NOTIMPL_   ((HRESULT)0x80004001L)
#define E_NOINTERFACE_ ((HRESULT)0x80004002L)
#define E_POINTER_   ((HRESULT)0x80004003L)
#define E_FAIL_      ((HRESULT)0x80004005L)
#define E_OUTOFMEMORY_ ((HRESULT)0x8007000EL)
#define CLASS_E_NOAGGREGATION ((HRESULT)0x80040110L)
#define CLASS_E_CLASSNOTAVAILABLE ((HRESULT)0x80040111L)
#define REGDB_E_CLASSNOTREG ((HRESULT)0x80040154L)

static const GUID CLSID_MsWinsock =
    {0x248DD896,0xBB45,0x11CF,{0x9A,0xBC,0x00,0x80,0xC7,0xE7,0xB7,0x8D}};
static const GUID IID_IUnknown_ =
    {0x00000000,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static const GUID IID_IClassFactory_ =
    {0x00000001,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static const GUID IID_IPersistPropertyBag_ =
    {0x37D84F60,0x42CB,0x11CE,{0x81,0x35,0x00,0xAA,0x00,0x4B,0xB8,0x51}};
static const GUID IID_IClassFactory2_ =
    {0xB196B28F,0xBAB4,0x101A,{0xB6,0x9C,0x00,0xAA,0x00,0x34,0x1D,0x07}};
static const GUID IID_IOleObject_ =
    {0x00000112,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static const GUID IID_IViewObject_ =
    {0x0000010D,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static const GUID IID_IViewObject2_ =
    {0x00000127,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static const GUID IID_IPersistStreamInit_ =
    {0x7FD52380,0x4E07,0x101B,{0xAE,0x2D,0x08,0x00,0x2B,0x2E,0xC7,0x13}};
static const GUID IID_IOleControl_ =
    {0xB196B288,0xBAB4,0x101A,{0xB6,0x9C,0x00,0xAA,0x00,0x34,0x1D,0x07}};
static const GUID IID_IDispatch_ =
    {0x00020400,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};

static int guid_eq(const GUID *a, const GUID *b)
{
    const unsigned char *p = (const unsigned char *)a;
    const unsigned char *q = (const unsigned char *)b;
    int i;
    if (!a || !b) return 0;
    for (i = 0; i < (int)sizeof(GUID); i++)
        if (p[i] != q[i]) return 0;
    return 1;
}

/*------------------------------------------------------------------------------
 * Trace log (next to the host EXE)
 *
 * WINSHIM_NO_LOG compiles the log out entirely: no file, no formatting, no
 * strings in the DLL. The VS2019 Release configuration defines it; Debug and
 * the TCC reference build keep the log.
 *----------------------------------------------------------------------------*/
#ifndef WINSHIM_NO_LOG
static void trace_append(const char *line)
{
    char path[MAX_PATH];
    HANDLE h;
    DWORD written;

    GetModuleFileNameA(0, path, MAX_PATH);
    {
        char *q = path;
        while (*q) q++;
        while (q > path && *q != '\\') q--;
        if (*q == '\\') q++;
        lstrcpyA(q, "winshim.log");
    }
    h = CreateFileA(path, FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE,
                    0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
    if (h == INVALID_HANDLE_VALUE) return;
    WriteFile(h, line, lstrlenA(line), &written, 0);
    WriteFile(h, "\r\n", 2, &written, 0);
    CloseHandle(h);
}

static void loghr(const char *what, HRESULT hr)
{
    char b[160];
    wsprintfA(b, "%s -> 0x%08lX", what, (unsigned long)hr);
    trace_append(b);
}

static void logiid(const char *what, const GUID *g)
{
    char b[200];
    char h[80];
    wsprintfA(h, "{%08lX-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}",
              (unsigned long)g->Data1, g->Data2, g->Data3,
              g->Data4[0], g->Data4[1], g->Data4[2], g->Data4[3],
              g->Data4[4], g->Data4[5], g->Data4[6], g->Data4[7]);
    lstrcpyA(b, what); lstrcatA(b, " "); lstrcatA(b, h);
    trace_append(b);
}
#else
#define trace_append(line)      ((void)0)
#define loghr(what, hr)         ((void)0)
#define logiid(what, guid)      ((void)0)
#endif

/*------------------------------------------------------------------------------
 * The control object
 *----------------------------------------------------------------------------*/
typedef struct Ctl Ctl;
typedef struct { const IOleObjectVtbl *lpVtbl; Ctl *c; } OlePart;
typedef struct { const IViewObjectVtbl *lpVtbl; Ctl *c; } ViewPart;
typedef struct { const IPersistStreamInitVtbl *lpVtbl; Ctl *c; } PersistPart;
typedef struct { const IPersistPropertyBagVtbl *lpVtbl; Ctl *c; } PersistBagPart;
typedef struct { const IOleControlVtbl *lpVtbl; Ctl *c; } CtlCtlPart;
typedef struct { const IDispatchVtbl *lpVtbl; Ctl *c; } DispPart;
struct Ctl
{
    LONG ref;
    IUnknown *inner;
    IOleClientSite *site;
    IAdviseSink *advise;
    DWORD adviseCookie;
    DWORD viewAspect;
    IAdviseSink *viewSink;
    SIZEL extent;
    OlePart ole;
    ViewPart view;
    PersistPart persist;
    PersistBagPart bag;
    CtlCtlPart ctlctl;
    DispPart disp;
    IDispatch *innerDisp;
};

static HRESULT CtlQI(Ctl *c, const GUID *riid, void **ppv);

/* ---- IOleObject ----------------------------------------------------------*/
static HRESULT __stdcall O_QI(IOleObject *p, const GUID *riid, void **ppv)
{ return CtlQI(((OlePart *)p)->c, riid, ppv); }
static ULONG __stdcall O_AddRef(IOleObject *p)
{ Ctl *c = ((OlePart *)p)->c; return (ULONG)InterlockedIncrement(&c->ref); }
static ULONG __stdcall O_Release(IOleObject *p)
{ Ctl *c = ((OlePart *)p)->c; return (ULONG)InterlockedDecrement(&c->ref); }

static HRESULT __stdcall O_SetClientSite(IOleObject *p, IOleClientSite *site)
{
    Ctl *c = ((OlePart *)p)->c;
    c->site = site;
    loghr("IOleObject::SetClientSite", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall O_GetClientSite(IOleObject *p, IOleClientSite **pp)
{
    Ctl *c = ((OlePart *)p)->c;
    if (!pp) return E_POINTER_;
    *pp = c->site;
    if (c->site) /* AddRef through the site vtable */
        ((IUnknown *)c->site)->lpVtbl->AddRef((IUnknown *)c->site);
    loghr("IOleObject::GetClientSite", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall O_SetHostNames(IOleObject *p, const wchar_t *app, const wchar_t *obj)
{ loghr("IOleObject::SetHostNames", S_OK_); return S_OK_; }

static HRESULT __stdcall O_Close(IOleObject *p, DWORD flags)
{
    Ctl *c = ((OlePart *)p)->c;
    c->site = 0;
    loghr("IOleObject::Close", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall O_SetMoniker(IOleObject *p, DWORD a, void *b)
{ loghr("IOleObject::SetMoniker", E_NOTIMPL_); return E_NOTIMPL_; }

static HRESULT __stdcall O_GetMoniker(IOleObject *p, DWORD a, DWORD b, void **pp)
{ loghr("IOleObject::GetMoniker", E_NOTIMPL_); return E_NOTIMPL_; }

static HRESULT __stdcall O_InitFromData(IOleObject *p, void *d, BOOL f, DWORD r)
{ loghr("IOleObject::InitFromData", E_NOTIMPL_); return E_NOTIMPL_; }

static HRESULT __stdcall O_GetClipboardData(IOleObject *p, DWORD r, void **pp)
{ loghr("IOleObject::GetClipboardData", E_NOTIMPL_); return E_NOTIMPL_; }

static HRESULT __stdcall O_DoVerb(IOleObject *p, LONG verb, void *msg, DWORD idx,
                                  HWND parent, const RECT *rect)
{
#ifndef WINSHIM_NO_LOG
    char b[80];
    wsprintfA(b, "IOleObject::DoVerb verb=%ld", (long)verb);
    loghr(b, S_OK_);
#endif
    return S_OK_;
}

static HRESULT __stdcall O_EnumVerbs(IOleObject *p, void **pp)
{ loghr("IOleObject::EnumVerbs", E_NOTIMPL_); return E_NOTIMPL_; }

static HRESULT __stdcall O_Update(IOleObject *p)
{ loghr("IOleObject::Update", S_OK_); return S_OK_; }

static HRESULT __stdcall O_IsUpToDate(IOleObject *p)
{ loghr("IOleObject::IsUpToDate", S_OK_); return S_OK_; }

static HRESULT __stdcall O_GetUserClassID(IOleObject *p, GUID *pclsid)
{
    if (!pclsid) return E_POINTER_;
    *pclsid = CLSID_MsWinsock;
    loghr("IOleObject::GetUserClassID", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall O_GetUserType(IOleObject *p, DWORD form, wchar_t **pp)
{ loghr("IOleObject::GetUserType", E_NOTIMPL_); return E_NOTIMPL_; }

static HRESULT __stdcall O_SetExtent(IOleObject *p, DWORD aspect, SIZEL *size)
{
    Ctl *c = ((OlePart *)p)->c;
    if (!size) return E_POINTER_;
    c->extent = *size;
    loghr("IOleObject::SetExtent", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall O_GetExtent(IOleObject *p, DWORD aspect, SIZEL *size)
{
    Ctl *c = ((OlePart *)p)->c;
    if (!size) return E_POINTER_;
    if (c->extent.cx == 0 && c->extent.cy == 0) { c->extent.cx = 2540; c->extent.cy = 2540; }
    *size = c->extent;
    loghr("IOleObject::GetExtent", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall O_Advise(IOleObject *p, IAdviseSink *sink, DWORD *cookie)
{
    Ctl *c = ((OlePart *)p)->c;
    c->advise = sink;
    c->adviseCookie = 1;
    if (cookie) *cookie = 1;
    loghr("IOleObject::Advise", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall O_Unadvise(IOleObject *p, DWORD cookie)
{
    Ctl *c = ((OlePart *)p)->c;
    c->advise = 0;
    loghr("IOleObject::Unadvise", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall O_EnumAdvise(IOleObject *p, void **pp)
{ loghr("IOleObject::EnumAdvise", E_NOTIMPL_); return E_NOTIMPL_; }

static HRESULT __stdcall O_GetMiscStatus(IOleObject *p, DWORD aspect, DWORD *pdw)
{
    if (pdw) *pdw = 0;
    loghr("IOleObject::GetMiscStatus", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall O_SetColorScheme(IOleObject *p, void *pal)
{ loghr("IOleObject::SetColorScheme", S_OK_); return S_OK_; }

static const IOleObjectVtbl g_oleVtbl =
{
    O_QI, O_AddRef, O_Release,
    O_SetClientSite, O_GetClientSite, O_SetHostNames, O_Close,
    O_SetMoniker, O_GetMoniker, O_InitFromData, O_GetClipboardData,
    O_DoVerb, O_EnumVerbs, O_Update, O_IsUpToDate, O_GetUserClassID,
    O_GetUserType, O_SetExtent, O_GetExtent, O_Advise, O_Unadvise,
    O_EnumAdvise, O_GetMiscStatus, O_SetColorScheme
};

/* ---- IViewObject ---------------------------------------------------------*/
static HRESULT __stdcall V_QI(IViewObject *p, const GUID *riid, void **ppv)
{ return CtlQI(((ViewPart *)p)->c, riid, ppv); }
static ULONG __stdcall V_AddRef(IViewObject *p)
{ Ctl *c = ((ViewPart *)p)->c; return (ULONG)InterlockedIncrement(&c->ref); }
static ULONG __stdcall V_Release(IViewObject *p)
{ Ctl *c = ((ViewPart *)p)->c; return (ULONG)InterlockedDecrement(&c->ref); }

static HRESULT __stdcall V_Draw(IViewObject *p, DWORD aspect, LONG idx, void *a,
                                void *b, void *c3, HDC hdc, RECT *prc, void *d, DWORD e)
{
    /* Empty shell: nothing to paint. */
    loghr("IViewObject::Draw", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall V_GetColorSet(IViewObject *p, DWORD a, LONG b, void *c1,
                                       void *d, HDC e, void **pp)
{ loghr("IViewObject::GetColorSet", E_NOTIMPL_); return E_NOTIMPL_; }

static HRESULT __stdcall V_Freeze(IViewObject *p, DWORD aspect, LONG idx, void *pv, DWORD *pdw)
{
    if (pdw) *pdw = 0;
    loghr("IViewObject::Freeze", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall V_Unfreeze(IViewObject *p, DWORD dw)
{ loghr("IViewObject::Unfreeze", S_OK_); return S_OK_; }

static HRESULT __stdcall V_SetAdvise(IViewObject *p, DWORD aspect, DWORD flags, IAdviseSink *sink)
{
    Ctl *c = ((ViewPart *)p)->c;
    c->viewAspect = aspect;
    c->viewSink = sink;
    loghr("IViewObject::SetAdvise", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall V_GetAdvise(IViewObject *p, DWORD *pa, DWORD *pf, IAdviseSink **pps)
{
    Ctl *c = ((ViewPart *)p)->c;
    if (pa) *pa = c->viewAspect;
    if (pf) *pf = 0;
    if (pps) *pps = c->viewSink;
    loghr("IViewObject::GetAdvise", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall V_GetExtent(IViewObject *p, DWORD aspect, LONG idx, void *dev, SIZEL *size)
{
    Ctl *c = ((ViewPart *)p)->c;
    if (!size) return E_POINTER_;
    if (c->extent.cx == 0 && c->extent.cy == 0) { c->extent.cx = 2540; c->extent.cy = 2540; }
    *size = c->extent;
    loghr("IViewObject2::GetExtent", S_OK_);
    return S_OK_;
}

static const IViewObjectVtbl g_viewVtbl =
{
    V_QI, V_AddRef, V_Release, V_Draw, V_GetColorSet,
    V_Freeze, V_Unfreeze, V_SetAdvise, V_GetAdvise, V_GetExtent
};

/* ---- IPersistStreamInit --------------------------------------------------*/
static HRESULT __stdcall P_QI(IPersistStreamInit *p, const GUID *riid, void **ppv)
{ return CtlQI(((PersistPart *)p)->c, riid, ppv); }
static ULONG __stdcall P_AddRef(IPersistStreamInit *p)
{ Ctl *c = ((PersistPart *)p)->c; return (ULONG)InterlockedIncrement(&c->ref); }
static ULONG __stdcall P_Release(IPersistStreamInit *p)
{ Ctl *c = ((PersistPart *)p)->c; return (ULONG)InterlockedDecrement(&c->ref); }

static HRESULT __stdcall P_GetClassID(IPersistStreamInit *p, GUID *pclsid)
{
    if (!pclsid) return E_POINTER_;
    *pclsid = CLSID_MsWinsock;
    loghr("IPersistStreamInit::GetClassID", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall P_IsDirty(IPersistStreamInit *p)
{
    loghr("IPersistStreamInit::IsDirty", S_FALSE_);
    return S_FALSE_;
}

static HRESULT __stdcall P_InitNew(IPersistStreamInit *p)
{
    loghr("IPersistStreamInit::InitNew", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall P_Load(IPersistStreamInit *p, void *stm)
{
    /* Design-time state is not interpreted in v1: the control has no
       persisted properties of its own beyond the generic extents. */
    loghr("IPersistStreamInit::Load", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall P_Save(IPersistStreamInit *p, void *stm, BOOL clearDirty)
{
    loghr("IPersistStreamInit::Save", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall P_GetSizeMax(IPersistStreamInit *p, ULARGE_INTEGER *pcb)
{
    if (!pcb) return E_POINTER_;
    pcb->QuadPart = 0;
    loghr("IPersistStreamInit::GetSizeMax", S_OK_);
    return S_OK_;
}

static const IPersistStreamInitVtbl g_persistVtbl =
{
    P_QI, P_AddRef, P_Release, P_GetClassID, P_IsDirty,
    P_InitNew, P_Load, P_Save, P_GetSizeMax
};

/* ---- IPersistPropertyBag ------------------------------------------------*/
/* VB6 asks for this when a control class is created from code. It is only
   offered once a client site exists (form hosting); see CtlQI. */
static HRESULT __stdcall PB_QI(IPersistPropertyBag *p, const GUID *riid, void **ppv)
{ return CtlQI(((PersistBagPart *)p)->c, riid, ppv); }
static ULONG __stdcall PB_AddRef(IPersistPropertyBag *p)
{ Ctl *c = ((PersistBagPart *)p)->c; return (ULONG)InterlockedIncrement(&c->ref); }
static ULONG __stdcall PB_Release(IPersistPropertyBag *p)
{ Ctl *c = ((PersistBagPart *)p)->c; return (ULONG)InterlockedDecrement(&c->ref); }

static HRESULT __stdcall PB_GetClassID(IPersistPropertyBag *p, GUID *pclsid)
{
    if (!pclsid) return E_POINTER_;
    *pclsid = CLSID_MsWinsock;
    loghr("IPersistPropertyBag::GetClassID", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall PB_InitNew(IPersistPropertyBag *p)
{ loghr("IPersistPropertyBag::InitNew", S_OK_); return S_OK_; }

static HRESULT __stdcall PB_Load(IPersistPropertyBag *p, void *bag, void *errlog)
{
    /* Property replay is not implemented; accepting the bag lets a host
       finish initialising the control. */
    loghr("IPersistPropertyBag::Load", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall PB_Save(IPersistPropertyBag *p, void *bag, BOOL clear, BOOL all)
{ loghr("IPersistPropertyBag::Save", S_OK_); return S_OK_; }

static const IPersistPropertyBagVtbl g_bagVtbl =
{
    PB_QI, PB_AddRef, PB_Release, PB_GetClassID, PB_InitNew, PB_Load, PB_Save
};


/* ---- IOleControl ---------------------------------------------------------*/
static HRESULT __stdcall OC_QI(IOleControl *p, const GUID *riid, void **ppv)
{ return CtlQI(((CtlCtlPart *)p)->c, riid, ppv); }
static ULONG __stdcall OC_AddRef(IOleControl *p)
{ Ctl *c = ((CtlCtlPart *)p)->c; return (ULONG)InterlockedIncrement(&c->ref); }
static ULONG __stdcall OC_Release(IOleControl *p)
{ Ctl *c = ((CtlCtlPart *)p)->c; return (ULONG)InterlockedDecrement(&c->ref); }

static HRESULT __stdcall OC_GetControlInfo(IOleControl *p, CONTROLINFO *ci)
{
    if (!ci) return E_POINTER_;
    ci->cb = sizeof(CONTROLINFO);
    ci->hAccel = 0;
    ci->cProps = 0;
    ci->pProps = 0;
    loghr("IOleControl::GetControlInfo", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall OC_OnMnemonic(IOleControl *p, MSG *msg)
{
    loghr("IOleControl::OnMnemonic", S_OK_);
    return S_OK_;
}

static HRESULT __stdcall OC_OnAmbientPropertyChange(IOleControl *p, DISPID id)
{
#ifndef WINSHIM_NO_LOG
    char b[80];
    wsprintfA(b, "IOleControl::OnAmbientPropertyChange id=%ld", (long)id);
    loghr(b, S_OK_);
#endif
    return S_OK_;
}

static HRESULT __stdcall OC_FreezeEvents(IOleControl *p, BOOL freeze)
{
#ifndef WINSHIM_NO_LOG
    char b[64];
    wsprintfA(b, "IOleControl::FreezeEvents %ld", (long)freeze);
    loghr(b, S_OK_);
#endif
    return S_OK_;
}

static const IOleControlVtbl g_ctlctlVtbl =
{
    OC_QI, OC_AddRef, OC_Release, OC_GetControlInfo,
    OC_OnMnemonic, OC_OnAmbientPropertyChange, OC_FreezeEvents
};


/* ---- shared QueryInterface -----------------------------------------------*/
static HRESULT CtlQI(Ctl *c, const GUID *riid, void **ppv)
{
    HRESULT hr;

    if (!ppv) return E_POINTER_;
    *ppv = 0;
    logiid("QI", riid);

    if (guid_eq(riid, &IID_IUnknown_) || guid_eq(riid, &IID_IOleObject_))
    {
        *ppv = &c->ole;
        InterlockedIncrement(&c->ref);
        return S_OK_;
    }
    if (guid_eq(riid, &IID_IViewObject_) || guid_eq(riid, &IID_IViewObject2_))
    {
        *ppv = &c->view;
        InterlockedIncrement(&c->ref);
        return S_OK_;
    }
    /* Persistence is only offered once a client site exists. A form host
       calls SetClientSite first; a control created from code (New Winsock in
       a VB6 EXE) does not, and aborts in the persistence handshake. */
    if (guid_eq(riid, &IID_IPersistStreamInit_) && c->site)
    {
        *ppv = &c->persist;
        InterlockedIncrement(&c->ref);
        return S_OK_;
    }
    if (guid_eq(riid, &IID_IPersistPropertyBag_) && c->site)
    {
        *ppv = &c->bag;
        InterlockedIncrement(&c->ref);
        return S_OK_;
    }
    if (guid_eq(riid, &IID_IOleControl_))
    {
        *ppv = &c->ctlctl;
        InterlockedIncrement(&c->ref);
        return S_OK_;
    }



    /* Socket interfaces live on the inner object. */
    hr = ((IUnknown *)c->inner)->lpVtbl->QueryInterface((IUnknown *)c->inner, riid, ppv);
    if (hr != S_OK_) logiid("QI-unhandled", riid);
    return hr;
}

/*------------------------------------------------------------------------------
 * Factory and DLL entry points
 *----------------------------------------------------------------------------*/
typedef void *(__stdcall *PFN_NEW)(void);
static HMODULE g_dll;
static PFN_NEW g_new;
static HINSTANCE g_self;

static HRESULT make_inner(void **out)
{
    *out = 0;
    if (!g_new)
    {
        g_dll = LoadLibraryA("WinsockVB6.dll");
        if (!g_dll) return REGDB_E_CLASSNOTREG;
        g_new = (PFN_NEW)GetProcAddress(g_dll, "New_WinsockVB6");
        if (!g_new) return REGDB_E_CLASSNOTREG;
    }
    *out = g_new();
    return (*out) ? S_OK_ : E_FAIL_;
}

static HRESULT __stdcall F_QueryInterface(IClassFactory *p, const GUID *riid, void **ppv)
{
    if (guid_eq(riid, &IID_IUnknown_) || guid_eq(riid, &IID_IClassFactory_))
    { *ppv = p; return S_OK_; }
    *ppv = 0;
    return E_NOINTERFACE_;
}
static ULONG __stdcall F_AddRef(IClassFactory *p)  { return 2; }
static ULONG __stdcall F_Release(IClassFactory *p) { return 1; }

static HRESULT __stdcall F_CreateInstance(IClassFactory *p, IUnknown *outer,
                                          const GUID *riid, void **ppv)
{
    void *inner = 0;
    Ctl *c;
    HRESULT hr;

    if (outer) return CLASS_E_NOAGGREGATION;
    hr = make_inner(&inner);
    if (hr != S_OK_) return hr;

    c = (Ctl *)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(Ctl));
    if (!c) return E_OUTOFMEMORY_;
    c->ref = 1;
    c->inner = (IUnknown *)inner;
    c->ole.lpVtbl = &g_oleVtbl;  c->ole.c = c;
    c->view.lpVtbl = &g_viewVtbl; c->view.c = c;
    c->persist.lpVtbl = &g_persistVtbl; c->persist.c = c;
    c->bag.lpVtbl = &g_bagVtbl; c->bag.c = c;
    c->ctlctl.lpVtbl = &g_ctlctlVtbl; c->ctlctl.c = c;

    trace_append("CreateInstance");
    hr = CtlQI(c, riid, ppv);
    InterlockedDecrement(&c->ref);
    return hr;
}

static HRESULT __stdcall F_LockServer(IClassFactory *p, BOOL lock) { return S_OK_; }

/* IClassFactory2: VB6 asks for it when creating a licensed control. Without
 * it the control is treated as a fresh, unlicensed instance and the form's
 * saved design-time properties are never loaded. */
typedef struct { LONG cbSize; BOOL fRuntimeKeyAvail; BOOL fLicVerified; } LICINFO_;

typedef wchar_t *(__stdcall *PFN_SYSALLOC)(const wchar_t *);
static PFN_SYSALLOC pSysAllocString;

static void lic_init(void)
{
    static HMODULE oleaut;
    if (pSysAllocString) return;
    if (!oleaut) oleaut = LoadLibraryA("oleaut32.dll");
    if (oleaut) pSysAllocString = (PFN_SYSALLOC)GetProcAddress(oleaut, "SysAllocString");
}

static HRESULT __stdcall F2_QueryInterface(IClassFactory2 *p, const GUID *riid, void **ppv)
{
    if (!ppv) return E_POINTER_;
    if (guid_eq(riid, &IID_IUnknown_) || guid_eq(riid, &IID_IClassFactory_) ||
        guid_eq(riid, &IID_IClassFactory2_))
    { *ppv = p; return S_OK_; }
    logiid("factory2 QI-unhandled", riid);
    *ppv = 0;
    return E_NOINTERFACE_;
}
static ULONG __stdcall F2_AddRef(IClassFactory2 *p)  { return 2; }
static ULONG __stdcall F2_Release(IClassFactory2 *p) { return 1; }
static HRESULT __stdcall F2_CreateInstance(IClassFactory2 *p, IUnknown *outer,
                                           const GUID *riid, void **ppv)
{ return F_CreateInstance((IClassFactory *)p, outer, riid, ppv); }
static HRESULT __stdcall F2_LockServer(IClassFactory2 *p, BOOL lock) { return S_OK_; }

static HRESULT __stdcall F2_GetLicInfo(IClassFactory2 *p, void *lic)
{
    LICINFO_ *li = (LICINFO_ *)lic;
#ifndef WINSHIM_NO_LOG
    char b[80];
#endif
    if (!li) return E_POINTER_;
    li->cbSize = sizeof(LICINFO_);
    li->fRuntimeKeyAvail = TRUE;
    li->fLicVerified = TRUE;
#ifndef WINSHIM_NO_LOG
    wsprintfA(b, "IClassFactory2::GetLicInfo cb=%ld", (long)li->cbSize);
    trace_append(b);
#endif
    return S_OK_;
}

static HRESULT __stdcall F2_RequestLicKey(IClassFactory2 *p, DWORD reserved, wchar_t **key)
{
    static const wchar_t k[] = L"WinsockVB6";
    if (!key) return E_POINTER_;
    lic_init();
    *key = pSysAllocString ? pSysAllocString(k) : 0;
    trace_append("IClassFactory2::RequestLicKey");
    return *key ? S_OK_ : E_OUTOFMEMORY_;
}

static HRESULT __stdcall F2_CreateInstanceLic(IClassFactory2 *p, IUnknown *outer,
                                              IUnknown *reserved, const GUID *riid,
                                              wchar_t *key, void **ppv)
{
    /* The key is not a real license: any key (or none) is accepted. */
    trace_append("IClassFactory2::CreateInstanceLic");
    (void)reserved; (void)key;
    return F_CreateInstance((IClassFactory *)p, outer, riid, ppv);
}

static const IClassFactory2Vtbl g_factory2Vtbl =
{
    F2_QueryInterface, F2_AddRef, F2_Release, F2_CreateInstance, F2_LockServer,
    F2_GetLicInfo, F2_RequestLicKey, F2_CreateInstanceLic
};
static IClassFactory2 g_factory2 = { &g_factory2Vtbl };

static const IClassFactoryVtbl g_factoryVtbl =
{
    F_QueryInterface, F_AddRef, F_Release, F_CreateInstance, F_LockServer
};
static IClassFactory g_factory = { &g_factoryVtbl };

WINSHIM_API HRESULT __stdcall DllGetClassObject(const GUID *rclsid,
                                                const GUID *riid,
                                                void **ppv)
{
    if (!guid_eq(rclsid, &CLSID_MsWinsock)) return CLASS_E_CLASSNOTAVAILABLE;
    return g_factory2.lpVtbl->QueryInterface(&g_factory2, riid, ppv);
}

WINSHIM_API HRESULT __stdcall DllCanUnloadNow(void) { return S_FALSE_; }

/*------------------------------------------------------------------------------
 * Registration (advapi32 resolved at run time)
 *----------------------------------------------------------------------------*/
typedef LONG (__stdcall *PFN_CREATE)(HKEY, LPCSTR, DWORD, LPSTR, DWORD, DWORD,
                                     LPSECURITY_ATTRIBUTES, PHKEY, LPDWORD);
typedef LONG (__stdcall *PFN_SET)(HKEY, LPCSTR, DWORD, DWORD, const BYTE *, DWORD);
typedef LONG (__stdcall *PFN_CLOSE)(HKEY);

static PFN_CREATE pRegCreate;
static PFN_SET    pRegSet;
static PFN_CLOSE  pRegClose;

static int advapi_ready(void)
{
    HMODULE adv;
    if (pRegCreate) return 1;
    adv = LoadLibraryA("advapi32.dll");
    if (!adv) return 0;
    pRegCreate = (PFN_CREATE)GetProcAddress(adv, "RegCreateKeyExA");
    pRegSet    = (PFN_SET)GetProcAddress(adv, "RegSetValueExA");
    pRegClose  = (PFN_CLOSE)GetProcAddress(adv, "RegCloseKey");
    return pRegCreate && pRegSet && pRegClose;
}

static void reg_sz(HKEY root, const char *sub, const char *name, const char *value)
{
    HKEY key;
    DWORD disp;
    if (pRegCreate(root, sub, 0, 0, 0, KEY_WRITE, 0, &key, &disp) != 0) return;
    if (value)
        pRegSet(key, name, 0, REG_SZ, (const BYTE *)value,
                (DWORD)(lstrlenA(value) + 1));
    pRegClose(key);
}

WINSHIM_API HRESULT __stdcall DllRegisterServer(void)
{
    char path[MAX_PATH];
    const char *clsid = "{248DD896-BB45-11CF-9ABC-0080C7E7B78D}";

    if (!advapi_ready()) return E_FAIL_;
    if (!GetModuleFileNameA(g_self, path, MAX_PATH)) return E_FAIL_;

    reg_sz(HKEY_LOCAL_MACHINE,
           "Software\\Classes\\CLSID\\{248DD896-BB45-11CF-9ABC-0080C7E7B78D}\\InprocServer32",
           0, path);
    reg_sz(HKEY_LOCAL_MACHINE,
           "Software\\Classes\\CLSID\\{248DD896-BB45-11CF-9ABC-0080C7E7B78D}\\InprocServer32",
           "ThreadingModel", "Apartment");
    reg_sz(HKEY_LOCAL_MACHINE, "Software\\Classes\\MSWinsock.Winsock",
           "CurVer", "MSWinsock.Winsock.1");
    reg_sz(HKEY_LOCAL_MACHINE, "Software\\Classes\\MSWinsock.Winsock\\CLSID", 0, clsid);
    reg_sz(HKEY_LOCAL_MACHINE, "Software\\Classes\\MSWinsock.Winsock.1\\CLSID", 0, clsid);
    return S_OK_;
}

WINSHIM_API HRESULT __stdcall DllUnregisterServer(void) { return S_OK_; }

BOOL __stdcall DllMain(HINSTANCE self, DWORD reason, LPVOID reserved)
{
    if (reason == DLL_PROCESS_ATTACH)
    {
        g_self = self;
        DisableThreadLibraryCalls(self);
    }
    return TRUE;
}