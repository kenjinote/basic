#define UNICODE
#pragma comment(linker,"/opt:nowin98")
#pragma comment(lib,"comsupp")
#include<windows.h>
#include<comutil.h>
#include<activscp.h>

TCHAR szClassName[] = TEXT("Window");
HWND hEdit2;

#define SAFE_RELEASE(p) if((p)) { (p)->Release(); (p)=NULL; }

#define IMPL_IUNKNOWN(IF_NAME) \
public:\
	STDMETHODIMP_(ULONG) AddRef() {\
		return ++refCount;\
	}\
	STDMETHODIMP_(ULONG) Release() {\
		--refCount;\
		if (refCount <= 0) {\
			delete this;\
			return 0;\
				}\
		return refCount;\
	}\
	STDMETHODIMP QueryInterface(const IID & riid,void **ppvObj) {\
		if (!ppvObj)\
			return E_POINTER;\
		if (IsEqualIID(riid, __uuidof(IF_NAME)) ||\
			IsEqualIID(riid, IID_IUnknown)) {\
			AddRef();\
			*ppvObj=this;\
			return S_OK;\
				}\
		return E_NOINTERFACE;\
	}\
private:\
	ULONG refCount;

class PRINT : public IDispatch
{
	IMPL_IUNKNOWN(IDispatch)
public:
	static LPCOLESTR GetClassName()
	{
		return L"PRINT";
	}
	PRINT() :refCount(1)
	{
	}
	virtual ~PRINT()
	{
	}
	STDMETHODIMP GetTypeInfoCount(UINT __RPC_FAR *pctinfo)
	{
		return E_NOTIMPL;
	}
	STDMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo)
	{
		return E_NOTIMPL;
	}
	STDMETHODIMP GetIDsOfNames(REFIID riid, LPOLESTR __RPC_FAR *rgszNames, UINT cNames, LCID lcid, DISPID __RPC_FAR *rgDispId)
	{
		HRESULT hr = NOERROR;
		for (UINT i = 0; i<cNames; i++)
		{
			*(rgDispId + i) = DISPID_UNKNOWN;
			hr = DISP_E_MEMBERNOTFOUND;
		}
		return hr;
	}
	STDMETHODIMP Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS FAR *pDispParams, VARIANT FAR *pVarResult, EXCEPINFO FAR *pExcepInfo, unsigned int FAR *puArgErr)
	{
		if (wFlags != DISPATCH_METHOD)
		{
			return DISP_E_MEMBERNOTFOUND;
		}
		if (dispIdMember == 0)
		{
			for (int i = pDispParams->cArgs - 1; i >= 0; i--)
			{
				_variant_t str(pDispParams->rgvarg[i]);
				str.ChangeType(VT_BSTR);
				const int ndx = GetWindowTextLength(hEdit2);
				SendMessage(hEdit2, EM_SETSEL, (WPARAM)ndx, (LPARAM)ndx);
				SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)((LPSTR)str.bstrVal));
				if (i)
				{
					SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)TEXT(","));
				}
				else
				{
					SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)TEXT("\r\n"));
				}
			}
			return S_OK;
		}
		return DISP_E_MEMBERNOTFOUND;
	}
};


class MyScriptSite : public IActiveScriptSite
{
	IMPL_IUNKNOWN(IActiveScriptSite)
	IActiveScript *pAs;
	IActiveScriptParse *pAsp;
	BOOL bError;
public:
	MyScriptSite() : refCount(1), pAs(0), pAsp(0), bError(FALSE){}
	virtual ~MyScriptSite(){}
	STDMETHODIMP GetLCID(LCID*pLcid){return E_NOTIMPL;}
	STDMETHODIMP GetDocVersionString(BSTR *pbstrVersionString){return E_NOTIMPL;}
	STDMETHODIMP OnScriptTerminate(const VARIANT *pvarResult, const EXCEPINFO *pexcepinfo){return S_OK;}
	STDMETHODIMP OnStateChange(SCRIPTSTATE ssScriptState){return S_OK;}
	STDMETHODIMP OnScriptError(IActiveScriptError *pAse)
	{
		if (pAse == 0) return E_POINTER;
		EXCEPINFO ei;
		const HRESULT hr = pAse->GetExceptionInfo(&ei);
		if (SUCCEEDED(hr))
		{
			const int ndx = GetWindowTextLength(hEdit2);
			SendMessage(hEdit2, EM_SETSEL, (WPARAM)ndx, (LPARAM)ndx);
			SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)((LPSTR)ei.bstrDescription));
			bError = TRUE;
		}
		return hr;
	}
	STDMETHODIMP OnEnterScript(void) {return S_OK;}
	STDMETHODIMP OnLeaveScript(void)
	{
		const int ndx = GetWindowTextLength(hEdit2);
		SendMessage(hEdit2, EM_SETSEL, (WPARAM)ndx, (LPARAM)ndx);
		SendMessage(hEdit2, EM_REPLACESEL, 0, (LPARAM)TEXT("OK"));
		return S_OK;
	}
	STDMETHODIMP GetItemInfo(LPCOLESTR pstrName, DWORD dwReturnMask, IUnknown **ppunkItem, ITypeInfo **ppTypeInfo)
	{
		if ((dwReturnMask & SCRIPTINFO_IUNKNOWN))
		{
			if (wcscmp(pstrName, PRINT::GetClassName()) == 0)
			{
				*ppunkItem = (IUnknown*)new PRINT();
				return S_OK;
			}
		}
		return TYPE_E_ELEMENTNOTFOUND;
	}
	void Prepare()
	{
		CLSID clsid;
		CLSIDFromProgID(L"VBScript", &clsid);
		CoCreateInstance(clsid, 0, CLSCTX_ALL, IID_IActiveScript, (void **)&pAs);
		pAs->SetScriptSite(this);
		pAs->QueryInterface(IID_IActiveScriptParse, (void **)&pAsp);
		pAsp->InitNew();
		pAs->AddNamedItem(PRINT::GetClassName(), SCRIPTITEM_GLOBALMEMBERS | SCRIPTITEM_ISVISIBLE);
	}
	void Run(LPCOLESTR pScript)
	{
		pAsp->ParseScriptText(pScript, 0, 0, 0, 0, 0, SCRIPTTEXT_ISPERSISTENT, 0, 0);
		pAs->SetScriptState(SCRIPTSTATE_CONNECTED);
	}
	void Close()
	{
		pAs->SetScriptState(SCRIPTSTATE_CLOSED);
		SAFE_RELEASE(pAs);
		SAFE_RELEASE(pAsp);
	}
};

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hEdit1;
	static HFONT hFont;
	static HBRUSH hBrush;
	switch (msg)
	{
	case WM_CREATE:
		hFont = CreateFont(48, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, TEXT("Consolas"));
		hBrush = CreateSolidBrush(RGB(0, 0, 0));
		hEdit1 = CreateWindow(TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | WS_HSCROLL | WS_VSCROLL | ES_MULTILINE | ES_AUTOHSCROLL | ES_AUTOVSCROLL | ES_WANTRETURN, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hEdit2 = CreateWindow(TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | WS_HSCROLL | WS_VSCROLL | ES_READONLY | ES_MULTILINE | ES_AUTOHSCROLL | ES_AUTOVSCROLL | ES_WANTRETURN, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SendMessage(hEdit1, EM_LIMITTEXT, 0, 0);
		SendMessage(hEdit2, EM_LIMITTEXT, 0, 0);
		SendMessage(hEdit1, WM_SETFONT, (WPARAM)hFont, 0);
		SendMessage(hEdit2, WM_SETFONT, (WPARAM)hFont, 0);
		break;
	case WM_CTLCOLOREDIT:
		SetBkColor((HDC)wParam, RGB(0, 0, 0));
		SetTextColor((HDC)wParam, RGB(255, 255, 255));
		return (long)hBrush;
	case WM_SIZE:
		MoveWindow(hEdit1, 0, 0, LOWORD(lParam)/2, HIWORD(lParam), 1);
		MoveWindow(hEdit2, LOWORD(lParam)/2, 0, LOWORD(lParam)/2, HIWORD(lParam), 1);
		break;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case 100:
			{
				SetWindowText(hEdit2, 0);
				const int nTextLength = GetWindowTextLength(hEdit1);
				LPTSTR lpszText = (LPTSTR)GlobalAlloc(GMEM_FIXED, sizeof(TCHAR)*(nTextLength + 1));
				GetWindowText(hEdit1, lpszText, nTextLength + 1);
				CoInitialize(NULL);
				MyScriptSite *my_site = new MyScriptSite;
				my_site->Prepare();
				my_site->Run(lpszText);
				my_site->Close();
				SAFE_RELEASE(my_site);
				CoUninitialize();
				GlobalFree(lpszText);
			}
			break;
		case 101:
			SendMessage(hEdit1, EM_SETSEL, 0, -1);
			break;
		case IDCANCEL:
			return 0;
		case IDOK:
			return 0;
		}
		break;
	case WM_ERASEBKGND:
		return 1;  
	case WM_CLOSE:
		DestroyWindow(hWnd);
		break;
	case WM_DESTROY:
		DeleteObject(hBrush);
		DeleteObject(hFont);
		PostQuitMessage(0);
		break;
	default:
		return DefDlgProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		0,
		WndProc,
		0,
		DLGWINDOWEXTRA,
		hInstance,
		0,
		LoadCursor(0, IDC_ARROW),
		0,
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("LAZY BASIC (F5ÉLÅ[Ç≈é¿çs)"),
		WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
		);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	ACCEL Accel[] = { { FVIRTKEY, VK_F5, 100 }, { FVIRTKEY | FCONTROL, 'A', 101 } };
	HACCEL hAccel = CreateAcceleratorTable(Accel, sizeof(Accel) / sizeof(ACCEL));
	while (GetMessage(&msg, 0, 0, 0))
	{
		if (!TranslateAccelerator(hWnd, hAccel, &msg) && !IsDialogMessage(hWnd, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	DestroyAcceleratorTable(hAccel);
	return msg.wParam;
}
