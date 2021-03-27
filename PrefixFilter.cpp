#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <vector>
#include <string>
#include <cstdio>
#include "resource.h"

DWORD s_dwOptions;

class CEnumString : public IEnumString, public IACList2
{
public:
    CEnumString() : m_cRefs(1), m_nIndex(0)
    {
        if (FILE *fp = fopen("list.txt", "rb"))
        {
            char buf[256];
            WCHAR szText[256];
            while (fgets(buf, 256, fp))
            {
                MultiByteToWideChar(CP_UTF8, 0, buf, _countof(buf), szText, _countof(szText));
                StrTrimW(szText, L" \t\r\n");
                m_data.push_back(szText);
            }
            fclose(fp);
        }
    }

    STDMETHODIMP QueryInterface(REFIID iid, VOID** ppv) override
    {
        if (iid == IID_IUnknown || iid == IID_IEnumString)
        {
            AddRef();
            *ppv = static_cast<IEnumString *>(this);
            return S_OK;
        }
        if (iid == IID_IACList || iid == IID_IACList2)
        {
            AddRef();
            *ppv = static_cast<IACList2 *>(this);
            return S_OK;
        }
        return E_NOINTERFACE;
    }
    STDMETHODIMP_(ULONG) AddRef() override
    {
        ++m_cRefs;
        return m_cRefs;
    }
    STDMETHODIMP_(ULONG) Release() override
    {
        --m_cRefs;
        if (m_cRefs == 0)
        {
            delete this;
            return 0;
        }
        return m_cRefs;
    }

    // *** IEnumString methods ***
    STDMETHODIMP Next(ULONG celt, LPOLESTR *rgelt, ULONG *pceltFetched) override
    {
        if (rgelt)
            *rgelt = NULL;
        if (pceltFetched)
            *pceltFetched = 0;
        if (celt != 1 || !rgelt || !pceltFetched)
            return E_INVALIDARG;
        if (m_nIndex >= m_data.size())
            return S_FALSE;

        SHStrDupW(m_data[m_nIndex].c_str(), rgelt);
        ++m_nIndex;
        if (!*rgelt)
            return E_OUTOFMEMORY;
        *pceltFetched = 1;
        return S_OK;
    }
    STDMETHODIMP Skip(ULONG celt) override
    {
        return E_NOTIMPL;
    }
    STDMETHODIMP Reset() override
    {
        m_nIndex = 0;
        return S_OK;
    }
    STDMETHODIMP Clone(IEnumString **ppenum) override
    {
        return E_NOTIMPL;
    }

    STDMETHODIMP Expand(PCWSTR pszExpand) override
    {
        return S_OK;
    }
    STDMETHODIMP GetOptions(DWORD *pdwFlag) override
    {
        return S_OK;
    }
    STDMETHODIMP SetOptions(DWORD dwFlag) override
    {
        return S_OK;
    }

protected:
    ULONG m_cRefs, m_nIndex;
    std::vector<std::wstring> m_data;
};

CEnumString *s_pEnum = NULL;

BOOL OnInitDialog(HWND hwnd, HWND hwndFocus, LPARAM lParam)
{
    IAutoComplete2 *pAC = NULL;
    CoCreateInstance(CLSID_AutoComplete, NULL, CLSCTX_INPROC_SERVER,
                     IID_IAutoComplete2, (VOID **)&pAC);
    if (pAC == NULL)
    {
        DestroyWindow(hwnd);
        return FALSE;
    }

    pAC->SetOptions(s_dwOptions);

    s_pEnum = new CEnumString();
    IUnknown *punk = static_cast<IEnumString *>(s_pEnum);
    pAC->Init(GetDlgItem(hwnd, edt1), punk, NULL, NULL); // IAutoComplete::Init

    return TRUE;
}

void OnCommand(HWND hwnd, int id, HWND hwndCtl, UINT codeNotify)
{
    switch (id)
    {
    case IDOK:
    case IDCANCEL:
        EndDialog(hwnd, id);
        break;
    }
}

INT_PTR CALLBACK
DialogProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
        HANDLE_MSG(hwnd, WM_INITDIALOG, OnInitDialog);
        HANDLE_MSG(hwnd, WM_COMMAND, OnCommand);
    }
    return 0;
}

DWORD GetOptions(void)
{
    DWORD ret = 0;

    WCHAR szPath[MAX_PATH];
    GetModuleFileNameW(NULL, szPath, _countof(szPath));
    PathRemoveFileSpecW(szPath);
    PathAppendW(szPath, L"Settings.ini");

    if (GetPrivateProfileIntW(L"Settings", L"AUTOAPPEND", 0, szPath))
    {
        ret |= ACO_AUTOAPPEND;
    }

    if (GetPrivateProfileIntW(L"Settings", L"AUTOSUGGEST", 0, szPath))
    {
        ret |= ACO_AUTOSUGGEST;
    }

    if (GetPrivateProfileIntW(L"Settings", L"FILTERPREFIXES", 0, szPath))
    {
        ret |= ACO_FILTERPREFIXES;
    }

    if (GetPrivateProfileIntW(L"Settings", L"NOPREFIXFILTERING", 0, szPath))
    {
        ret |= ACO_NOPREFIXFILTERING;
    }

    return ret;
}

struct CCoInit
{
    CCoInit() { hr = CoInitialize(NULL); }
    ~CCoInit() { if (SUCCEEDED(hr)) { CoUninitialize(); } }
    HRESULT hr;
};

INT WINAPI
WinMain(HINSTANCE   hInstance,
        HINSTANCE   hPrevInstance,
        LPSTR       lpCmdLine,
        INT         nCmdShow)
{
    CCoInit co;
    InitCommonControls();
    s_dwOptions = GetOptions();
    DialogBox(hInstance, MAKEINTRESOURCE(IDD_MAIN), NULL, DialogProc);
    return 0;
}
