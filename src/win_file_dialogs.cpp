#ifdef _WIN32
#include <windows.h>
#include <shobjidl.h>
#include <vector>
#include <string>

static std::string narrow(const std::wstring& ws) {
    if (ws.empty()) return {};
    int len = ::WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), nullptr, 0, nullptr, nullptr);
    std::string out(len, '\0');
    ::WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), out.data(), len, nullptr, nullptr);
    return out;
}

std::vector<std::string> open_logs_dialog(const wchar_t* title, bool allowMulti) {
    std::vector<std::string> result;

    IFileOpenDialog* dlg = nullptr;
    HRESULT hr = ::CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dlg));
    if (FAILED(hr)) return result;

    DWORD opts = 0;
    if (SUCCEEDED(dlg->GetOptions(&opts))) {
        if (allowMulti) opts |= FOS_ALLOWMULTISELECT;
        dlg->SetOptions(opts);
    }
    if (title) dlg->SetTitle(title);

    COMDLG_FILTERSPEC filters[] = {
        { L"Log files (*.log;*.txt)", L"*.log;*.txt" },
        { L"All files (*.*)",         L"*.*" }
    };
    dlg->SetFileTypes((UINT)(sizeof(filters)/sizeof(filters[0])), filters);

    hr = dlg->Show(nullptr);
    if (SUCCEEDED(hr)) {
        IShellItemArray* items = nullptr;
        if (SUCCEEDED(dlg->GetResults(&items))) {
            DWORD count = 0;
            items->GetCount(&count);
            for (DWORD i = 0; i < count; ++i) {
                IShellItem* it = nullptr;
                if (SUCCEEDED(items->GetItemAt(i, &it))) {
                    PWSTR psz = nullptr;
                    if (SUCCEEDED(it->GetDisplayName(SIGDN_FILESYSPATH, &psz)) && psz) {
                        result.push_back(narrow(psz));
                        ::CoTaskMemFree(psz);
                    }
                    it->Release();
                }
            }
            items->Release();
        }
    }
    dlg->Release();
    return result;
}

std::string save_xlsx_dialog(const wchar_t* title, const std::wstring& suggestedName) {
    std::string out;

    IFileSaveDialog* dlg = nullptr;
    HRESULT hr = ::CoCreateInstance(CLSID_FileSaveDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dlg));
    if (FAILED(hr)) return out;

    if (title) dlg->SetTitle(title);

    COMDLG_FILTERSPEC fs[] = { { L"Excel Workbook (*.xlsx)", L"*.xlsx" } };
    dlg->SetFileTypes(1, fs);
    dlg->SetDefaultExtension(L"xlsx");
    dlg->SetFileName(suggestedName.c_str());

    hr = dlg->Show(nullptr);
    if (SUCCEEDED(hr)) {
        IShellItem* item = nullptr;
        if (SUCCEEDED(dlg->GetResult(&item))) {
            PWSTR psz = nullptr;
            if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &psz)) && psz) {
                out = narrow(psz);
                ::CoTaskMemFree(psz);
            }
            item->Release();
        }
    }
    dlg->Release();
    return out;
}
#endif

