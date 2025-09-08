#pragma once
#include <string>
#include <vector>

// Windows-only helpers. No-ops on non-Windows builds.
#ifdef _WIN32
std::vector<std::string> open_logs_dialog(const wchar_t* title, bool allowMulti);
std::string save_xlsx_dialog(const wchar_t* title, const std::wstring& suggestedName);
#endif

