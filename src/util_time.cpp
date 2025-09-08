// util_time.cpp
#include "util_time.hpp"

#include <algorithm>
#include <chrono>
#include <cctype>
#include <ctime>      // _mkgmtime / timegm
#include <iomanip>
#include <optional>
#include <regex>
#include <sstream>
#include <string>

namespace util {

    // --- UTC tm* -> time_t portability shim ---
    static std::time_t timegm_portable(std::tm* tm) {
#ifdef _WIN32
        return _mkgmtime(tm);   // MSVC / Windows
#else
        return timegm(tm);      // POSIX
#endif
    }

    // Trim leading/trailing ASCII whitespace
    std::string trim(const std::string& s) {
        static constexpr char ws[] = " \t\r\n";
        const auto b = s.find_first_not_of(ws);
        if (b == std::string::npos) return "";
        const auto e = s.find_last_not_of(ws);
        return s.substr(b, e - b + 1);
    }

    // Parse a few common UTC timestamp formats into a time_point
    std::optional<std::chrono::system_clock::time_point>
        parse_time(const std::string& s) {
        std::tm tm{};
        std::istringstream ss(s);

        // ISO 8601 with trailing Z
        ss >> std::get_time(&tm, "%Y-%m-%dT%H:%M:%SZ");
        if (!ss.fail()) return std::chrono::system_clock::from_time_t(timegm_portable(&tm));
        ss.clear(); ss.str(s);

        // "YYYY-MM-DD HH:MM:SS"
        ss >> std::get_time(&tm, "%Y-%m-%d %H:%M:%S");
        if (!ss.fail()) return std::chrono::system_clock::from_time_t(timegm_portable(&tm));
        ss.clear(); ss.str(s);

        // "MM/DD/YYYY HH:MM:SS"
        ss >> std::get_time(&tm, "%m/%d/%Y %H:%M:%S");
        if (!ss.fail()) return std::chrono::system_clock::from_time_t(timegm_portable(&tm));

        return std::nullopt;
    }

    std::string compute_duration(const std::string& start, const std::string& end) {
        auto s = parse_time(start);
        auto e = parse_time(end);
        if (!s || !e) return "";
        auto diff = std::chrono::duration_cast<std::chrono::seconds>(*e - *s);
        return seconds_to_hhmmss(static_cast<int>(diff.count()));
    }

    std::string seconds_to_hhmmss(int seconds) {
        if (seconds < 0) seconds = 0;
        const int h = seconds / 3600;
        const int m = (seconds % 3600) / 60;
        const int s = seconds % 60;

        std::ostringstream os;
        os << std::setw(2) << std::setfill('0') << h << ':'
            << std::setw(2) << std::setfill('0') << m << ':'
            << std::setw(2) << std::setfill('0') << s;
        return os.str();
    }

    std::string size_to_gb(const std::string& text) {
        // capture "123", "123.45" with units B/KB/MB/GB (case-insensitive)
        static const std::regex re(R"(([0-9]+(?:\.[0-9]+)?)\s*(B|KB|MB|GB))",
            std::regex::icase);

        std::smatch m;
        if (!std::regex_search(text, m, re)) return text;

        double val = std::stod(m[1].str());
        std::string unit = m[2].str();
        std::transform(unit.begin(), unit.end(), unit.begin(),
            [](unsigned char c) { return static_cast<char>(std::toupper(c)); });

        if (unit == "B")       val /= (1024.0 * 1024.0 * 1024.0);
        else if (unit == "KB") val /= (1024.0 * 1024.0);
        else if (unit == "MB") val /= 1024.0;
        // GB: leave as-is

        std::ostringstream os;
        os.setf(std::ios::fixed);
        os << std::setprecision(2) << val;
        return os.str();
    }

    std::string extract_date(const std::string& ts) {
        return ts.size() >= 10 ? ts.substr(0, 10) : std::string{};
    }

} // namespace util
