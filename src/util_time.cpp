#include "util_time.hpp"
#include <sstream>
#include <iomanip>
#include <regex>

namespace util {

std::string trim(const std::string &s) {
    size_t b = s.find_first_not_of(" \t\r\n");
    size_t e = s.find_last_not_of(" \t\r\n");
    if (b == std::string::npos) return "";
    return s.substr(b, e - b + 1);
}

std::optional<std::chrono::system_clock::time_point> parse_time(const std::string &s) {
    std::tm tm{};
    std::istringstream ss(s);
    if (ss >> std::get_time(&tm, "%Y-%m-%dT%H:%M:%SZ")) {
        return std::chrono::system_clock::from_time_t(timegm(&tm));
    }
    ss.clear(); ss.str(s);
    if (ss >> std::get_time(&tm, "%Y-%m-%d %H:%M:%S")) {
        return std::chrono::system_clock::from_time_t(timegm(&tm));
    }
    ss.clear(); ss.str(s);
    if (ss >> std::get_time(&tm, "%m/%d/%Y %H:%M:%S")) {
        return std::chrono::system_clock::from_time_t(timegm(&tm));
    }
    return std::nullopt;
}

std::string compute_duration(const std::string &start, const std::string &end) {
    auto s = parse_time(start);
    auto e = parse_time(end);
    if (!s || !e) return "";
    auto diff = std::chrono::duration_cast<std::chrono::seconds>(*e - *s);
    return seconds_to_hhmmss((int)diff.count());
}

std::string seconds_to_hhmmss(int seconds) {
    int h = seconds / 3600;
    int m = (seconds % 3600) / 60;
    int s = seconds % 60;
    std::ostringstream os;
    os << std::setw(2) << std::setfill('0') << h << ':'
       << std::setw(2) << std::setfill('0') << m << ':'
       << std::setw(2) << std::setfill('0') << s;
    return os.str();
}

std::string size_to_gb(const std::string &text) {
    std::regex re(R"(([0-9.]+)\s*(B|KB|MB|GB))", std::regex::icase);
    std::smatch m;
    if (std::regex_search(text, m, re)) {
        double val = std::stod(m[1]);
        std::string unit = m[2];
        if (unit == "B" || unit == "b") val /= (1024.0*1024*1024);
        else if (unit == "KB" || unit == "kb") val /= (1024.0*1024);
        else if (unit == "MB" || unit == "mb") val /= 1024.0;
        // GB stays
        std::ostringstream os; os.setf(std::ios::fixed); os<<std::setprecision(2)<<val; return os.str();
    }
    return text;
}

std::string extract_date(const std::string &ts) {
    if (ts.size() >= 10) return ts.substr(0,10);
    return "";
}

} // namespace util

