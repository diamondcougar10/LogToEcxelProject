#pragma once
#include <string>
#include <optional>
#include <chrono>

namespace util {
std::string trim(const std::string &s);
std::optional<std::chrono::system_clock::time_point> parse_time(const std::string &s);
std::string compute_duration(const std::string &start, const std::string &end);
std::string seconds_to_hhmmss(int seconds);
std::string size_to_gb(const std::string &text);
std::string extract_date(const std::string &ts);
}

