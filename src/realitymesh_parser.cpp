#include "realitymesh_parser.hpp"
#include "util_time.hpp"
#include <fstream>
#include <regex>
#include <filesystem>

RealityMeshRow parse_realitymesh(const std::string &path) {
    RealityMeshRow row;
    row.logPath = path;
    row.projectName = std::filesystem::path(path).stem().string();
    std::ifstream in(path);
    if (!in) return row;
    std::string line;
    std::regex dataset(R"delim(-command_file\s+"([^"]+)")delim");
    std::regex exitR(R"(Process completed with exit code:\s*(\d+))");
    std::regex runR(R"(Time to run TT project:\s*([0-9]+)\s*seconds)");
    std::regex inputOffset(R"(Input offset:\s*([\-0-9\.]+)\s+([\-0-9\.]+)\s+([\-0-9\.]+))");
    std::regex convOffset(R"(Converted offset:\s*([\-0-9\.]+)\s+([\-0-9\.]+)\s+([\-0-9\.]+))");
    int errCount=0;
    while (std::getline(in,line)) {
        std::smatch m;
        if (std::regex_search(line,m,dataset)) row.datasetName = m[1];
        if (std::regex_search(line,m,inputOffset)) {
            row.offsetX = m[1]; row.offsetY = m[2]; row.offsetZ = m[3];
        }
        if (std::regex_search(line,m,convOffset)) {
            row.offsetX = m[1]; row.offsetY = m[2]; row.offsetZ = m[3];
        }
        if (std::regex_search(line,m,exitR)) row.success = (m[1]=="0")?"True":"False";
        if (std::regex_search(line,m,runR)) row.duration = util::seconds_to_hhmmss(std::stoi(m[1]));
        auto pos = line.find("Error:");
        if (pos != std::string::npos) {
            errCount++;
            if (!row.errors.empty()) row.errors += ";";
            row.errors += util::trim(line.substr(pos+6));
        }
    }
    if (row.success.empty()) row.success = errCount==0 ? "True":"False";
    if (row.errors.empty() && errCount) row.errors = std::to_string(errCount);
    return row;
}

