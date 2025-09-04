#include "photomesh_parser.hpp"
#include "util_time.hpp"
#include <fstream>
#include <regex>
#include <filesystem>

PhotoMeshRow parse_photomesh(const std::string &path) {
    PhotoMeshRow row;
    row.logPath = path;
    row.projectName = std::filesystem::path(path).stem().string();
    std::ifstream in(path);
    if (!in) return row;
    std::string line;
    std::regex kv(R"(SLDEFAULT=>\s*(\w+)\s*:\s*(.*))");
    std::regex machine(R"delim("MachineName"\s*:\s*"([^"]+)")delim");
    std::regex timeR(R"delim("MsgTime"\s*:\s*"([^"]+)")delim");
    std::regex exitR(R"(Finished with exit code\s*\((\d+)\))");
    int warn=0, err=0;
    while (std::getline(in,line)) {
        std::smatch m;
        if (std::regex_search(line,m,machine)) row.machine = m[1];
        if (std::regex_search(line,m,timeR)) {
            if (row.startTime.empty()) row.startTime = m[1];
            row.endTime = m[1];
        }
        if (std::regex_search(line,m,kv)) {
            std::string key = m[1];
            std::string val = util::trim(m[2]);
            if (key == "ExportType") row.exportType = val;
            else if (key == "Resolution") row.resolution = val;
            else if (key == "TileScheme") row.tileScheme = val;
            else if (key == "PhotosUsed") row.photosUsed = val;
            else if (key == "PhotoFolders") row.photoFolders = val;
            else if (key == "PhotoCoverage") row.photoCoverage = val;
            else if (key == "FusersUsed") row.fusersUsed = val;
            else if (key == "CPUThreads") row.cpuThreads = val;
            else if (key == "GPUCount") row.gpuCount = val;
            else if (key == "OutputFolder") row.outputFolder = val;
            else if (key == "TotalFiles") row.totalFiles = val;
            else if (key == "TotalSize") row.totalSizeGB = util::size_to_gb(val);
            else if (key == "Offset_CoordSys") row.offsetCoordSys = val;
            else if (key == "Offset_HDatum") row.offsetHDatum = val;
            else if (key == "Offset_VDatum") row.offsetVDatum = val;
            else if (key == "OffsetX") row.offsetX = val;
            else if (key == "OffsetY") row.offsetY = val;
            else if (key == "OffsetZ") row.offsetZ = val;
            else if (key == "PivotCenterX") row.pivotCenterX = val;
            else if (key == "PivotCenterY") row.pivotCenterY = val;
            else if (key == "PivotCenterZ") row.pivotCenterZ = val;
            else if (key == "FlipYZ") row.flipYZ = val;
            else if (key == "Trim") row.trim = val;
            else if (key == "Collision") row.collision = val;
            else if (key == "VisualLOD") row.visualLOD = val;
        }
        if (line.find("Warning") != std::string::npos) warn++;
        if (line.find("Error") != std::string::npos) err++;
        if (std::regex_search(line,m,exitR)) row.success = (m[1]=="0")?"True":"False";
    }
    if (row.success.empty()) row.success = err==0 ? "True" : "False";
    row.warnings = warn?std::to_string(warn):"";
    row.errors = err?std::to_string(err):"";
    row.duration = util::compute_duration(row.startTime,row.endTime);
    return row;
}

