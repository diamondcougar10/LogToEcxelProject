#pragma once
#include <string>
#include <vector>

struct PhotoMeshRow {
    std::string projectName;
    std::string buildID;
    std::string machine;
    std::string hostIP;
    std::string user;
    std::string startTime;
    std::string endTime;
    std::string duration;
    std::string exportType;
    std::string resolution;
    std::string tileScheme;
    std::string photosUsed;
    std::string photoFolders;
    std::string photoCoverage;
    std::string fusersUsed;
    std::string cpuThreads;
    std::string gpuCount;
    std::string outputFolder;
    std::string totalFiles;
    std::string totalSizeGB;
    std::string offsetCoordSys;
    std::string offsetHDatum;
    std::string offsetVDatum;
    std::string offsetX;
    std::string offsetY;
    std::string offsetZ;
    std::string pivotCenterX;
    std::string pivotCenterY;
    std::string pivotCenterZ;
    std::string flipYZ;
    std::string trim;
    std::string collision;
    std::string visualLOD;
    std::string success;
    std::string warnings;
    std::string errors;
    std::string logPath;
};

struct RealityMeshRow {
    std::string projectName;
    std::string datasetName;
    std::string machine;
    std::string hostIP;
    std::string user;
    std::string startTime;
    std::string endTime;
    std::string duration;
    std::string processPreset;
    std::string exportType;
    std::string selAreaSize;
    std::string resolution;
    std::string tileScheme;
    std::string offsetCoordSys;
    std::string offsetHDatum;
    std::string offsetVDatum;
    std::string offsetX;
    std::string offsetY;
    std::string offsetZ;
    std::string pivotCenterX;
    std::string pivotCenterY;
    std::string pivotCenterZ;
    std::string flipYZ;
    std::string trim;
    std::string collision;
    std::string visualLOD;
    std::string outputFolder;
    std::string totalFiles;
    std::string totalSizeGB;
    std::string success;
    std::string warnings;
    std::string errors;
    std::string logPath;
};

struct SummaryRow {
    std::string projectName;
    std::string runDate;
    std::string tool;
    std::string exportType;
    std::string duration;
    std::string totalSizeGB;
    std::string photosUsed;
    std::string fusersUsed;
    std::string machine;
    std::string success;
    std::string errors;
};

SummaryRow make_summary(const PhotoMeshRow &row);
SummaryRow make_summary(const RealityMeshRow &row);

