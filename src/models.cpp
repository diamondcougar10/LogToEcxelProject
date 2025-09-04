#include "models.hpp"
#include "util_time.hpp"
#include <filesystem>

SummaryRow make_summary(const PhotoMeshRow &row) {
    SummaryRow s;
    s.projectName = row.projectName;
    s.runDate = util::extract_date(row.startTime.empty() ? row.endTime : row.startTime);
    s.tool = "PhotoMesh";
    s.exportType = row.exportType;
    s.duration = row.duration;
    s.totalSizeGB = row.totalSizeGB;
    s.photosUsed = row.photosUsed;
    s.fusersUsed = row.fusersUsed;
    s.machine = row.machine;
    s.success = row.success;
    s.errors = row.errors;
    return s;
}

SummaryRow make_summary(const RealityMeshRow &row) {
    SummaryRow s;
    s.projectName = row.projectName.empty() ? row.datasetName : row.projectName;
    s.runDate = util::extract_date(row.startTime.empty() ? row.endTime : row.startTime);
    s.tool = "RealityMesh";
    s.exportType = row.exportType;
    s.duration = row.duration;
    s.totalSizeGB = row.totalSizeGB;
    s.photosUsed = "";
    s.fusersUsed = "";
    s.machine = row.machine;
    s.success = row.success;
    s.errors = row.errors;
    return s;
}

