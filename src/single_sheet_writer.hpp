#pragma once
#include <string>
#include <vector>
#include "models.hpp"

namespace excel {

// Unified row: superset of PM/RM fields + Tool + IngestedAt
struct UnifiedRow {
  std::string ProjectName, Tool, DatasetName, BuildID;
  std::string StartTime, EndTime, Duration, RunDate;
  std::string ProcessPreset, ExportType, SelAreaSize, Resolution, TileScheme;
  std::string PhotosUsed, PhotoFolders, PhotoCoverage, FusersUsed, CPUThreads, GPUCount;
  std::string Machine, HostIP, User;
  std::string OutputFolder, TotalFiles, TotalSizeGB;
  std::string Offset_CoordSys, Offset_HDatum, Offset_VDatum;
  std::string OffsetX, OffsetY, OffsetZ;
  std::string PivotCenterX, PivotCenterY, PivotCenterZ;
  std::string FlipYZ, Trim, Collision, VisualLOD;
  std::string Success, Warnings, Errors, LogPath, IngestedAt;
};

// Build unified rows from parsed structs
std::vector<UnifiedRow> unify(const std::vector<PhotoMeshRow>& pm,
                              const std::vector<RealityMeshRow>& rm);

// Append into <outputs_dir>/All_Exports.tsv (create + header if missing; skip duplicates by LogPath),
// then rebuild <outputs_dir>/All_Exports.xlsx (single-sheet) from the TSV.
void append_to_master_and_rebuild_xlsx(const std::string& outputs_dir,
                                       const std::vector<UnifiedRow>& new_rows);

} // namespace excel

