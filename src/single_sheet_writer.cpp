#include "single_sheet_writer.hpp"
#include "util_time.hpp"
#include <xlsxwriter.h>

#include <algorithm>
#include <chrono>
#include <filesystem>
#include <fstream>
#include <optional>
#include <sstream>
#include <string>
#include <unordered_set>
#include <vector>

namespace fs = std::filesystem;

namespace {

inline std::string now_utc() {
  using namespace std::chrono;
  auto t = system_clock::to_time_t(system_clock::now());
  std::tm tm{};
#ifdef _WIN32
  gmtime_s(&tm, &t);
#else
  gmtime_r(&t, &tm);
#endif
  char buf[20]; // YYYY-MM-DD HH:MM:SS
  std::snprintf(buf, sizeof(buf), "%04d-%02d-%02d %02d:%02d:%02d",
                tm.tm_year + 1900, tm.tm_mon + 1, tm.tm_mday,
                tm.tm_hour, tm.tm_min, tm.tm_sec);
  return std::string(buf);
}

// Order of columns in the single sheet
static const std::vector<std::string> kHeaders = {
 "ProjectName","Tool","DatasetName","BuildID",
 "StartTime","EndTime","Duration(hh:mm:ss)","RunDate",
 "ProcessPreset","ExportType","SelAreaSize(km²)","Resolution","TileScheme",
 "PhotosUsed","PhotoFolders","PhotoCoverage(km²)","FusersUsed","CPUThreads","GPUCount",
 "Machine","HostIP","User",
 "OutputFolder","TotalFiles","TotalSize(GB)",
 "Offset_CoordSys","Offset_HDatum","Offset_VDatum",
 "OffsetX","OffsetY","OffsetZ",
 "PivotCenterX","PivotCenterY","PivotCenterZ",
 "FlipYZ","Trim","Collision","VisualLOD",
 "Success","Warnings","Errors","LogPath","IngestedAt"
};

inline std::string to_tsv(const std::vector<std::string>& cells) {
  std::ostringstream os;
  for (size_t i=0;i<cells.size();++i) {
    // Use tab-separated with basic sanitization (no tabs/newlines)
    std::string v = cells[i];
    for (char& c : v) if (c=='\t' || c=='\r' || c=='\n') c = ' ';
    os << v;
    if (i+1<cells.size()) os << '\t';
  }
  return os.str();
}

inline std::vector<std::string> split_tsv(const std::string& line) {
  std::vector<std::string> out;
  std::string cur;
  for (char c : line) {
    if (c=='\t') { out.push_back(cur); cur.clear(); }
    else if (c!='\r' && c!='\n') cur.push_back(c);
  }
  out.push_back(cur);
  return out;
}

} // namespace

namespace excel {

std::vector<UnifiedRow> unify(const std::vector<PhotoMeshRow>& pm,
                              const std::vector<RealityMeshRow>& rm) {
  std::vector<UnifiedRow> out;
  const std::string ingested = now_utc();

  for (const auto& r : pm) {
    UnifiedRow u{};
    u.ProjectName = r.projectName;
    u.Tool = "PhotoMesh";
    u.DatasetName = "";
    u.BuildID = r.buildID;
    u.StartTime = r.startTime; u.EndTime = r.endTime; u.Duration = r.duration;
    u.RunDate = util::extract_date(r.startTime.empty()? r.endTime : r.startTime);
    u.ProcessPreset = "";
    u.ExportType = r.exportType; u.SelAreaSize = ""; u.Resolution = r.resolution; u.TileScheme = r.tileScheme;
    u.PhotosUsed = r.photosUsed; u.PhotoFolders = r.photoFolders; u.PhotoCoverage = r.photoCoverage;
    u.FusersUsed = r.fusersUsed; u.CPUThreads = r.cpuThreads; u.GPUCount = r.gpuCount;
    u.Machine = r.machine; u.HostIP = r.hostIP; u.User = r.user;
    u.OutputFolder = r.outputFolder; u.TotalFiles = r.totalFiles; u.TotalSizeGB = r.totalSizeGB;
    u.Offset_CoordSys = r.offsetCoordSys; u.Offset_HDatum = r.offsetHDatum; u.Offset_VDatum = r.offsetVDatum;
    u.OffsetX = r.offsetX; u.OffsetY = r.offsetY; u.OffsetZ = r.offsetZ;
    u.PivotCenterX = ""; u.PivotCenterY = ""; u.PivotCenterZ = "";
    u.FlipYZ = r.flipYZ; u.Trim = r.trim; u.Collision = r.collision; u.VisualLOD = "";
    u.Success = r.success; u.Warnings = r.warnings; u.Errors = r.errors; u.LogPath = r.logPath;
    u.IngestedAt = ingested;
    out.push_back(std::move(u));
  }
  for (const auto& r : rm) {
    UnifiedRow u{};
    u.ProjectName = r.projectName.empty()? r.datasetName : r.projectName;
    u.Tool = "RealityMesh";
    u.DatasetName = r.datasetName; u.BuildID = "";
    u.StartTime = r.startTime; u.EndTime = r.endTime; u.Duration = r.duration;
    u.RunDate = util::extract_date(r.startTime.empty()? r.endTime : r.startTime);
    u.ProcessPreset = r.processPreset; u.ExportType = r.exportType;
    u.SelAreaSize = r.selAreaSize; u.Resolution = r.resolution; u.TileScheme = r.tileScheme;
    u.PhotosUsed = ""; u.PhotoFolders = ""; u.PhotoCoverage = "";
    u.FusersUsed = ""; u.CPUThreads = ""; u.GPUCount = "";
    u.Machine = r.machine; u.HostIP = r.hostIP; u.User = r.user;
    u.OutputFolder = r.outputFolder; u.TotalFiles = r.totalFiles; u.TotalSizeGB = r.totalSizeGB;
    u.Offset_CoordSys = r.offsetCoordSys; u.Offset_HDatum = r.offsetHDatum; u.Offset_VDatum = r.offsetVDatum;
    u.OffsetX = r.offsetX; u.OffsetY = r.offsetY; u.OffsetZ = r.offsetZ;
    u.PivotCenterX = ""; u.PivotCenterY = ""; u.PivotCenterZ = "";
    u.FlipYZ = r.flipYZ; u.Trim = r.trim; u.Collision = r.collision; u.VisualLOD = "";
    u.Success = r.success; u.Warnings = r.warnings; u.Errors = r.errors; u.LogPath = r.logPath;
    u.IngestedAt = ingested;
    out.push_back(std::move(u));
  }
  return out;
}

static void ensure_dir(const std::string& d) {
  std::error_code ec;
  fs::create_directories(fs::path(d), ec);
}

static std::string tsv_line_from_unified(const UnifiedRow& u) {
  std::vector<std::string> cells = {
    u.ProjectName,u.Tool,u.DatasetName,u.BuildID,
    u.StartTime,u.EndTime,u.Duration,u.RunDate,
    u.ProcessPreset,u.ExportType,u.SelAreaSize,u.Resolution,u.TileScheme,
    u.PhotosUsed,u.PhotoFolders,u.PhotoCoverage,u.FusersUsed,u.CPUThreads,u.GPUCount,
    u.Machine,u.HostIP,u.User,
    u.OutputFolder,u.TotalFiles,u.TotalSizeGB,
    u.Offset_CoordSys,u.Offset_HDatum,u.Offset_VDatum,
    u.OffsetX,u.OffsetY,u.OffsetZ,
    u.PivotCenterX,u.PivotCenterY,u.PivotCenterZ,
    u.FlipYZ,u.Trim,u.Collision,u.VisualLOD,
    u.Success,u.Warnings,u.Errors,u.LogPath,u.IngestedAt
  };
  return to_tsv(cells);
}

static void append_unique_to_tsv(const std::string& tsv_path,
                                 const std::vector<UnifiedRow>& rows) {
  std::unordered_set<std::string> seen;
  bool exists = fs::exists(tsv_path);

  // Build 'seen' from existing TSV by LogPath column
  if (exists) {
    std::ifstream in(tsv_path, std::ios::binary);
    std::string line;
    if (std::getline(in, line)) {
      auto hdr = split_tsv(line);
      int idx = -1;
      for (size_t i=0;i<hdr.size();++i) if (hdr[i] == "LogPath") { idx = (int)i; break; }
      if (idx >= 0) {
        while (std::getline(in, line)) {
          if (line.empty()) continue;
          auto cols = split_tsv(line);
          if ((size_t)idx < cols.size()) seen.insert(cols[idx]);
        }
      }
    }
  }

  std::ofstream out(tsv_path, std::ios::app | std::ios::binary);
  out.seekp(0, std::ios::end);
  if (!exists || out.tellp() == std::streampos(0)) {
    out << to_tsv(kHeaders) << "\n";
  }

  for (const auto& u : rows) {
    if (seen.find(u.LogPath) == seen.end()) {
      out << tsv_line_from_unified(u) << "\n";
    }
  }
}

static void rebuild_xlsx_from_tsv(const std::string& tsv_path,
                                  const std::string& xlsx_path) {
  std::ifstream in(tsv_path, std::ios::binary);
  if (!in) return;

  std::vector<std::string> headers;
  std::vector<std::vector<std::string>> rows;

  std::string line;
  if (std::getline(in, line)) headers = split_tsv(line);
  while (std::getline(in, line)) {
    if (!line.empty()) rows.push_back(split_tsv(line));
  }

  lxw_workbook* wb = workbook_new(xlsx_path.c_str());
  lxw_worksheet* ws = workbook_add_worksheet(wb, "All_Exports");

  // Write header & rows
  for (size_t c=0;c<headers.size();++c)
    worksheet_write_string(ws, 0, (lxw_col_t)c, headers[c].c_str(), nullptr);
  for (size_t r=0;r<rows.size();++r)
    for (size_t c=0;c<rows[r].size();++c)
      worksheet_write_string(ws, (lxw_row_t)(r+1), (lxw_col_t)c, rows[r][c].c_str(), nullptr);

  // Format as table, freeze header, set width
  if (!headers.empty()) {
    worksheet_add_table(ws, 0, 0, (lxw_row_t)rows.size(), (lxw_col_t)(headers.size()-1), nullptr);
    worksheet_freeze_panes(ws, 1, 0);
    worksheet_set_column(ws, 0, (lxw_col_t)(headers.size()-1), 22, nullptr);
  }

  // Conditional format for Success column
  int success_col = -1;
  for (size_t i=0;i<headers.size();++i) if (headers[i] == "Success") { success_col = (int)i; break; }
  if (success_col >= 0) {
    lxw_format* green = workbook_add_format(wb); format_set_bg_color(green, LXW_COLOR_GREEN);
    lxw_conditional_format cf1{}; cf1.type = LXW_CONDITIONAL_TYPE_CELL; cf1.criteria = LXW_CONDITIONAL_CRITERIA_EQUAL_TO;
    cf1.value_string = const_cast<char*>("True"); cf1.format = green;
    worksheet_conditional_format_range(ws, 1, success_col, (lxw_row_t)rows.size(), success_col, &cf1);

    lxw_format* red = workbook_add_format(wb); format_set_bg_color(red, LXW_COLOR_RED);
    lxw_conditional_format cf2{}; cf2.type = LXW_CONDITIONAL_TYPE_CELL; cf2.criteria = LXW_CONDITIONAL_CRITERIA_EQUAL_TO;
    cf2.value_string = const_cast<char*>("False"); cf2.format = red;
    worksheet_conditional_format_range(ws, 1, success_col, (lxw_row_t)rows.size(), success_col, &cf2);
  }

  workbook_close(wb);
}

void append_to_master_and_rebuild_xlsx(const std::string& outputs_dir,
                                       const std::vector<UnifiedRow>& new_rows) {
  ensure_dir(outputs_dir);
  const std::string tsv = (fs::path(outputs_dir) / "All_Exports.tsv").string();
  const std::string xlsx = (fs::path(outputs_dir) / "All_Exports.xlsx").string();
  append_unique_to_tsv(tsv, new_rows);
  rebuild_xlsx_from_tsv(tsv, xlsx);
}

} // namespace excel

