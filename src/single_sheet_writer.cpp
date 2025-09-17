#include "single_sheet_writer.hpp"
#include "util_time.hpp"
#include <xlsxwriter.h>

#include <algorithm>
#include <cctype>
#include <chrono>
#include <filesystem>
#include <fstream>
#include <optional>
#include <sstream>
#include <string>
#include <unordered_map>
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

inline std::string normalize_header(const std::string& s) {
  std::string out;
  out.reserve(s.size());
  for (unsigned char ch : s) {
    if (std::isalnum(ch)) out.push_back(static_cast<char>(std::tolower(ch)));
  }
  return out;
}

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

struct ColumnSpec {
  std::string title;
  std::string group;
  double      width;
  bool        wrap;
};

struct ColumnLayout {
  std::string key;
  ColumnSpec spec;
};

static const std::vector<ColumnLayout> kMasterLayout = {
  {"ProjectName",      {"Project Name",         "At a Glance",       28.0, false}},
  {"Machine",          {"Computer Name",        "At a Glance",       22.0, false}},
  {"Tool",             {"Tool",                 "At a Glance",       16.0, false}},
  {"DatasetName",      {"Dataset Name",         "At a Glance",       24.0, false}},
  {"ExportType",       {"Export Type",          "At a Glance",       20.0, false}},
  {"ProcessPreset",    {"Process Preset",       "At a Glance",       18.0, false}},
  {"StartTime",        {"Start Time",           "At a Glance",       20.0, false}},
  {"EndTime",          {"End Time",             "At a Glance",       20.0, false}},
  {"Duration(hh:mm:ss)",{"Duration (hh:mm:ss)", "At a Glance",       16.0, false}},
  {"RunDate",          {"Run Date",             "At a Glance",       14.0, false}},
  {"TotalSize(GB)",    {"Total Size (GB)",      "At a Glance",       14.0, false}},

  {"BuildID",          {"Build ID",             "Project Setup",     16.0, false}},
  {"SelAreaSize(km²)", {"Selection Area (km²)", "Project Setup",     18.0, false}},
  {"Resolution",       {"Resolution",           "Project Setup",     14.0, false}},
  {"TileScheme",       {"Tile Scheme",          "Project Setup",     14.0, false}},

  {"PhotosUsed",       {"Photos Used",          "Resources",         12.0, false}},
  {"PhotoFolders",     {"Photo Folders",        "Resources",         28.0, true }},
  {"PhotoCoverage(km²)",{"Photo Coverage (km²)","Resources",         18.0, false}},
  {"FusersUsed",       {"Fusers Used",          "Resources",         12.0, false}},
  {"CPUThreads",       {"CPU Threads",          "Resources",         12.0, false}},
  {"GPUCount",         {"GPU Count",            "Resources",         12.0, false}},

  {"HostIP",           {"Host IP",              "Environment",       16.0, false}},
  {"User",             {"User",                 "Environment",       16.0, false}},

  {"OutputFolder",     {"Output Folder",        "Output",            36.0, true }},
  {"TotalFiles",       {"Total Files",          "Output",            12.0, false}},
  {"LogPath",          {"Log Path",             "Output",            40.0, true }},

  {"Offset_CoordSys",  {"Offset Coord Sys",     "Offsets & Settings",22.0, false}},
  {"Offset_HDatum",    {"Offset H Datum",       "Offsets & Settings",18.0, false}},
  {"Offset_VDatum",    {"Offset V Datum",       "Offsets & Settings",18.0, false}},
  {"OffsetX",          {"Offset X",             "Offsets & Settings",12.0, false}},
  {"OffsetY",          {"Offset Y",             "Offsets & Settings",12.0, false}},
  {"OffsetZ",          {"Offset Z",             "Offsets & Settings",12.0, false}},
  {"PivotCenterX",     {"Pivot Center X",       "Offsets & Settings",14.0, false}},
  {"PivotCenterY",     {"Pivot Center Y",       "Offsets & Settings",14.0, false}},
  {"PivotCenterZ",     {"Pivot Center Z",       "Offsets & Settings",14.0, false}},
  {"FlipYZ",           {"Flip YZ",              "Offsets & Settings",10.0, false}},
  {"Trim",             {"Trim",                 "Offsets & Settings",10.0, false}},
  {"Collision",        {"Collision",            "Offsets & Settings",12.0, false}},
  {"VisualLOD",        {"Visual LOD",           "Offsets & Settings",12.0, false}},

  {"Success",         {"Success",              "Status",            10.0, false}},
  {"Warnings",        {"Warnings",             "Status",            26.0, true }},
  {"Errors",          {"Errors",               "Status",            26.0, true }},

  {"IngestedAt",      {"Ingested At",          "Metadata",          20.0, false}},
};

void write_sectioned_table(lxw_workbook* wb,
                           lxw_worksheet* ws,
                           const std::vector<ColumnSpec>& columns,
                           const std::vector<std::vector<std::string>>& rows) {
  if (columns.empty()) return;

  lxw_format* group_fmt = workbook_add_format(wb);
  format_set_align(group_fmt, LXW_ALIGN_CENTER);
  format_set_align(group_fmt, LXW_ALIGN_VERTICAL_CENTER);
  format_set_bold(group_fmt);
  format_set_bg_color(group_fmt, 0xD9E1F2);

  lxw_format* header_fmt = workbook_add_format(wb);
  format_set_align(header_fmt, LXW_ALIGN_CENTER);
  format_set_align(header_fmt, LXW_ALIGN_VERTICAL_CENTER);
  format_set_bold(header_fmt);
  format_set_bg_color(header_fmt, 0xEEF2F7);

  lxw_format* wrap_fmt = workbook_add_format(wb);
  format_set_text_wrap(wrap_fmt);
  format_set_align(wrap_fmt, LXW_ALIGN_TOP);

  worksheet_set_row(ws, 0, 22.0, nullptr);
  worksheet_set_row(ws, 1, 20.0, nullptr);

  size_t col = 0;
  while (col < columns.size()) {
    const auto& spec = columns[col];
    if (!spec.group.empty()) {
      size_t end = col;
      while (end + 1 < columns.size() && columns[end + 1].group == spec.group) {
        ++end;
      }
      worksheet_merge_range(ws, 0, col, 0, end, spec.group.c_str(), group_fmt);
      col = end + 1;
    } else {
      worksheet_write_string(ws, 0, col, "", group_fmt);
      ++col;
    }
  }

  for (size_t c = 0; c < columns.size(); ++c) {
    worksheet_write_string(ws, 1, static_cast<lxw_col_t>(c),
                           columns[c].title.c_str(), header_fmt);
    worksheet_set_column(ws, static_cast<lxw_col_t>(c), static_cast<lxw_col_t>(c),
                         columns[c].width, columns[c].wrap ? wrap_fmt : nullptr);
  }

  for (size_t r = 0; r < rows.size(); ++r) {
    for (size_t c = 0; c < rows[r].size(); ++c) {
      worksheet_write_string(ws, static_cast<lxw_row_t>(r + 2),
                             static_cast<lxw_col_t>(c), rows[r][c].c_str(), nullptr);
    }
  }

  lxw_row_t last_row = rows.empty() ? 1 : static_cast<lxw_row_t>(rows.size() + 1);
  worksheet_add_table(ws, 1, 0, last_row, static_cast<lxw_col_t>(columns.size() - 1), nullptr);
  worksheet_freeze_panes(ws, 2, 0);
}

size_t find_column_index(const std::vector<ColumnSpec>& cols, const std::string& title) {
  for (size_t i = 0; i < cols.size(); ++i) {
    if (cols[i].title == title) return i;
  }
  return cols.size();
}

void add_success_format(lxw_workbook* wb,
                        lxw_worksheet* ws,
                        size_t success_col,
                        size_t row_count) {
  if (success_col >= LXW_COL_MAX || row_count == 0) return;

  lxw_format* green = workbook_add_format(wb);
  format_set_bg_color(green, LXW_COLOR_GREEN);
  lxw_conditional_format cf1{};
  cf1.type = LXW_CONDITIONAL_TYPE_CELL;
  cf1.criteria = LXW_CONDITIONAL_CRITERIA_EQUAL_TO;
  cf1.value_string = const_cast<char*>("True");
  cf1.format = green;
  worksheet_conditional_format_range(
      ws, 2, static_cast<lxw_col_t>(success_col),
      static_cast<lxw_row_t>(row_count + 1), static_cast<lxw_col_t>(success_col), &cf1);

  lxw_format* red = workbook_add_format(wb);
  format_set_bg_color(red, LXW_COLOR_RED);
  lxw_conditional_format cf2{};
  cf2.type = LXW_CONDITIONAL_TYPE_CELL;
  cf2.criteria = LXW_CONDITIONAL_CRITERIA_EQUAL_TO;
  cf2.value_string = const_cast<char*>("False");
  cf2.format = red;
  worksheet_conditional_format_range(
      ws, 2, static_cast<lxw_col_t>(success_col),
      static_cast<lxw_row_t>(row_count + 1), static_cast<lxw_col_t>(success_col), &cf2);
}

size_t index_in_layout(const std::vector<ColumnLayout>& layout, const std::string& key) {
  const std::string norm = normalize_header(key);
  for (size_t i = 0; i < layout.size(); ++i) {
    if (normalize_header(layout[i].key) == norm) return i;
  }
  return layout.size();
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
    u.PivotCenterX = r.pivotCenterX; u.PivotCenterY = r.pivotCenterY; u.PivotCenterZ = r.pivotCenterZ;
    u.FlipYZ = r.flipYZ; u.Trim = r.trim; u.Collision = r.collision; u.VisualLOD = r.visualLOD;
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
    u.PivotCenterX = r.pivotCenterX; u.PivotCenterY = r.pivotCenterY; u.PivotCenterZ = r.pivotCenterZ;
    u.FlipYZ = r.flipYZ; u.Trim = r.trim; u.Collision = r.collision; u.VisualLOD = r.visualLOD;
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
      const std::string target = normalize_header("LogPath");
      for (size_t i=0;i<hdr.size();++i) {
        if (normalize_header(hdr[i]) == target) { idx = static_cast<int>(i); break; }
      }
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
  std::vector<std::vector<std::string>> raw_rows;

  std::string line;
  if (std::getline(in, line)) headers = split_tsv(line);
  while (std::getline(in, line)) {
    if (!line.empty()) raw_rows.push_back(split_tsv(line));
  }

  std::unordered_map<std::string, size_t> header_index;
  for (size_t i = 0; i < headers.size(); ++i) {
    header_index.emplace(normalize_header(headers[i]), i);
  }

  std::vector<std::vector<std::string>> rows;
  rows.reserve(raw_rows.size());
  for (const auto& row : raw_rows) {
    std::vector<std::string> ordered;
    ordered.reserve(kMasterLayout.size());
    for (const auto& col : kMasterLayout) {
      auto it = header_index.find(normalize_header(col.key));
      if (it != header_index.end() && it->second < row.size())
        ordered.push_back(row[it->second]);
      else
        ordered.emplace_back();
    }
    rows.push_back(std::move(ordered));
  }

  const auto start_idx   = index_in_layout(kMasterLayout, "StartTime");
  const auto end_idx     = index_in_layout(kMasterLayout, "EndTime");
  const auto run_idx     = index_in_layout(kMasterLayout, "RunDate");
  const auto proj_idx    = index_in_layout(kMasterLayout, "ProjectName");
  const auto machine_idx = index_in_layout(kMasterLayout, "Machine");

  auto time_key = [&](const std::vector<std::string>& row) {
    if (start_idx < row.size()) {
      if (auto t = util::parse_time(row[start_idx])) return *t;
    }
    if (end_idx < row.size()) {
      if (auto t = util::parse_time(row[end_idx])) return *t;
    }
    if (run_idx < row.size() && !row[run_idx].empty()) {
      if (auto t = util::parse_time(row[run_idx] + " 00:00:00")) return *t;
    }
    return std::chrono::system_clock::time_point::min();
  };

  auto field = [](const std::vector<std::string>& row, size_t idx) -> const std::string& {
    static const std::string empty;
    return idx < row.size() ? row[idx] : empty;
  };

  std::sort(rows.begin(), rows.end(), [&](const auto& a, const auto& b) {
    auto ta = time_key(a);
    auto tb = time_key(b);
    if (ta != tb) return ta > tb;
    const std::string& pa = field(a, proj_idx);
    const std::string& pb = field(b, proj_idx);
    if (pa != pb) return pa < pb;
    const std::string& ma = field(a, machine_idx);
    const std::string& mb = field(b, machine_idx);
    return ma < mb;
  });

  std::vector<ColumnSpec> specs;
  specs.reserve(kMasterLayout.size());
  for (const auto& col : kMasterLayout) specs.push_back(col.spec);

  lxw_workbook* wb = workbook_new(xlsx_path.c_str());
  lxw_worksheet* ws = workbook_add_worksheet(wb, "All_Exports");

  write_sectioned_table(wb, ws, specs, rows);
  add_success_format(wb, ws, find_column_index(specs, "Success"), rows.size());

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

