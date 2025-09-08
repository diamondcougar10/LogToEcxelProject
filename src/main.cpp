#include "cli.hpp"
#include "photomesh_parser.hpp"
#include "realitymesh_parser.hpp"
#include "excel_writer.hpp"
#include "single_sheet_writer.hpp"
#include "models.hpp"

#include <fmt/printf.h>
#include <filesystem>
#include <string>
#include <vector>

#ifdef _WIN32
#include <combaseapi.h>
#include "win_file_dialogs.hpp" // optional, if you added GUI earlier
#endif

namespace fs = std::filesystem;

static void parse_extra_flags(int argc, char** argv,
                              std::string& outputsDir,
                              bool& doMaster,
                              bool& doReport) {
  outputsDir = "EXCEL OUTPUTS";
  doMaster = true; doReport = true;
  for (int i=1;i<argc;++i) {
    std::string a = argv[i];
    if (a == "--outputs-dir" && i+1 < argc)      outputsDir = argv[++i];
    else if (a == "--no-master")                 doMaster = false;
    else if (a == "--no-report")                 doReport = false;
    else if (a == "--single-only")               { doMaster = true; doReport = false; }
  }
}

int main(int argc, char **argv) {
  Options opt = parse_cli(argc, argv);

  // Also accept raw paths (drag & drop). Classify by extension; parser is robust anyway.
  for (int i=1;i<argc;++i) {
    std::string a = argv[i];
    if (!a.empty() && a[0] != '-') {
      // push to both; parser won't crash if log doesn't match
      opt.photomeshLogs.push_back(a);
      opt.realitymeshLogs.push_back(a);
    }
  }

  std::string outputsDir; bool doMaster, doReport;
  parse_extra_flags(argc, argv, outputsDir, doMaster, doReport);

#ifdef _WIN32
  if (opt.photomeshLogs.empty() && opt.realitymeshLogs.empty()) {
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
    auto pm = open_logs_dialog(L"Select PhotoMesh log(s)", true);
    auto rm = open_logs_dialog(L"Select RealityMesh log(s)", true);
    for (auto& s : pm) opt.photomeshLogs.push_back(s);
    for (auto& s : rm) opt.realitymeshLogs.push_back(s);
    if (opt.output.empty()) {
      std::string save = save_xlsx_dialog(L"Save per-run report (optional)", L"Report.xlsx");
      if (!save.empty()) opt.output = save; else doReport = false;
    }
    CoUninitialize();
  }
#endif

  if (opt.photomeshLogs.empty() && opt.realitymeshLogs.empty()) {
    fmt::print(stderr,
      "Usage: logtoExcel --photomesh <pm.log> --realitymesh <rm.log> -o out.xlsx "
      "[--outputs-dir <folder>] [--no-master] [--no-report|--single-only]\n");
    return 2;
  }

  // Parse logs
  std::vector<PhotoMeshRow> pm_rows;
  for (const auto &p : opt.photomeshLogs) pm_rows.push_back(parse_photomesh(p));
  std::vector<RealityMeshRow> rm_rows;
  for (const auto &p : opt.realitymeshLogs) rm_rows.push_back(parse_realitymesh(p));

  // Master single-sheet: unify -> append -> rebuild xlsx
  if (doMaster) {
    auto unified = excel::unify(pm_rows, rm_rows);
    excel::append_to_master_and_rebuild_xlsx(outputsDir, unified);
  }

  // Per-run multi-sheet report (existing)
  if (doReport) {
    if (opt.output.empty()) {
      fs::create_directories(outputsDir);
      opt.output = (fs::path(outputsDir)/"Report.xlsx").string();
    }
    std::vector<SummaryRow> summary;
    for (const auto &r : pm_rows) summary.push_back(make_summary(r));
    for (const auto &r : rm_rows) summary.push_back(make_summary(r));
    excel::write_workbook(opt.output, pm_rows, rm_rows, summary); // existing function
  }

  fmt::print("Done. Master: {}/All_Exports.xlsx{}{}\n",
             outputsDir,
             doReport ? ", Per-run: " : "",
             doReport ? opt.output : "");
  return 0;
}

