#include "cli.hpp"
#include "photomesh_parser.hpp"
#include "realitymesh_parser.hpp"
#include "excel_writer.hpp"
#include "models.hpp"

#include <fmt/printf.h>
#include <fstream>
#include <string>
#include <vector>

#ifdef _WIN32
#include <combaseapi.h> // CoInitializeEx/CoUninitialize
#include "win_file_dialogs.hpp"
#endif

// Simple classifier to detect which parser to use for a raw .log path
enum class LogKind { PhotoMesh, RealityMesh, Unknown };

static LogKind detect_log_type(const std::string& path) {
    std::ifstream in(path);
    if (!in) return LogKind::Unknown;
    std::string line;
    size_t limit = 4000; // scan up to ~4k lines
    while (limit-- && std::getline(in, line)) {
        // PhotoMesh markers (see photomesh_parser)  [scan order matters]
        if (line.find("SLDEFAULT=>") != std::string::npos ||
            line.find("\"MachineName\"") != std::string::npos ||
            line.find("\"MsgTime\"")    != std::string::npos) {
            return LogKind::PhotoMesh;
        }
        // RealityMesh markers (see realitymesh_parser)
        if (line.find("Time to run TT project:") != std::string::npos ||
            line.find("-command_file")           != std::string::npos ||
            line.find("Process completed with exit code:") != std::string::npos ||
            line.find("Converted offset:")       != std::string::npos) {
            return LogKind::RealityMesh;
        }
    }
    return LogKind::Unknown;
}

static void classify_paths_into_options(const std::vector<std::string>& paths, Options& opt) {
    for (const auto& p : paths) {
        switch (detect_log_type(p)) {
        case LogKind::PhotoMesh:   opt.photomeshLogs.push_back(p);   break;
        case LogKind::RealityMesh: opt.realitymeshLogs.push_back(p); break;
        case LogKind::Unknown:     // default to PhotoMesh if unknown
            opt.photomeshLogs.push_back(p);
            break;
        }
    }
}

int main(int argc, char** argv) {
    // 1) Parse CLI flags as before (kept compatible)
    Options opt = parse_cli(argc, argv);

    // 2) Drag & drop support: any bare arg that doesn't start with '-' is treated as a raw log path.
    std::vector<std::string> barePaths;
    for (int i = 1; i < argc; ++i) {
        std::string a = argv[i];
        if (!a.empty() && a[0] != '-') barePaths.push_back(a);
    }
    if (!barePaths.empty()) {
        classify_paths_into_options(barePaths, opt);
    }

#ifdef _WIN32
    // 3) GUI path: if user double-clicked the exe (no logs via flags or bare paths), show dialogs
    if (opt.photomeshLogs.empty() && opt.realitymeshLogs.empty()) {
        CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

        auto pmSel = open_logs_dialog(L"Select PhotoMesh log(s)", true);
        auto rmSel = open_logs_dialog(L"Select RealityMesh log(s)", true);

        classify_paths_into_options(pmSel, opt);
        classify_paths_into_options(rmSel, opt);

        if (opt.output.empty()) {
            std::string chosen = save_xlsx_dialog(L"Save Excel report as...", L"Report.xlsx");
            if (!chosen.empty()) opt.output = chosen;
        }

        CoUninitialize();
    }
#endif

    // 4) Final validation / defaults
    if (opt.photomeshLogs.empty() && opt.realitymeshLogs.empty()) {
        fmt::print(stderr,
            "Drag one or more PhotoMesh/RealityMesh logs onto the EXE,\n"
            "or run with: logtoExcel --photomesh <pm.log> --realitymesh <rm.log> -o out.xlsx\n");
        return 2;
    }
    if (opt.output.empty()) opt.output = "Report.xlsx";

    // 5) Parse and write the workbook
    std::vector<PhotoMeshRow> pm_rows;
    for (const auto& p : opt.photomeshLogs) pm_rows.push_back(parse_photomesh(p));

    std::vector<RealityMeshRow> rm_rows;
    for (const auto& p : opt.realitymeshLogs) rm_rows.push_back(parse_realitymesh(p));

    std::vector<SummaryRow> summary;
    for (const auto& r : pm_rows) summary.push_back(make_summary(r));
    for (const auto& r : rm_rows) summary.push_back(make_summary(r));

    excel::write_workbook(opt.output, pm_rows, rm_rows, summary);
    fmt::print("Parsed {} PhotoMesh row(s), {} RealityMesh row(s) -> {}\n",
               pm_rows.size(), rm_rows.size(), opt.output);
    return 0;
}

