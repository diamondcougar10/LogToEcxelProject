#include "cli.hpp"
#include "photomesh_parser.hpp"
#include "realitymesh_parser.hpp"
#include "excel_writer.hpp"
#include "models.hpp"
#include <fmt/printf.h>

int main(int argc, char **argv) {
    Options opt = parse_cli(argc, argv);
    if (opt.photomeshLogs.empty() && opt.realitymeshLogs.empty()) {
        fmt::print(stderr, "Usage: logtoExcel --photomesh <pm.log> --realitymesh <rm.log> -o out.xlsx\n");
        return 2;
    }
    std::vector<PhotoMeshRow> pm_rows;
    for (const auto &p : opt.photomeshLogs) {
        pm_rows.push_back(parse_photomesh(p));
    }
    std::vector<RealityMeshRow> rm_rows;
    for (const auto &p : opt.realitymeshLogs) {
        rm_rows.push_back(parse_realitymesh(p));
    }
    std::vector<SummaryRow> summary;
    for (const auto &r : pm_rows) summary.push_back(make_summary(r));
    for (const auto &r : rm_rows) summary.push_back(make_summary(r));
    excel::write_workbook(opt.output, pm_rows, rm_rows, summary);
    fmt::print("Parsed {} PhotoMesh rows, {} RealityMesh rows -> {}\n", pm_rows.size(), rm_rows.size(), opt.output);
    return 0;
}

