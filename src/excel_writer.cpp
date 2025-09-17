#include "excel_writer.hpp"
#include "util_time.hpp"

#include <xlsxwriter.h>

#include <algorithm>
#include <chrono>
#include <set>
#include <string>
#include <vector>

namespace {

struct ColumnSpec {
    std::string title;
    std::string group;
    double      width;
    bool        wrap;
};

template <typename Row>
struct ColumnWithMember {
    ColumnSpec spec;
    const std::string Row::* member;
};

using Clock = std::chrono::system_clock;

Clock::time_point parse_time_or_min(const std::string& primary,
                                    const std::string& fallback = {}) {
    if (auto t = util::parse_time(primary)) return *t;
    if (!fallback.empty())
        if (auto t = util::parse_time(fallback)) return *t;
    return Clock::time_point::min();
}

Clock::time_point parse_date_or_min(const std::string& date_only) {
    if (date_only.empty()) return Clock::time_point::min();
    return util::parse_time(date_only + " 00:00:00").value_or(Clock::time_point::min());
}

template <typename Row>
std::vector<std::vector<std::string>>
materialize_rows(const std::vector<Row>& rows,
                 const std::vector<ColumnWithMember<Row>>& columns) {
    std::vector<std::vector<std::string>> out;
    out.reserve(rows.size());
    for (const auto& r : rows) {
        std::vector<std::string> row;
        row.reserve(columns.size());
        for (const auto& c : columns) {
            row.push_back(r.*(c.member));
        }
        out.push_back(std::move(row));
    }
    return out;
}

template <typename Row>
std::vector<ColumnSpec> collect_specs(const std::vector<ColumnWithMember<Row>>& columns) {
    std::vector<ColumnSpec> specs;
    specs.reserve(columns.size());
    for (const auto& c : columns) specs.push_back(c.spec);
    return specs;
}

static const std::vector<ColumnWithMember<PhotoMeshRow>> kPmColumns = {
    {{"Project Name",            "Project Overview",  26.0, false}, &PhotoMeshRow::projectName},
    {{"Export Type",             "Project Overview",  18.0, false}, &PhotoMeshRow::exportType},
    {{"Resolution",              "Project Overview",  14.0, false}, &PhotoMeshRow::resolution},
    {{"Tile Scheme",             "Project Overview",  14.0, false}, &PhotoMeshRow::tileScheme},
    {{"Build ID",                "Project Overview",  16.0, false}, &PhotoMeshRow::buildID},

    {{"Start Time",              "Timing",            20.0, false}, &PhotoMeshRow::startTime},
    {{"End Time",                "Timing",            20.0, false}, &PhotoMeshRow::endTime},
    {{"Duration (hh:mm:ss)",     "Timing",            18.0, false}, &PhotoMeshRow::duration},

    {{"Computer Name",           "Machine",           22.0, false}, &PhotoMeshRow::machine},
    {{"Host IP",                 "Machine",           16.0, false}, &PhotoMeshRow::hostIP},
    {{"User",                    "Machine",           16.0, false}, &PhotoMeshRow::user},

    {{"Photos Used",             "Resources",         12.0, false}, &PhotoMeshRow::photosUsed},
    {{"Photo Folders",           "Resources",         28.0, true }, &PhotoMeshRow::photoFolders},
    {{"Photo Coverage (km²)",    "Resources",         16.0, false}, &PhotoMeshRow::photoCoverage},
    {{"Fusers Used",             "Resources",         12.0, false}, &PhotoMeshRow::fusersUsed},
    {{"CPU Threads",             "Resources",         12.0, false}, &PhotoMeshRow::cpuThreads},
    {{"GPU Count",               "Resources",         12.0, false}, &PhotoMeshRow::gpuCount},

    {{"Output Folder",           "Output",            36.0, true }, &PhotoMeshRow::outputFolder},
    {{"Total Files",             "Output",            12.0, false}, &PhotoMeshRow::totalFiles},
    {{"Total Size (GB)",         "Output",            14.0, false}, &PhotoMeshRow::totalSizeGB},

    {{"Offset Coord Sys",        "Offsets & Settings",22.0, false}, &PhotoMeshRow::offsetCoordSys},
    {{"Offset H Datum",          "Offsets & Settings",18.0, false}, &PhotoMeshRow::offsetHDatum},
    {{"Offset V Datum",          "Offsets & Settings",18.0, false}, &PhotoMeshRow::offsetVDatum},
    {{"Offset X",                "Offsets & Settings",12.0, false}, &PhotoMeshRow::offsetX},
    {{"Offset Y",                "Offsets & Settings",12.0, false}, &PhotoMeshRow::offsetY},
    {{"Offset Z",                "Offsets & Settings",12.0, false}, &PhotoMeshRow::offsetZ},
    {{"Pivot Center X",          "Offsets & Settings",14.0, false}, &PhotoMeshRow::pivotCenterX},
    {{"Pivot Center Y",          "Offsets & Settings",14.0, false}, &PhotoMeshRow::pivotCenterY},
    {{"Pivot Center Z",          "Offsets & Settings",14.0, false}, &PhotoMeshRow::pivotCenterZ},
    {{"Flip YZ",                 "Offsets & Settings",10.0, false}, &PhotoMeshRow::flipYZ},
    {{"Trim",                    "Offsets & Settings",10.0, false}, &PhotoMeshRow::trim},
    {{"Collision",               "Offsets & Settings",12.0, false}, &PhotoMeshRow::collision},
    {{"Visual LOD",              "Offsets & Settings",12.0, false}, &PhotoMeshRow::visualLOD},

    {{"Success",                 "Status",            10.0, false}, &PhotoMeshRow::success},
    {{"Warnings",                "Status",            26.0, true }, &PhotoMeshRow::warnings},
    {{"Errors",                  "Status",            26.0, true }, &PhotoMeshRow::errors},
    {{"Log Path",                "Status",            42.0, true }, &PhotoMeshRow::logPath},
};

static const std::vector<ColumnWithMember<RealityMeshRow>> kRmColumns = {
    {{"Project Name",            "Project Overview",  26.0, false}, &RealityMeshRow::projectName},
    {{"Dataset Name",            "Project Overview",  24.0, false}, &RealityMeshRow::datasetName},
    {{"Export Type",             "Project Overview",  18.0, false}, &RealityMeshRow::exportType},
    {{"Process Preset",          "Project Overview",  18.0, false}, &RealityMeshRow::processPreset},
    {{"Selection Area (km²)",    "Project Overview",  18.0, false}, &RealityMeshRow::selAreaSize},
    {{"Resolution",              "Project Overview",  14.0, false}, &RealityMeshRow::resolution},
    {{"Tile Scheme",             "Project Overview",  14.0, false}, &RealityMeshRow::tileScheme},

    {{"Start Time",              "Timing",            20.0, false}, &RealityMeshRow::startTime},
    {{"End Time",                "Timing",            20.0, false}, &RealityMeshRow::endTime},
    {{"Duration (hh:mm:ss)",     "Timing",            18.0, false}, &RealityMeshRow::duration},

    {{"Computer Name",           "Machine",           22.0, false}, &RealityMeshRow::machine},
    {{"Host IP",                 "Machine",           16.0, false}, &RealityMeshRow::hostIP},
    {{"User",                    "Machine",           16.0, false}, &RealityMeshRow::user},

    {{"Output Folder",           "Output",            36.0, true }, &RealityMeshRow::outputFolder},
    {{"Total Files",             "Output",            12.0, false}, &RealityMeshRow::totalFiles},
    {{"Total Size (GB)",         "Output",            14.0, false}, &RealityMeshRow::totalSizeGB},

    {{"Offset Coord Sys",        "Offsets & Settings",22.0, false}, &RealityMeshRow::offsetCoordSys},
    {{"Offset H Datum",          "Offsets & Settings",18.0, false}, &RealityMeshRow::offsetHDatum},
    {{"Offset V Datum",          "Offsets & Settings",18.0, false}, &RealityMeshRow::offsetVDatum},
    {{"Offset X",                "Offsets & Settings",12.0, false}, &RealityMeshRow::offsetX},
    {{"Offset Y",                "Offsets & Settings",12.0, false}, &RealityMeshRow::offsetY},
    {{"Offset Z",                "Offsets & Settings",12.0, false}, &RealityMeshRow::offsetZ},
    {{"Pivot Center X",          "Offsets & Settings",14.0, false}, &RealityMeshRow::pivotCenterX},
    {{"Pivot Center Y",          "Offsets & Settings",14.0, false}, &RealityMeshRow::pivotCenterY},
    {{"Pivot Center Z",          "Offsets & Settings",14.0, false}, &RealityMeshRow::pivotCenterZ},
    {{"Flip YZ",                 "Offsets & Settings",10.0, false}, &RealityMeshRow::flipYZ},
    {{"Trim",                    "Offsets & Settings",10.0, false}, &RealityMeshRow::trim},
    {{"Collision",               "Offsets & Settings",12.0, false}, &RealityMeshRow::collision},

    {{"Success",                 "Status",            10.0, false}, &RealityMeshRow::success},
    {{"Warnings",                "Status",            26.0, true }, &RealityMeshRow::warnings},
    {{"Errors",                  "Status",            26.0, true }, &RealityMeshRow::errors},
    {{"Log Path",                "Status",            42.0, true }, &RealityMeshRow::logPath},
};

static const std::vector<ColumnWithMember<SummaryRow>> kSummaryColumns = {
    {{"Project Name",            "At a Glance",       28.0, false}, &SummaryRow::projectName},
    {{"Computer Name",           "At a Glance",       22.0, false}, &SummaryRow::machine},
    {{"Run Date",                "At a Glance",       14.0, false}, &SummaryRow::runDate},
    {{"Duration (hh:mm:ss)",     "At a Glance",       16.0, false}, &SummaryRow::duration},
    {{"Total Size (GB)",         "At a Glance",       14.0, false}, &SummaryRow::totalSizeGB},

    {{"Tool",                    "Details",           16.0, false}, &SummaryRow::tool},
    {{"Export Type",             "Details",           20.0, false}, &SummaryRow::exportType},

    {{"Photos Used",             "Resources",         12.0, false}, &SummaryRow::photosUsed},
    {{"Fusers Used",             "Resources",         12.0, false}, &SummaryRow::fusersUsed},

    {{"Success",                 "Status",            10.0, false}, &SummaryRow::success},
    {{"Errors",                  "Status",            26.0, true }, &SummaryRow::errors},
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
        worksheet_write_string(ws, 1, c, columns[c].title.c_str(), header_fmt);
        worksheet_set_column(ws, c, c, columns[c].width,
                             columns[c].wrap ? wrap_fmt : nullptr);
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

} // namespace

namespace excel {

void write_workbook(const std::string& path,
                    const std::vector<PhotoMeshRow>& pm,
                    const std::vector<RealityMeshRow>& rm,
                    const std::vector<SummaryRow>& summary) {
    lxw_workbook* wb = workbook_new(path.c_str());
    lxw_worksheet* ws_pm = workbook_add_worksheet(wb, "PhotoMesh_Exports");
    lxw_worksheet* ws_rm = workbook_add_worksheet(wb, "RealityMesh_Exports");
    lxw_worksheet* ws_sum = workbook_add_worksheet(wb, "Summary");

    std::vector<PhotoMeshRow> pm_sorted(pm.begin(), pm.end());
    std::sort(pm_sorted.begin(), pm_sorted.end(), [](const PhotoMeshRow& a, const PhotoMeshRow& b) {
        auto ta = parse_time_or_min(a.startTime, a.endTime);
        auto tb = parse_time_or_min(b.startTime, b.endTime);
        if (ta != tb) return ta > tb;
        return a.projectName < b.projectName;
    });
    auto pm_rows = materialize_rows(pm_sorted, kPmColumns);
    auto pm_specs = collect_specs(kPmColumns);
    write_sectioned_table(wb, ws_pm, pm_specs, pm_rows);
    add_success_format(wb, ws_pm, find_column_index(pm_specs, "Success"), pm_rows.size());

    std::vector<RealityMeshRow> rm_sorted(rm.begin(), rm.end());
    std::sort(rm_sorted.begin(), rm_sorted.end(), [](const RealityMeshRow& a, const RealityMeshRow& b) {
        auto ta = parse_time_or_min(a.startTime, a.endTime);
        auto tb = parse_time_or_min(b.startTime, b.endTime);
        if (ta != tb) return ta > tb;
        return a.projectName < b.projectName;
    });
    auto rm_rows = materialize_rows(rm_sorted, kRmColumns);
    auto rm_specs = collect_specs(kRmColumns);
    write_sectioned_table(wb, ws_rm, rm_specs, rm_rows);
    add_success_format(wb, ws_rm, find_column_index(rm_specs, "Success"), rm_rows.size());

    std::vector<SummaryRow> summary_sorted(summary.begin(), summary.end());
    std::sort(summary_sorted.begin(), summary_sorted.end(), [](const SummaryRow& a, const SummaryRow& b) {
        auto ta = parse_date_or_min(a.runDate);
        auto tb = parse_date_or_min(b.runDate);
        if (ta != tb) return ta > tb;
        if (a.projectName != b.projectName) return a.projectName < b.projectName;
        return a.machine < b.machine;
    });
    auto sum_rows = materialize_rows(summary_sorted, kSummaryColumns);
    auto sum_specs = collect_specs(kSummaryColumns);
    write_sectioned_table(wb, ws_sum, sum_specs, sum_rows);
    add_success_format(wb, ws_sum, find_column_index(sum_specs, "Success"), sum_rows.size());

    lxw_worksheet* ws_how = workbook_add_worksheet(wb, "HowTo");
    worksheet_write_string(ws_how, 0, 0, "Usage:", nullptr);
    worksheet_write_string(ws_how, 1, 0,
                           "logtoExcel --photomesh pm.log --realitymesh rm.log -o Report.xlsx",
                           nullptr);

    lxw_worksheet* ws_dict = workbook_add_worksheet(wb, "Data_Dictionary");
    worksheet_write_string(ws_dict, 0, 0, "Field", nullptr);
    worksheet_write_string(ws_dict, 0, 1, "Description", nullptr);
    std::set<std::string> fields;
    for (const auto& c : kPmColumns) fields.insert(c.spec.title);
    for (const auto& c : kRmColumns) fields.insert(c.spec.title);
    for (const auto& c : kSummaryColumns) fields.insert(c.spec.title);
    int row = 1;
    for (const auto& f : fields) {
        worksheet_write_string(ws_dict, row, 0, f.c_str(), nullptr);
        worksheet_write_string(ws_dict, row, 1, "See README", nullptr);
        ++row;
    }
    worksheet_set_column(ws_dict, 0, 1, 32.0, nullptr);

    workbook_close(wb);
}

} // namespace excel

