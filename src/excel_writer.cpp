#include "excel_writer.hpp"
#include <xlsxwriter.h>
#include <fmt/format.h>
#include <set>

namespace {
std::vector<std::string> pm_headers = {
"ProjectName","BuildID","Machine","HostIP","User","StartTime","EndTime","Duration(hh:mm:ss)","ExportType","Resolution","TileScheme","PhotosUsed","PhotoFolders","PhotoCoverage(km²)","FusersUsed","CPUThreads","GPUCount","OutputFolder","TotalFiles","TotalSize(GB)","Offset_CoordSys","Offset_HDatum","Offset_VDatum","OffsetX","OffsetY","OffsetZ","PivotCenterX","PivotCenterY","PivotCenterZ","FlipYZ","Trim","Collision","VisualLOD","Success","Warnings","Errors","LogPath"};

std::vector<std::string> rm_headers = {
"ProjectName","DatasetName","Machine","HostIP","User","StartTime","EndTime","Duration(hh:mm:ss)","ProcessPreset","ExportType","SelAreaSize(km²)","Resolution","TileScheme","Offset_CoordSys","Offset_HDatum","Offset_VDatum","OffsetX","OffsetY","OffsetZ","FlipYZ","Trim","Collision","OutputFolder","TotalFiles","TotalSize(GB)","Success","Warnings","Errors","LogPath"};

std::vector<std::string> summary_headers = {
"ProjectName","RunDate","Tool","ExportType","Duration(hh:mm:ss)","TotalSize(GB)","PhotosUsed","FusersUsed","Machine","Success","Errors"};
}

namespace excel {

static void write_rows(lxw_worksheet *ws, const std::vector<std::string> &headers, const std::vector<std::vector<std::string>> &rows) {
    for (size_t c=0;c<headers.size();++c)
        worksheet_write_string(ws,0,c,headers[c].c_str(),NULL);
    for (size_t r=0;r<rows.size();++r)
        for (size_t c=0;c<rows[r].size();++c)
            worksheet_write_string(ws,r+1,c,rows[r][c].c_str(),NULL);
    worksheet_add_table(ws,0,0,rows.size(),headers.size()-1,NULL);
    worksheet_freeze_panes(ws,1,0);
    worksheet_set_column(ws,0,headers.size()-1,20,NULL);
}

static void add_success_format(lxw_workbook *wb, lxw_worksheet *ws, int col, size_t rows) {
    lxw_format *green = workbook_add_format(wb);
    format_set_bg_color(green, LXW_COLOR_GREEN);
    lxw_conditional_format cf1{};
    cf1.type = LXW_CONDITIONAL_TYPE_CELL;
    cf1.criteria = LXW_CONDITIONAL_CRITERIA_EQUAL_TO;
    cf1.value_string = const_cast<char*>("True");
    cf1.format = green;
    worksheet_conditional_format_range(ws,1,col,rows,col,&cf1);

    lxw_format *red = workbook_add_format(wb);
    format_set_bg_color(red, LXW_COLOR_RED);
    lxw_conditional_format cf2{};
    cf2.type = LXW_CONDITIONAL_TYPE_CELL;
    cf2.criteria = LXW_CONDITIONAL_CRITERIA_EQUAL_TO;
    cf2.value_string = const_cast<char*>("False");
    cf2.format = red;
    worksheet_conditional_format_range(ws,1,col,rows,col,&cf2);
}

void write_workbook(const std::string &path,
                    const std::vector<PhotoMeshRow> &pm,
                    const std::vector<RealityMeshRow> &rm,
                    const std::vector<SummaryRow> &summary) {
    lxw_workbook *wb = workbook_new(path.c_str());
    lxw_worksheet *ws_pm = workbook_add_worksheet(wb, "PhotoMesh_Exports");
    lxw_worksheet *ws_rm = workbook_add_worksheet(wb, "RealityMesh_Exports");
    lxw_worksheet *ws_sum = workbook_add_worksheet(wb, "Summary");

    std::vector<std::vector<std::string>> pm_rows;
    for (const auto &r: pm) {
        pm_rows.push_back({r.projectName,r.buildID,r.machine,r.hostIP,r.user,r.startTime,r.endTime,r.duration,r.exportType,r.resolution,r.tileScheme,r.photosUsed,r.photoFolders,r.photoCoverage,r.fusersUsed,r.cpuThreads,r.gpuCount,r.outputFolder,r.totalFiles,r.totalSizeGB,r.offsetCoordSys,r.offsetHDatum,r.offsetVDatum,r.offsetX,r.offsetY,r.offsetZ,r.pivotCenterX,r.pivotCenterY,r.pivotCenterZ,r.flipYZ,r.trim,r.collision,r.visualLOD,r.success,r.warnings,r.errors,r.logPath});
    }
    write_rows(ws_pm,pm_headers,pm_rows);
    add_success_format(wb,ws_pm,34,pm_rows.size());

    std::vector<std::vector<std::string>> rm_rows;
    for (const auto &r: rm) {
        rm_rows.push_back({r.projectName,r.datasetName,r.machine,r.hostIP,r.user,r.startTime,r.endTime,r.duration,r.processPreset,r.exportType,r.selAreaSize,r.resolution,r.tileScheme,r.offsetCoordSys,r.offsetHDatum,r.offsetVDatum,r.offsetX,r.offsetY,r.offsetZ,r.flipYZ,r.trim,r.collision,r.outputFolder,r.totalFiles,r.totalSizeGB,r.success,r.warnings,r.errors,r.logPath});
    }
    write_rows(ws_rm,rm_headers,rm_rows);
    add_success_format(wb,ws_rm,25,rm_rows.size());

    std::vector<std::vector<std::string>> sum_rows;
    for (const auto &r: summary) {
        sum_rows.push_back({r.projectName,r.runDate,r.tool,r.exportType,r.duration,r.totalSizeGB,r.photosUsed,r.fusersUsed,r.machine,r.success,r.errors});
    }
    write_rows(ws_sum,summary_headers,sum_rows);
    add_success_format(wb,ws_sum,9,sum_rows.size());

    lxw_worksheet *ws_how = workbook_add_worksheet(wb, "HowTo");
    worksheet_write_string(ws_how,0,0,"Usage:",NULL);
    worksheet_write_string(ws_how,1,0,"logtoExcel --photomesh pm.log --realitymesh rm.log -o Report.xlsx",NULL);

    lxw_worksheet *ws_dict = workbook_add_worksheet(wb, "Data_Dictionary");
    worksheet_write_string(ws_dict,0,0,"Field",NULL);
    worksheet_write_string(ws_dict,0,1,"Description",NULL);
    std::set<std::string> fields(pm_headers.begin(),pm_headers.end());
    fields.insert(rm_headers.begin(),rm_headers.end());
    fields.insert(summary_headers.begin(),summary_headers.end());
    int r=1; for (auto &f: fields) {
        worksheet_write_string(ws_dict,r,0,f.c_str(),NULL);
        worksheet_write_string(ws_dict,r,1,"See README",NULL);
        ++r;
    }
    worksheet_set_column(ws_dict,0,1,30,NULL);

    workbook_close(wb);
}

} // namespace excel

