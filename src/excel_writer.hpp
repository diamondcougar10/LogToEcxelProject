#pragma once
#include "models.hpp"
#include <vector>
#include <string>

namespace excel {
void write_workbook(const std::string &path,
                    const std::vector<PhotoMeshRow> &pm,
                    const std::vector<RealityMeshRow> &rm,
                    const std::vector<SummaryRow> &summary);
}

