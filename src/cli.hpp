#pragma once
#include <string>
#include <vector>

struct Options {
    std::vector<std::string> photomeshLogs;
    std::vector<std::string> realitymeshLogs;
    std::string output = "Photomesh_RealityMesh_LogReport.xlsx";
};

inline Options parse_cli(int argc, char **argv) {
    Options opt;
    for (int i = 1; i < argc; ++i) {
        std::string a = argv[i];
        if (a == "--photomesh") {
            while (i + 1 < argc && argv[i + 1][0] != '-') {
                opt.photomeshLogs.emplace_back(argv[++i]);
            }
        } else if (a == "--realitymesh") {
            while (i + 1 < argc && argv[i + 1][0] != '-') {
                opt.realitymeshLogs.emplace_back(argv[++i]);
            }
        } else if (a == "-o" || a == "--output") {
            if (i + 1 < argc) opt.output = argv[++i];
        }
    }
    return opt;
}

