cmake_minimum_required(VERSION 3.20)
project(logtoexcel LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(fmt CONFIG REQUIRED)
find_package(PkgConfig REQUIRED)
pkg_check_modules(XLSXWRITER REQUIRED xlsxwriter)

add_library(logtoexcel_lib
    src/excel_writer.cpp
    src/photomesh_parser.cpp
    src/realitymesh_parser.cpp
    src/util_time.cpp
    src/models.cpp
)

# Headers only
set_target_properties(logtoexcel_lib PROPERTIES
    PUBLIC_HEADER "src/cli.hpp;src/excel_writer.hpp;src/photomesh_parser.hpp;src/realitymesh_parser.hpp;src/models.hpp;src/util_time.hpp"
)

target_include_directories(logtoexcel_lib PUBLIC ${CMAKE_CURRENT_SOURCE_DIR}/src ${XLSXWRITER_INCLUDE_DIRS})

target_link_libraries(logtoexcel_lib PUBLIC fmt::fmt ${XLSXWRITER_LIBRARIES})

add_executable(logtoExcel src/main.cpp)

target_link_libraries(logtoExcel PRIVATE logtoexcel_lib)

enable_testing()
add_subdirectory(tests)
