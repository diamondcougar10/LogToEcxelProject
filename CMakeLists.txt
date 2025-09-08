cmake_minimum_required(VERSION 3.20)
project(logtoexcel LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# ---------- vcpkg installed root (used for lookups & DLL copy) ----------
set(_VCPKG_INSTALLED_ROOT "$ENV{VCPKG_ROOT}/installed/x64-windows")
if(NOT EXISTS "${_VCPKG_INSTALLED_ROOT}" AND EXISTS "C:/vcpkg/installed/x64-windows")
  set(_VCPKG_INSTALLED_ROOT "C:/vcpkg/installed/x64-windows")
endif()

# ---------- fmt (header-only => no fmt DLLs) ----------
find_package(fmt CONFIG REQUIRED)

# ---------- libxlsxwriter: try pkg-config, then fallback ----------
find_package(PkgConfig QUIET)
if(PkgConfig_FOUND)
  # Help pkgconf find the .pc files from vcpkg
  set(ENV{PKG_CONFIG_PATH}
      "${_VCPKG_INSTALLED_ROOT}/lib/pkgconfig;${_VCPKG_INSTALLED_ROOT}/debug/lib/pkgconfig;${_VCPKG_INSTALLED_ROOT}/share/pkgconfig")
  if(NOT PKG_CONFIG_EXECUTABLE AND EXISTS "${_VCPKG_INSTALLED_ROOT}/tools/pkgconf/pkgconf.exe")
    set(PKG_CONFIG_EXECUTABLE "${_VCPKG_INSTALLED_ROOT}/tools/pkgconf/pkgconf.exe" CACHE FILEPATH "" FORCE)
  endif()
  pkg_check_modules(XLSXWRITER QUIET IMPORTED_TARGET xlsxwriter)
endif()

if(NOT TARGET PkgConfig::XLSXWRITER)
  # Fallback: import directly from vcpkg's installed tree (Debug/Release)
  find_path(XLSXWRITER_INCLUDE_DIR NAMES xlsxwriter.h
            PATHS "${_VCPKG_INSTALLED_ROOT}/include" REQUIRED)
  find_library(XLSXWRITER_LIBRARY_RELEASE NAMES xlsxwriter
               PATHS "${_VCPKG_INSTALLED_ROOT}/lib")
  find_library(XLSXWRITER_LIBRARY_DEBUG   NAMES xlsxwriter
               PATHS "${_VCPKG_INSTALLED_ROOT}/debug/lib")

  add_library(libxlsxwriter UNKNOWN IMPORTED)
  set_target_properties(libxlsxwriter PROPERTIES
    INTERFACE_INCLUDE_DIRECTORIES "${XLSXWRITER_INCLUDE_DIR}"
    IMPORTED_LOCATION_RELEASE "${XLSXWRITER_LIBRARY_RELEASE}"
    IMPORTED_LOCATION_DEBUG   "${XLSXWRITER_LIBRARY_DEBUG}"
  )
  set(XLSXWRITER_TARGET libxlsxwriter)
else()
  set(XLSXWRITER_TARGET PkgConfig::XLSXWRITER)
endif()

# ---- shared library stays the same ----
add_library(logtoexcel_lib
  src/excel_writer.cpp
  src/photomesh_parser.cpp
  src/realitymesh_parser.cpp
  src/util_time.cpp
  src/models.cpp
  src/single_sheet_writer.cpp
  $<$<PLATFORM_ID:Windows>:src/win_file_dialogs.cpp>
)
target_include_directories(logtoexcel_lib PUBLIC ${CMAKE_CURRENT_SOURCE_DIR}/src)
target_link_libraries(logtoexcel_lib PUBLIC fmt::fmt-header-only ${XLSXWRITER_TARGET})

# ---- GUI exe (default app on Windows; no startup prompts) ----
if (WIN32)
  add_executable(logtoExcel_gui WIN32 src/gui_main.cpp)
  target_link_libraries(logtoExcel_gui PRIVATE logtoexcel_lib Ole32 Shell32 Uuid Comctl32)

  # copy vcpkg runtime DLLs (debug/release) next to the GUI exe
  if(EXISTS "${_VCPKG_INSTALLED_ROOT}/debug/bin")
    file(GLOB _VCPKG_DLLS_DEBUG   "${_VCPKG_INSTALLED_ROOT}/debug/bin/*.dll")
  endif()
  if(EXISTS "${_VCPKG_INSTALLED_ROOT}/bin")
    file(GLOB _VCPKG_DLLS_RELEASE "${_VCPKG_INSTALLED_ROOT}/bin/*.dll")
  endif()

  if(_VCPKG_DLLS_DEBUG)
    add_custom_command(TARGET logtoExcel_gui POST_BUILD
      COMMAND ${CMAKE_COMMAND} -E copy_if_different
              ${_VCPKG_DLLS_DEBUG}
              "$<TARGET_FILE_DIR:logtoExcel_gui>"
      COMMENT "Copying vcpkg DEBUG DLLs for GUI"
      VERBATIM
      CONFIGURATIONS Debug)
  endif()
  if(_VCPKG_DLLS_RELEASE)
    add_custom_command(TARGET logtoExcel_gui POST_BUILD
      COMMAND ${CMAKE_COMMAND} -E copy_if_different
              ${_VCPKG_DLLS_RELEASE}
              "$<TARGET_FILE_DIR:logtoExcel_gui>"
      COMMENT "Copying vcpkg RELEASE DLLs for GUI"
      VERBATIM
      CONFIGURATIONS Release RelWithDebInfo MinSizeRel)
  endif()
endif()

# ---- CLI exe (for scripts/automation; unchanged flags) ----
add_executable(logtoExcel_cli src/main.cpp)
target_link_libraries(logtoExcel_cli PRIVATE logtoexcel_lib)

# copy DLLs for CLI too
if(_VCPKG_DLLS_DEBUG)
  add_custom_command(TARGET logtoExcel_cli POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
            ${_VCPKG_DLLS_DEBUG}
            "$<TARGET_FILE_DIR:logtoExcel_cli>"
    COMMENT "Copying vcpkg DEBUG DLLs for CLI"
    VERBATIM
    CONFIGURATIONS Debug)
endif()
if(_VCPKG_DLLS_RELEASE)
  add_custom_command(TARGET logtoExcel_cli POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
            ${_VCPKG_DLLS_RELEASE}
            "$<TARGET_FILE_DIR:logtoExcel_cli>"
    COMMENT "Copying vcpkg RELEASE DLLs for CLI"
    VERBATIM
    CONFIGURATIONS Release RelWithDebInfo MinSizeRel)
endif()

