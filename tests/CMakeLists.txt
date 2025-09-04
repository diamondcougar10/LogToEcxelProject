add_test(NAME basic COMMAND ${CMAKE_COMMAND} -DTEST_EXE=$<TARGET_FILE:logtoExcel> -P ${CMAKE_CURRENT_SOURCE_DIR}/run_basic_test.cmake)
