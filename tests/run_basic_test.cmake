execute_process(
  COMMAND ${TEST_EXE} --photomesh sample_pm.log --realitymesh sample_rm.log -o out.xlsx
  WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
  RESULT_VARIABLE result)
if(NOT result EQUAL 0)
  message(FATAL_ERROR "logtoExcel returned ${result}")
endif()
if(NOT EXISTS ${CMAKE_CURRENT_SOURCE_DIR}/out.xlsx)
  message(FATAL_ERROR "No output workbook")
endif()
file(SIZE ${CMAKE_CURRENT_SOURCE_DIR}/out.xlsx size)
if(size LESS 1000)
  message(FATAL_ERROR "Workbook too small")
endif()
