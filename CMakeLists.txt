cmake_minimum_required(VERSION 2.8)
project(opc_helper)

set(CMAKE_BUILD_TYPE CACHE STRING "Release")
set(PYTHON_DIR "C:/Python34" CACHE STRING "Set the location of the Python interpreter")
set(PYTHON_LIB_FILE "python34" CACHE STRING "Set the name of Python link library")
set(CMAKE_INSTALL_PREFIX "${PYTHON_DIR}/Lib/site-packages"
            CACHE STRING "Location where to install python module" FORCE)

set(SOURCES
	opc/opccomn_i.c
	opc/opcda_i.c
	opc/OpcEnum_i.c
	com_funcs.cpp
	opcenumquery.cpp
	opcserver.cpp
	module_main.cpp)

include_directories("${PYTHON_DIR}/include")
link_directories("${PYTHON_DIR}/libs")

add_library(opc_helper SHARED ${SOURCES})
target_link_libraries(opc_helper ${PYTHON_LIB_FILE})
set_target_properties(opc_helper PROPERTIES PREFIX "")
set_target_properties(opc_helper PROPERTIES SUFFIX ".pyd")

# don't specify ARCHIVE DESTINATION to not install .lib file
install(TARGETS "opc_helper"
    RUNTIME DESTINATION "./"
    LIBRARY DESTINATION "./"
#    ARCHIVE DESTINATION "./"
)
