# CMakeLists.txt --- CMake project settings
#    ex) cmake -G "Visual Studio 9 2008" .
#    ex) cmake -DCMAKE_BUILD_TYPE=Release -G "MSYS Makefiles" .
##############################################################################

# CMake minimum version
cmake_minimum_required(VERSION 3.0)

# project name and languages
project(PrefixFilter CXX RC)

##############################################################################

# PrefixFilter.exe
add_executable(PrefixFilter WIN32 PrefixFilter.cpp PrefixFilter_res.rc)
target_link_libraries(PrefixFilter comctl32 shlwapi)

##############################################################################
