cmake_minimum_required(VERSION 3.14)

project(wordx LANGUAGES CXX)

file(GLOB_RECURSE QRC_SOURCE_FILES ${CMAKE_CURRENT_SOURCE_DIR}/*.qrc)

include_directories(zipper)

set(CMAKE_INCLUDE_CURRENT_DIR ON)

set(CMAKE_AUTOUIC ON)
set(CMAKE_AUTOMOC ON)
set(CMAKE_AUTORCC ON)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(public_include ${CMAKE_CURRENT_SOURCE_DIR}/xml_models)
set(save_src_dir ${CMAKE_CURRENT_SOURCE_DIR}/xml_models)

set(QRC_SOURCE_FILES xml_models.qrc)

find_package(QT NAMES Qt6 Qt5 COMPONENTS Core REQUIRED)
find_package(Qt${QT_VERSION_MAJOR} COMPONENTS Core REQUIRED)


qt5_add_resources(QRC_FILES ${QRC_SOURCE_FILES})

SOURCE_GROUP("Resource Files" FILES ${QRC_SOURCE_FILES})

add_subdirectory(zipper)

add_executable(wordx
  main.cpp
  fileop.cpp
  wordop.cpp
  
  ${QRC_FILES}
)

target_link_libraries(wordx Qt${QT_VERSION_MAJOR}::Core zipper)



