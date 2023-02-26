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

set(src_dir ${CMAKE_CURRENT_SOURCE_DIR}/demo_files)
set(save_src_dir ${CMAKE_CURRENT_BINARY_DIR}/demo_files)

execute_process(COMMAND ${CMAKE_COMMAND} -E make_directory "${save_src_dir}")
execute_process(COMMAND ${CMAKE_COMMAND} -E copy_directory  "${src_dir}" "${save_src_dir}")

set(QRC_SOURCE_FILES xml_models.qrc)

find_package(QT NAMES Qt6 Qt5 COMPONENTS Core Gui REQUIRED)
find_package(Qt${QT_VERSION_MAJOR} COMPONENTS Core Gui REQUIRED)
find_package(OpenCV REQUIRED)

include_directories(${OpenCV_INCLUDE_DIRS})

qt5_add_resources(QRC_FILES ${QRC_SOURCE_FILES})

SOURCE_GROUP("Resource Files" FILES ${QRC_SOURCE_FILES})

add_subdirectory(zipper)

add_executable(wordx
  demo.cpp
  fileop.cpp
  wordop.cpp
  
  info.cpp
  ${QRC_FILES}
)

link_directories(${OpenCV_LIBRARY_DIRS})
target_link_libraries(wordx Qt${QT_VERSION_MAJOR}::Core Qt${QT_VERSION_MAJOR}::Gui ${OpenCV_LIBS} zipper)



