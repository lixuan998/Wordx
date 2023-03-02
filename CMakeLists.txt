cmake_minimum_required(VERSION 3.14)

project(wordx LANGUAGES CXX)

include_directories(zipper)

set(CMAKE_INCLUDE_CURRENT_DIR ON)

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(QT NAMES Qt6 Qt5 COMPONENTS Core Gui REQUIRED)
find_package(Qt${QT_VERSION_MAJOR} COMPONENTS Core Gui REQUIRED)
find_package(OpenCV REQUIRED)

include_directories(${OpenCV_INCLUDE_DIRS})

add_subdirectory(zipper)

add_library(Wordx SHARED   
  fileop.cpp
  wordop.cpp
  info.cpp
)

link_directories(${OpenCV_LIBRARY_DIRS})
target_link_libraries(Wordx Qt${QT_VERSION_MAJOR}::Core Qt${QT_VERSION_MAJOR}::Gui ${OpenCV_LIBS} zipper)


