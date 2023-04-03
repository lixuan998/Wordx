cmake_minimum_required(VERSION 3.0)
project(Wordx LANGUAGES CXX)
include_directories(zipper)

set(CMAKE_INCLUDE_CURRENT_DIR ON)

IF (CMAKE_SYSTEM_NAME MATCHES "Darwin") 
  set(CMAKE_OSX_ARCHITECTURES arm64)
ENDIF (CMAKE_SYSTEM_NAME MATCHES "Darwin")

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(Qt5 COMPONENTS Core Gui Widgets REQUIRED)
find_package(OpenCV REQUIRED)

include_directories(${OpenCV_INCLUDE_DIRS})

add_subdirectory(zipper)

add_library(Wordx SHARED fileop.cpp wordop.cpp info.cpp add_image_thread.cpp)

target_link_libraries(Wordx Qt5::Core Qt5::Gui Qt5::Widgets ${OpenCV_LIBS} zipper)
