cmake_minimum_required(VERSION 3.0)

include_directories(zipper)

set(CMAKE_INCLUDE_CURRENT_DIR ON)

IF (CMAKE_SYSTEM_NAME MATCHES "Darwin") 
  set(CMAKE_OSX_ARCHITECTURES arm64)
ENDIF (CMAKE_SYSTEM_NAME MATCHES "Darwin")

set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

file(
    GLOB_RECURSE
    SOURCE_FILE_SET 
    ${CMAKE_CURRENT_SOURCE_DIR}/src/*
)

file(
    GLOB_RECURSE
    HEADER_FILE_SET
    ${CMAKE_CURRENT_SOURCE_DIR}/include/*
)

include_directories(
    ${CMAKE_CURRENT_SOURCE_DIR}/include
)

find_package(Qt5 COMPONENTS Core Gui Widgets REQUIRED)
find_package(OpenCV REQUIRED)

include_directories(${OpenCV_INCLUDE_DIRS})

add_subdirectory(zipper)

add_library(Wordx STATIC ${SOURCE_FILE_SET})

target_link_libraries(Wordx Qt5::Core Qt5::Gui Qt5::Widgets ${OpenCV_LIBS} zipper)
