


list(APPEND CMAKE_MODULE_PATH  /home/climatex/Lx_Workspace/Qt_workspace/wordx/zipper/cmake)

if(EXISTS ${CMAKE_CURRENT_LIST_DIR}/zipperTargets.cmake AND NOT TARGET )
  include(${CMAKE_CURRENT_LIST_DIR}/zipperTargets.cmake)
  if(NOT TARGET zipper::zipper)
    if(TARGET Zipper-static)
      set_target_properties(Zipper-static PROPERTIES IMPORTED_GLOBAL TRUE)
      add_library(zipper::static ALIAS Zipper-static)
    endif()
    if(TARGET Zipper)
      set_target_properties(Zipper PROPERTIES IMPORTED_GLOBAL TRUE)
      add_library(zipper::shared ALIAS Zipper)
      add_library(zipper::zipper ALIAS Zipper)
    elseif(TARGET Zipper-static)
      add_library(zipper::zipper ALIAS Zipper-static)
    endif()
  endif()
endif()
