find_program(TTEXE TextTransformCore.exe)

add_custom_command(OUTPUT ${mapistublibrary_SOURCE_DIR}/include/ScrubHeaders.cpp
	COMMAND cmake -E copy_if_different ${CMAKE_CURRENT_SOURCE_DIR}/ScrubHeaders.tt ${mapistublibrary_SOURCE_DIR}/include
	COMMAND cmake -E chdir ${mapistublibrary_SOURCE_DIR}/include ${TTEXE} ScrubHeaders.tt
	DEPENDS ${CMAKE_CURRENT_SOURCE_DIR}/ScrubHeaders.tt ${TTEXE})

add_executable(scrubber ${mapistublibrary_SOURCE_DIR}/include/ScrubHeaders.cpp)
target_include_directories(scrubber PRIVATE ${mapistublibrary_SOURCE_DIR}/include)
target_compile_options(scrubber PRIVATE /EHsc)

add_custom_command(OUTPUT ${mapistublibrary_SOURCE_DIR}/include/MAPIConstants.h
	COMMAND scrubber > ${mapistublibrary_SOURCE_DIR}/include/MAPIConstants.h VERBATIM
	DEPENDS scrubber)
add_custom_target(scrubconstants ALL DEPENDS ${mapistublibrary_SOURCE_DIR}/include/MAPIConstants.h)
