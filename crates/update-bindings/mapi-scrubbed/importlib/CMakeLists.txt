add_custom_command(OUTPUT ${CMAKE_BINARY_DIR}/lib/x64/mapi32.lib
	COMMAND cmake -E make_directory ${CMAKE_BINARY_DIR}/lib/x64
	COMMAND cmake -E chdir ${CMAKE_BINARY_DIR}/lib/x64 lib /nologo /machine:x64 /out:mapi32.lib /def:${CMAKE_CURRENT_SOURCE_DIR}/mapi32.def
	DEPENDS mapi32.def)

add_custom_command(OUTPUT ${CMAKE_BINARY_DIR}/lib/x86/mapi32.lib
	COMMAND cmake -E make_directory ${CMAKE_BINARY_DIR}/lib/x86
	COMMAND cmake -E chdir ${CMAKE_BINARY_DIR}/lib/x86 lib /nologo /machine:x86 /out:mapi32.lib /def:${CMAKE_CURRENT_SOURCE_DIR}/mapi32.def
	DEPENDS mapi32.def)

add_custom_command(OUTPUT ${CMAKE_BINARY_DIR}/lib/x64/olmapi32.lib
	COMMAND cmake -E make_directory ${CMAKE_BINARY_DIR}/lib/x64
	COMMAND cmake -E chdir ${CMAKE_BINARY_DIR}/lib/x64 lib /nologo /machine:x64 /out:olmapi32.lib /def:${CMAKE_CURRENT_SOURCE_DIR}/olmapi32.def
	DEPENDS olmapi32.def)

add_custom_command(OUTPUT ${CMAKE_BINARY_DIR}/lib/x86/olmapi32.lib
	COMMAND cmake -E make_directory ${CMAKE_BINARY_DIR}/lib/x86
	COMMAND cmake -E chdir ${CMAKE_BINARY_DIR}/lib/x86 lib /nologo /machine:x86 /out:olmapi32.lib /def:${CMAKE_CURRENT_SOURCE_DIR}/olmapi32.def
	DEPENDS olmapi32.def)

add_custom_target(buildimportlibs ALL DEPENDS
	${CMAKE_BINARY_DIR}/lib/x64/mapi32.lib
	${CMAKE_BINARY_DIR}/lib/x86/mapi32.lib
	${CMAKE_BINARY_DIR}/lib/x64/olmapi32.lib
	${CMAKE_BINARY_DIR}/lib/x86/olmapi32.lib)

install(FILES
		${CMAKE_BINARY_DIR}/lib/x64/mapi32.lib
		${CMAKE_BINARY_DIR}/lib/x64/olmapi32.lib
	DESTINATION lib/x64)
install(FILES
		${CMAKE_BINARY_DIR}/lib/x86/mapi32.lib
		${CMAKE_BINARY_DIR}/lib/x86/olmapi32.lib
	DESTINATION lib/x86)
