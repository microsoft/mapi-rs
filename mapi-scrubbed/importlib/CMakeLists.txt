add_custom_command(OUTPUT ${CMAKE_CURRENT_SOURCE_DIR}/x64/mapi32.lib
	COMMAND cmake -E make_directory x64
	COMMAND cmake -E chdir x64 lib /nologo /machine:x64 /out:mapi32.lib /def:${CMAKE_CURRENT_SOURCE_DIR}/mapi32.def
	COMMAND cmake -E copy_if_different x64/mapi32.lib ${CMAKE_CURRENT_SOURCE_DIR}/x64
	DEPENDS mapi32.def)

add_custom_command(OUTPUT ${CMAKE_CURRENT_SOURCE_DIR}/x86/mapi32.lib
	COMMAND cmake -E make_directory x86
	COMMAND cmake -E chdir x86 lib /nologo /machine:x86 /out:mapi32.lib /def:${CMAKE_CURRENT_SOURCE_DIR}/mapi32.def
	COMMAND cmake -E copy_if_different x86/mapi32.lib ${CMAKE_CURRENT_SOURCE_DIR}/x86
	DEPENDS mapi32.def)

add_custom_command(OUTPUT ${CMAKE_CURRENT_SOURCE_DIR}/x64/olmapi32.lib
	COMMAND cmake -E make_directory x64
	COMMAND cmake -E chdir x64 lib /nologo /machine:x64 /out:olmapi32.lib /def:${CMAKE_CURRENT_SOURCE_DIR}/olmapi32.def
	COMMAND cmake -E copy_if_different x64/olmapi32.lib ${CMAKE_CURRENT_SOURCE_DIR}/x64
	DEPENDS olmapi32.def)

add_custom_command(OUTPUT ${CMAKE_CURRENT_SOURCE_DIR}/x86/olmapi32.lib
	COMMAND cmake -E make_directory x86
	COMMAND cmake -E chdir x86 lib /nologo /machine:x86 /out:olmapi32.lib /def:${CMAKE_CURRENT_SOURCE_DIR}/olmapi32.def
	COMMAND cmake -E copy_if_different x86/olmapi32.lib ${CMAKE_CURRENT_SOURCE_DIR}/x86
	DEPENDS olmapi32.def)

add_custom_target(copyimportlib ALL DEPENDS
	${CMAKE_CURRENT_SOURCE_DIR}/x64/mapi32.lib
	${CMAKE_CURRENT_SOURCE_DIR}/x86/mapi32.lib
	${CMAKE_CURRENT_SOURCE_DIR}/x64/olmapi32.lib
	${CMAKE_CURRENT_SOURCE_DIR}/x86/olmapi32.lib)
