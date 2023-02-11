set(CMAKE_CXX_STANDARD 17)
cmake_minimum_required (VERSION 3.9)
cmake_policy(VERSION 3.9)

project(xloil)

set(XLOIL_INCLUDE_DIR ${CMAKE_SOURCE_DIR}/include)
set(XLOIL_HEADERS
    ${XLOIL_INCLUDE_DIR}/xloil/AppObjects.h
    ${XLOIL_INCLUDE_DIR}/xloil/ArrayBuilder.h
    ${XLOIL_INCLUDE_DIR}/xloil/Async.h
    ${XLOIL_INCLUDE_DIR}/xloil/AutoBind.h
    ${XLOIL_INCLUDE_DIR}/xloil/Caller.h
    ${XLOIL_INCLUDE_DIR}/xloil/Date.h
    ${XLOIL_INCLUDE_DIR}/xloil/DynamicRegister.h
    ${XLOIL_INCLUDE_DIR}/xloil/EnumHelper.h
    ${XLOIL_INCLUDE_DIR}/xloil/Events.h
    ${XLOIL_INCLUDE_DIR}/xloil/ExcelArray.h
    ${XLOIL_INCLUDE_DIR}/xloil/ExcelCall.h
    ${XLOIL_INCLUDE_DIR}/xloil/ExcelObj.h
    ${XLOIL_INCLUDE_DIR}/xloil/ExcelObjCache.h
    ${XLOIL_INCLUDE_DIR}/xloil/ExcelRef.h
    ${XLOIL_INCLUDE_DIR}/xloil/ExcelThread.h
    ${XLOIL_INCLUDE_DIR}/xloil/ExcelTypeLib.h
    ${XLOIL_INCLUDE_DIR}/xloil/ExcelUI.h
    ${XLOIL_INCLUDE_DIR}/xloil/ExportMacro.h
    ${XLOIL_INCLUDE_DIR}/xloil/FPArray.h
    ${XLOIL_INCLUDE_DIR}/xloil/FuncSpec.h
    ${XLOIL_INCLUDE_DIR}/xloil/Interface.h
    ${XLOIL_INCLUDE_DIR}/xloil/Log.h
    ${XLOIL_INCLUDE_DIR}/xloil/LogWindow.h
    ${XLOIL_INCLUDE_DIR}/xloil/NumericTypeConverters.h
    ${XLOIL_INCLUDE_DIR}/xloil/ObjectCache.h
    ${XLOIL_INCLUDE_DIR}/xloil/Plugin.h
    ${XLOIL_INCLUDE_DIR}/xloil/Preprocessor.h
    ${XLOIL_INCLUDE_DIR}/xloil/PString.h
    ${XLOIL_INCLUDE_DIR}/xloil/Range.h
    ${XLOIL_INCLUDE_DIR}/xloil/Register.h
    ${XLOIL_INCLUDE_DIR}/xloil/RtdServer.h
    ${XLOIL_INCLUDE_DIR}/xloil/State.h
    ${XLOIL_INCLUDE_DIR}/xloil/StaticRegister.h
    ${XLOIL_INCLUDE_DIR}/xloil/StringUtils.h
    ${XLOIL_INCLUDE_DIR}/xloil/Throw.h
    ${XLOIL_INCLUDE_DIR}/xloil/TypeConverters.h
    ${XLOIL_INCLUDE_DIR}/xloil/Version.h
    ${XLOIL_INCLUDE_DIR}/xloil/WindowsSlim.h
    ${XLOIL_INCLUDE_DIR}/xloil/XlCallSlim.h
    ${XLOIL_INCLUDE_DIR}/xloil/XllEntryPoint.h
)

set(XLOIL_SRCS_DIR ${CMAKE_CURRENT_SOURCE_DIR}/src)
set(XLOIL_PRIVATE_HEADERS
    ${XLOIL_SRCS_DIR}/xlOil/Loaders/AddinLoader.h
    ${XLOIL_SRCS_DIR}/xlOil/Loaders/CoreEntryPoint.h
    ${XLOIL_SRCS_DIR}/xlOil/Loaders/PluginLoader.h
    ${XLOIL_SRCS_DIR}/xlOil-COM/ClassFactory.h
    ${XLOIL_SRCS_DIR}/xlOil-COM/ComAddin.h
    ${XLOIL_SRCS_DIR}/xlOil-COM/ComEventSink.h
    ${XLOIL_SRCS_DIR}/xlOil-COM/ComVariant.h
    ${XLOIL_SRCS_DIR}/xlOil-COM/CustomTaskPane.h
    ${XLOIL_SRCS_DIR}/xlOil-COM/RibbonExtensibility.h
    ${XLOIL_SRCS_DIR}/xlOil-COM/RtdAsyncManager.h
    ${XLOIL_SRCS_DIR}/xlOil-COM/RtdManager.h
    ${XLOIL_SRCS_DIR}/xlOil-COM/RtdServerWorker.h
    ${XLOIL_SRCS_DIR}/xlOil-COM/TaskPaneHostControl.h
    ${XLOIL_SRCS_DIR}/xlOil-COM/WorkbookScopeFunctions.h
    ${XLOIL_SRCS_DIR}/xlOil-COM/XllContextInvoke.h
    ${XLOIL_SRCS_DIR}/xlOil-Dynamic/ExternalRegionAllocator.h
    ${XLOIL_SRCS_DIR}/xlOil-Dynamic/LocalFunctions.h
    ${XLOIL_SRCS_DIR}/xlOil-Dynamic/PEHelper.h
    ${XLOIL_SRCS_DIR}/xlOil-Dynamic/Thunker.h
    ${XLOIL_SRCS_DIR}/xlOil-XLL/ExcelCallMapping.h
    ${XLOIL_SRCS_DIR}/xlOil-XLL/FuncRegistry.h
    ${XLOIL_SRCS_DIR}/xlOil-XLL/Intellisense.h
    ${XLOIL_SRCS_DIR}/xlOil-XLL/LogWindowSink.h
    ${XLOIL_SRCS_DIR}
)
set(XLOIL_SRCS
    ${XLOIL_SRCS_DIR}/xlOil/Interface.cpp
    ${XLOIL_SRCS_DIR}/xlOil/Log.cpp
    ${XLOIL_SRCS_DIR}/xlOil/Loaders/AddinLoader.cpp
    ${XLOIL_SRCS_DIR}/xlOil/Loaders/CoreEntryPoint.cpp
    ${XLOIL_SRCS_DIR}/xlOil/Loaders/PluginLoader.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/AppObjects.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/API/ComUtils.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/API/Events.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/API/ExcelRange.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/API/ExcelThread.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/API/RtdServer.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/ClassFactory.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/ComAddin.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/ComEventSink.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/ComVariant.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/Connect.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/CustomTaskPane.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/RibbonExtensibility.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/RtdAsyncManager.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/rtdmanager.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/TaskPaneHostControl.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/WorkbookScopeFunctions.cpp
    ${XLOIL_SRCS_DIR}/xlOil-COM/xllcontextinvoke.cpp
    ${XLOIL_SRCS_DIR}/xlOil-Dynamic/DynamicRegistration.cpp
    ${XLOIL_SRCS_DIR}/xlOil-Dynamic/LocalFunctions.cpp
    ${XLOIL_SRCS_DIR}/xlOil-Dynamic/PEHelper.cpp
    ${XLOIL_SRCS_DIR}/xlOil-Dynamic/Thunker.cpp
    ${XLOIL_SRCS_DIR}/xlOil-Funcs/ExcelObjCache.cpp
    ${XLOIL_SRCS_DIR}/xlOil-Funcs/xloHelp.cpp
    ${XLOIL_SRCS_DIR}/xlOil-Funcs/xloLog.cpp
    # ${XLOIL_SRCS_DIR}/xlOil-Funcs/xloReload.cpp
    ${XLOIL_SRCS_DIR}/xlOil-Funcs/xloVersion.cpp
    # ${XLOIL_SRCS_DIR}/xlOil-Loader/LoaderEntryPoint.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/Async.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/Caller.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/Date.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/ExcelArray.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/ExcelCall.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/ExcelObj.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/ExcelRef.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/FPArray.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/FuncRegistry.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/Intellisense.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/Log.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/LogWindow.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/LogWindowSink.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/State.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/StaticRegister.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/Throw.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/XlCall.cpp
    ${XLOIL_SRCS_DIR}/xlOil-XLL/XllEvents.cpp
    ${XLOIL_SRCS_DIR}/xlOilHelpers/Environment.cpp
    ${XLOIL_SRCS_DIR}/xlOilHelpers/GuidUtils.cpp
    ${XLOIL_SRCS_DIR}/xlOilHelpers/Settings.cpp
)

set(XLOIL_3RD_PARTIES_DIR ${CMAKE_CURRENT_SOURCE_DIR}/external)
set(XLOIL_3RD_PARTIES
    ${XLOIL_3RD_PARTIES_DIR}/rdcfswatcher/rdc_fs_watcher.cpp
    ${XLOIL_3RD_PARTIES_DIR}/rdcfswatcher/rdc_fs_watcher.h
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/arch.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/assembler.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/builder.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/callconv.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/codeholder.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/compiler.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/constpool.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/cpuinfo.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/emitter.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/func.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/globals.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/inst.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/jitallocator.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/jitruntime.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/logging.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/operand.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/osutils.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/ralocal.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/rapass.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/rastack.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/string.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/support.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/target.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/type.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/virtmem.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/zone.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/zonehash.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/zonelist.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/zonestack.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/zonetree.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/core/zonevector.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/x86/x86assembler.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/x86/x86builder.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/x86/x86callconv.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/x86/x86compiler.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/x86/x86features.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/x86/x86instapi.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/x86/x86instdb.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/x86/x86internal.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/x86/x86logging.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/x86/x86operand.cpp
    ${XLOIL_3RD_PARTIES_DIR}/asmjit/src/asmjit/x86/x86rapass.cpp
    
)

set(Excel2013SDK ${CMAKE_CURRENT_SOURCE_DIR}/external/Excel2013SDK/INCLUDE)

set(XLOIL_DEF_FILE ${XLOIL_SRCS_DIR}/xlOil-Loader/xlOilAddin.def)

add_library(xloil STATIC ${XLOIL_HEADERS} ${XLOIL_PRIVATE_HEADERS} ${XLOIL_SRCS} ${XLOIL_3RD_PARTIES})
# add_library(xloil STATIC ${CMAKE_CURRENT_SOURCE_DIR}/src/test.cpp)

find_package(fmt REQUIRED)
find_package(spdlog REQUIRED)

target_compile_features(xloil PUBLIC cxx_std_17)
set_target_properties(xloil
    PROPERTIES
        CXX_VISIBILITY_PRESET hidden
        VISIBILITY_INLINES_HIDDEN ON
)
target_include_directories(
    xloil
    PUBLIC
        $<BUILD_INTERFACE:${XLOIL_INCLUDE_DIR}>
    PRIVATE
        ${XLOIL_PRIVATE_HEADERS}
        ${XLOIL_3RD_PARTIES_DIR}
        ${XLOIL_3RD_PARTIES_DIR}/asmjit/src
        ${Excel2013SDK}
)

target_compile_definitions(xloil PUBLIC )
# target_compile_definitions(xloil PRIVATE XLOIL_EXPORTS)
target_compile_definitions(xloil PRIVATE _WINDLL )
target_compile_definitions(xloil PUBLIC
    # SPDLOG_FMT_EXTERNAL
    XLOIL_STATIC_LIB
    _USRDLL
    _WINDOWS
    _NDEBUG
    SPDLOG_ACTIVE_LEVEL=SPDLOG_LEVEL_TRACE
    UNICODE
    NOMINMAX
    ASMJIT_STATIC
    _SILENCE_CXX17_CODECVT_HEADER_DEPRECATION_WARNING
)

target_link_libraries(xloil
    PUBLIC
        # fmt::fmt-header-only
        spdlog::spdlog_header_only
    PRIVATE
        oleacc
        user32
        delayimp
)

add_subdirectory(example)