@rem
@rem Copyright 2015 the original author or authors.
@rem
@rem Licensed under the Apache License, Version 2.0 (the "License");
@rem you may not use this file except in compliance with the License.
@rem You may obtain a copy of the License at
@rem
@rem      https://www.apache.org/licenses/LICENSE-2.0
@rem
@rem Unless required by applicable law or agreed to in writing, software
@rem distributed under the License is distributed on an "AS IS" BASIS,
@rem WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
@rem See the License for the specific language governing permissions and
@rem limitations under the License.
@rem

@if "%DEBUG%" == "" @echo off
@rem ##########################################################################
@rem
@rem  teamsChannel4Record startup script for Windows
@rem
@rem ##########################################################################

@rem Set local scope for the variables with windows NT shell
if "%OS%"=="Windows_NT" setlocal

set DIRNAME=%~dp0
if "%DIRNAME%" == "" set DIRNAME=.
set APP_BASE_NAME=%~n0
set APP_HOME=%DIRNAME%..

@rem Resolve any "." and ".." in APP_HOME to make it shorter.
for %%i in ("%APP_HOME%") do set APP_HOME=%%~fi

@rem Add default JVM options here. You can also use JAVA_OPTS and TEAMS_CHANNEL4_RECORD_OPTS to pass JVM options to this script.
set DEFAULT_JVM_OPTS=

@rem Find java.exe
if defined JAVA_HOME goto findJavaFromJavaHome

set JAVA_EXE=java.exe
%JAVA_EXE% -version >NUL 2>&1
if "%ERRORLEVEL%" == "0" goto execute

echo.
echo ERROR: JAVA_HOME is not set and no 'java' command could be found in your PATH.
echo.
echo Please set the JAVA_HOME variable in your environment to match the
echo location of your Java installation.

goto fail

:findJavaFromJavaHome
set JAVA_HOME=%JAVA_HOME:"=%
set JAVA_EXE=%JAVA_HOME%/bin/java.exe

if exist "%JAVA_EXE%" goto execute

echo.
echo ERROR: JAVA_HOME is set to an invalid directory: %JAVA_HOME%
echo.
echo Please set the JAVA_HOME variable in your environment to match the
echo location of your Java installation.

goto fail

:execute
@rem Setup the command line

set CLASSPATH=%APP_HOME%\lib\teamsChannel4Record-1.0.0.jar;%APP_HOME%\lib\unified-agent-12.0.0.jar;%APP_HOME%\lib\agent-interfaces-12.0.0.jar;%APP_HOME%\lib\agent-utils-12.0.0.jar;%APP_HOME%\lib\sednaclient-blueline-12.0.0.jar;%APP_HOME%\lib\blmetadata-12.0.0.jar;%APP_HOME%\lib\blconfig-12.0.0.jar;%APP_HOME%\lib\bluelineutil-12.0.0.jar;%APP_HOME%\lib\blueline-overlay-12.0.0.jar;%APP_HOME%\lib\blueline-int-12.0.0.jar;%APP_HOME%\lib\blueline-12.0.0.jar;%APP_HOME%\lib\log4j-slf4j-impl-2.17.2.jar;%APP_HOME%\lib\log4j-core-2.17.2.jar;%APP_HOME%\lib\log4j-api-2.17.2.jar;%APP_HOME%\lib\json-20230227.jar;%APP_HOME%\lib\okhttp-4.10.0.jar;%APP_HOME%\lib\agentserver-remote-interfaces-12.0.0.jar;%APP_HOME%\lib\sednaclient-api-nodep-12.0.0.jar;%APP_HOME%\lib\doxis4-descriptor-resolver-api-12.0.0.jar;%APP_HOME%\lib\dx4-commons-clients-12.0.0.jar;%APP_HOME%\lib\httpclient-4.5.13.jar;%APP_HOME%\lib\sercommon-xml-3.0.0.jar;%APP_HOME%\lib\sercommon-collections-3.0.0.jar;%APP_HOME%\lib\sercommon-lang-3.0.0.jar;%APP_HOME%\lib\pdfbox-2.0.26.jar;%APP_HOME%\lib\fontbox-2.0.26.jar;%APP_HOME%\lib\sercommon-lang-base-3.0.0.jar;%APP_HOME%\lib\commons-logging-1.2.jar;%APP_HOME%\lib\slf4j-api-1.7.25.jar;%APP_HOME%\lib\okio-jvm-3.0.0.jar;%APP_HOME%\lib\kotlin-stdlib-jdk8-1.5.31.jar;%APP_HOME%\lib\kotlin-stdlib-jdk7-1.5.31.jar;%APP_HOME%\lib\kotlin-stdlib-1.6.20.jar;%APP_HOME%\lib\dx4-commons-service-interfaces-12.0.0.jar;%APP_HOME%\lib\imagingservice-nodep-12.0.0.jar;%APP_HOME%\lib\blueline-tools-12.0.0.jar;%APP_HOME%\lib\commons-text-1.9.jar;%APP_HOME%\lib\commons-lang3-3.12.0.jar;%APP_HOME%\lib\commons-codec-1.15.jar;%APP_HOME%\lib\commons-collections4-4.4.jar;%APP_HOME%\lib\commons-io-2.11.0.jar;%APP_HOME%\lib\annotations-3.0.1.jar;%APP_HOME%\lib\jcifs-1.3.17.jar;%APP_HOME%\lib\kotlin-stdlib-common-1.6.20.jar;%APP_HOME%\lib\annotations-13.0.jar;%APP_HOME%\lib\dx4-migration-commons-12.0.0.jar;%APP_HOME%\lib\dx4-commons-feel-scala-12.0.0.jar;%APP_HOME%\lib\sedna-system-definitions-12.0.0.jar;%APP_HOME%\lib\commons-csv-1.9.0.jar;%APP_HOME%\lib\blueline-localization-12.0.0.jar;%APP_HOME%\lib\bcmail-jdk15on-1.64.jar;%APP_HOME%\lib\bcpkix-jdk15on-1.64.jar;%APP_HOME%\lib\bcprov-jdk15on-1.64.jar;%APP_HOME%\lib\jai_codec-1.1.2.jar;%APP_HOME%\lib\icu4j-67.1.jar;%APP_HOME%\lib\commons-imaging-1.0-alpha2.jar;%APP_HOME%\lib\httpcore-4.4.13.jar;%APP_HOME%\lib\jaxb-api-2.3.1.jar;%APP_HOME%\lib\byte-buddy-1.12.1.jar;%APP_HOME%\lib\jai_core-1.1.2.jar;%APP_HOME%\lib\javax.activation-api-1.2.0.jar


@rem Execute teamsChannel4Record
"%JAVA_EXE%" %DEFAULT_JVM_OPTS% %JAVA_OPTS% %TEAMS_CHANNEL4_RECORD_OPTS%  -classpath "%CLASSPATH%" junit.AgentTester %*

:end
@rem End local scope for the variables with windows NT shell
if "%ERRORLEVEL%"=="0" goto mainEnd

:fail
rem Set variable TEAMS_CHANNEL4_RECORD_EXIT_CONSOLE if you need the _script_ return code instead of
rem the _cmd.exe /c_ return code!
if  not "" == "%TEAMS_CHANNEL4_RECORD_EXIT_CONSOLE%" exit 1
exit /b 1

:mainEnd
if "%OS%"=="Windows_NT" endlocal

:omega
