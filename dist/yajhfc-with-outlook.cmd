@echo off
rem Executes YajHFC including the outlook plugin
set PATH=%PATH%;%~dp0\lib

javaw -jar yajhfc.jar --load-plugin=yajhfc-outlook-pb-plugin.jar %*
