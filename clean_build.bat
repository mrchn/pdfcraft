@echo off
echo cleaning old builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist pdfcraft_mrchn.egg-info rmdir /s /q pdfcraft_mrchn.egg-info
echo done!