@echo off
title install occ Requirements
echo Install occ project dependencies with pip?
pause
pip install -r requirements.txt
TIMEOUT 5