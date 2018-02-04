@echo off
title run_coverage
pytest --cov-report term --cov-report xml --cov-report html --cov=bt_expense tests/
pause