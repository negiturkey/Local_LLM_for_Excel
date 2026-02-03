@echo off
set CUDA_VISIBLE_DEVICES=0

call .venv\Scripts\activate
set UV_PYTHON=.venv\Scripts\python.exe
python server/host.py