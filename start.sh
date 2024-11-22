#!/bin/bash
apt-get update
apt-get install -y python3-tk xvfb
Xvfb :99 -screen 0 1024x768x16 &
export DISPLAY=:99
exec "$@"