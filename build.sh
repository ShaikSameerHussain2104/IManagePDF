#!/usr/bin/env bash
apt-get update
apt-get install -y libicu-dev libssl-dev
pip install -r requirements.txt
