#! /usr/bin/env bash

set -e

cd "$(dirname "$0")"

uv sync --upgrade
uv lock --upgrade

uvx black .
uvx ruff check --fix .
uvx ruff format .
