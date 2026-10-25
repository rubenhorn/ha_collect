# Home Assistant Collect

A simple script to collect entity data over time as Excel, CSV, or JSON-lines files.

**Getting Started:**  
1. Install [uv](https://docs.astral.sh/uv/)
2. Run `./setup_project.sh`
3. Export `HA_URL` and `HA_TOKEN` in the shell environment
4. Create `config.yaml` based on `config.yaml.template`  
(Use `uv run ha_collect.py --list > info.md` to check entities)
5. Run the data collection using `uv run ha_collect`
