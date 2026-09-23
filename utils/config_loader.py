# -*- coding: utf-8 -*-
"""
utils/config_loader.py
======================
Loads and caches config/assumptions.yaml.

Usage::

    from utils.config_loader import load_config
    cfg = load_config()
    gdp = cfg["real_sector"]["nominal_gdp_kes_bn"]
"""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path

import yaml

from utils.logger import get_logger

log = get_logger(__name__)

CONFIG_PATH = Path(__file__).resolve().parent.parent / "config" / "assumptions.yaml"


@lru_cache(maxsize=1)
def load_config(path: str | None = None) -> dict:
    """
    Load assumptions.yaml and return as a dict.

    The result is cached — subsequent calls return the same object
    without re-reading the file.

    Parameters
    ----------
    path : str, optional
        Override the default config path (useful in tests).

    Returns
    -------
    dict
        Parsed YAML config.

    Raises
    ------
    FileNotFoundError
        If the config file does not exist.
    yaml.YAMLError
        If the file contains invalid YAML.
    """
    config_file = Path(path) if path else CONFIG_PATH

    if not config_file.exists():
        raise FileNotFoundError(
            f"Config file not found: {config_file}\n"
            "Expected at: config/assumptions.yaml"
        )

    log.debug("Loading config from %s", config_file)
    with open(config_file, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    log.debug("Config loaded: %d top-level keys", len(cfg))
    return cfg


def reload_config(path: str | None = None) -> dict:
    """Force reload by clearing the cache, then loading fresh."""
    load_config.cache_clear()
    return load_config(path)
