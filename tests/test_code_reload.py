"""Regression test: an update of the code must reach modules already imported.

Run:  py -3 -m pytest tests -q
"""
import importlib
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "soil_lab_tool"))

from code_reload import drop_stale_modules  # noqa: E402


def test_changed_code_is_loaded_on_the_next_import(tmp_path, monkeypatch):
    module = tmp_path / "reload_probe.py"
    module.write_text("VALUE = 'old'\n")
    monkeypatch.syspath_prepend(str(tmp_path))
    monkeypatch.delattr(sys, "_lab_code_version", raising=False)

    first = drop_stale_modules(str(tmp_path))
    assert importlib.import_module("reload_probe").VALUE == "old"

    # Nothing changed: the module stays loaded
    assert drop_stale_modules(str(tmp_path)) == first
    assert "reload_probe" in sys.modules

    module.write_text("VALUE = 'new'\n")
    os.utime(module, (first + 10, first + 10))
    assert drop_stale_modules(str(tmp_path)) == first + 10
    assert "reload_probe" not in sys.modules
    assert importlib.import_module("reload_probe").VALUE == "new"

    sys.modules.pop("reload_probe", None)


def test_modules_outside_the_folder_are_left_alone(tmp_path, monkeypatch):
    monkeypatch.delattr(sys, "_lab_code_version", raising=False)
    (tmp_path / "x.py").write_text("")
    drop_stale_modules(str(tmp_path))
    assert "os" in sys.modules and "pytest" in sys.modules
