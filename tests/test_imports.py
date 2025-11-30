"""Basic import tests."""


def test_version():
    """Test that we can read the version."""
    import tomllib
    from pathlib import Path

    pyproject = Path(__file__).parent.parent / "pyproject.toml"
    with open(pyproject, "rb") as f:
        data = tomllib.load(f)

    assert "project" in data
    assert "version" in data["project"]
    assert data["project"]["version"]
