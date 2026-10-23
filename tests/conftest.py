def pytest_configure(config):
    config.addinivalue_line("markers", "network: uses the real SEC cache/network (skipped without .cache)")
