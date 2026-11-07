"""
Pytest global fixtures for Agent 365 test suite.
Automatically seeds the mock client with benchmark test fixtures so unit tests pass reliably.
"""

import pytest
from agent365.config import settings
from agent365.m365_client import get_m365_client, MockM365Client, reset_client


@pytest.fixture(autouse=True)
def setup_benchmark_fixtures():
    """Seeds benchmark fixtures for tests requiring multi-event or multi-file datasets."""
    settings.fresh_out_of_box = False
    reset_client()
    client = get_m365_client(force_mode="mock")
    if isinstance(client, MockM365Client):
        client.seed_demo_dataset()
    yield
    settings.fresh_out_of_box = True
    reset_client()
