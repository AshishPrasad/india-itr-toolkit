"""Shared pytest fixtures."""
import os

import pytest

from generate_sample_data import DATA_DIR, generate


@pytest.fixture(scope="session", autouse=True)
def sample_data():
    """Generate the sample workbooks into a gitignored temp folder."""
    generate(DATA_DIR)
    return DATA_DIR


@pytest.fixture
def data_dir(sample_data):
    return sample_data


@pytest.fixture
def dividends_path(data_dir):
    return os.path.join(data_dir, "dividends_sample.xlsx")


@pytest.fixture
def stocks_path(data_dir):
    return os.path.join(data_dir, "stocks_sample.xlsx")


@pytest.fixture
def stocks_renamed_path(data_dir):
    return os.path.join(data_dir, "stocks_renamed.xlsx")
