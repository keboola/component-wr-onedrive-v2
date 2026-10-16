"""Pytest configuration for the keboola.wr-onedrive-v2 test suite.

No shared fixtures or environment variables are hoisted here yet — every test module in this
suite is currently self-contained (each of ``test_unit.py``/``test_datadir.py``/etc. defines its
own small local fixtures/doubles right next to the tests that use them, matching the rest of the
suite's style). This file exists as the conventional place to put anything that later needs to be
shared across modules (see ``component-ex-medallia``'s ``tests/conftest.py`` for an example of
what that looks like — an environment variable that a VCR replay cap must share with its recorder).
"""
