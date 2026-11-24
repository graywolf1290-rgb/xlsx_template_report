# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
# Make ``engine`` importable as a top-level package so these tests never
# import the Odoo module (``xlsx_template_report/__init__.py`` imports odoo).
#
# Run from the module folder:  python -m pytest engine/tests
# (pytest.ini next to this file keeps pytest's rootdir here.)
import os
import sys

MODULE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if MODULE_DIR not in sys.path:
    sys.path.insert(0, MODULE_DIR)
