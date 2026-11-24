=========================
Excel Template Reports
=========================

Design a report in your own Excel file and let Odoo fill in the data. LGPL-3, Gray Bob.
Scheduling and delivery by email or Telegram are in ``xlsx_template_report_pro``.
The store page for Odoo Apps is ``static/description/index.html``.

Layout
======

* ``engine/``: pure Python (no ``odoo`` import), shared by every Odoo version.

  * ``render.py``: fills a template. Each data source writes a sheet and an Excel table with the
    same name; other sheets only get ``{{placeholder}}`` replacements, rows are never inserted, so
    user formulas such as ``=SUM(sales[amount_total])`` stay valid. Text starting with ``=`` is
    written as text (no formula injection). ``fix_core_properties()`` repairs the
    ``...+00:00Z`` dates written by openpyxl 3.1.2 (Odoo 18-20), which make Excel repair the file.
  * ``periods.py``: relative periods (yesterday, last week, month to date, ...).
  * ``schedule.py``: next run of a daily / weekly / monthly schedule in a timezone.

* ``models/xlsx_report.py``: the report; data is always read with the rights of the "Run as"
  user, never as superuser. ``_deliver()`` is the hook the Pro addon overrides.
* ``models/xlsx_report_source.py``: data sources (records with ordered columns, grouped totals with
  ``_read_group``, read-only SQL for the SQL administrator group, in its own cursor with a timeout).
* ``models/compat.py``: the few API differences between Odoo 17, 18, 19 and 20.
* ``data/sample_report.xml``: a ready-to-run "Contacts overview" report.

Notes
=====

* The web client reads binaries with ``bin_size=True``: use ``_template_bytes()``, never the field.
* Odoo 17 does not ship ``openpyxl``: ``pip install openpyxl``. Odoo 18 and later include it.

Tests
=====

::

    odoo -d test -i xlsx_template_report --test-enable --test-tags /xlsx_template_report --stop-after-init
    python -m pytest engine/tests        # the engine alone, without Odoo
