# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
import io
from datetime import datetime, timedelta

from openpyxl import load_workbook

from odoo.exceptions import AccessError, UserError, ValidationError
from odoo.tests import TransactionCase, new_test_user, tagged


class XlsxReportCommon(TransactionCase):

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.company_a = cls.env.company
        cls.company_b = cls.env["res.company"].create({"name": "Company B"})
        cls.user = new_test_user(
            cls.env, login="xlsx_user", groups="base.group_user,xlsx_template_report.group_xlsx_report_user",
            company_id=cls.company_a.id, company_ids=[(6, 0, cls.company_a.ids)])
        cls.sql_admin = new_test_user(
            cls.env, login="xlsx_sql", groups="base.group_user,xlsx_template_report.group_xlsx_report_sql")
        cls.partner_model = cls.env["ir.model"]._get("res.partner")
        cls.field_name = cls.env["ir.model.fields"]._get("res.partner", "name")
        cls.field_country = cls.env["ir.model.fields"]._get("res.partner", "country_id")
        Partner = cls.env["res.partner"]
        cls.partner_a = Partner.create({"name": "XLSX Partner A", "company_id": cls.company_a.id})
        cls.partner_b = Partner.create({"name": "XLSX Partner B", "company_id": cls.company_b.id})

    def _report(self, user=None, **vals):
        env = self.env(user=user) if user else self.env
        return env["xlsx.report"].create(dict({
            "name": "Partners",
            "period_type": "this_month",
            "source_ids": [(0, 0, {
                "name": "partners",
                "source_type": "search",
                "model_id": self.partner_model.id,
                "domain": "[('name', 'like', 'XLSX Partner')]",
                "column_ids": [
                    (0, 0, {"sequence": 1, "field_id": self.field_name.id}),
                    (0, 0, {"sequence": 2, "field_id": self.field_country.id}),
                ],
            })],
        }, **vals))

    def _rows(self, run, sheet="partners"):
        wb = load_workbook(io.BytesIO(run.attachment_id.raw))
        return [[c.value for c in row] for row in wb[sheet].iter_rows()]


@tagged("post_install", "-at_install")
class TestXlsxReport(XlsxReportCommon):

    def test_render_respects_multi_company(self):
        report = self._report(user=self.user)
        self.assertEqual(report.owner_id, self.user)
        run = report._execute("manual")
        self.assertEqual(run.state, "done")
        names = [row[0] for row in self._rows(run)[1:]]
        self.assertIn("XLSX Partner A", names)
        self.assertNotIn("XLSX Partner B", names)  # other company, invisible to the owner

    def test_column_order_and_custom_header(self):
        report = self._report()
        name_col, country_col = report.source_ids.column_ids.sorted("sequence")
        name_col.write({"sequence": 3, "header": "Customer"})
        run = report._execute("manual")
        self.assertEqual(self._rows(run)[0], ["country_id", "Customer"])

    def test_columns_must_match_model(self):
        user_login = self.env["ir.model.fields"]._get("res.users", "login")
        with self.assertRaises(ValidationError):
            self._report().source_ids.column_ids = [(0, 0, {"field_id": user_login.id})]

    def test_template_info_as_web_client(self):
        """The web client reads binaries with bin_size=True: the info must still parse the file."""
        report = self.env.ref("xlsx_template_report.sample_report_contacts")
        info = report.with_context(bin_size=True).template_info
        self.assertIn("contacts", info)
        self.assertNotIn("Invalid", info)
        run = report.with_context(bin_size=True)._execute("manual")
        self.assertEqual(run.state, "done")

    def test_sample_report_renders(self):
        report = self.env.ref("xlsx_template_report.sample_report_contacts")
        self.assertFalse(any(line.startswith("⚠") for line in (report.template_info or "").splitlines()))
        run = report._execute("manual")
        self.assertEqual(run.state, "done")
        wb = load_workbook(io.BytesIO(run.attachment_id.raw))
        summary = wb["Summary"]
        self.assertEqual(summary["A1"].value, "Sample: Contacts overview")
        self.assertIsInstance(summary["B5"].value, int)  # {{contacts_count}} stays a number
        self.assertEqual(summary["B6"].value, "=COUNTIF(contacts[is_company],TRUE)")
        self.assertEqual([c.value for c in wb["contacts"][1]],
                         ["name", "email", "country_id", "is_company", "create_date"])
        self.assertFalse(run.channel_log)  # no unknown placeholders

    def test_generated_file_opens_without_repair(self):
        """openpyxl 3.1.2 (Odoo 18-20) writes '...+00:00Z' dates that make Excel repair the file."""
        import re
        import zipfile
        run = self._report()._execute("manual")
        core = zipfile.ZipFile(io.BytesIO(run.attachment_id.raw)).read("docProps/core.xml").decode()
        dates = re.findall(r"<dcterms:(?:created|modified)[^>]*>([^<]*)<", core)
        self.assertTrue(dates)
        for value in dates:
            self.assertRegex(value, r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})$")

    def test_download_action(self):
        action = self._report().action_download()
        self.assertEqual(action["type"], "ir.actions.act_url")
        self.assertIn("/web/content/", action["url"])

    def test_formula_injection_from_data(self):
        self.partner_a.name = '=HYPERLINK("http://evil.example","XLSX Partner")'
        run = self._report()._execute("manual")
        wb = load_workbook(io.BytesIO(run.attachment_id.raw))
        cells = [row[0] for row in wb["partners"].iter_rows(min_row=2)]
        evil = [c for c in cells if str(c.value).startswith("=")]
        self.assertTrue(evil)
        self.assertTrue(all(c.data_type == "s" for c in evil))

    def test_user_cannot_set_other_owner(self):
        report = self._report(user=self.user, owner_id=self.env.uid)
        self.assertEqual(report.owner_id, self.user, "create forces the owner for non-managers")
        with self.assertRaises(AccessError):
            report.write({"owner_id": self.env.uid})

    def test_user_cannot_create_sql_source(self):
        with self.assertRaises(AccessError):
            self.env(user=self.user)["xlsx.report"].create({
                "name": "SQL",
                "source_ids": [(0, 0, {"name": "q", "source_type": "sql"})],
            })

    def test_sql_is_read_only(self):
        report = self.env(user=self.sql_admin)["xlsx.report"].create({
            "name": "SQL",
            "source_ids": [(0, 0, {
                "name": "q",
                "source_type": "sql",
                "sql_query": "WITH d AS (DELETE FROM res_partner WHERE id = %s RETURNING id) "
                             "SELECT * FROM d" % self.partner_a.id,
            })],
        })
        with self.assertRaises(UserError):
            report._render()
        self.assertTrue(self.partner_a.exists())

    def test_sql_rejects_multiple_statements(self):
        report = self.env(user=self.sql_admin)["xlsx.report"].create({
            "name": "SQL",
            "source_ids": [(0, 0, {"name": "q", "source_type": "sql",
                                   "sql_query": "SELECT 1; SELECT 2"})],
        })
        with self.assertRaises(UserError):
            report._render()

    def test_sql_select(self):
        report = self.env(user=self.sql_admin)["xlsx.report"].create({
            "name": "SQL",
            "source_ids": [(0, 0, {"name": "q", "source_type": "sql",
                                   "sql_query": "SELECT %(date_from)s::date AS d, 42 AS answer"})],
        })
        content, _filename, rows, _warnings = report._render()
        ws = load_workbook(io.BytesIO(content))["q"]
        self.assertEqual(rows, 1)
        self.assertEqual(ws["B2"].value, 42)

    def test_group_source(self):
        report = self._report()
        report.source_ids.write({
            "source_type": "group",
            "groupby_ids": [(6, 0, self.field_country.ids)],
        })
        content, *_ = report._render()
        ws = load_workbook(io.BytesIO(content))["partners"]
        self.assertEqual([c.value for c in ws[1]], ["country_id", "count"])

    def test_failed_manual_run_raises(self):
        report = self._report()
        report.source_ids.domain = "[('no_such_field', '=', 1)]"
        with self.assertRaises(Exception):
            report._execute("manual")

    def test_invalid_source_name(self):
        with self.assertRaises(ValidationError):
            self._report().source_ids.write({"name": "my sales"})

    def test_period_preview(self):
        report = self._report(period_type="last_month")
        self.assertIn("→", report.period_preview)

    def test_datetime_period_bounds_use_report_timezone(self):
        report = self._report(timezone="Asia/Ho_Chi_Minh")
        source = report.source_ids
        field = self.env["res.partner"]._fields["create_date"]
        period = report._get_period()
        domain = source._period_bounds(field, period, "Asia/Ho_Chi_Minh")
        # Local midnight (UTC+7) is 17:00 UTC the previous day.
        self.assertEqual(domain[0][2], datetime.combine(period.date_from, datetime.min.time()) - timedelta(hours=7))
