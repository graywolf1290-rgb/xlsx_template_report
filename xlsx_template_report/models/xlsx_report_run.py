# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
from datetime import timedelta

from odoo import api, fields, models

from .compat import get_config_int

RETENTION_PARAM = "xlsx_template_report.run_retention_days"
DEFAULT_RETENTION_DAYS = 90


class XlsxReportRun(models.Model):
    _name = "xlsx.report.run"
    _description = "Excel Report Run"
    _order = "create_date desc, id desc"
    _rec_name = "report_id"

    report_id = fields.Many2one("xlsx.report", required=True, ondelete="cascade", index=True, readonly=True)
    trigger = fields.Selection([("manual", "Manual"), ("schedule", "Scheduled"), ("test", "Test")],
                               required=True, readonly=True)
    state = fields.Selection([("running", "Running"), ("done", "Done"), ("failed", "Failed")],
                             required=True, default="running", readonly=True)
    duration = fields.Float(string="Duration (s)", readonly=True, digits=(16, 2))
    row_count = fields.Integer(string="Rows", readonly=True)
    file_size = fields.Integer(readonly=True)
    attachment_id = fields.Many2one("ir.attachment", string="File", readonly=True, ondelete="set null")
    channel_log = fields.Text(string="Delivery", readonly=True)
    error_message = fields.Text(readonly=True)

    def action_download(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_url",
            "url": "/web/content/%s?download=true" % self.attachment_id.id,
            "target": "self",
        }

    @api.autovacuum
    def _gc_old_runs(self):
        days = get_config_int(self.env, RETENTION_PARAM) or DEFAULT_RETENTION_DAYS
        old = self.sudo().search([("create_date", "<", fields.Datetime.now() - timedelta(days=days))])
        old.attachment_id.unlink()
        old.unlink()
