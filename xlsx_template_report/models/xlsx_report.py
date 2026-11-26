# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
import logging
import time
from datetime import datetime

import pytz

from odoo import _, api, fields, models
from odoo.addons.base.models.res_partner import _tz_get
from odoo.exceptions import AccessError, UserError, ValidationError

from .. import engine
from .compat import LIST_VIEW, binary_bytes  # noqa: F401 - binary_bytes re-exported for the Pro addon

_logger = logging.getLogger(__name__)

XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GROUP_USER = "xlsx_template_report.group_xlsx_report_user"
GROUP_MANAGER = "xlsx_template_report.group_xlsx_report_manager"
GROUP_SQL = "xlsx_template_report.group_xlsx_report_sql"
# Context keys set while rendering a file for one recipient (used by the Pro addon).
RECIPIENT_PARTNER_KEY = "xlsx_recipient_partner_id"
RECIPIENT_USER_KEY = "xlsx_recipient_user_id"

WEEKDAYS = [("0", "Monday"), ("1", "Tuesday"), ("2", "Wednesday"), ("3", "Thursday"),
            ("4", "Friday"), ("5", "Saturday"), ("6", "Sunday")]


class XlsxReport(models.Model):
    _name = "xlsx.report"
    _description = "Excel Template Report"
    _inherit = ["mail.thread"]
    _order = "name"

    name = fields.Char(required=True, tracking=True)
    active = fields.Boolean(default=True)
    owner_id = fields.Many2one(
        "res.users", string="Run as", required=True, tracking=True,
        default=lambda self: self.env.user,
        help="Data is read with the access rights of this user. Everyone who receives the "
             "file sees what this user is allowed to see.")
    company_id = fields.Many2one("res.company", required=True, default=lambda self: self.env.company)
    company_ids = fields.Many2many(
        "res.company", string="Data from companies",
        default=lambda self: self.env.company,
        help="Companies whose data is included. Limited to the companies of the 'Run as' user.")

    template = fields.Binary(attachment=True, help="Excel (.xlsx) template. Leave empty to "
                                                    "get a plain workbook with one sheet per data source.")
    template_filename = fields.Char()
    template_info = fields.Text(compute="_compute_template_info")

    timezone = fields.Selection(_tz_get, required=True,
                                default=lambda self: self.env.user.tz or "UTC")
    period_type = fields.Selection(engine.PERIOD_TYPES, string="Period", required=True,
                                   default="last_week")
    period_days = fields.Integer(string="Number of days", default=7)
    week_start = fields.Selection(WEEKDAYS, required=True, default="0")
    period_preview = fields.Char(string="Period if run now", compute="_compute_period_preview")
    filename_pattern = fields.Char(
        string="File name", required=True, default="{report} {date_from} - {date_to}",
        help="Available keys: {report}, {date_from}, {date_to}, {period}, {today}")

    source_ids = fields.One2many("xlsx.report.source", "report_id", string="Data sources", copy=True)
    run_ids = fields.One2many("xlsx.report.run", "report_id", string="Runs")
    run_count = fields.Integer(compute="_compute_run_count")

    # ------------------------------------------------------------------
    # Compute / constraints
    # ------------------------------------------------------------------
    @api.depends("template")
    def _compute_template_info(self):
        for report in self:
            if not report.template:
                report.template_info = False
                continue
            try:
                info = engine.inspect_template(report._template_bytes())
            except engine.TemplateError as exc:
                report.template_info = _("Invalid template: %s", exc)
                continue
            lines = [
                _("Sheets: %s", ", ".join(info.sheet_names)),
                _("Excel tables: %s", ", ".join(info.table_names) or "-"),
                _("Placeholders: %s", ", ".join(info.placeholders) or "-"),
            ] + ["⚠ " + w for w in info.warnings]
            report.template_info = "\n".join(lines)

    def _template_bytes(self):
        """The template file content. The web client reads binaries with
        ``bin_size=True`` (it then gets "12.3 Kb"), so always read the real bytes."""
        self.ensure_one()
        return binary_bytes(self.with_context(bin_size=False, bin_size_template=False).template)

    @api.depends("period_type", "period_days", "week_start", "timezone")
    def _compute_period_preview(self):
        for report in self:
            try:
                report.period_preview = report._get_period().label
            except ValueError:
                report.period_preview = False

    def _compute_run_count(self):
        counts = dict(self.env["xlsx.report.run"]._read_group(
            [("report_id", "in", self.ids)], ["report_id"], ["__count"]))
        for report in self:
            report.run_count = counts.get(report, 0)

    @api.constrains("template", "template_filename")
    def _check_template(self):
        for report in self.filtered("template"):
            name = (report.template_filename or "").lower()
            if name and not name.endswith(".xlsx"):
                raise ValidationError(_("The template must be an .xlsx file."))
            try:
                engine.inspect_template(report._template_bytes())
            except engine.TemplateError as exc:
                raise ValidationError(_("Invalid template: %s", exc)) from exc

    @api.constrains("period_type", "period_days")
    def _check_period_days(self):
        for report in self:
            if report.period_type == "last_n_days" and not 1 <= report.period_days <= 366:
                raise ValidationError(_("Number of days must be between 1 and 366."))

    # ------------------------------------------------------------------
    # Ownership: nobody may borrow another user's access rights
    # ------------------------------------------------------------------
    def _is_manager(self):
        return self.env.su or self.env.user.has_group(GROUP_MANAGER)

    @api.model_create_multi
    def create(self, vals_list):
        if not self._is_manager():
            for vals in vals_list:
                vals["owner_id"] = self.env.uid
        return super().create(vals_list)

    def write(self, vals):
        if not self._is_manager():
            if "owner_id" in vals and vals["owner_id"] != self.env.uid:
                raise AccessError(_("Only Excel Report managers can change the 'Run as' user."))
            if any(r.owner_id != self.env.user for r in self):
                # Editing someone else's report makes you its owner.
                vals = dict(vals, owner_id=self.env.uid)
        return super().write(vals)

    # ------------------------------------------------------------------
    # Rendering
    # ------------------------------------------------------------------
    def _get_period(self):
        self.ensure_one()
        today = datetime.now(pytz.timezone(self.timezone or "UTC")).date()
        return engine.compute_period(self.period_type or "last_week", today,
                                     week_start=int(self.week_start or 0),
                                     n_days=self.period_days or 7)

    def _get_data_env(self, recipient=None):
        """Environment used to read data: the owner's rights, never superuser.

        ``recipient`` (a ``res.partner``) only adds its ids to the context, so data
        source domains can filter on ``recipient_partner_id`` / ``recipient_user_id``.
        It never grants more rights than the owner has.
        """
        self.ensure_one()
        owner = self.owner_id
        companies = (self.company_ids or self.company_id) & owner.company_ids
        if not companies:
            raise UserError(_("The 'Run as' user has no access to the selected companies."))
        main_company = self.company_id if self.company_id in companies else companies[0]
        context = dict(self.env.context, tz=self.timezone, lang=owner.lang,
                       allowed_company_ids=[main_company.id] + (companies - main_company).ids)
        context.pop(RECIPIENT_PARTNER_KEY, None)
        context.pop(RECIPIENT_USER_KEY, None)
        if recipient:
            user = self.env["res.users"].sudo().with_context(active_test=False).search(
                [("partner_id", "=", recipient.id)], limit=1)
            context[RECIPIENT_PARTNER_KEY] = recipient.id
            context[RECIPIENT_USER_KEY] = user.id or False
        return self.env(user=owner.id, su=False, context=context)

    def _build_filename(self, period):
        self.ensure_one()
        keys = {
            "report": self.name,
            "date_from": period.date_from.isoformat(),
            "date_to": period.date_to.isoformat(),
            "period": period.label,
            "today": fields.Date.context_today(self.with_context(tz=self.timezone)).isoformat(),
        }
        try:
            name = self.filename_pattern.format_map(keys)
        except (KeyError, ValueError, IndexError):
            name = "%s %s" % (self.name, period.label)
        for char in '\\/:*?"<>|':
            name = name.replace(char, "-")
        return name.strip()[:150] + ".xlsx"

    def _render(self, recipient=None):
        """Return ``(content, filename, row_count, warnings)``.

        ``recipient`` (a ``res.partner``): render the file for that person only,
        see :meth:`_get_data_env`.
        """
        self.ensure_one()
        if not self.source_ids and not self.template:
            raise UserError(_("Add a template or at least one data source."))
        period = self._get_period()
        data_env = self._get_data_env(recipient)
        tables, row_count = {}, 0
        for source in self.source_ids.sorted("sequence"):
            data = source._fetch(data_env, period)
            tables[source.name] = data
            row_count += len(data.rows)

        tz = pytz.timezone(self.timezone or "UTC")
        values = {
            "report_name": self.name,
            "company_name": self.company_id.name,
            "date_from": period.date_from,
            "date_to": period.date_to,
            "period_label": period.label,
            "generated_at": datetime.now(tz).replace(tzinfo=None, microsecond=0),
            "run_as": self.owner_id.name,
            "recipient_name": recipient.sudo().name if recipient else "",
        }
        for name, data in tables.items():
            values["%s_count" % name] = len(data.rows)

        template = self._template_bytes()
        try:
            result = engine.render(template, values, tables)
        except engine.TemplateError as exc:
            raise UserError(str(exc)) from exc
        return result.content, self._build_filename(period), row_count, result.warnings

    def _deliver(self, run, attachment, **kwargs):
        """Hook for delivery add-ons. Returns a list of status lines for the run log.

        The free module only stores the file; ``xlsx_template_report_pro`` sends it.
        """
        return []

    def _execute(self, trigger, **deliver_kwargs):
        """Render, store and deliver the report. Returns the ``xlsx.report.run``.

        Errors are recorded on the run and re-raised for manual triggers.
        """
        self.ensure_one()
        run = self.env["xlsx.report.run"].sudo().create(
            {"report_id": self.id, "trigger": trigger, "state": "running"})
        started = time.monotonic()
        try:
            # Savepoint: a database error while rendering must not poison the
            # transaction, so the failure can still be recorded on the run.
            with self.env.cr.savepoint():
                content, filename, row_count, warnings = self._render()
                attachment = self.env["ir.attachment"].sudo().create({
                    "name": filename,
                    "raw": content,
                    "mimetype": XLSX_MIMETYPE,
                    "res_model": run._name,
                    "res_id": run.id,
                })
                run.row_count = row_count
                results = self._deliver(run, attachment, **deliver_kwargs)
            run.write({
                "state": "done",
                "file_size": len(content),
                "attachment_id": attachment.id,
                "channel_log": "\n".join(results + ["⚠ " + w for w in warnings]) or False,
                "duration": time.monotonic() - started,
            })
        except Exception as exc:
            _logger.warning("Excel report %s (id %s) failed", self.name, self.id, exc_info=True)
            run.write({
                "state": "failed",
                "error_message": str(exc),
                "duration": time.monotonic() - started,
            })
            if trigger != "schedule":
                raise
        return run

    # ------------------------------------------------------------------
    # Buttons
    # ------------------------------------------------------------------
    def action_download(self):
        self.ensure_one()
        run = self._execute("manual")
        return {
            "type": "ir.actions.act_url",
            "url": "/web/content/%s?download=true" % run.attachment_id.id,
            "target": "self",
        }

    def action_view_runs(self):
        self.ensure_one()
        return {
            "type": "ir.actions.act_window",
            "name": _("Runs"),
            "res_model": "xlsx.report.run",
            "view_mode": "%s,form" % LIST_VIEW,
            "domain": [("report_id", "=", self.id)],
            "context": {"default_report_id": self.id},
        }

    def _notify(self, message, notif_type="success"):
        return {
            "type": "ir.actions.client",
            "tag": "display_notification",
            "params": {"message": message, "type": notif_type, "sticky": False},
        }
