# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
import re
from datetime import datetime, time, timedelta

import pytz

from odoo import _, api, fields, models
from odoo.exceptions import AccessError, UserError, ValidationError
from odoo.tools.safe_eval import safe_eval

from .. import engine
from .xlsx_report import GROUP_SQL, RECIPIENT_PARTNER_KEY, RECIPIENT_USER_KEY

UNSUPPORTED_FIELD_TYPES = ("binary", "image", "properties", "properties_definition", "json")
SQL_START_RE = re.compile(r"^\s*(select|with)\b", re.IGNORECASE)
SQL_COMMENT_RE = re.compile(r"(--[^\n]*)|(/\*.*?\*/)", re.DOTALL)


class XlsxReportSource(models.Model):
    _name = "xlsx.report.source"
    _description = "Excel Report Data Source"
    _order = "sequence, id"

    report_id = fields.Many2one("xlsx.report", required=True, ondelete="cascade", index=True)
    sequence = fields.Integer(default=10)
    name = fields.Char(
        required=True,
        help="Also the name of the data sheet and of the Excel table. Use it in formulas, "
             "e.g. =SUM(sales[amount_total]). Letters, digits and _ only.")
    source_type = fields.Selection(
        [("search", "Records"), ("group", "Grouped totals"), ("sql", "SQL query (admin)")],
        string="Type", required=True, default="search")

    model_id = fields.Many2one("ir.model", string="Model", ondelete="cascade",
                               domain=[("transient", "=", False)])
    model_name = fields.Char(related="model_id.model", string="Model name")
    domain = fields.Char(
        default="[]",
        help="Variables: uid, user, date_from, date_to, company_ids, and for per-recipient "
             "delivery (Pro) recipient_partner_id, recipient_user_id.")
    date_field_id = fields.Many2one(
        "ir.model.fields", string="Filter period on",
        domain="[('model_id', '=', model_id), ('ttype', 'in', ('date', 'datetime')), ('store', '=', True)]",
        ondelete="set null",
        help="Only keep records whose date falls in the report period.")
    column_ids = fields.One2many("xlsx.report.source.column", "source_id", string="Columns", copy=True)
    header_mode = fields.Selection(
        [("name", "Technical name (stable for formulas)"), ("label", "Field label")],
        default="name", required=True,
        help="Header used for columns without a custom header.")

    groupby_ids = fields.Many2many(
        "ir.model.fields", "xlsx_report_source_groupby_rel", "source_id", "field_id",
        string="Group by", domain="[('model_id', '=', model_id), ('store', '=', True)]")
    date_granularity = fields.Selection(
        [("day", "Day"), ("week", "Week"), ("month", "Month"), ("quarter", "Quarter"), ("year", "Year")],
        default="month", help="Used when grouping by a date field.")
    measure_ids = fields.Many2many(
        "ir.model.fields", "xlsx_report_source_measure_rel", "source_id", "field_id",
        string="Measures",
        domain="[('model_id', '=', model_id), ('store', '=', True), "
               "('ttype', 'in', ('integer', 'float', 'monetary'))]")
    aggregator = fields.Selection(
        [("sum", "Sum"), ("avg", "Average"), ("min", "Min"), ("max", "Max")],
        default="sum", required=True)

    sql_query = fields.Text(
        groups=GROUP_SQL,
        help="Read-only SELECT. Parameters: %(date_from)s, %(date_to)s, %(company_ids)s, "
             "%(recipient_partner_id)s, %(recipient_user_id)s (NULL unless per-recipient). "
             "Write a literal % as %%. SQL ignores record rules: anyone receiving the report "
             "sees everything the query returns.")
    row_limit = fields.Integer(default=50000, required=True)
    sql_timeout = fields.Integer(string="Timeout (s)", default=30, groups=GROUP_SQL)

    # ------------------------------------------------------------------
    # Constraints & security
    # ------------------------------------------------------------------
    @api.constrains("name")
    def _check_name(self):
        for source in self:
            if not engine.valid_table_name(source.name):
                raise ValidationError(_(
                    "Invalid data source name %r: use letters, digits and _ (max 31 characters, "
                    "not starting with a digit, not looking like a cell such as A1).", source.name))
            siblings = source.report_id.source_ids - source
            if source.name.lower() in {s.name.lower() for s in siblings}:
                raise ValidationError(_("Data source names must be unique in a report."))

    @api.constrains("source_type", "model_id", "column_ids", "measure_ids", "groupby_ids", "row_limit")
    def _check_config(self):
        for source in self.sudo():  # ir.model.fields is metadata
            if not 1 <= source.row_limit <= 500000:
                raise ValidationError(_("Row limit must be between 1 and 500,000."))
            if source.source_type == "sql":
                continue
            if not source.model_id:
                raise ValidationError(_("Choose a model for data source %s.", source.name))
            if source.source_type == "search" and not source.column_ids:
                raise ValidationError(_("Choose at least one column for data source %s.", source.name))
            if source.source_type == "group" and not source.groupby_ids:
                raise ValidationError(_("Choose at least one 'Group by' field for %s.", source.name))
            wrong_model = source.column_ids.field_id.filtered(lambda f: f.model_id != source.model_id)
            if source.source_type == "search" and wrong_model:
                raise ValidationError(_("Columns must belong to the model %s.", source.model_id.model))
            bad = (source.column_ids.field_id | source.groupby_ids).filtered(
                lambda f: f.ttype in UNSUPPORTED_FIELD_TYPES)
            if bad:
                raise ValidationError(_("Unsupported field types: %s", ", ".join(bad.mapped("name"))))

    def _check_sql_rights(self, vals_list):
        if self.env.su or self.env.user.has_group(GROUP_SQL):
            return
        if any(v.get("source_type") == "sql" or "sql_query" in v for v in vals_list):
            raise AccessError(_("Only SQL report administrators can create SQL data sources."))

    @api.model_create_multi
    def create(self, vals_list):
        self._check_sql_rights(vals_list)
        return super().create(vals_list)

    def write(self, vals):
        self._check_sql_rights([vals])
        if not (self.env.su or self.env.user.has_group(GROUP_SQL)) \
                and any(s.source_type == "sql" for s in self):
            raise AccessError(_("Only SQL report administrators can edit SQL data sources."))
        return super().write(vals)

    # ------------------------------------------------------------------
    # Fetching
    # ------------------------------------------------------------------
    def _period_bounds(self, field, period, tz_name):
        """Domain on ``field`` for the (local, inclusive) report period."""
        if field.type == "date":
            return [(field.name, ">=", period.date_from), (field.name, "<=", period.date_to)]
        tz = pytz.timezone(tz_name or "UTC")

        def utc(day):
            local = tz.localize(datetime.combine(day, time.min), is_dst=False)
            return local.astimezone(pytz.utc).replace(tzinfo=None)

        return [(field.name, ">=", utc(period.date_from)),
                (field.name, "<", utc(period.date_to + timedelta(days=1)))]

    def _get_domain(self, Model, period):
        context = {
            "uid": Model.env.uid,
            "user": Model.env.user,
            "date_from": period.date_from,
            "date_to": period.date_to,
            "company_ids": Model.env.companies.ids,
            # Per-recipient reports (Pro): ids of the person the file is for,
            # False when the report is rendered for everyone.
            "recipient_partner_id": Model.env.context.get(RECIPIENT_PARTNER_KEY) or False,
            "recipient_user_id": Model.env.context.get(RECIPIENT_USER_KEY) or False,
            "datetime": datetime,
            "timedelta": timedelta,
        }
        try:
            domain = safe_eval(self.domain or "[]", context)
        except Exception as exc:
            raise UserError(_("Invalid domain on data source %s: %s", self.name, exc)) from exc
        if self.date_field_id:
            field = Model._fields.get(self.date_field_id.name)
            if field is None:
                raise UserError(_("Field %s no longer exists.", self.date_field_id.name))
            domain = list(domain) + self._period_bounds(field, period, Model.env.context.get("tz"))
        return domain

    def _fetch(self, env, period):
        """Return :class:`engine.TableData`. ``env`` carries the owner's rights.

        The source *configuration* (model, columns...) is read as superuser because
        it is metadata (ir.model.fields); the *data* is always read through ``env``.
        """
        self.ensure_one()
        self = self.sudo()  # noqa: PLW0642 - configuration only, see docstring
        if self.source_type == "sql":
            return self._fetch_sql(env, period)
        if self.model_name not in env:
            raise UserError(_("Model %s is not installed.", self.model_name))
        Model = env[self.model_name]
        domain = self._get_domain(Model, period)
        if self.source_type == "search":
            return self._fetch_search(Model, domain)
        return self._fetch_group(Model, domain)

    def _header(self, Model, fname):
        if self.header_mode == "label":
            return Model._fields[fname]._description_string(Model.env)
        return fname

    def _fetch_search(self, Model, domain):
        columns = self.column_ids.sorted("sequence").filtered(lambda c: c.field_id.name in Model._fields)
        fnames = columns.mapped("field_id.name")
        records = Model.search(domain, limit=self.row_limit + 1)
        if len(records) > self.row_limit:
            raise UserError(_("Data source %s returns more than %s rows. Narrow the domain or "
                              "raise the row limit.", self.name, self.row_limit))
        tz = pytz.timezone(Model.env.context.get("tz") or "UTC")
        rows = []
        # read() applies field-level access (groups=) of the owner.
        for values in records.read(fnames):
            rows.append([self._convert(Model, fname, values[fname], tz) for fname in fnames])
        headers = [c.header or self._header(Model, c.field_id.name) for c in columns]
        return engine.TableData(headers, rows)

    def _fetch_group(self, Model, domain):
        groupby, headers = [], []
        for field in self.groupby_ids:
            if field.name not in Model._fields:
                continue
            spec = field.name
            if field.ttype in ("date", "datetime"):
                spec = "%s:%s" % (field.name, self.date_granularity)
            groupby.append(spec)
            headers.append(self._header(Model, field.name))
        aggregates = ["%s:%s" % (f.name, self.aggregator) for f in self.measure_ids
                      if f.name in Model._fields]
        headers += [self._header(Model, f.name) for f in self.measure_ids if f.name in Model._fields]
        headers.append("count")
        aggregates.append("__count")

        groups = Model._read_group(domain, groupby, aggregates, limit=self.row_limit + 1)
        if len(groups) > self.row_limit:
            raise UserError(_("Data source %s returns more than %s groups.", self.name, self.row_limit))
        rows = []
        for group in groups:
            row = []
            for value in group:
                if isinstance(value, models.BaseModel):
                    value = value.display_name if value else None
                row.append(value)
            rows.append(row)
        return engine.TableData(headers, rows)

    def _convert(self, Model, fname, value, tz):
        field = Model._fields[fname]
        if field.type == "boolean":
            return bool(value)
        if value is False or value is None:
            return None
        if field.type == "many2one":
            return value[1] if isinstance(value, (list, tuple)) else value
        if field.type in ("one2many", "many2many"):
            return ", ".join(Model.env[field.comodel_name].browse(value).mapped("display_name"))
        if field.type == "selection":
            return dict(field._description_selection(Model.env)).get(value, value)
        if field.type == "datetime":
            return pytz.utc.localize(value).astimezone(tz).replace(tzinfo=None)
        if field.type == "html":
            return str(value)
        return value

    def _fetch_sql(self, env, period):
        source = self.sudo()  # sql_query is restricted by groups=
        owner = env.user
        if not owner.has_group(GROUP_SQL):
            raise AccessError(_("The 'Run as' user of this report is not an SQL report administrator."))
        query = SQL_COMMENT_RE.sub(" ", source.sql_query or "").strip().rstrip(";").strip()
        if not SQL_START_RE.match(query):
            raise UserError(_("SQL data sources must start with SELECT or WITH."))
        if ";" in query:
            raise UserError(_("Only one SQL statement is allowed."))
        params = {
            "date_from": period.date_from,
            "date_to": period.date_to,
            "company_ids": tuple(env.companies.ids) or (0,),
            "recipient_partner_id": env.context.get(RECIPIENT_PARTNER_KEY) or None,
            "recipient_user_id": env.context.get(RECIPIENT_USER_KEY) or None,
        }
        # A separate cursor so the READ ONLY transaction and the timeout never
        # affect the main transaction. PostgreSQL itself rejects any write.
        with env.registry.cursor() as cr:
            cr.execute("SET TRANSACTION READ ONLY")
            cr.execute("SET LOCAL statement_timeout = %s", (max(1, source.sql_timeout or 30) * 1000,))
            try:
                cr.execute(query, params)
            except Exception as exc:
                cr.rollback()
                raise UserError(_("SQL error in data source %s: %s", self.name, exc)) from exc
            headers = [col.name for col in cr.description]
            rows = cr.fetchmany(self.row_limit + 1)
            cr.rollback()
        if len(rows) > self.row_limit:
            raise UserError(_("Data source %s returns more than %s rows.", self.name, self.row_limit))
        return engine.TableData(headers, [list(r) for r in rows])


class XlsxReportSourceColumn(models.Model):
    """One column of a 'Records' data source, in the order the user chose.

    The order matters: the template's row-2 style is applied by column position.
    """

    _name = "xlsx.report.source.column"
    _description = "Excel Report Data Source Column"
    _order = "sequence, id"

    source_id = fields.Many2one("xlsx.report.source", required=True, ondelete="cascade", index=True)
    sequence = fields.Integer(default=10)
    field_id = fields.Many2one("ir.model.fields", string="Field", required=True, ondelete="cascade")
    header = fields.Char(help="Column header in the Excel table. Empty: technical name or label, "
                              "depending on the data source setting. Formulas refer to it, "
                              "e.g. =SUM(sales[Amount]).")
