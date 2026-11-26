# Part of Excel Template Reports. License LGPL-3, see LICENSE file. Copyright (c) 2026 Gray Bob.
"""Small helpers that hide API differences between Odoo 17, 18, 19 and 20.

Python code of both addons is identical for every Odoo version; only XML and
security files are rewritten per version by scripts/build_versions.py.
"""
import base64

from odoo import release

ODOO_MAJOR = release.version_info[0]

# Odoo 17 calls list views "tree".
LIST_VIEW = "list" if ODOO_MAJOR >= 18 else "tree"


def binary_bytes(value):
    """Raw bytes of a Binary field value.

    Odoo 20 returns a ``BinaryValue`` (raw bytes in ``.content``); Odoo <= 19
    returns base64.
    """
    if not value:
        return None
    content = getattr(value, "content", None)
    if content is not None:
        return bytes(content)
    return base64.b64decode(value)


def attachment_bytes(attachment):
    """Content of an ``ir.attachment``.

    Unlike other Binary fields, ``raw`` is already raw bytes on Odoo <= 19 (never
    base64), so do not use :func:`binary_bytes` for it. Odoo 20 wraps it in a
    ``BinaryValue``.
    """
    value = attachment.raw
    if not value:
        return b""
    content = getattr(value, "content", None)
    return bytes(content if content is not None else value)


def get_config_str(env, key):
    params = env["ir.config_parameter"].sudo()
    if hasattr(params, "get_str"):  # Odoo 20
        return params.get_str(key)
    return params.get_param(key)


def get_config_int(env, key):
    params = env["ir.config_parameter"].sudo()
    if hasattr(params, "get_int"):  # Odoo 20
        return params.get_int(key)
    try:
        return int(params.get_param(key) or 0)
    except ValueError:
        return 0


def set_config_str(env, key, value):
    params = env["ir.config_parameter"].sudo()
    if hasattr(params, "set_str"):  # Odoo 20
        return params.set_str(key, value)
    return params.set_param(key, value)
