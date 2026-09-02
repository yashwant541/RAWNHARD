# -*- coding: utf-8 -*-
"""
RWA Calculation webapp - backend additions
==========================================

Paste this block at the END of your existing backend (after CATEGORY_CONFIG and all the
existing routes / helpers).  It adds:

  * a tiny admin layer  - sign in, edit each category's folders + recipe name/id, that's
    all.  The edits are persisted as `category_overrides.json` in a config managed folder
    and merged over the hard-coded CATEGORY_CONFIG on start-up and on save, so the rest of
    your code keeps reading CATEGORY_CONFIG unchanged.
  * the deep-trace routes that drive the Trace page:
        /excel-trace/open-output      open a produced workbook straight from a folder
        /excel-trace/columns          header names of a sheet
        /excel-trace/mismatch-rows    rows where two columns disagree
        /excel-trace/expression-tree  the node-by-node formula tree for one cell

Requires: `webapp_core` and `excel_deep_trace` on the project's Python libraries
(you already need them for ExcelDeepTraceEngine), and a project variable
`RWA_WEBAPP_CONFIG_FOLDER` holding the id of an (empty) managed folder for the config
JSON.  Nothing is hand-authored in that folder - the backend writes admin.json /
category_overrides.json itself.
"""

from webapp_core import auth as _auth

_ADMIN_FILE = "admin.json"
_OVERRIDES_FILE = "category_overrides.json"
_EDITABLE_KEYS = (
    "display_name", "description", "icon", "recipe_id", "recipe_name",
    "data_folder_id", "mapping_folder_id", "output_folder_id", "template_folder_id",
)


def _rwa_cfg_folder():
    folder_id = None
    try:
        folder_id = dataiku.get_custom_variables().get("RWA_WEBAPP_CONFIG_FOLDER")
    except Exception:  # noqa: BLE001
        pass
    folder_id = folder_id or os.environ.get("RWA_WEBAPP_CONFIG_FOLDER")
    if not folder_id:
        raise RuntimeError("project variable RWA_WEBAPP_CONFIG_FOLDER is not set")
    return dataiku.Folder(folder_id)


def _rwa_read_json(name, default):
    try:
        folder = _rwa_cfg_folder()
        present = {normalize_folder_path(p) for p in folder.list_paths_in_partition()}
        if ("/" + name) not in present:
            return default
        with folder.get_download_stream("/" + name) as stream:
            raw = stream.read()
        if not raw:
            return default
        return json.loads(raw.decode("utf-8") if isinstance(raw, bytes) else raw)
    except Exception:  # noqa: BLE001
        return default


def _rwa_write_json(name, obj):
    folder = _rwa_cfg_folder()
    folder.upload_stream(
        "/" + name,
        io.BytesIO(json.dumps(obj, ensure_ascii=False, indent=2, default=str).encode("utf-8")),
    )


def _rwa_apply_overrides():
    overrides = _rwa_read_json(_OVERRIDES_FILE, {}) or {}
    for key, patch in overrides.items():
        if key in CATEGORY_CONFIG and isinstance(patch, dict):
            CATEGORY_CONFIG[key].update(
                {k: v for k, v in patch.items() if k in _EDITABLE_KEYS})
    return overrides


try:
    _rwa_apply_overrides()
except Exception as _exc:  # noqa: BLE001
    LOG.warning("RWA admin overrides not applied: %s", _exc)


def _rwa_admin_record():
    rec = _rwa_read_json(_ADMIN_FILE, None)
    if not rec:
        salt, digest = _auth.hash_password("changeme")
        rec = {"username": "admin", "salt": salt, "hash": digest,
               "must_change": True, "secret": _auth.new_secret()}
        try:
            _rwa_write_json(_ADMIN_FILE, rec)
        except Exception:  # noqa: BLE001
            pass
    return rec


def _rwa_require_admin():
    rec = _rwa_admin_record()
    token = request.headers.get("X-Admin-Token", "")
    return rec if _auth.verify_token(token, rec.get("secret", "")) else None


# --------------------------------------------------------------------------- admin API
@app.route("/api/whoami", methods=["GET"])
def api_whoami():
    return success_response(data={
        "app_title": APP_TITLE,
        "user": os.environ.get("DKU_CURRENT_USER") or None,
    })


@app.route("/api/login", methods=["POST"])
def api_admin_login():
    body = request.get_json(force=True, silent=True) or {}
    rec = _rwa_admin_record()
    ok = (body.get("username") == rec["username"]
          and _auth.verify_password(body.get("password", ""), rec["salt"], rec["hash"]))
    if not ok:
        return error_response("Invalid credentials.", status_code=401)
    return success_response(data={
        "token": _auth.make_token(rec["username"], rec["secret"]),
        "must_change": rec.get("must_change", False),
    })


@app.route("/api/admin/change-credentials", methods=["POST"])
def api_admin_change_credentials():
    rec = _rwa_require_admin()
    if not rec:
        return error_response("Admin authentication required.", status_code=401)
    body = request.get_json(force=True, silent=True) or {}
    if not _auth.verify_password(body.get("current_password", ""), rec["salt"], rec["hash"]):
        return error_response("Current password is wrong.", status_code=403)
    new_pw = body.get("new_password", "")
    if len(new_pw) < 6:
        return error_response("New password must be at least 6 characters.")
    salt, digest = _auth.hash_password(new_pw)
    _rwa_write_json(_ADMIN_FILE, {
        "username": body.get("new_username") or rec["username"],
        "salt": salt, "hash": digest, "must_change": False, "secret": rec["secret"],
    })
    return success_response(message="Credentials updated - sign in again.")


@app.route("/api/admin/config", methods=["GET"])
def api_admin_config():
    if not _rwa_require_admin():
        return error_response("Admin authentication required.", status_code=401)
    out = []
    for key, cfg in CATEGORY_CONFIG.items():
        def clean(field):
            val = cfg.get(field, "")
            return "" if is_placeholder(val) else val
        out.append({
            "key": key,
            "display_name": cfg.get("display_name", key),
            "description": cfg.get("description", ""),
            "icon": cfg.get("icon", "layers"),
            "recipe_id": clean("recipe_id"),
            "recipe_name": cfg.get("recipe_name", ""),
            "folders": {
                "data": clean("data_folder_id"),
                "mapping": clean("mapping_folder_id"),
                "output": clean("output_folder_id"),
                "template": clean("template_folder_id"),
            },
        })
    return success_response(data={"app_title": APP_TITLE, "categories": out})


@app.route("/api/admin/category/<key>", methods=["POST"])
def api_admin_save_category(key):
    if not _rwa_require_admin():
        return error_response("Admin authentication required.", status_code=401)
    if key not in CATEGORY_CONFIG:
        return error_response("Unknown category.", status_code=404)
    body = request.get_json(force=True, silent=True) or {}
    folders = body.get("folders", {}) or {}
    current = CATEGORY_CONFIG[key]
    patch = {
        "display_name": (body.get("display_name") or current.get("display_name") or key).strip(),
        "description": (body.get("description") or current.get("description", "")).strip(),
        "recipe_id": (body.get("recipe_id") or "").strip(),
        "recipe_name": (body.get("recipe_name") or "").strip(),
        "data_folder_id": (folders.get("data") or "").strip(),
        "mapping_folder_id": (folders.get("mapping") or "").strip(),
        "output_folder_id": (folders.get("output") or "").strip(),
        "template_folder_id": (folders.get("template") or "").strip(),
    }
    overrides = _rwa_read_json(_OVERRIDES_FILE, {}) or {}
    overrides[key] = dict(overrides.get(key, {}), **patch)
    _rwa_write_json(_OVERRIDES_FILE, overrides)
    CATEGORY_CONFIG[key].update(patch)
    append_audit_log(key, "ADMIN_UPDATE_CATEGORY", details=patch)
    return success_response(message="Category configuration saved.", data={"key": key})


# --------------------------------------------------------------- deep-trace extra routes
@app.route("/excel-trace/open-output", methods=["POST"])
def excel_trace_open_output():
    body = request.get_json(force=True, silent=True) or {}
    category = body.get("category")
    path = body.get("path")
    folder_type = body.get("folder_type", "output")

    if not category or not path:
        return json_error("category and path are required.")
    try:
        category_key, _ = get_category_config(category)
        folder = get_folder(category_key, folder_type)
        content = read_managed_folder_file(folder, normalize_folder_path(path))

        engine = ExcelDeepTraceEngine(
            workbook_bytes=content,
            filename=os.path.basename(path),
            max_expanded_range=DEFAULT_MAX_EXPANDED_RANGE,
        )
        summary = engine.scan()

        session_id = uuid.uuid4().hex
        with TRACE_SESSION_LOCK:
            TRACE_SESSIONS[session_id] = {
                "session_id": session_id, "engine": engine, "directory": None,
                "filename": os.path.basename(path), "category": category_key,
            }
        return jsonify({"success": True, "session_id": session_id,
                        "filename": os.path.basename(path), "summary": summary})
    except Exception as exc:  # noqa: BLE001
        return json_error("Could not open the output workbook.", status=500, details=str(exc))


@app.route("/excel-trace/columns", methods=["POST"])
def excel_trace_columns():
    body = request.get_json(force=True, silent=True) or {}
    try:
        session = require_session(body.get("session_id"))
        return jsonify({"success": True, **session["engine"].sheet_columns(body.get("sheet"))})
    except KeyError as exc:
        return json_error(str(exc), status=404)
    except Exception as exc:  # noqa: BLE001
        return json_error("Could not read the column list.", status=500, details=str(exc))


@app.route("/excel-trace/mismatch-rows", methods=["POST"])
def excel_trace_mismatch_rows():
    body = request.get_json(force=True, silent=True) or {}
    try:
        session = require_session(body.get("session_id"))
        result = session["engine"].mismatch_rows(
            body.get("sheet"), body.get("left"), body.get("right"),
            float(body.get("tolerance", 0.01)))
        return jsonify({"success": True, **result})
    except KeyError as exc:
        return json_error(str(exc), status=404)
    except Exception as exc:  # noqa: BLE001
        return json_error("Mismatch scan failed.", status=500, details=str(exc))


@app.route("/excel-trace/expression-tree", methods=["POST"])
def excel_trace_expression_tree():
    body = request.get_json(force=True, silent=True) or {}
    try:
        session = require_session(body.get("session_id"))
        tree = session["engine"].expression_tree(
            body.get("cell"), int(body.get("max_depth", 8)))
        return jsonify({"success": True, "tree": tree})
    except KeyError as exc:
        return json_error(str(exc), status=404)
    except Exception as exc:  # noqa: BLE001
        return json_error("Expression tree failed.", status=500, details=str(exc))
