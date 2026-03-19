#!/usr/bin/env python3
"""olook — Local Outlook CLI via COM automation. No cloud, no OAuth, no telemetry."""

import datetime
import io
import json
import os
import re
import subprocess
import sys
import time

# Force UTF-8 stdout on Windows (cp1252 chokes on emoji folder names)
if sys.stdout.encoding != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import html as html_mod

import click
import pythoncom
import win32com.client


# ── Security: output sanitization ──────────────────────────────────────
# Email content is untrusted. Sanitize before terminal display to prevent
# ANSI escape injection, bidi spoofing, and control character attacks.

_ANSI_RE = re.compile(r'\x1b(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
_BIDI_CHARS = set('\u200b\u200c\u200d\u200e\u200f\u202a\u202b\u202c\u202d\u202e'
                  '\u2066\u2067\u2068\u2069\ufeff')


def _sanitize(s: str) -> str:
    """Strip ANSI escapes, bidi overrides, and C0/C1 control chars from untrusted content."""
    if not isinstance(s, str):
        return str(s)
    s = _ANSI_RE.sub('', s)
    s = ''.join(c for c in s if c not in _BIDI_CHARS)
    # Strip C0 controls (except tab/newline) and C1 controls (U+0080-009F)
    # C1 includes U+009B (CSI) which terminals interpret as ESC[
    s = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\x80-\x9f]', '', s)
    return s


def _sanitize_dict(d: dict) -> dict:
    """Sanitize all string values in a message dict."""
    return {k: _sanitize(v) if isinstance(v, str) else v for k, v in d.items()}


_ENTRYID_RE = re.compile(r'^[0-9A-Fa-f]{32,}$')


def _validate_entry_id(entry_id: str) -> str:
    """Validate that an EntryID looks like a hex MAPI identifier."""
    entry_id = entry_id.strip()
    if not _ENTRYID_RE.match(entry_id):
        raise click.BadParameter(
            f"Invalid EntryID format (expected hex string, got {len(entry_id)} chars)"
        )
    return entry_id


_IMPORTANCE_MAP = {0: "Low", 1: "Normal", 2: "High"}

# ── MAPI folder constants ──────────────────────────────────────────────

FOLDER_MAP = {
    "inbox": 6, "outbox": 4, "sent": 5, "deleted": 3,
    "drafts": 16, "junk": 23, "calendar": 9, "contacts": 10,
    "tasks": 13, "notes": 12,
}


# ── Outlook connection ─────────────────────────────────────────────────

_OUTLOOK_PATHS = [
    os.path.join(os.environ.get("ProgramFiles", ""), "Microsoft Office", "root", "Office16", "OUTLOOK.EXE"),
    os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Microsoft Office", "root", "Office16", "OUTLOOK.EXE"),
    os.path.join(os.environ.get("ProgramFiles", ""), "Microsoft Office", "Office16", "OUTLOOK.EXE"),
]


def _find_outlook_exe() -> str:
    """Locate OUTLOOK.EXE on disk. Fails if not found (no PATH fallback to avoid hijack)."""
    # Allow explicit override via environment variable
    env_exe = os.environ.get("OLOOK_OUTLOOK_EXE")
    if env_exe:
        if '"' in env_exe or '&' in env_exe or '|' in env_exe:
            raise click.ClickException("OLOOK_OUTLOOK_EXE contains invalid characters")
        if os.path.isfile(env_exe):
            return env_exe
    for p in _OUTLOOK_PATHS:
        if os.path.isfile(p):
            return p
    raise click.ClickException(
        "OUTLOOK.EXE not found at standard Office paths. "
        "Set OLOOK_OUTLOOK_EXE environment variable to specify the path."
    )


_com_initialized = False


def _ensure_com():
    """Initialize COM exactly once per process."""
    global _com_initialized
    if not _com_initialized:
        pythoncom.CoInitialize()
        _com_initialized = True


def _is_outlook_running() -> bool:
    """Check if Outlook is registered in the COM Running Object Table."""
    try:
        _ensure_com()
        win32com.client.GetActiveObject("Outlook.Application")
        return True
    except Exception:
        return False


def _launch_outlook_hidden():
    """Start Outlook in COM-server mode (/embedding) with no visible window."""
    exe = _find_outlook_exe()
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 6  # SW_MINIMIZE
    subprocess.Popen(
        [exe, "/embedding"],
        startupinfo=startupinfo,
        creationflags=subprocess.DETACHED_PROCESS,
    )
    for _ in range(30):  # poll up to 15s
        time.sleep(0.5)
        if _is_outlook_running():
            return
    raise click.ClickException("Outlook failed to start within 15 seconds")


def get_outlook():
    """Connect to a running Outlook instance, or silently launch one."""
    _ensure_com()
    try:
        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        pass
    _launch_outlook_hidden()
    return win32com.client.GetActiveObject("Outlook.Application")


# ── Helpers ────────────────────────────────────────────────────────────

def get_folder(ns, folder_name: str):
    """Resolve a folder by name ('Inbox') or path ('Inbox/Projects')."""
    folder_name = folder_name.strip()
    key = folder_name.lower().split("/")[0]
    if key in FOLDER_MAP:
        base = ns.GetDefaultFolder(FOLDER_MAP[key])
        for part in folder_name.split("/")[1:]:
            base = base.Folders[part]
        return base
    store_count = 0
    for store in ns.Stores:
        store_count += 1
        try:
            target = store.GetRootFolder()
            for part in folder_name.split("/"):
                target = target.Folders[part]
            return target
        except Exception:
            continue
    raise click.ClickException(
        f"Folder '{folder_name}' not found (searched {store_count} store(s))"
    )


def msg_to_dict(msg, body_len: int = 500) -> dict:
    """Extract message fields into a plain dict."""
    d = {}
    for attr, default in [
        ("Subject", ""), ("SenderName", ""), ("SenderEmailAddress", ""),
        ("ReceivedTime", ""), ("EntryID", ""), ("UnRead", None),
        ("FlagRequest", ""), ("Categories", ""), ("Importance", 1),
        ("Size", 0), ("Attachments", None),
    ]:
        try:
            val = getattr(msg, attr)
            if attr == "ReceivedTime":
                val = str(val)
            elif attr == "Attachments":
                val = val.Count if val else 0
            d[attr] = val
        except Exception:
            d[attr] = default
    if body_len > 0:
        try:
            d["Body"] = msg.Body[:body_len]
        except Exception:
            d["Body"] = ""
    return _sanitize_dict(d)


def format_msg(d: dict, compact: bool = False) -> str:
    """Format a message dict for terminal output."""
    if compact:
        subj = d.get("Subject", "")[:60]
        sender = d.get("SenderName", "")[:20]
        date = d.get("ReceivedTime", "")[:19]
        unread = "*" if d.get("UnRead") else " "
        return f"[{unread}] {date}  {sender:<20}  {subj}"
    lines = [
        f"Subject: {d.get('Subject', '')}",
        f"From: {d.get('SenderName', '')} <{d.get('SenderEmailAddress', '')}>",
        f"Date: {d.get('ReceivedTime', '')}",
        f"Unread: {d.get('UnRead', '')}",
        f"Importance: {_IMPORTANCE_MAP.get(d.get('Importance', 1), 'Unknown')}",
        f"Categories: {d.get('Categories', '')}",
        f"Flag: {d.get('FlagRequest', '')}",
        f"Attachments: {d.get('Attachments', 0)}",
        f"ID: {d.get('EntryID', '')}",
    ]
    if "Body" in d:
        lines.append(f"Body: {d['Body']}")
    return "\n".join(lines)


_PIPED_WARNING_SHOWN = False


def output(data, as_json: bool):
    """Emit data as JSON or plain text. Warns when piped to signal untrusted content."""
    global _PIPED_WARNING_SHOWN
    if not sys.stdout.isatty() and not _PIPED_WARNING_SHOWN:
        click.echo(
            "# OLOOK WARNING: Output contains untrusted email content. "
            "Do not feed to an LLM without a trust boundary.",
            err=True,
        )
        _PIPED_WARNING_SHOWN = True
    if as_json:
        click.echo(json.dumps(data, indent=2, default=str))
    elif isinstance(data, str):
        click.echo(data)
    elif isinstance(data, list):
        click.echo("\n---\n".join(data) if data else "No results.")


# ── CLI ────────────────────────────────────────────────────────────────

@click.group()
@click.option("--json", "as_json", is_flag=True, help="Output as JSON")
@click.pass_context
def cli(ctx, as_json):
    """olook - Local Outlook CLI. Talks directly to Outlook Desktop via COM.

    \b
    No cloud APIs. No OAuth. No telemetry. Just your mailbox.
    Requires: Windows, Outlook Desktop (classic), Python 3.8+, pywin32, click.
    """
    ctx.ensure_object(dict)
    ctx.obj["json"] = as_json


# ── Reading ────────────────────────────────────────────────────────────

@cli.command()
@click.option("-n", "--count", default=10, help="Number of emails")
@click.option("-f", "--folder", default="Inbox", help="Folder name or path")
@click.pass_context
def inbox(ctx, count, folder):
    """List recent emails."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    fld = get_folder(ns, folder)
    messages = fld.Items
    messages.Sort("[ReceivedTime]", True)
    results = []
    for i, msg in enumerate(messages):
        if i >= count:
            break
        try:
            results.append(msg_to_dict(msg, body_len=0))
        except Exception:
            continue
    if ctx.obj["json"]:
        output(results, True)
    else:
        for d in results:
            click.echo(format_msg(d, compact=True))
        if not results:
            click.echo("No emails found.")


@cli.command()
@click.argument("entry_id")
@click.option("--body-limit", default=0, help="Max body chars (0 = unlimited)")
@click.pass_context
def read(ctx, entry_id, body_limit):
    """Read full email by EntryID."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    msg = ns.GetItemFromID(_validate_entry_id(entry_id))
    d = msg_to_dict(msg, body_len=0)
    raw_body = msg.Body[:body_limit] if body_limit > 0 else msg.Body
    d["Body"] = _sanitize(raw_body)
    if ctx.obj["json"]:
        output(d, True)
    else:
        click.echo(format_msg(d))


# ── Search ─────────────────────────────────────────────────────────────

@cli.command()
@click.argument("query")
@click.option("-f", "--folder", default="Inbox")
@click.option("-n", "--count", default=20)
@click.option("--field", default="all", type=click.Choice(["subject", "body", "from", "all"]))
@click.pass_context
def search(ctx, query, folder, count, field):
    """Search emails using Outlook's native DASL index."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    fld = get_folder(ns, folder)
    if re.search(r'[\x00-\x1f\x7f]', query):
        raise click.BadParameter("Query contains control characters")
    q = query.replace("'", "''").replace("%", "[%]").replace("_", "[_]")
    filters = {
        "subject": f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{q}%'",
        "body": f"@SQL=\"urn:schemas:httpmail:textdescription\" LIKE '%{q}%'",
        "from": f"@SQL=\"urn:schemas:httpmail:fromemail\" LIKE '%{q}%' OR \"urn:schemas:httpmail:fromname\" LIKE '%{q}%'",
        "all": f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{q}%' OR \"urn:schemas:httpmail:textdescription\" LIKE '%{q}%' OR \"urn:schemas:httpmail:fromname\" LIKE '%{q}%'",
    }
    items = fld.Items.Restrict(filters.get(field, filters["all"]))
    items.Sort("[ReceivedTime]", True)
    results = []
    for i, msg in enumerate(items):
        if i >= count:
            break
        try:
            results.append(msg_to_dict(msg, 300))
        except Exception:
            continue
    if ctx.obj["json"]:
        output(results, True)
    else:
        for d in results:
            click.echo(format_msg(d, compact=True))
        if not results:
            click.echo("No emails found.")


# ── Composing ──────────────────────────────────────────────────────────

@cli.command()
@click.option("--to", required=True, help="Recipient(s), semicolon-separated")
@click.option("--subject", required=True)
@click.option("--body", required=True)
@click.option("--cc", default="")
@click.option("--bcc", default="")
@click.pass_context
def send(ctx, to, subject, body, cc, bcc):
    """Send an email."""
    ol = get_outlook()
    mail = ol.CreateItem(0)
    mail.To = to
    mail.Subject = subject
    mail.Body = body
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc
    mail.Send()
    result = {"status": "sent", "to": to, "subject": subject}
    output(result, ctx.obj["json"]) if ctx.obj["json"] else click.echo(f"Sent to {to}")


@cli.command()
@click.argument("entry_id")
@click.option("--body", required=True)
@click.option("--all", "reply_all", is_flag=True, help="Reply to all recipients")
@click.pass_context
def reply(ctx, entry_id, body, reply_all):
    """Reply to an email by EntryID."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    msg = ns.GetItemFromID(_validate_entry_id(entry_id))
    r = msg.ReplyAll() if reply_all else msg.Reply()
    # Preserve HTML formatting if original is HTML (BodyFormat 2)
    try:
        if msg.BodyFormat == 2:  # olFormatHTML
            r.HTMLBody = f"<p>{html_mod.escape(body)}</p>" + r.HTMLBody
        else:
            r.Body = body + r.Body
    except Exception:
        r.Body = body + r.Body
    r.Send()
    mode = "all" if reply_all else "sender only"
    result = {"status": "replied", "mode": mode, "entry_id": entry_id}
    output(result, ctx.obj["json"]) if ctx.obj["json"] else click.echo(f"Reply sent ({mode})")


@cli.command()
@click.argument("entry_id")
@click.option("--to", required=True)
@click.option("--body", default="")
@click.pass_context
def forward(ctx, entry_id, to, body):
    """Forward an email by EntryID."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    msg = ns.GetItemFromID(_validate_entry_id(entry_id))
    fwd = msg.Forward()
    fwd.To = to
    if body:
        try:
            if msg.BodyFormat == 2:  # olFormatHTML
                fwd.HTMLBody = f"<p>{html_mod.escape(body)}</p>" + fwd.HTMLBody
            else:
                fwd.Body = body + fwd.Body
        except Exception:
            fwd.Body = body + fwd.Body
    fwd.Send()
    result = {"status": "forwarded", "to": to, "entry_id": entry_id}
    output(result, ctx.obj["json"]) if ctx.obj["json"] else click.echo(f"Forwarded to {to}")


# ── Organizing ─────────────────────────────────────────────────────────

@cli.command()
@click.argument("entry_id")
@click.option("--to", "dest", required=True, help="Destination folder")
@click.pass_context
def move(ctx, entry_id, dest):
    """Move an email to another folder."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    msg = ns.GetItemFromID(_validate_entry_id(entry_id))
    target = get_folder(ns, dest)
    msg.Move(target)
    result = {"status": "moved", "destination": dest, "entry_id": entry_id}
    output(result, ctx.obj["json"]) if ctx.obj["json"] else click.echo(f"Moved to {dest}")


@cli.command()
@click.argument("entry_id")
@click.option("--text", default="Follow up")
@click.pass_context
def flag(ctx, entry_id, text):
    """Flag an email for follow-up."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    msg = ns.GetItemFromID(_validate_entry_id(entry_id))
    msg.FlagRequest = text
    msg.Save()
    result = {"status": "flagged", "flag": text, "entry_id": entry_id}
    output(result, ctx.obj["json"]) if ctx.obj["json"] else click.echo(f"Flagged: {text}")


@cli.command("mark-read")
@click.argument("entry_id")
@click.option("--unread", is_flag=True, help="Mark as unread instead")
@click.pass_context
def mark_read(ctx, entry_id, unread):
    """Mark an email as read (or --unread)."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    msg = ns.GetItemFromID(_validate_entry_id(entry_id))
    msg.UnRead = unread
    msg.Save()
    state = "unread" if unread else "read"
    result = {"status": state, "entry_id": entry_id}
    output(result, ctx.obj["json"]) if ctx.obj["json"] else click.echo(f"Marked as {state}")


@cli.command("categorize")
@click.argument("entry_id")
@click.option("--categories", required=True, help="Comma-separated categories")
@click.pass_context
def categorize(ctx, entry_id, categories):
    """Set categories on an email."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    msg = ns.GetItemFromID(_validate_entry_id(entry_id))
    msg.Categories = categories
    msg.Save()
    result = {"status": "categorized", "categories": categories, "entry_id": entry_id}
    output(result, ctx.obj["json"]) if ctx.obj["json"] else click.echo(f"Categories set: {categories}")


# ── Folders ────────────────────────────────────────────────────────────

@cli.command()
@click.option("--root", default="", help="Start from a specific folder")
@click.pass_context
def folders(ctx, root):
    """List the mail folder tree."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    lines = []

    def walk(folder, prefix=""):
        try:
            lines.append(f"{prefix}{_sanitize(folder.Name)} ({folder.Items.Count})")
            for i in range(folder.Folders.Count):
                walk(folder.Folders.Item(i + 1), prefix + "  ")
        except Exception:
            pass

    if root:
        walk(get_folder(ns, root))
    else:
        for store in ns.Stores:
            try:
                lines.append(f"[{_sanitize(store.DisplayName)}]")
                rf = store.GetRootFolder()
                for i in range(rf.Folders.Count):
                    walk(rf.Folders.Item(i + 1), "  ")
            except Exception:
                continue
    click.echo("\n".join(lines) if lines else "No folders found.")


@cli.command()
@click.option("-f", "--folder", default="Inbox")
def unread(folder):
    """Show unread count for a folder."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    fld = get_folder(ns, folder)
    click.echo(f"{fld.UnReadItemCount} unread in {folder}")


# ── Stats ──────────────────────────────────────────────────────────────

@cli.command()
@click.option("-f", "--folder", default="Inbox")
@click.pass_context
def stats(ctx, folder):
    """Folder statistics: total, unread, top senders."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    fld = get_folder(ns, folder)
    messages = fld.Items
    messages.Sort("[ReceivedTime]", True)
    total = fld.Items.Count
    unread_ct = fld.UnReadItemCount
    senders = {}
    oldest = newest = None
    count = 0
    for msg in messages:
        if count >= 500:
            break
        try:
            name = _sanitize(str(msg.SenderName))
            senders[name] = senders.get(name, 0) + 1
            dt = str(msg.ReceivedTime)
            if newest is None:
                newest = dt
            oldest = dt
            count += 1
        except Exception:
            continue
    top = sorted(senders.items(), key=lambda x: -x[1])[:10]
    info = {
        "folder": folder, "total": total, "unread": unread_ct,
        "sampled": count, "newest": newest, "oldest": oldest,
        "top_senders": dict(top),
    }
    if ctx.obj["json"]:
        output(info, True)
    else:
        click.echo(f"Folder: {folder}")
        click.echo(f"Total: {total}  Unread: {unread_ct}")
        click.echo(f"Newest: {newest}")
        click.echo(f"Oldest: {oldest}")
        click.echo("Top senders:")
        for name, c in top:
            click.echo(f"  {name}: {c}")


# ── Scrape ─────────────────────────────────────────────────────────────

@cli.command()
@click.option("-f", "--folder", default="Inbox")
@click.option("-n", "--count", default=100)
@click.option("--fields", default="subject,from,date",
              help="Comma-separated: subject,from,date,unread,categories,flag,importance,size,id,attachments")
@click.pass_context
def scrape(ctx, folder, count, fields):
    """Bulk export selected fields from a folder."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    fld = get_folder(ns, folder)
    messages = fld.Items
    messages.Sort("[ReceivedTime]", True)
    field_list = [f.strip().lower() for f in fields.split(",")]
    attr_map = {
        "subject": "Subject", "from": "SenderName", "date": "ReceivedTime",
        "unread": "UnRead", "categories": "Categories", "flag": "FlagRequest",
        "importance": "Importance", "size": "Size", "id": "EntryID",
        "attachments": "Attachments",
    }
    rows = []
    for i, msg in enumerate(messages):
        if i >= count:
            break
        try:
            row = {}
            for f in field_list:
                attr = attr_map.get(f)
                if not attr:
                    continue
                val = getattr(msg, attr, "")
                if attr == "ReceivedTime":
                    val = str(val)
                elif attr == "Attachments":
                    val = val.Count if val else 0
                elif attr == "Importance":
                    val = _IMPORTANCE_MAP.get(val, f"Unknown({val})")
                else:
                    val = _sanitize(str(val))
                row[f] = val
            rows.append(row)
        except Exception:
            continue
    if ctx.obj["json"]:
        output(rows, True)
    else:
        for row in rows:
            click.echo(" | ".join(f"{k}={v}" for k, v in row.items()))
        if not rows:
            click.echo("No emails found.")


# ── Calendar ───────────────────────────────────────────────────────────

@cli.command()
@click.option("-d", "--days", default=7, help="Days ahead to show")
@click.pass_context
def cal(ctx, days):
    """Show upcoming calendar events."""
    ol = get_outlook()
    ns = ol.GetNamespace("MAPI")
    calendar = ns.GetDefaultFolder(9)
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    now = datetime.datetime.now()
    end = now + datetime.timedelta(days=days)
    restrict = (f"[Start] >= '{now.strftime('%m/%d/%Y %I:%M %p')}'"
                f" AND [Start] <= '{end.strftime('%m/%d/%Y %I:%M %p')}'")
    filtered = items.Restrict(restrict)
    results = []
    for item in filtered:
        try:
            evt = {
                "subject": _sanitize(item.Subject),
                "start": str(item.Start),
                "end": str(item.End),
                "location": _sanitize(item.Location),
                "body": _sanitize(str(item.Body)[:200]),
            }
            results.append(evt)
        except Exception:
            continue
    if ctx.obj["json"]:
        output(results, True)
    else:
        for e in results:
            click.echo(f"  {e['start'][:16]}  {e['subject']}")
            if e["location"]:
                click.echo(f"    @ {e['location']}")
        if not results:
            click.echo("No upcoming events.")


@cli.command("cal-add")
@click.option("--subject", required=True)
@click.option("--start", required=True, help="YYYY-MM-DD HH:MM")
@click.option("--end", required=True, help="YYYY-MM-DD HH:MM")
@click.option("--location", default="")
@click.option("--body", default="")
@click.pass_context
def cal_add(ctx, subject, start, end, location, body):
    """Create a calendar event."""
    ol = get_outlook()
    item = ol.CreateItem(1)
    item.Subject = subject
    item.Start = start
    item.End = end
    if location:
        item.Location = location
    if body:
        item.Body = body
    item.Save()
    result = {"status": "created", "subject": subject, "start": start, "end": end}
    output(result, ctx.obj["json"]) if ctx.obj["json"] else click.echo(f"Event created: {subject}")


# ── Ghost Mode ─────────────────────────────────────────────────────────

@cli.command()
@click.argument("action", type=click.Choice(["install", "remove", "status"]))
def ghost(action):
    """Manage Outlook ghost mode (hidden startup on login).

    \b
    install  - Register a scheduled task to start Outlook hidden at logon
    remove   - Remove the scheduled task
    status   - Check if ghost mode is active
    """
    task_name = "OlookGhostOutlook"
    exe = _find_outlook_exe()

    if action == "install":
        result = subprocess.run(
            ["schtasks", "/create", "/tn", task_name,
             "/tr", f'"{exe}" /embedding',
             "/sc", "onlogon", "/rl", "limited", "/f"],
            capture_output=True, text=True,
        )
        if result.returncode == 0:
            click.echo("Ghost mode installed. Outlook will start hidden at logon.")
            click.echo(f"  Task: {task_name}")
            click.echo(f"  Exe:  {exe}")
        else:
            raise click.ClickException(f"Failed to create task: {result.stderr.strip()}")

    elif action == "remove":
        result = subprocess.run(
            ["schtasks", "/delete", "/tn", task_name, "/f"],
            capture_output=True, text=True,
        )
        if result.returncode == 0:
            click.echo("Ghost mode removed.")
        else:
            click.echo("Ghost mode was not installed (or already removed).")

    elif action == "status":
        result = subprocess.run(
            ["schtasks", "/query", "/tn", task_name, "/fo", "csv", "/nh"],
            capture_output=True, text=True,
        )
        running = "YES" if _is_outlook_running() else "NO"
        if result.returncode == 0:
            # /tn provides exact match — returncode 0 means task exists
            click.echo("Ghost mode: ACTIVE")
        else:
            click.echo("Ghost mode: NOT INSTALLED")
        click.echo(f"Outlook running: {running}")


# ── Entry point ────────────────────────────────────────────────────────

if __name__ == "__main__":
    cli()
