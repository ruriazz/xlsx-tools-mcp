"""Pre/post checks around a single write: rules, rollback snapshot and ripple.

Lifecycle (all inside the caller's per-file lock):

    guard = WriteGuard(path, ctx, recalc_requested)
    try:
        guard.prepare()      # before the file changes: load rules, evaluate, snapshot
        ... save + recalc ...
        guard.finish(result) # evaluate again, maybe roll back, attach report keys
    except BaseException:
        guard.abort()        # after a failed save/recalc: restore the snapshot if any
        raise
    finally:
        guard.cleanup()      # the snapshot never outlives the call (unless restoring failed)

Ripple is informational: any failure while computing it is reported as skipped and
never fails the write.
"""

from __future__ import annotations

import contextlib
import os
import shutil
import uuid
from typing import Any

from . import recalc, ripple, rules, settings
from .context import WriteContext
from .errors import FileInUseError
from .locking import safe_replace


class WriteGuard:
    def __init__(self, path: str, ctx: WriteContext, recalc_requested: bool):
        self.path = path
        self.ctx = ctx
        self.recalc_requested = recalc_requested
        self.ruleset: rules.Ruleset | None = None
        self.pre: rules.Evaluation | None = None
        self.snapshot_path: str | None = None
        self.snapshot_ready = False  # True only once the copy completed
        self.file_modified = False   # set by the writer right before it saves
        self.keep_snapshot = False
        self.ripple_before: ripple.Snapshot | None = None
        self.ripple_skip: str | None = None

    @property
    def _rules_active(self) -> bool:
        return bool(self.ruleset and self.ruleset.rules and self.ctx.changes_values)

    # ── before the write ──────────────────────────────────────────────────────

    def prepare(self) -> None:
        # A corrupt rules file fails closed here, before anything is written.
        self.ruleset = rules.load_ruleset(self.path)
        if not self.ctx.changes_values:
            return
        if self._rules_active:
            self.pre = rules.evaluate(self.path, self.ruleset)
            if self.ruleset.needs_snapshot:
                # Set before copying so cleanup() removes a partial copy; a partial
                # copy is never restored (snapshot_ready stays False).
                self.snapshot_path = f"{self.path}.snap-{os.getpid()}-{uuid.uuid4().hex[:8]}"
                shutil.copy2(self.path, self.snapshot_path)
                self.snapshot_ready = True
        self._prepare_ripple()

    def _prepare_ripple(self) -> None:
        if not (settings.RIPPLE_ENABLED and self.recalc_requested):
            return
        if self.ctx.structural:
            self.ripple_skip = "structural change shifts cell coordinates"
            return
        if recalc.find_soffice() is None:
            self.ripple_skip = "LibreOffice not available; formulas were not recalculated"
            return
        try:
            before = ripple.snapshot(self.path)
        except Exception as exc:  # informational only — never block the write
            self.ripple_skip = f"could not read formula results: {exc}"
            return
        if before is None:
            self.ripple_skip = f"more than XLSX_MCP_RIPPLE_MAX_CELLS={settings.RIPPLE_MAX_CELLS} formula cells"
        else:
            self.ripple_before = before

    # ── after the write ───────────────────────────────────────────────────────

    def finish(self, result: dict[str, Any]) -> None:
        if self.ruleset and self.ruleset.rules:
            if self.ctx.changes_values:
                self._check_rules(result)
            else:
                result["rules"] = {
                    "status": "skipped",
                    "reason": "operation does not change cell values; run check_rules to re-check",
                }
        self._report_ripple(result)

    def _check_rules(self, result: dict[str, Any]) -> None:
        if self.ruleset is None or self.pre is None:
            return
        post = rules.evaluate(
            self.path,
            self.ruleset,
            baseline=self.pre,
            values_current=bool(result.get("recalculated")),
            structural=self.ctx.structural,
        )
        c = rules.classify(self.pre.results, post.results)
        result["rules"] = {
            "status": rules.status_of(post.results),
            "results": post.results,
            "regressions": c.regressions,
            "pre_existing": c.pre_existing,
        }
        if c.unknown_before:
            result["rules"]["unknown_before"] = c.unknown_before
        blocking = []
        if self.ruleset.on_violation == "rollback":
            blocking += c.regressions
        if self.ruleset.on_unverifiable == "rollback":
            blocking += c.unverifiable + c.unknown_before
        if blocking:
            self._rollback(result, blocking)
            return
        result["rolled_back"] = False
        notes = []
        if c.regressions:
            notes.append(f"violated {c.regressions}")
        if c.unverifiable:
            notes.append(f"unverifiable {c.unverifiable}")
        if c.unknown_before:
            notes.append(
                f"failing {c.unknown_before}, but their state before this write was unknown "
                "(run recalculate_workbook to refresh formula results)"
            )
        if notes:
            result["message"] = f"{result.get('message', '')} Rule check: {'; '.join(notes)}.".strip()

    def _restore(self) -> str | None:
        """Put the snapshot back; returns an error message when that fails."""
        if self.snapshot_path is None or not self.snapshot_ready:
            return "no snapshot was taken"
        try:
            safe_replace(self.snapshot_path, self.path)
        except (FileInUseError, OSError) as exc:
            self.keep_snapshot = True
            return str(exc)
        self.snapshot_path = None  # consumed by the replace
        self.snapshot_ready = False
        return None

    def _rollback(self, result: dict[str, Any], rule_ids: list[str]) -> None:
        snapshot = self.snapshot_path
        error = self._restore()
        if error is not None:
            result.update(rolled_back=False, rollback_error=error)
            if self.keep_snapshot:
                result["snapshot_path"] = snapshot
            result["message"] = (
                f"Rules {rule_ids} were violated and the rollback failed: {error}. "
                f"The pre-write copy is kept at '{snapshot}'; use restore_backup or copy it back manually."
            )
            return
        result.update(
            saved=False,
            rolled_back=True,
            errors_found=[],
            message=f"Write rolled back: rule(s) {rule_ids} failed after the change. The file is unchanged.",
        )

    def abort(self) -> str | None:
        """Called when saving or recalculating raised: with a rules snapshot, undo the write.

        Without this, rules would fail open — the file would stay modified while the
        caller sees an error and no rule was checked. Only a complete snapshot is
        restored, and only once the save may have touched the file. Returns a message
        when the restore itself failed (the snapshot is then kept and named).
        """
        if not (self.snapshot_ready and self.file_modified):
            return None
        snapshot = self.snapshot_path
        error = self._restore()
        if error is None:
            return None
        return f"restoring the pre-write copy failed ({error}); it is kept at '{snapshot}'"

    def _report_ripple(self, result: dict[str, Any]) -> None:
        if not (settings.RIPPLE_ENABLED and self.recalc_requested and self.ctx.changes_values):
            return
        if result.get("rolled_back"):
            result["ripple"] = {"skipped": "write was rolled back"}
        elif self.ripple_skip:
            result["ripple"] = {"skipped": self.ripple_skip}
        elif self.ripple_before is None or not result.get("recalculated"):
            result["ripple"] = {"skipped": "recalculation did not run"}
        else:
            try:
                after = ripple.snapshot(self.path)
            except Exception as exc:  # informational only — the write already succeeded
                result["ripple"] = {"skipped": f"could not read formula results: {exc}"}
                return
            if after is None:
                result["ripple"] = {"skipped": "too many formula cells after the write"}
            else:
                result["ripple"] = ripple.diff(self.ripple_before, after, self.ctx.touched)

    def cleanup(self) -> None:
        if self.snapshot_path and not self.keep_snapshot:
            with contextlib.suppress(OSError):
                os.remove(self.snapshot_path)
