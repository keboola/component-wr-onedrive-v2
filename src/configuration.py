"""Pydantic v2 configuration models for keboola.wr-onedrive-v2.

The component receives a single *merged* ``parameters`` object (root account
parameters + row parameters combined by the platform) — see design spec §5/§6.
``RowConfig`` is the top-level model built from that merged object.
"""

from enum import StrEnum
from typing import Any, Self

from pydantic import BaseModel, ConfigDict, Field, model_validator


def _pop_blank_strings(data: Any, fields: tuple[str, ...]) -> Any:
    """Normalize the Keboola UI's untouched-text-field convention: a field the user never typed
    into is submitted as ``""``, not omitted entirely and not ``null``. Popping a blank/
    whitespace-only string among ``fields`` lets the field's own default apply — ``None`` for a
    genuinely optional field, or the field's declared default otherwise (e.g.
    ``Destination.conflict_behavior``) — exactly as if the key had never been sent.

    Without this, a UI-submitted row like ``{"path": "", "drive_id": "b!...", "file_id": "01..."}``
    (path untouched, ids picked via the dropdowns) would fail :class:`Workbook`'s "path cannot be
    combined with drive_id/file_id" mutual-exclusivity check, since an empty string still counts
    as "the user set this field" under a naive ``is not None`` presence check.
    """
    if not isinstance(data, dict):
        return data
    normalized = dict(data)
    for field in fields:
        if isinstance(normalized.get(field), str) and not normalized[field].strip():
            del normalized[field]
    return normalized


class AccountType(StrEnum):
    """OneDrive/SharePoint account kind — drives both UI visibility and API dispatch."""

    PRIVATE_ONEDRIVE = "private_onedrive"
    ONEDRIVE_FOR_BUSINESS = "onedrive_for_business"
    SHAREPOINT = "sharepoint"


class ConflictBehavior(StrEnum):
    """File/CSV upload conflict resolution (Excel mode has its own append/overwrite semantics)."""

    FAIL = "fail"
    REPLACE = "replace"
    RENAME = "rename"


class Mode(StrEnum):
    """Row output mode.

    Only two modes exist going forward — ``file`` (uploads mapped files as-is *and* writes mapped
    tables as CSV, merged from the old ``file``/``table_csv`` split) and ``worksheet`` (writes one
    mapped table into an Excel worksheet, the old ``table_excel``). The pre-merge names are still
    accepted as input (see :data:`_LEGACY_MODE_ALIASES`) — every already-recorded VCR cassette and
    every platform row created before this change uses them — but are normalized to one of these
    two members before any other validation runs, so the rest of the codebase only ever sees
    ``FILE``/``WORKSHEET``.
    """

    FILE = "file"
    WORKSHEET = "worksheet"


# Maps a pre-merge mode value to its current equivalent (`RowConfig._normalize_mode_alias`).
# `"table_csv"` and `"file"` merge into one mode (`FILE` now processes both a file input mapping
# and a table input mapping); `"table_excel"` is simply renamed to `"worksheet"`.
_LEGACY_MODE_ALIASES: dict[str, str] = {
    "table_csv": Mode.FILE.value,
    "table_excel": Mode.WORKSHEET.value,
}


class Account(BaseModel):
    """Root-config account/tenant parameters."""

    model_config = ConfigDict(extra="ignore")

    account_type: AccountType
    tenant_id: str | None = None
    site_url: str | None = None

    @model_validator(mode="after")
    def _validate_requirements(self) -> Self:
        if (
            self.account_type in (AccountType.ONEDRIVE_FOR_BUSINESS, AccountType.SHAREPOINT)
            and not self.tenant_id
        ):
            raise ValueError(
                "account.tenant_id is required when account_type is "
                "'onedrive_for_business' or 'sharepoint'."
            )
        if self.account_type == AccountType.SHAREPOINT and not self.site_url:
            raise ValueError("account.site_url is required when account_type is 'sharepoint'.")
        return self


class Destination(BaseModel):
    """Mode 'file' target: document library / folder / conflict handling."""

    model_config = ConfigDict(extra="ignore")

    drive_id: str | None = None
    folder_path: str | None = None
    # A relative ("yesterday", "3 days ago") or absolute ("2026-01-31") date, resolved via
    # `dateparser` (component.py's `_resolve_now`) and fed into `folder_path`'s
    # `{date:<strftime-format>}` placeholders in place of the job's own start time. Empty/unset
    # preserves today's behavior (job start, UTC).
    date: str | None = None
    conflict_behavior: ConflictBehavior = ConflictBehavior.FAIL

    @model_validator(mode="before")
    @classmethod
    def _normalize_blank_strings(cls, data: Any) -> Any:
        return _pop_blank_strings(data, ("drive_id", "folder_path", "date", "conflict_behavior"))


class CsvOptions(BaseModel):
    """CSV file formatting options for the tables mode 'file' maps (mapped files upload as-is,
    unaffected by these options)."""

    model_config = ConfigDict(extra="ignore")

    file_name: str | None = None
    delimiter: str = ","
    enclosure: str = '"'
    include_header: bool = True


class WorkbookTargeting(StrEnum):
    """How a row's ``workbook`` section targets a workbook (Change B UX addition).

    ``PICK`` (UI default): target via the ``drive_id``/``file_id`` async-select dropdowns;
    ``PATH``: target via a typed/looked-up ``path``. Explicitly choosing one makes the *other*
    field's value irrelevant no matter what it holds — a stale value left over from switching
    ``targeting`` back and forth in the UI can never break validation or get used by mistake.
    """

    PICK = "pick"
    PATH = "path"


class Workbook(BaseModel):
    """Excel workbook target (mode=worksheet) — v1-compatible field names.

    v1 targeting rules: either ``path`` alone, or both ``drive_id`` and ``file_id`` together. A
    lone id or a path combined with an id is invalid. ``metadata`` is opaque UI file-picker
    storage — accepted and ignored.

    ``targeting`` (Change B UX addition) makes that choice explicit instead of inferring it from
    which fields happen to be set: ``"pick"`` uses ``drive_id``/``file_id`` and ignores ``path``
    entirely (cleared below, regardless of any stale value); ``"path"`` uses ``path`` and ignores
    the ids the same way. ``targeting`` is absent for configs created before this change (row API
    payloads, already-recorded VCR cassettes) — those keep today's v1 mutual-exclusivity check,
    just with a friendlier message.
    """

    model_config = ConfigDict(extra="ignore")

    targeting: WorkbookTargeting | None = None
    drive_id: str | None = None
    file_id: str | None = None
    path: str | None = None
    metadata: Any | None = None

    @model_validator(mode="before")
    @classmethod
    def _normalize_blank_strings(cls, data: Any) -> Any:
        return _pop_blank_strings(data, ("path", "drive_id", "file_id", "targeting"))

    @model_validator(mode="after")
    def _validate_target(self) -> Self:
        if self.targeting == WorkbookTargeting.PICK:
            if not self.drive_id or not self.file_id:
                raise ValueError(
                    'Targeting is set to "Pick via dropdowns": both workbook.drive_id and '
                    "workbook.file_id are required."
                )
            self.path = None  # Ignore any stale/hidden Path value — pick mode never reads it.
            return self

        if self.targeting == WorkbookTargeting.PATH:
            if not self.path:
                raise ValueError('Targeting is set to "By path": workbook.path is required.')
            self.drive_id = None  # Ignore any stale/hidden Library/Workbook picker values.
            self.file_id = None
            return self

        # Legacy (no `targeting`): v1's own mutual-exclusivity rule, humanized (Change B).
        has_path = self.path is not None
        has_drive_id = self.drive_id is not None
        has_file_id = self.file_id is not None

        if has_path and (has_drive_id or has_file_id):
            raise ValueError(
                "Choose one way to target the workbook: either Library + Workbook (drive_id + "
                "file_id), or Path — please clear the other. Got both."
            )

        if not has_path:
            if has_drive_id != has_file_id:
                raise ValueError("workbook.drive_id and workbook.file_id must be provided together.")
            if not has_drive_id:
                raise ValueError(
                    "workbook requires either workbook.path, or both "
                    "workbook.drive_id and workbook.file_id."
                )
        return self


class WorksheetSelection(StrEnum):
    """How a row's ``worksheet`` section picks its target sheet (Change C UX addition).

    ``PICK`` (UI default): target the sheet by ``id`` (the ``listWorksheets`` async select) with
    no rename; ``NAME``: target/create the sheet by ``name`` alone. ``position`` has been dropped
    from the UI entirely — it's kept on the model only so v1-parity/API-created rows that still
    set it (without ``selection``) keep working. Explicitly choosing one makes the *other*
    field(s) irrelevant no matter what they hold — switching ``selection`` back and forth in the
    UI can leave a stale hidden ``id``/``name``/``position`` that must never get used by mistake.
    """

    PICK = "pick"
    NAME = "name"


class Worksheet(BaseModel):
    """Excel worksheet target (mode=worksheet).

    v1/legacy targeting (no ``selection``): ``id`` and ``position`` are mutually exclusive;
    ``name`` may be combined with either (renaming that sheet) or given alone (select-or-create by
    name). At least one of ``id``/``name``/``position`` must be provided. ``position`` accepts
    numeric strings (v1 configs hold both string and int forms).

    ``selection`` (Change C UX addition) replaces that inference with an explicit choice: ``pick``
    uses ``id`` and ignores ``name``/``position`` (cleared below — no rename, matching the "Pick
    existing" UI label); ``name`` uses ``name`` alone (creating the sheet if missing) and ignores
    ``id``/``position`` the same way. ``selection`` is absent for configs created before this
    change (row API payloads, already-recorded VCR cassettes, and any row still holding a
    v1-parity ``position``) — those keep today's rename-on-combine behavior, just with a
    friendlier mutual-exclusivity message.
    """

    model_config = ConfigDict(extra="ignore")

    selection: WorksheetSelection | None = None
    id: str | None = None
    name: str | None = None
    position: int | None = None
    metadata: Any | None = None

    @model_validator(mode="before")
    @classmethod
    def _normalize_inputs(cls, data: Any) -> Any:
        """Pop blank ``id``/``name``/``position``/``selection`` strings (UI untouched-field
        convention), then coerce a still-present numeric-string ``position`` to ``int`` (v1
        configs hold both string and int forms). Blank-popping must run first: an empty
        ``position`` string must become ``None``, not be handed to ``int("")`` and crash.
        """
        data = _pop_blank_strings(data, ("id", "name", "position", "selection"))
        if isinstance(data, dict) and isinstance(data.get("position"), str):
            raw = data["position"]
            try:
                data = {**data, "position": int(raw)}
            except ValueError as exc:
                raise ValueError(f"worksheet.position must be an integer, got '{raw}'.") from exc
        return data

    @model_validator(mode="after")
    def _validate_selector(self) -> Self:
        if self.selection == WorksheetSelection.PICK:
            if not self.id:
                raise ValueError('Selection is set to "Pick existing": worksheet.id is required.')
            self.name = None  # Ignore any stale/hidden Name value — pick mode never renames.
            self.position = None
            return self

        if self.selection == WorksheetSelection.NAME:
            if not self.name:
                raise ValueError(
                    'Selection is set to "By name (creates if missing)": worksheet.name is required.'
                )
            self.id = None  # Ignore any stale/hidden ID/Position values.
            self.position = None
            return self

        # Legacy (no `selection`): v1's own mutual-exclusivity rule, humanized (Change C).
        if self.id is not None and self.position is not None:
            raise ValueError(
                "Choose one way to target the worksheet: either ID, or Position — please clear "
                "the other. Got both."
            )
        if self.id is None and self.position is None and self.name is None:
            raise ValueError("worksheet requires at least one of id, name, or position.")
        return self


class RowConfig(BaseModel):
    """Top-level model for the merged (root + row) ``parameters`` object."""

    model_config = ConfigDict(extra="ignore")

    mode: Mode
    account: Account
    destination: Destination = Field(default_factory=Destination)
    csv: CsvOptions = Field(default_factory=CsvOptions)
    workbook: Workbook | None = None
    worksheet: Worksheet | None = None
    append: bool = False
    # `gt=0`: 0 would silently write nothing (batched in chunks of zero rows) and a negative
    # value produces an unmapped `ValueError` deep inside the Excel writer's batching helper —
    # neither is a sensible configuration, so both are rejected here as a normal validation error
    # (phase 8 audit IMPORTANT-4).
    batch_size: int = Field(default=5000, gt=0)

    @model_validator(mode="before")
    @classmethod
    def _normalize_mode_alias(cls, data: Any) -> Any:
        """Silently map a pre-merge ``mode`` value (``table_csv``/``table_excel``) to its current
        equivalent (:data:`_LEGACY_MODE_ALIASES`) before ``Mode`` itself ever sees it.

        Every already-recorded VCR cassette config and every row created on the platform before
        this change uses the old names — this keeps them working unchanged, with no re-recording
        and no forced re-save, rather than requiring every existing config to be touched.
        """
        if isinstance(data, dict) and isinstance(data.get("mode"), str) and data["mode"] in _LEGACY_MODE_ALIASES:
            data = {**data, "mode": _LEGACY_MODE_ALIASES[data["mode"]]}
        return data

    @model_validator(mode="after")
    def _validate_mode_requirements(self) -> Self:
        if self.mode == Mode.WORKSHEET:
            if self.workbook is None:
                raise ValueError("workbook configuration is required when mode is 'worksheet'.")
            if self.worksheet is None:
                raise ValueError("worksheet configuration is required when mode is 'worksheet'.")
        return self
