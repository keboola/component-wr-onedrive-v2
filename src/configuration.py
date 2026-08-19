"""Pydantic v2 configuration models for keboola.wr-onedrive-v2.

The component receives a single *merged* ``parameters`` object (root account
parameters + row parameters combined by the platform) — see design spec §5/§6.
``RowConfig`` is the top-level model built from that merged object.
"""

from enum import StrEnum
from typing import Any, Self

from pydantic import BaseModel, ConfigDict, Field, model_validator


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
    """Row output mode."""

    FILE = "file"
    TABLE_CSV = "table_csv"
    TABLE_EXCEL = "table_excel"


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
    """File + CSV mode target: document library / folder / conflict handling."""

    model_config = ConfigDict(extra="ignore")

    drive_id: str | None = None
    folder_path: str | None = None
    conflict_behavior: ConflictBehavior = ConflictBehavior.FAIL


class CsvOptions(BaseModel):
    """CSV file formatting options (mode=table_csv)."""

    model_config = ConfigDict(extra="ignore")

    file_name: str | None = None
    delimiter: str = ","
    enclosure: str = '"'
    include_header: bool = True


class Workbook(BaseModel):
    """Excel workbook target (mode=table_excel) — v1-compatible field names.

    v1 targeting rules: either ``path`` alone, or both ``drive_id`` and
    ``file_id`` together. A lone id or a path combined with an id is invalid.
    ``metadata`` is opaque UI file-picker storage — accepted and ignored.
    """

    model_config = ConfigDict(extra="ignore")

    drive_id: str | None = None
    file_id: str | None = None
    path: str | None = None
    metadata: Any | None = None

    @model_validator(mode="after")
    def _validate_target(self) -> Self:
        has_path = self.path is not None
        has_drive_id = self.drive_id is not None
        has_file_id = self.file_id is not None

        if has_path and (has_drive_id or has_file_id):
            raise ValueError("workbook.path cannot be combined with workbook.drive_id/workbook.file_id.")

        if not has_path:
            if has_drive_id != has_file_id:
                raise ValueError("workbook.drive_id and workbook.file_id must be provided together.")
            if not has_drive_id:
                raise ValueError(
                    "workbook requires either workbook.path, or both "
                    "workbook.drive_id and workbook.file_id."
                )
        return self


class Worksheet(BaseModel):
    """Excel worksheet target (mode=table_excel).

    ``id`` and ``position`` are mutually exclusive; ``name`` may be combined
    with either (or given alone, e.g. for worksheet creation). At least one of
    ``id``/``name``/``position`` must be provided. ``position`` accepts
    numeric strings (v1 configs hold both string and int forms).
    """

    model_config = ConfigDict(extra="ignore")

    id: str | None = None
    name: str | None = None
    position: int | None = None
    metadata: Any | None = None

    @model_validator(mode="before")
    @classmethod
    def _coerce_position(cls, data: Any) -> Any:
        if isinstance(data, dict) and isinstance(data.get("position"), str):
            raw = data["position"]
            try:
                data = {**data, "position": int(raw)}
            except ValueError as exc:
                raise ValueError(f"worksheet.position must be an integer, got '{raw}'.") from exc
        return data

    @model_validator(mode="after")
    def _validate_selector(self) -> Self:
        if self.id is not None and self.position is not None:
            raise ValueError("worksheet.id and worksheet.position are mutually exclusive.")
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

    @model_validator(mode="after")
    def _validate_mode_requirements(self) -> Self:
        if self.mode == Mode.TABLE_EXCEL:
            if self.workbook is None:
                raise ValueError("workbook configuration is required when mode is 'table_excel'.")
            if self.worksheet is None:
                raise ValueError("worksheet configuration is required when mode is 'table_excel'.")
        return self
