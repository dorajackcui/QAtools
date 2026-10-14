"""Store the original file without exposing its worksheet ranges to importers."""
from __future__ import annotations

import base64
import binascii
import hashlib
from io import BytesIO
from pathlib import Path

from tools.excel_rows import iter_stored_rows, stored_row

from . import workbook_schema as schema
from .errors import WorkbookFormatError


# Base64 expands each block to at most 30,000 characters, below Excel's 32,767.
CHUNK_BYTES = 22_500
CHUNK_CHARACTERS = 30_000
MAX_CHUNKS = 1_048_575  # One worksheet row is reserved for the header.


def write_original_file(workbook, input_path: Path) -> dict[str, object]:
    worksheet = workbook.create_sheet(schema.ORIGINAL_FILE_SHEET)
    worksheet.sheet_state = "hidden"
    worksheet.append(schema.ORIGINAL_FILE_COLUMNS)
    digest = hashlib.sha256()
    size = count = 0
    with input_path.open("rb") as source:
        while block := source.read(CHUNK_BYTES):
            count += 1
            if count > MAX_CHUNKS:
                raise WorkbookFormatError("Original file exceeds the backup worksheet capacity.")
            worksheet.append([count, base64.b64encode(block).decode("ascii")])
            digest.update(block)
            size += len(block)
    return {
        schema.ORIGINAL_FILE_SIZE_KEY: size,
        schema.ORIGINAL_FILE_SHA256_KEY: digest.hexdigest(),
        schema.ORIGINAL_FILE_CHUNKS_KEY: count,
    }


def read_original_file(workbook, metadata: dict[str, object]) -> BytesIO:
    """Return a verified file stream. The caller owns and must close it."""
    size = _positive_integer(metadata.get(schema.ORIGINAL_FILE_SIZE_KEY),
                             CHUNK_BYTES * MAX_CHUNKS, "file size")
    count = _positive_integer(metadata.get(schema.ORIGINAL_FILE_CHUNKS_KEY),
                              MAX_CHUNKS, "chunk count")
    checksum = metadata.get(schema.ORIGINAL_FILE_SHA256_KEY)
    if (
        count != (size + CHUNK_BYTES - 1) // CHUNK_BYTES
        or not isinstance(checksum, str) or len(checksum) != 64
        or any(char not in "0123456789abcdef" for char in checksum)
    ):
        raise _invalid_backup("invalid size, chunk count or checksum metadata")
    if schema.ORIGINAL_FILE_SHEET not in workbook.sheetnames:
        raise _invalid_backup("missing backup sheet")
    worksheet = workbook[schema.ORIGINAL_FILE_SHEET]
    headers = stored_row(worksheet)
    if [headers[col].value if col in headers else None for col in (1, 2)] != schema.ORIGINAL_FILE_COLUMNS:
        raise _invalid_backup("invalid backup columns")

    chunks: dict[int, str] = {}
    for _, cells in iter_stored_rows(worksheet, min_row=2, max_col=2):
        index = cells[1].value if 1 in cells else None
        data = cells[2].value if 2 in cells else None
        if index is None and data is None:
            continue
        index = _positive_integer(index, count, "chunk index")
        if (
            index in chunks
            or not isinstance(data, str) or not 0 < len(data) <= CHUNK_CHARACTERS
        ):
            raise _invalid_backup("invalid or duplicate chunk")
        chunks[index] = data
    if len(chunks) != count:
        raise _invalid_backup("missing chunks")

    stream = BytesIO()
    digest = hashlib.sha256()
    try:
        for index in range(1, count + 1):
            try:
                block = base64.b64decode(chunks[index], validate=True)
            except (ValueError, binascii.Error) as exc:
                raise _invalid_backup("invalid base64 data") from exc
            expected_size = CHUNK_BYTES if index < count else size - CHUNK_BYTES * (count - 1)
            if len(block) != expected_size:
                raise _invalid_backup("truncated or oversized chunk")
            stream.write(block)
            digest.update(block)
        if digest.hexdigest() != checksum:
            raise _invalid_backup("checksum mismatch")
        stream.seek(0)
        return stream
    except Exception:
        stream.close()
        raise


def _positive_integer(value: object, maximum: int, field: str) -> int:
    # Numeric cells can use decimal or exponent notation without changing value.
    if type(value) not in (int, float) or not 0 < value <= maximum or value != int(value):
        raise _invalid_backup(f"invalid {field}")
    return int(value)


def _invalid_backup(detail: str) -> WorkbookFormatError:
    return WorkbookFormatError(f"Strings workbook has a damaged original file backup: {detail}.")
