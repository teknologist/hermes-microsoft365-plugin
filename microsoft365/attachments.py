"""Profile-scoped attachment artifacts; remote filenames never determine local paths."""
from __future__ import annotations

import hashlib
import os
import stat
import uuid
from contextlib import ExitStack
from pathlib import Path

from .errors import GraphError

#: Fixed root under the active Hermes home: no arbitrary output paths are accepted.
ARTIFACT_SUBPATH = ("downloads", "microsoft365")
#: A closed MIME -> suffix map for inspection tools. The suffix derives from the
#: authenticated content type only; a remote filename's extension is never trusted as
#: proof of format, and an unknown type gets the neutral ``.bin`` suffix.
SAFE_MIME_SUFFIXES = {
    "application/pdf": ".pdf",
    "application/zip": ".zip",
    "application/gzip": ".gz",
    "application/xml": ".xml",
    "text/xml": ".xml",
    "application/json": ".json",
    "text/csv": ".csv",
    "text/plain": ".txt",
    "text/html": ".html",
    "image/png": ".png",
    "image/jpeg": ".jpeg",
    "image/gif": ".gif",
    "image/webp": ".webp",
    "image/tiff": ".tiff",
    "application/msword": ".doc",
    "application/vnd.ms-excel": ".xls",
    "application/vnd.ms-powerpoint": ".ppt",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx",
}
DEFAULT_SUFFIX = ".bin"


def safe_mime_suffix(content_type: object) -> str:
    """A safe, closed MIME-derived file suffix; anything unknown stays ``.bin``."""
    if not isinstance(content_type, str):
        return DEFAULT_SUFFIX
    mime = content_type.split(";", 1)[0].strip().lower()
    return SAFE_MIME_SUFFIXES.get(mime, DEFAULT_SUFFIX)


def _refuse_unsafe_directory(descriptor: int) -> None:
    """Fail closed when an artifact-path directory is group- or world-writable.

    The artifact must live on a path only we can write: a group- or world-writable
    directory lets another principal plant or rename entries (a symlink, a swapped
    partial file), which the per-component ``O_NOFOLLOW`` alone does not stop. Directories
    this function creates are always ``0700`` (``fchmod`` after each create), so this only
    ever fires on a pre-existing -- or concurrently modified -- directory in the chain.
    """
    if os.fstat(descriptor).st_mode & (stat.S_IWGRP | stat.S_IWOTH):
        raise GraphError(
            "configuration_error",
            "attachment artifact path is group/world-writable and cannot be trusted",
        )


def save_attachment(data: bytes, content_type: str | None = None) -> dict[str, str | int]:
    """Create a private, unique artifact through no-follow directory descriptors.

    Fail closed on platforms without secure dir_fd/O_NOFOLLOW support. The caller
    has already bounded and verified the bytes. No tenant-provided path is used.

    Safety properties, each pinned by a test:

    * the name is a generated UUID plus a safe MIME-derived suffix -- the remote
      filename never appears in the path;
    * files are ``0600`` and created directories ``0700`` regardless of the process
      umask (an explicit ``fchmod`` after each creation); when the umask is so
      restrictive that a freshly created directory cannot even be opened (e.g. umask
      ``0777``), the writer fails closed -- dropping the partial directory it created --
      rather than claiming every umask is safe;
    * creation is exclusive (``O_EXCL``), so an existing file -- including a planted
      symlink -- is never followed or overwritten;
    * a pre-existing (or concurrently modified) directory in the path that is group-
      or world-writable is refused, so the artifact never lands on a path another
      principal can plant entries in;
    * every failure after the exclusive create (fchmod, fdopen, write, hash) closes
      the file descriptor and removes the partial file again; a distinct email's
      provenance is never deleted;
    * any local failure of the save (platform gaps included) is reported as the typed
      ``configuration_error`` -- never as a raw exception, never as a partial file.
    """
    from hermes_constants import get_hermes_home

    try:
        root = Path(get_hermes_home()).resolve(strict=True)
        name = uuid.uuid4().hex + safe_mime_suffix(content_type)
        with ExitStack() as stack:
            flags = os.O_RDONLY | os.O_DIRECTORY | os.O_NOFOLLOW
            directory = os.open(root, flags)
            stack.callback(os.close, directory)
            _refuse_unsafe_directory(directory)
            for component in ARTIFACT_SUBPATH:
                created = False
                try:
                    os.mkdir(component, mode=0o700, dir_fd=directory)
                    created = True
                except FileExistsError:
                    pass
                try:
                    opened = os.open(component, flags, dir_fd=directory)
                except OSError:
                    # A directory we just created cannot be opened: the process umask
                    # stripped even the owner's read/execute bits (e.g. umask 0777), so its
                    # mode cannot be verified or pinned. Fail closed instead of writing into
                    # a directory we cannot control, and drop the partial directory we
                    # created so no uncontrolled path is left behind.
                    if created:
                        try:
                            os.rmdir(component, dir_fd=directory)
                        except OSError:
                            pass
                    raise
                directory = opened
                stack.callback(os.close, directory)
                if created:
                    os.fchmod(directory, 0o700)
                _refuse_unsafe_directory(directory)
            descriptor = os.open(
                name, os.O_WRONLY | os.O_CREAT | os.O_EXCL | os.O_NOFOLLOW, 0o600, dir_fd=directory
            )
            # Explicit ownership transfer: until ``os.fdopen`` succeeds, the raw
            # descriptor is still ours to close. Both fchmod and fdopen can fail before
            # the file object takes ownership, and each failure must close the descriptor
            # *and* remove the file the exclusive open just created -- no leak, no partial.
            try:
                os.fchmod(descriptor, 0o600)
                output = os.fdopen(descriptor, "wb")
            except BaseException:
                os.close(descriptor)
                try:
                    os.unlink(name, dir_fd=directory)
                except OSError:
                    pass
                raise
            try:
                with output:
                    output.write(data)
                # The hash must be computed inside the guard: any failure between the
                # exclusive create and a fully verified artifact removes the file.
                digest = hashlib.sha256(data).hexdigest()
            except BaseException:
                # ``with`` (or the hash) failed: the file object already closed the
                # descriptor, so only the partial file is removed.
                try:
                    os.unlink(name, dir_fd=directory)
                finally:
                    raise
            return {
                "path": str(root / Path(*ARTIFACT_SUBPATH) / name),
                "sha256": digest,
                "size": len(data),
            }
    except Exception as exc:  # fail closed: any local save failure is a typed error
        raise GraphError("configuration_error", "Attachment artifact could not be saved safely") from exc
