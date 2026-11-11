"""Process isolation + parent-only transactional commit. Workers cannot publish."""

import asyncio
import contextlib
import ctypes
import json
import os
import shutil
import signal
import sys
import tempfile
import uuid
from pathlib import Path

from astra_plugin_sdk import (
    BadArguments,
    InternalError,
    NotFound,
    Timeout,
    Unauthorized,
    Unavailable,
)

from .safety import check_size, checked_path, revision

ERRORS = {
    c.__name__: c
    for c in [BadArguments, NotFound, Timeout, Unauthorized, Unavailable, InternalError]
}


class ProcessTree:
    """Windows kill-on-close job; POSIX process group. Includes LibreOffice."""

    def __init__(self, pid):
        self.pid, self.job = pid, None
        if os.name == "nt":
            from ctypes import wintypes as w

            class Basic(ctypes.Structure):
                _fields_ = [
                    ("PerProcessUserTimeLimit", ctypes.c_int64),
                    ("PerJobUserTimeLimit", ctypes.c_int64),
                    ("LimitFlags", w.DWORD),
                    ("MinimumWorkingSetSize", ctypes.c_size_t),
                    ("MaximumWorkingSetSize", ctypes.c_size_t),
                    ("ActiveProcessLimit", w.DWORD),
                    ("Affinity", ctypes.c_size_t),
                    ("PriorityClass", w.DWORD),
                    ("SchedulingClass", w.DWORD),
                ]

            class IO(ctypes.Structure):
                _fields_ = [
                    (k, ctypes.c_uint64)
                    for k in [
                        "ReadOperationCount",
                        "WriteOperationCount",
                        "OtherOperationCount",
                        "ReadTransferCount",
                        "WriteTransferCount",
                        "OtherTransferCount",
                    ]
                ]

            class Extended(ctypes.Structure):
                _fields_ = [
                    ("BasicLimitInformation", Basic),
                    ("IoInfo", IO),
                    ("ProcessMemoryLimit", ctypes.c_size_t),
                    ("JobMemoryLimit", ctypes.c_size_t),
                    ("PeakProcessMemoryUsed", ctypes.c_size_t),
                    ("PeakJobMemoryUsed", ctypes.c_size_t),
                ]

            k = ctypes.WinDLL("kernel32", use_last_error=True)
            k.CreateJobObjectW.argtypes, k.CreateJobObjectW.restype = (
                [w.LPVOID, w.LPCWSTR],
                w.HANDLE,
            )
            k.OpenProcess.argtypes, k.OpenProcess.restype = (
                [w.DWORD, w.BOOL, w.DWORD],
                w.HANDLE,
            )
            k.SetInformationJobObject.argtypes = [
                w.HANDLE,
                ctypes.c_int,
                w.LPVOID,
                w.DWORD,
            ]
            k.AssignProcessToJobObject.argtypes = [w.HANDLE, w.HANDLE]
            k.CloseHandle.argtypes = [w.HANDLE]
            job = k.CreateJobObjectW(None, None)
            info = Extended()
            info.BasicLimitInformation.LimitFlags = (
                0x2000  # JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            )
            proc = k.OpenProcess(0x0100 | 0x0001, False, pid)
            try:
                if (
                    not job
                    or not proc
                    or not k.SetInformationJobObject(
                        job, 9, ctypes.byref(info), ctypes.sizeof(info)
                    )
                    or not k.AssignProcessToJobObject(job, proc)
                ):
                    if job:
                        k.CloseHandle(job)
                    raise OSError(
                        ctypes.get_last_error(), "Cannot isolate worker process tree"
                    )
                self.job, self.kernel = job, k
            finally:
                if proc:
                    k.CloseHandle(proc)

    def close(self):
        if self.job:
            self.kernel.CloseHandle(self.job)
            self.job = None
        elif os.name != "nt":
            with contextlib.suppress(ProcessLookupError):
                os.killpg(self.pid, signal.SIGKILL)


class Runner:
    def __init__(self):
        self.lock = asyncio.Lock()
        self.processes = set()
        self.closed = False

    async def close(self):
        self.closed = True
        for proc, tree in list(self.processes):
            tree.close()
            with contextlib.suppress(ProcessLookupError):
                proc.kill()
        await asyncio.gather(
            *(p.wait() for p, _ in list(self.processes)), return_exceptions=True
        )

    async def run(self, name, args, settings):
        async with self.lock:
            if self.closed:
                raise Unavailable("Plugin is shutting down")
            return await self._run(name, args, settings or {})

    async def _run(self, name, args, settings):
        from .registry import help_result

        if name == "office_help":
            return help_result(args)
        settings = dict(settings)
        snapshots, output, target_hash = {}, None, None
        # Resolve all input resources before giving the worker its private copies.
        with tempfile.TemporaryDirectory(prefix="office-files-") as tmp:
            job = Path(tmp)

            def snapshot(value):
                p = checked_path(value, settings, exists=True)
                check_size(p, settings)
                if str(p) not in snapshots:
                    before = revision(p)
                    dest = job / (str(len(snapshots)) + p.suffix.lower())
                    shutil.copyfile(p, dest)
                    if revision(dest) != before or revision(p) != before:
                        raise BadArguments("Input changed during snapshot: " + str(p))
                    snapshots[str(p)] = (before, str(dest))
                return snapshots[str(p)][1]

            def resolve(value, schema):
                if "oneOf" in schema:
                    from jsonschema import Draft202012Validator

                    schema = next(
                        s
                        for s in schema["oneOf"]
                        if Draft202012Validator(s).is_valid(value)
                    )
                if isinstance(value, dict):
                    properties = schema.get("properties", {})
                    return {
                        k: snapshot(v)
                        if k in properties
                        and k in {"input_path", "template_path", "path"}
                        else resolve(v, properties[k])
                        if k in properties
                        else v
                        for k, v in value.items()
                    }
                if isinstance(value, list):
                    return [resolve(v, schema.get("items", {})) for v in value]
                return value

            from .registry import TOOLS

            resolved = resolve(args, TOOLS[name]["schema"])
            if (
                "revision" in args
                and snapshots[str(checked_path(args["input_path"], settings))][0]
                != args["revision"]
            ):
                raise BadArguments("Stale revision; inspect the source again")
            output_key = "output_dir" if name == "office_render" else "output_path"
            if output_key in args:
                output = checked_path(args[output_key], settings)
                if output.exists():
                    if (
                        name == "office_render"
                        or not output.is_file()
                        or not args.get("overwrite", False)
                    ):
                        raise BadArguments(
                            "Output already exists; choose another path or explicitly overwrite"
                        )
                    target_hash = revision(output)
                resolved[output_key] = str(
                    job
                    / (
                        "result"
                        if name == "office_render"
                        else "result" + output.suffix.lower()
                    )
                )
            request = {"name": name, "args": resolved, "settings": settings}
            options = (
                {"creationflags": 0x08000000}
                if os.name == "nt"
                else {"start_new_session": True}
            )
            proc = await asyncio.create_subprocess_exec(
                sys.executable,
                "-m",
                "src.worker",
                str(job),
                stdin=asyncio.subprocess.PIPE,
                stdout=asyncio.subprocess.DEVNULL,
                stderr=asyncio.subprocess.DEVNULL,
                cwd=str(Path(__file__).resolve().parent.parent),
                **options,
            )
            tree = None
            try:
                tree = ProcessTree(proc.pid)
                self.processes.add((proc, tree))
                await asyncio.wait_for(
                    proc.communicate(json.dumps(request).encode("utf-8")),
                    timeout=min(120, max(1, int(settings.get("timeout_secs", 120)))),
                )
                result_file = job / "response.json"
                if not result_file.is_file():
                    raise InternalError(
                        "Worker exited without a result; inputs and outputs unchanged"
                    )
                result = json.loads(result_file.read_text(encoding="utf-8"))
                if "error" in result:
                    raise ERRORS.get(result["error"]["type"], InternalError)(
                        result["error"]["message"]
                    )
                if not output:
                    return result
                if self.closed:
                    raise Unavailable("Plugin shut down before commit")
                for raw, (digest, _) in snapshots.items():
                    if (
                        checked_path(raw, settings, True) != Path(raw)
                        or revision(raw) != digest
                    ):
                        raise BadArguments("Input changed before save: " + raw)
                if checked_path(str(output), settings) != output:
                    raise Unauthorized("Output directory changed before commit")
                if output.exists() and (
                    target_hash is None or revision(output) != target_hash
                ):
                    raise BadArguments("Output changed before save")
                if target_hash and not output.exists():
                    raise BadArguments("Output disappeared before save")
                # Same-filesystem staging is required for an atomic commit.
                output.parent.mkdir(parents=True, exist_ok=True)
                if checked_path(str(output), settings) != output:
                    raise Unauthorized("Output path escaped before commit")
                stage = output.parent / (
                    ".office-files-" + uuid.uuid4().hex + output.suffix
                )
                backup = None
                try:
                    private = Path(resolved[output_key])
                    if name == "office_render":
                        shutil.copytree(private, stage)
                        if output.exists():
                            raise BadArguments("Preview directory appeared before save")
                        os.rename(stage, output)
                    else:
                        shutil.copyfile(private, stage)
                        if target_hash is not None:
                            if revision(output) != target_hash:
                                raise BadArguments("Output changed before replace")
                            backup = output.with_name(
                                output.name + ".backup-" + uuid.uuid4().hex
                            )
                            with (
                                backup.open("xb") as stream,
                                output.open("rb") as source,
                            ):
                                shutil.copyfileobj(source, stream)
                            if (
                                revision(backup) != target_hash
                                or revision(output) != target_hash
                            ):
                                raise BadArguments(
                                    "Output changed while creating backup"
                                )
                            os.replace(stage, output)
                        else:
                            # link is atomic AND fails if a file appears; replace would clobber it.
                            os.link(stage, output)
                            stage.unlink()
                finally:
                    if stage.is_dir():
                        shutil.rmtree(stage)
                    elif stage.exists():
                        stage.unlink()
                old = str(private)

                def remap(value):
                    if isinstance(value, str) and (
                        value == old or value.startswith(old + os.sep)
                    ):
                        return str(output) + value[len(old) :]
                    if isinstance(value, dict):
                        return {k: remap(v) for k, v in value.items()}
                    if isinstance(value, list):
                        return [remap(v) for v in value]
                    return value

                result = remap(result)
                if backup:
                    result["backup_path"] = str(backup)
                return result
            except asyncio.TimeoutError as exc:
                raise Timeout(
                    "Operation exceeded timeout_secs; no result was committed"
                ) from exc
            finally:
                if tree:
                    tree.close()
                    self.processes.discard((proc, tree))
                if proc.returncode is None:
                    with contextlib.suppress(ProcessLookupError):
                        proc.kill()
                await proc.wait()
