import io
import os
import tarfile
import time
from pathlib import Path

import docker
from docker.types import Mount
from termcolor import cprint


def _print_container_tail(container, lines: int = 200) -> None:
    try:
        logs = container.logs(tail=lines)
        try:
            text = logs.decode(errors="ignore")
        except Exception:
            text = str(logs)
        cprint("---- container logs (tail) ----", "cyan")
        cprint(text, "cyan")
        cprint("---- end logs ----", "cyan")
    except Exception:
        pass


def _wait_healthy_or_port(container, timeout: int, port: int = 2002) -> bool:
    """
    Prefer Docker HEALTH status; fallback to /dev/tcp probe inside container.
    """
    start = time.time()
    while time.time() - start < timeout:
        # Try health
        try:
            container.reload()
            health = container.attrs.get("State", {}).get("Health", {}).get("Status")
            if health == "healthy":
                return True
        except Exception:
            pass
        # Fallback: probe inside the container
        res = container.exec_run(
            ["bash", "-lc", f"exec 3<>/dev/tcp/127.0.0.1/{port}"], demux=False
        )
        if res.exit_code == 0:
            return True
        time.sleep(1)
    return False


def _bytes_tar_from_file(host_path: str, arcname: str) -> bytes:
    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode="w") as tar:
        tar.add(host_path, arcname=arcname)
    buf.seek(0)
    return buf.getvalue()


def _extract_single_from_tar_to_host(tar_bytes: bytes, host_out_path: str):
    buf = io.BytesIO(tar_bytes)
    with tarfile.open(fileobj=buf, mode="r:*") as tar:
        for member in tar.getmembers():
            if member.isfile():
                with tar.extractfile(member) as src, open(host_out_path, "wb") as dst:
                    dst.write(src.read())
                break


def _make_mounts(host_in_dir: str, host_out_dir: str):
    """
    - If IN and OUT are the same host folder, bind it ONCE to /work (single=True).
    - If different, bind to /work/in and /work/out (single=False).
    Returns (mounts_list, single_boolean).
    """
    host_in_dir = str(Path(host_in_dir).resolve())
    host_out_dir = str(Path(host_out_dir).resolve())
    if host_in_dir == host_out_dir:
        # One mount to /work
        mounts = [
            Mount(target="/work", source=host_in_dir, type="bind", read_only=False)
        ]
        return mounts, True
    else:
        mounts = [
            Mount(target="/work/in", source=host_in_dir, type="bind", read_only=False),
            Mount(
                target="/work/out", source=host_out_dir, type="bind", read_only=False
            ),
        ]
        return mounts, False


def start_lo_container(
    *,
    docker_image: str = "lo-headless:latest",
    container_name: str = "lo_headless",
    uno_port: int = 2002,
    host_in_dir: str,
    host_out_dir: str,
    recreate: bool = False,
) -> None:
    """
    Start headless LibreOffice container by name, WITHOUT waiting for readiness.
    """
    client = docker.from_env()

    # Optionally recreate
    try:
        existing = client.containers.get(container_name)
        if recreate:
            cprint(f"Stopping existing container '{container_name}'...", "cyan")
            existing.stop()
            try:
                existing.remove()
            except Exception:
                pass
        else:
            if existing.status != "running":
                cprint(f"Starting existing container '{container_name}'...", "cyan")
                existing.start()
            cprint(f"Container '{container_name}' is already present.", "cyan")
            return
    except docker.errors.NotFound:
        pass

    mounts, single = _make_mounts(host_in_dir, host_out_dir)
    cprint(
        f"Starting new container '{container_name}' from image '{docker_image}'...",
        "cyan",
    )
    container = client.containers.run(
        docker_image,
        name=container_name,
        detach=True,
        ports={f"{uno_port}/tcp": uno_port},
        mounts=mounts,
    )

    container.reload()
    m = container.attrs.get("Mounts", [])
    cprint("==== Container mounts ====", "cyan")
    for x in m:
        cprint(
            f"Source: {x.get('Source')} -> Dest: {x.get('Destination')} (Type:{x.get('Type')}, RW:{x.get('RW')})",
            "cyan",
        )
    cprint("==========================", "cyan")

    cprint("Container started.", "cyan")


def stop_lo_container(
    container_name: str = "lo_headless", remove: bool = False
) -> None:
    client = docker.from_env()
    try:
        ctr = client.containers.get(container_name)
    except docker.errors.NotFound:
        cprint(f"Container '{container_name}' not found.", "cyan")
        return
    try:
        cprint(f"Stopping container '{container_name}'...", "cyan")
        ctr.stop()
        if remove:
            cprint(f"Removing container '{container_name}'...", "cyan")
            ctr.remove()
    except Exception as e:
        cprint(f"Warning: error stopping/removing container: {e}", "red")


def process_docx_via_libreoffice(
    docx_path: str,
    output_folder: str,
    docker_image: str = "lo-headless:latest",
    container_name: str = "lo_headless",
    uno_port: int = 2002,
    start_timeout: int = 20,
) -> bool:
    """
    Process a DOCX using a named container.
    - If container missing or not running: start it here with proper mounts and WAIT.
    - If running: verify listener readiness.
    - In-container script:
        * updates fields (currently turned off), saves DOCX over itself (currently turned off), exports PDF.
    - Optional: convert PDF to images (host-side).
    """
    cprint(f"Processing via LibreOffice: {docx_path}", "cyan")

    Path(output_folder).mkdir(parents=True, exist_ok=True)

    client = docker.from_env()
    host_in_dir = str(Path(docx_path).resolve().parent)
    host_out_dir = str(Path(output_folder).resolve())
    base = os.path.splitext(os.path.basename(docx_path))[0]
    host_pdf_path = os.path.join(host_out_dir, f"{base}.pdf")

    # Get or start container with correct mounts
    try:
        container = client.containers.get(container_name)
        if container.status != "running":
            container.start()
            # NOTE: if mounts are wrong (e.g., pre-existing), we recreate it
            container.reload()
    except docker.errors.NotFound:
        container = None

    # If missing or mounts don't match, (re)create with proper mounts
    recreate = False
    if container is None:
        recreate = True
    else:
        container.reload()
        dests = {m.get("Destination") for m in container.attrs.get("Mounts", [])}
        # We need either /work OR both /work/in and /work/out
        if "/work" not in dests and not ("/work/in" in dests and "/work/out" in dests):
            recreate = True

    if recreate:
        if container is not None:
            try:
                container.remove(force=True)
            except Exception:
                pass
        mounts, single = _make_mounts(host_in_dir, host_out_dir)
        cprint(
            f"Starting new container '{container_name}' from image '{docker_image}'...",
            "cyan",
        )
        container = client.containers.run(
            docker_image,
            name=container_name,
            detach=True,
            ports={f"{uno_port}/tcp": uno_port},
            mounts=mounts,
        )
        # Print mounts
        container.reload()
        m = container.attrs.get("Mounts", [])
        cprint("==== Container mounts ====", "cyan")
        for x in m:
            cprint(
                f"Source: {x.get('Source')} -> Dest: {x.get('Destination')} (Type:{x.get('Type')}, RW:{x.get('RW')})",
                "cyan",
            )
        cprint("==========================", "cyan")
    else:
        # Determine if we’re single-mount or double
        dests = {m.get("Destination") for m in container.attrs.get("Mounts", [])}
        single = "/work" in dests

    cprint("Waiting for LibreOffice listener (health/port)...", "cyan")
    if not _wait_healthy_or_port(container, timeout=start_timeout, port=uno_port):
        cprint("Error: LibreOffice listener did not become ready in time.", "red")
        _print_container_tail(container)
        return False

    # In-container file paths
    if single:
        ctr_in_file = f"/work/{os.path.basename(docx_path)}"
        ctr_out_file = f"/work/{base}.pdf"
    else:
        ctr_in_file = f"/work/in/{os.path.basename(docx_path)}"
        ctr_out_file = f"/work/out/{base}.pdf"

    # Pre-flight: verify input is visible (and non-zero) INSIDE container
    stat_res = container.exec_run(
        [
            "bash",
            "-lc",
            f"test -f {ctr_in_file} && stat -c '%s' {ctr_in_file} || echo MISSING",
        ],
        demux=True,
    )
    stat_out = (stat_res.output[0] or b"").decode(errors="ignore").strip()
    if "MISSING" in stat_out or stat_out == "" or stat_out == "0":
        cprint(
            f"Error: {ctr_in_file} is missing or empty inside the container. stat: {stat_out}",
            "red",
        )
        cprint("Falling back to docker copy API (put_archive/get_archive).", "cyan")

        # Fallback tmp paths (no bind required)
        ctr_tmp_in_dir = "/tmp/in"
        ctr_tmp_out_dir = "/tmp/out"
        ctr_tmp_in_file = f"{ctr_tmp_in_dir}/{os.path.basename(docx_path)}"
        ctr_tmp_out_file = f"{ctr_tmp_out_dir}/{base}.pdf"

        container.exec_run(
            [
                "bash",
                "-lc",
                f"mkdir -p {ctr_tmp_in_dir} {ctr_tmp_out_dir} && rm -f {ctr_tmp_out_file}",
            ],
            demux=False,
        )
        # Upload DOCX
        tar_bytes = _bytes_tar_from_file(docx_path, arcname=os.path.basename(docx_path))
        container.put_archive(ctr_tmp_in_dir, tar_bytes)

        # Run UNO script on tmp paths
        exec_res = container.exec_run(
            [
                "/usr/bin/python3",
                "/app/update_and_export.py",
                ctr_tmp_in_file,
                ctr_tmp_out_file,
            ],
            demux=True,
        )
        if exec_res.output:
            stdout, stderr = exec_res.output
            if stdout:
                try:
                    cprint(stdout.decode(errors="ignore"), "cyan")
                except Exception:
                    cprint(str(stdout), "cyan")
            if stderr:
                try:
                    cprint(stderr.decode(errors="ignore"), "red")
                except Exception:
                    cprint(str(stderr), "red")
        if exec_res.exit_code != 0:
            cprint(
                f"Error: LibreOffice export failed with exit code {exec_res.exit_code}",
                "red",
            )
            _print_container_tail(container)
            return False

        # Download the PDF back to host_out_dir
        stream, _ = container.get_archive(ctr_tmp_out_file)
        pdf_tar = b"".join(chunk for chunk in stream)
        _extract_single_from_tar_to_host(pdf_tar, host_pdf_path)

    else:
        # Normal path (bind mounts OK)
        cprint("Mounts look good; proceeding with bound paths.", "cyan")
        exec_res = container.exec_run(
            [
                "/usr/bin/python3",
                "/app/update_and_export.py",
                ctr_in_file,
                ctr_out_file,
            ],
            demux=True,
        )
        if exec_res.output:
            stdout, stderr = exec_res.output
            if stdout:
                try:
                    cprint(stdout.decode(errors="ignore"), "cyan")
                except Exception:
                    cprint(str(stdout), "cyan")
            if stderr:
                try:
                    cprint(stderr.decode(errors="ignore"), "red")
                except Exception:
                    cprint(str(stderr), "red")
        if exec_res.exit_code != 0:
            cprint(
                f"Error: LibreOffice export failed with exit code {exec_res.exit_code}",
                "red",
            )
            _print_container_tail(container)
            return False

        if not os.path.exists(host_pdf_path):
            cprint(f"Error: PDF not found at expected path: {host_pdf_path}", "red")
            _print_container_tail(container)
            return False

    cprint(f"Converted DOCX to PDF: {host_pdf_path}", "cyan")

    return True
