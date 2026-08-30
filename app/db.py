from __future__ import annotations

from contextlib import contextmanager
import os
import select
import socket
import socketserver
import threading
from urllib.parse import quote_plus

import paramiko
import pandas as pd
from sqlalchemy import create_engine, text

from app.config import ModuleConfig, RegionConfig
from app.env import load_dotenv


def execute_module_query(module: ModuleConfig, region: RegionConfig, org_id: int) -> pd.DataFrame:
    load_dotenv()
    with _mysql_endpoint(region) as endpoint:
        engine = create_engine(_mysql_url(region, endpoint.host, endpoint.port))
        with engine.connect() as connection:
            return pd.read_sql_query(text(module.query_template), connection, params={"org_id": org_id})


class _MysqlEndpoint:
    def __init__(self, host: str, port: int) -> None:
        self.host = host
        self.port = port


class _ForwardHandler(socketserver.BaseRequestHandler):
    remote_host: str
    remote_port: int
    transport: paramiko.Transport

    def handle(self) -> None:
        channel = self.transport.open_channel(
            "direct-tcpip",
            (self.remote_host, self.remote_port),
            self.request.getpeername(),
        )
        if channel is None:
            return
        try:
            while True:
                readable, _, _ = select.select([self.request, channel], [], [])
                if self.request in readable:
                    data = self.request.recv(16384)
                    if not data:
                        break
                    channel.sendall(data)
                if channel in readable:
                    data = channel.recv(16384)
                    if not data:
                        break
                    self.request.sendall(data)
        finally:
            channel.close()


class _ForwardServer(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


@contextmanager
def _mysql_endpoint(region: RegionConfig):
    prefix = region.env_prefix
    host = _required_env(f"{prefix}_MYSQL_HOST")
    port = int(_required_env(f"{prefix}_MYSQL_PORT"))

    if not _has_env(f"{prefix}_SSH_HOST"):
        yield _MysqlEndpoint(host, port)
        return

    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    server = None
    thread = None
    try:
        client.connect(
            _required_env(f"{prefix}_SSH_HOST"),
            port=int(os.getenv(f"{prefix}_SSH_PORT", "22")),
            username=_required_env(f"{prefix}_SSH_USER"),
            password=_required_env(f"{prefix}_SSH_PASSWORD"),
            look_for_keys=False,
            allow_agent=False,
            timeout=20,
        )
        transport = client.get_transport()
        if transport is None:
            raise RuntimeError("SSH transport unavailable after authentication.")

        handler = type(
            f"{prefix.title()}ForwardHandler",
            (_ForwardHandler,),
            {"remote_host": host, "remote_port": port, "transport": transport},
        )
        server = _ForwardServer(("127.0.0.1", 0), handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        yield _MysqlEndpoint("127.0.0.1", int(server.server_address[1]))
    finally:
        if server is not None:
            server.shutdown()
            server.server_close()
        if thread is not None:
            thread.join(timeout=2)
        client.close()


def _mysql_url(region: RegionConfig, host: str, port: int) -> str:
    prefix = region.env_prefix
    database = _required_env(f"{prefix}_MYSQL_DATABASE")
    user = quote_plus(_required_env(f"{prefix}_MYSQL_USER"))
    password = quote_plus(_required_env(f"{prefix}_MYSQL_PASSWORD"))
    return (
        f"mysql+pymysql://{user}:{password}"
        f"@{host}:{port}/{database}"
    )


def _required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def _has_env(name: str) -> bool:
    return bool(os.getenv(name))
