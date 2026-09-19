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
from sqlalchemy import bindparam, create_engine, text

from app.config import ModuleConfig, RegionConfig
from app.env import load_dotenv


def execute_module_query(module: ModuleConfig, region: RegionConfig, org_id: int) -> pd.DataFrame:
    load_dotenv()
    with _mysql_endpoint(region) as endpoint:
        engine = create_engine(_mysql_url(region, endpoint.host, endpoint.port))
        with engine.connect() as connection:
            return pd.read_sql_query(text(module.query_template), connection, params={"org_id": org_id})


def execute_custom_field_queries(
    region: RegionConfig,
    org_id: int,
    field_ids: list[int],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if not field_ids:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    metadata_query = text(
        """
        SELECT id, entity_type, field_type, field_name
        FROM myprompts_custom_fields
        WHERE orgid = :org_id AND id IN :field_ids
        """
    ).bindparams(bindparam("field_ids", expanding=True))
    values_query = text(
        """
        SELECT field_id, entity_id, value
        FROM myprompts_custom_fields_values
        WHERE field_id IN :field_ids
        """
    ).bindparams(bindparam("field_ids", expanding=True))
    options_query = text(
        """
        SELECT id, field_id, text
        FROM myprompts_custom_select_options
        WHERE orgid = :org_id AND field_id IN :field_ids
        """
    ).bindparams(bindparam("field_ids", expanding=True))

    load_dotenv()
    with _mysql_endpoint(region) as endpoint:
        engine = create_engine(_mysql_url(region, endpoint.host, endpoint.port))
        with engine.connect() as connection:
            params = {"org_id": org_id, "field_ids": field_ids}
            metadata = pd.read_sql_query(metadata_query, connection, params=params)
            values = pd.read_sql_query(values_query, connection, params=params)
            options = pd.read_sql_query(options_query, connection, params=params)
            return metadata, values, options


def execute_history_queries(
    region: RegionConfig,
    org_id: int,
    history_ids: list[int],
) -> tuple[pd.DataFrame, pd.DataFrame]:
    if not history_ids:
        return pd.DataFrame(), pd.DataFrame()

    definitions_query = text(
        """
        SELECT id, orgid, history, names, sortby
        FROM preferences_histories
        WHERE orgid = :org_id AND history IN :history_ids
        """
    ).bindparams(bindparam("history_ids", expanding=True))
    values_query = text(
        """
        SELECT id, orgid, matterid, hid, date, country, code
        FROM myprompts_matterhistories
        WHERE orgid = :org_id AND hid IN :history_ids
        """
    ).bindparams(bindparam("history_ids", expanding=True))

    load_dotenv()
    with _mysql_endpoint(region) as endpoint:
        engine = create_engine(_mysql_url(region, endpoint.host, endpoint.port))
        with engine.connect() as connection:
            params = {"org_id": org_id, "history_ids": history_ids}
            definitions = pd.read_sql_query(definitions_query, connection, params=params)
            values = pd.read_sql_query(values_query, connection, params=params)
            return definitions, values


def execute_party_queries(
    region: RegionConfig,
    org_id: int,
    matter_ids: list[int],
) -> pd.DataFrame:
    if not matter_ids:
        return pd.DataFrame()
    query = text(
        """
                SELECT mc.matterid, mc.cid, mc.category, cs.id AS contact_id, cs.name
                FROM contacts_special cs
                LEFT JOIN myprompts_contacts mc
                    ON mc.orgid = cs.orgid
                 AND mc.cid = cs.id
                 AND mc.matterid IN :matter_ids
                 AND LOWER(mc.category) IN ('applicant', 'inventor')
                WHERE cs.orgid = :org_id
        """
    ).bindparams(bindparam("matter_ids", expanding=True))
    load_dotenv()
    with _mysql_endpoint(region) as endpoint:
        engine = create_engine(_mysql_url(region, endpoint.host, endpoint.port))
        with engine.connect() as connection:
            return pd.read_sql_query(
                query,
                connection,
                params={"org_id": org_id, "matter_ids": matter_ids},
            )


def execute_associate_queries(
    region: RegionConfig,
    org_id: int,
    matter_ids: list[int],
) -> pd.DataFrame:
    if not matter_ids:
        return pd.DataFrame()
    query = text(
        """
          SELECT mc.matterid, mc.sid AS cid, mc.category,
               cs.id AS contact_id, cs.sitename AS name
        FROM myprompts_contacts mc
        INNER JOIN contacts_sites cs
            ON cs.orgid = mc.orgid
              AND cs.id = mc.sid
        WHERE mc.orgid = :org_id
          AND mc.matterid IN :matter_ids
          AND LOWER(mc.category) = 'associate'
        """
    ).bindparams(bindparam("matter_ids", expanding=True))
    load_dotenv()
    with _mysql_endpoint(region) as endpoint:
        engine = create_engine(_mysql_url(region, endpoint.host, endpoint.port))
        with engine.connect() as connection:
            return pd.read_sql_query(
                query,
                connection,
                params={"org_id": org_id, "matter_ids": matter_ids},
            )


def execute_organisation_queries(
    region: RegionConfig,
    org_id: int,
) -> pd.DataFrame:
    query = text(
        """
        SELECT id, organisation
        FROM contacts_organisations
        WHERE orgid = :org_id
        """
    )
    load_dotenv()
    with _mysql_endpoint(region) as endpoint:
        engine = create_engine(_mysql_url(region, endpoint.host, endpoint.port))
        with engine.connect() as connection:
            return pd.read_sql_query(query, connection, params={"org_id": org_id})


def execute_site_queries(
    region: RegionConfig,
    org_id: int,
) -> pd.DataFrame:
    query = text(
        """
        SELECT id, sitename
        FROM contacts_sites
        WHERE orgid = :org_id
        """
    )
    load_dotenv()
    with _mysql_endpoint(region) as endpoint:
        engine = create_engine(_mysql_url(region, endpoint.host, endpoint.port))
        with engine.connect() as connection:
            return pd.read_sql_query(query, connection, params={"org_id": org_id})


def execute_update_lookup_queries(
    region: RegionConfig,
    org_id: int,
) -> dict[str, pd.DataFrame]:
    queries = {
        "matters": "SELECT id, mattercode FROM myprompts_matters WHERE orgid = :org_id",
        "users": "SELECT id, username FROM users WHERE orgid = :org_id",
        "families": "SELECT id, family FROM myprompts_family WHERE orgid = :org_id",
        "dropdowns": """
            SELECT oldid, orgid, type, description
            FROM preferences_dropdown
            WHERE orgid = :org_id
              AND (type LIKE '%MATTER_STATUS%' OR type LIKE '%MATTER_CATEGORIES%' OR type LIKE '%TYPEOFMARK%')
        """,
    }
    load_dotenv()
    with _mysql_endpoint(region) as endpoint:
        engine = create_engine(_mysql_url(region, endpoint.host, endpoint.port))
        with engine.connect() as connection:
            return {
                name: pd.read_sql_query(text(query), connection, params={"org_id": org_id})
                for name, query in queries.items()
            }


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
