"""开发启动入口：python -m app

主机、端口、是否热重载见 backend/config/server.json。
"""

from __future__ import annotations

import uvicorn

from app.core.config import BACKEND_ROOT
from app.core.constants import server_config


def main() -> None:
    config = server_config()
    reload = bool(config["reload"])
    reload_dirs = [
        str(BACKEND_ROOT / directory)
        for directory in config.get("reload_dirs", ())
        if (BACKEND_ROOT / directory).is_dir()
    ]
    uvicorn.run(
        "app.main:app",
        host=str(config["host"]),
        port=int(config["port"]),
        reload=reload,
        reload_dirs=reload_dirs if reload else None,
    )


if __name__ == "__main__":
    main()
