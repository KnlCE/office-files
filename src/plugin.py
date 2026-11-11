"""Astra lifecycle adapter; Office operations run in bounded worker processes."""

import json

from astra_plugin_sdk import BadArguments, Plugin, ToolDef

from .registry import TOOLS, normalize
from .runtime import Runner


class OfficeFiles(Plugin):
    def __init__(self):
        super().__init__()
        self.runner = Runner()
        self.settings = {}

    async def on_config_changed(self, config_json):
        config = (
            json.loads(config_json) if isinstance(config_json, str) else config_json
        )
        if not isinstance(config, dict):
            return
        # Older/malformed host settings must not crash the lifecycle hook.
        clean = {}
        for key in ["allowed_roots", "libreoffice_path"]:
            if isinstance(config.get(key), str):
                clean[key] = config[key]
        for key, maximum in [
            ("timeout_secs", 120),
            ("max_file_mb", 100),
            ("max_unpacked_mb", 512),
            ("max_result_kb", 28),
        ]:
            value = config.get(key)
            if (
                type(value) is int
                and (4 if key == "max_result_kb" else 1) <= value <= maximum
            ):
                clean[key] = value
        self.settings = clean

    async def list_tools(self):
        return [
            ToolDef(
                name=name,
                description=t["description"],
                parameters_json=json.dumps(t["schema"]),
            )
            for name, t in TOOLS.items()
        ]

    async def call_tool(self, name, arguments_json):
        try:
            args = json.loads(arguments_json or "{}")
        except (ValueError, TypeError) as exc:
            raise BadArguments("Expected a JSON object") from exc
        args = normalize(name, args)
        result = await self.runner.run(name, args, self.settings)
        return {"success": True, "result": json.dumps(result, ensure_ascii=False)}

    async def on_shutdown(self):
        await self.runner.close()


if __name__ == "__main__":
    OfficeFiles().run()
