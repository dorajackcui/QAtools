#!/usr/bin/env python3
"""Build or install the Finder "QA workflow" quick action on macOS."""

from __future__ import annotations

import argparse
import plistlib
import shlex
import subprocess
import sys
from pathlib import Path


QUICK_ACTION_NAME = "QA workflow"
QUICK_ACTION_BUNDLE_NAME = f"{QUICK_ACTION_NAME}.workflow"
EXCEL_UTIS = (
    "org.openxmlformats.spreadsheetml.sheet",
    "org.openxmlformats.spreadsheetml.sheet.macroenabled",
)


def build_shell_command(*, python_executable: Path, launcher: Path) -> str:
    python_arg = shlex.quote(str(python_executable))
    launcher_arg = shlex.quote(str(launcher))
    log_arg = shlex.quote("/tmp/tagexactor-qa-workflow.log")
    return f"""for qa_file in "$@"; do
  case "$qa_file" in
    *.xlsx|*.XLSX|*.xlsm|*.XLSM)
      /usr/bin/nohup {python_arg} {launcher_arg} --qa-workflow "$qa_file" >>{log_arg} 2>&1 </dev/null &
      exit 0
      ;;
  esac
done
exit 1
"""


def build_info_plist() -> dict[str, object]:
    return {
        "CFBundleDevelopmentRegion": "zh_CN",
        "CFBundleIdentifier": "com.tagexactor.services.qa-workflow",
        "CFBundleName": QUICK_ACTION_NAME,
        "CFBundleShortVersionString": "1.0",
        "NSServices": [
            {
                "NSMenuItem": {"default": QUICK_ACTION_NAME},
                "NSMessage": "runWorkflowAsService",
                "NSSendFileTypes": list(EXCEL_UTIS),
            }
        ],
    }


def build_document_plist(shell_command: str) -> dict[str, object]:
    action_uuid = "F37C9B40-3792-4DCB-BFA4-54B563B9DE62"
    input_uuid = "E46C53F2-B932-4B76-A6F5-C6C5997454F7"
    output_uuid = "9DE2345D-B130-4947-A86F-D708C066A2B6"
    return {
        "AMApplicationBuild": "523",
        "AMApplicationVersion": "2.10",
        "AMDocumentVersion": "2",
        "actions": [
            {
                "action": {
                    "ActionBundlePath": (
                        "/System/Library/Automator/Run Shell Script.action"
                    ),
                    "ActionName": "Run Shell Script",
                    "ActionParameters": {
                        "CheckedForUserDefaultShell": True,
                        "COMMAND_STRING": shell_command,
                        "inputMethod": 1,
                        "shell": "/bin/zsh",
                        "source": "",
                    },
                    "AMAccepts": {
                        "Container": "List",
                        "Optional": False,
                        "Types": ["com.apple.cocoa.path"],
                    },
                    "AMActionVersion": "2.0.3",
                    "AMApplication": ["Automator"],
                    "AMParameterProperties": {
                        "CheckedForUserDefaultShell": {},
                        "COMMAND_STRING": {},
                        "inputMethod": {},
                        "shell": {},
                        "source": {},
                    },
                    "AMProvides": {
                        "Container": "List",
                        "Types": ["com.apple.cocoa.string"],
                    },
                    "BundleIdentifier": "com.apple.RunShellScript",
                    "CanShowSelectedItemsWhenRun": False,
                    "CanShowWhenRun": True,
                    "Category": ["AMCategoryUtilities"],
                    "CFBundleVersion": "2.0.3",
                    "Class Name": "RunShellScriptAction",
                    "InputUUID": input_uuid,
                    "OutputUUID": output_uuid,
                    "UUID": action_uuid,
                },
                "isViewVisible": True,
            }
        ],
        "connectors": {},
        "workflowMetaData": {
            "serviceApplicationBundleID": "com.apple.finder",
            "serviceApplicationPath": "/System/Library/CoreServices/Finder.app",
            "serviceInputTypeIdentifier": "com.apple.Automator.fileSystemObject",
            "serviceOutputTypeIdentifier": "com.apple.Automator.nothing",
            "serviceProcessesInput": 0,
            "workflowTypeIdentifier": "com.apple.Automator.servicesMenu",
        },
    }


def build_quick_action_bundle(
    destination: Path,
    *,
    python_executable: Path,
    launcher: Path,
) -> Path:
    bundle_path = destination / QUICK_ACTION_BUNDLE_NAME
    resources_path = bundle_path / "Contents" / "Resources"
    resources_path.mkdir(parents=True, exist_ok=True)

    shell_command = build_shell_command(
        python_executable=python_executable,
        launcher=launcher,
    )
    with (bundle_path / "Contents" / "Info.plist").open("wb") as info_file:
        plistlib.dump(build_info_plist(), info_file, sort_keys=False)
    with (resources_path / "document.wflow").open("wb") as document_file:
        plistlib.dump(
            build_document_plist(shell_command),
            document_file,
            sort_keys=False,
        )
    return bundle_path


def build_argument_parser() -> argparse.ArgumentParser:
    project_root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(
        description='生成或安装 Finder 右键快速操作“QA workflow”。'
    )
    parser.add_argument(
        "--destination",
        type=Path,
        default=Path.home() / "Library" / "Services",
        help="快速操作输出目录；默认直接安装到当前用户。",
    )
    parser.add_argument(
        "--python",
        dest="python_executable",
        type=Path,
        default=Path(sys.executable),
        help="启动 Toolshub 使用的 Python。",
    )
    parser.add_argument(
        "--launcher",
        type=Path,
        default=project_root / "toolshub_gui.py",
        help="toolshub_gui.py 的路径。",
    )
    parser.add_argument(
        "--no-refresh",
        action="store_true",
        help="写入后不刷新 macOS 服务缓存。",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_argument_parser().parse_args(argv)
    launcher = args.launcher.expanduser().absolute()
    python_executable = args.python_executable.expanduser().absolute()
    if sys.platform != "darwin":
        print("该安装器仅支持 macOS。", file=sys.stderr)
        return 2
    if not launcher.is_file():
        print(f"找不到 Toolshub 入口：{launcher}", file=sys.stderr)
        return 2
    if not python_executable.is_file():
        print(f"找不到 Python：{python_executable}", file=sys.stderr)
        return 2

    bundle_path = build_quick_action_bundle(
        args.destination.expanduser().absolute(),
        python_executable=python_executable,
        launcher=launcher,
    )

    if not args.no_refresh:
        pbs = Path("/System/Library/CoreServices/pbs")
        if pbs.is_file():
            subprocess.run(
                [str(pbs), "-flush"],
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
    print(bundle_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
