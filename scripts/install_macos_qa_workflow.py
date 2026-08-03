#!/usr/bin/env python3
"""Build or install the Finder "QA workflow" quick action on macOS."""

from __future__ import annotations

import argparse
import plistlib
import shlex
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path


EXCEL_UTIS = (
    "org.openxmlformats.spreadsheetml.sheet",
    "org.openxmlformats.spreadsheetml.sheet.macroenabled",
)


@dataclass(frozen=True)
class QuickActionSpec:
    name: str
    menu_item_name: str
    bundle_identifier: str
    launcher_flag: str
    log_path: str
    action_uuid: str
    input_uuid: str
    output_uuid: str

    @property
    def bundle_name(self) -> str:
        return f"{self.name}.workflow"


QA_QUICK_ACTION = QuickActionSpec(
    name="QA workflow",
    menu_item_name="ABC · QA workflow",
    bundle_identifier="com.tagexactor.services.qa-workflow",
    launcher_flag="--qa-workflow",
    log_path="/tmp/tagexactor-qa-workflow.log",
    action_uuid="F37C9B40-3792-4DCB-BFA4-54B563B9DE62",
    input_uuid="E46C53F2-B932-4B76-A6F5-C6C5997454F7",
    output_uuid="9DE2345D-B130-4947-A86F-D708C066A2B6",
)
NBSP_QUICK_ACTION = QuickActionSpec(
    name="NBSP restore",
    menu_item_name="ABC · NBSP restore",
    bundle_identifier="com.tagexactor.services.nbsp-restore",
    launcher_flag="--nbsp-restore",
    log_path="/tmp/tagexactor-nbsp-restore.log",
    action_uuid="0EB6B477-F006-4CFF-A89C-6D6D96514731",
    input_uuid="49F12AC6-2289-482E-BD89-D8AD4500A6D1",
    output_uuid="F0098923-B758-45E7-82B1-48603188E53D",
)
QUICK_ACTION_NAME = QA_QUICK_ACTION.name
QUICK_ACTION_BUNDLE_NAME = QA_QUICK_ACTION.bundle_name


def build_shell_command(
    *,
    python_executable: Path,
    launcher: Path,
    spec: QuickActionSpec = QA_QUICK_ACTION,
) -> str:
    python_arg = shlex.quote(str(python_executable))
    launcher_arg = shlex.quote(str(launcher))
    log_arg = shlex.quote(spec.log_path)
    return f"""for qa_file in "$@"; do
  case "$qa_file" in
    *.xlsx|*.XLSX|*.xlsm|*.XLSM)
      /usr/bin/nohup {python_arg} {launcher_arg} {spec.launcher_flag} "$qa_file" >>{log_arg} 2>&1 </dev/null &
      exit 0
      ;;
  esac
done
exit 1
"""


def build_info_plist(
    spec: QuickActionSpec = QA_QUICK_ACTION,
) -> dict[str, object]:
    return {
        "CFBundleDevelopmentRegion": "zh_CN",
        "CFBundleIdentifier": spec.bundle_identifier,
        "CFBundleName": spec.name,
        "CFBundleShortVersionString": "1.0",
        "NSServices": [
            {
                "NSMenuItem": {"default": spec.menu_item_name},
                "NSMessage": "runWorkflowAsService",
                "NSSendFileTypes": list(EXCEL_UTIS),
            }
        ],
    }


def build_document_plist(
    shell_command: str,
    spec: QuickActionSpec = QA_QUICK_ACTION,
) -> dict[str, object]:
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
                    "InputUUID": spec.input_uuid,
                    "OutputUUID": spec.output_uuid,
                    "UUID": spec.action_uuid,
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
    spec: QuickActionSpec = QA_QUICK_ACTION,
) -> Path:
    bundle_path = destination / spec.bundle_name
    resources_path = bundle_path / "Contents" / "Resources"
    resources_path.mkdir(parents=True, exist_ok=True)

    shell_command = build_shell_command(
        python_executable=python_executable,
        launcher=launcher,
        spec=spec,
    )
    with (bundle_path / "Contents" / "Info.plist").open("wb") as info_file:
        plistlib.dump(build_info_plist(spec), info_file, sort_keys=False)
    with (resources_path / "document.wflow").open("wb") as document_file:
        plistlib.dump(
            build_document_plist(shell_command, spec),
            document_file,
            sort_keys=False,
        )
    return bundle_path


def build_argument_parser() -> argparse.ArgumentParser:
    project_root = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(
        description="生成或安装 QAtools 的 Finder Excel 快速操作。"
    )
    parser.add_argument(
        "--action",
        choices=("qa", "nbsp", "all"),
        default="qa",
        help="安装 QA workflow、NBSP restore，或同时安装两者。",
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

    specs = {
        "qa": (QA_QUICK_ACTION,),
        "nbsp": (NBSP_QUICK_ACTION,),
        "all": (QA_QUICK_ACTION, NBSP_QUICK_ACTION),
    }[args.action]
    bundle_paths = [
        build_quick_action_bundle(
            args.destination.expanduser().absolute(),
            python_executable=python_executable,
            launcher=launcher,
            spec=spec,
        )
        for spec in specs
    ]

    if not args.no_refresh:
        pbs = Path("/System/Library/CoreServices/pbs")
        if pbs.is_file():
            subprocess.run(
                [str(pbs), "-flush"],
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
    for bundle_path in bundle_paths:
        print(bundle_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
