from __future__ import annotations

import argparse
import datetime as dt
import os
import pathlib
import subprocess
import sys


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")


PROJECT_ROOT = pathlib.Path(__file__).resolve().parent.parent
SYNC_SCRIPT = PROJECT_ROOT / "agent" / "run_sync_and_notify.ps1"
EVENT_SCRIPT = PROJECT_ROOT / "agent" / "record_finance_event.ps1"
WINDOWS_POWERSHELL = pathlib.Path(r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe")


def normalize_receipt_status(value: str) -> str:
    text = value.strip()
    if text in {"已到账", "当天到账", "当日到账", "已经到账", "已收款", "当天收款"}:
        return "已到账"
    if text in {"未到账", "尚未到账", "待到账", "未收款", "赊账"}:
        return "未到账"
    raise argparse.ArgumentTypeError(f"不支持的到账状态：{value}；请使用已到账或未到账")


def normalize_platform(value: str) -> str:
    text = value.strip()
    aliases = {
        "美团外卖": "美团",
        "京东外卖": "京东到家",
        "京东秒送": "京东到家",
        "饿了么": "淘宝闪购",
    }
    return aliases.get(text, text)


def run_utf8_powershell(script_path: pathlib.Path, arguments: list[tuple[str, object | None, bool]]) -> int:
    if not script_path.is_file():
        print(f"PowerShell entry point was not found: {script_path}", file=sys.stderr)
        return 2

    command = [
        str(WINDOWS_POWERSHELL),
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script_path),
    ]
    for name, value, is_switch in arguments:
        if value is None or value is False:
            continue
        command.append(f"-{name}")
        if not is_switch:
            command.append(str(value))

    completed = subprocess.run(
        command,
        cwd=str(PROJECT_ROOT),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        check=False,
    )
    if completed.stdout.strip():
        print(completed.stdout.strip())
    if completed.stderr.strip():
        print(completed.stderr.strip(), file=sys.stderr)
    return completed.returncode


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="宠速星经营助手的 UTF-8 安全命令入口")
    subparsers = parser.add_subparsers(dest="command", required=True)

    sync = subparsers.add_parser("sync", help="同步或回读订单、库存和驾驶舱")
    sync.add_argument("--date", default=dt.date.today().isoformat())
    sync.add_argument("--mode", choices=["midday", "reminder", "night", "manual"], default="manual")
    sync.add_argument("--skip-sync", action="store_true")
    sync.add_argument("--skip-orders", action="store_true")
    sync.add_argument("--skip-inventory", action="store_true")
    sync.add_argument("--silent", action="store_true")
    sync.add_argument("--dry-run", action="store_true")
    sync.add_argument("--force-resend", action="store_true")
    sync.add_argument("--full-rebuild", action="store_true", help="强制执行全量成本、简账和完整性流程")

    event = subparsers.add_parser("event", help="录入一条结构化经营事件")
    event.add_argument(
        "--action",
        required=True,
        choices=[
            "AddIntervention",
            "RemoveIntervention",
            "RestoreInventory",
            "ConfirmSelfDelivery",
            "SetExternalDeliveryFee",
            "SetOrderDeliveryFee",
            "SetOrderActuals",
            "OfflineSale",
            "AddOrderGift",
            "Replenishment",
            "CashExpense",
            "OperatingExpense",
            "OwnerWithdrawal",
            "OwnerLabor",
            "MarketRent",
            "BeginningCash",
            "ActualReceipt",
            "AssetAmortization",
            "SetSkuExpiry",
            "SetSkuBatchDate",
            "SetSkuCost",
            "ChangeSkuCode",
            "SetRefundDisposition",
        ],
    )
    event.add_argument("--date", default=dt.date.today().isoformat())
    event.add_argument("--order-no")
    event.add_argument("--daily-order-no")
    event.add_argument("--person")
    event.add_argument("--items")
    event.add_argument("--goods-code")
    event.add_argument("--transition-code")
    event.add_argument("--new-goods-code")
    event.add_argument(
        "--refund-disposition",
        choices=["returned_restocked", "refund_no_return", "returned_damaged"],
    )
    event.add_argument("--retire-existing-new-code", action="store_true")
    event.add_argument("--expiry-date")
    event.add_argument("--clear-expiry", action="store_true")
    event.add_argument("--production-date")
    event.add_argument("--shelf-life-value", type=float)
    event.add_argument("--shelf-life-unit", choices=["天", "月", "年"], default="月")
    event.add_argument("--quantity", type=float)
    event.add_argument("--purchase-amount", type=float)
    event.add_argument("--payment-amount", type=float)
    event.add_argument("--payment-date")
    event.add_argument("--amount", type=float)
    event.add_argument("--delivery-fee", type=float)
    event.add_argument("--receipt-status", type=normalize_receipt_status, choices=["已到账", "未到账"], default="已到账")
    event.add_argument("--platform", type=normalize_platform, choices=["美团", "淘宝闪购", "京东到家", "线下"], default="美团")
    event.add_argument("--note")
    event.add_argument("--event-id")
    event.add_argument("--dry-run", action="store_true")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.command == "sync":
        return run_utf8_powershell(
            SYNC_SCRIPT,
            [
                ("Date", args.date, False),
                ("Mode", args.mode, False),
                ("SkipSync", args.skip_sync, True),
                ("SkipOrders", args.skip_orders, True),
                ("SkipInventory", args.skip_inventory, True),
                ("Silent", args.silent, True),
                ("DryRun", args.dry_run, True),
                ("ForceResend", args.force_resend, True),
                ("FullRebuild", args.full_rebuild, True),
                ("ProjectRoot", PROJECT_ROOT, False),
            ],
        )

    if os.environ.get("PETSTAR_CONFIRMATION_PHASE") == "preview":
        print('{"ok":false,"error":"待确认预览阶段禁止执行经营数据写入"}')
        return 2

    return run_utf8_powershell(
        EVENT_SCRIPT,
        [
            ("Action", args.action, False),
            ("Date", args.date, False),
            ("OrderNo", args.order_no, False),
            ("DailyOrderNo", args.daily_order_no, False),
            ("Person", args.person, False),
            ("Items", args.items, False),
            ("GoodsCode", args.goods_code, False),
            ("TransitionCode", args.transition_code, False),
            ("NewGoodsCode", args.new_goods_code, False),
            ("RefundDisposition", args.refund_disposition, False),
            ("RetireExistingNewCode", args.retire_existing_new_code, True),
            ("ExpiryDate", args.expiry_date, False),
            ("ClearExpiry", args.clear_expiry, True),
            ("ProductionDate", args.production_date, False),
            ("ShelfLifeValue", args.shelf_life_value, False),
            ("ShelfLifeUnit", args.shelf_life_unit, False),
            ("Quantity", args.quantity, False),
            ("PurchaseAmount", args.purchase_amount, False),
            ("PaymentAmount", args.payment_amount, False),
            ("PaymentDate", args.payment_date, False),
            ("Amount", args.amount, False),
            ("DeliveryFee", args.delivery_fee, False),
            ("ReceiptStatus", args.receipt_status, False),
            ("Platform", args.platform, False),
            ("Note", args.note, False),
            ("EventId", args.event_id, False),
            ("DryRun", args.dry_run, True),
            ("ProjectRoot", PROJECT_ROOT, False),
        ],
    )


if __name__ == "__main__":
    raise SystemExit(main())
