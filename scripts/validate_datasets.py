import argparse
import json
import sys

from app.services import datasets_validator


def _print_text(results):
    for r in results:
        print(f"[{r.status.upper()}] {r.name}: {r.records} records")
        for issue in r.issues:
            print(f"  - {issue}")


def _print_json(results):
    data = [r.to_dict() for r in results]
    print(json.dumps(data, ensure_ascii=False, indent=2))


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="Проверяет целостность JSON/CSV-справочников приложения."
    )
    parser.add_argument(
        "--format",
        choices=["text", "json"],
        default="text",
        help="Формат вывода отчёта",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Строгий режим: предупреждения считаются ошибками (код возврата 2).",
    )

    args = parser.parse_args(argv)

    results = datasets_validator.run_all_validations()
    summary = datasets_validator.summarize_status(results)

    if args.format == "json":
        _print_json(results)
    else:
        _print_text(results)

    if summary == "error":
        return 2
    if summary == "warning":
        return 2 if args.strict else 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())


