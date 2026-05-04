import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / "soil_lab_tool"))

from core.excel_output import build_kte_gw_btex_simple_from_xml  # noqa: E402


def main() -> None:
    input_path = Path(
        r"C:\Users\Asaf\claude\laboratory_results_analsys\Laboratory_results\KTE\groundwater\PR2605239_0_EXCEL_GENERIC.XLS"
    )
    output_path = Path(
        r"C:\Users\Asaf\claude\laboratory_results_analsys\Laboratory_results\KTE\groundwater\PR2605239_0_BTEX_simple.xlsx"
    )

    out = build_kte_gw_btex_simple_from_xml(input_path, output_path)
    print("נוצר קובץ:", out)


if __name__ == "__main__":
    main()

