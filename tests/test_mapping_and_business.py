from datetime import date
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pandas as pd

from app.core.business import analyze_eso
from app.core.field_mapper import map_fields, missing_required_fields
from app.core.schemas import ESO_ARCHIVE_FIELDS, ESO_MAIN_FIELDS


def test_eso_mapping_and_unfinished_count():
    primary = pd.DataFrame(
        [
            {"PN": "1001", "ESO Plan Date": "2026-05-10", "ESO Actual Date": "", "Function Group": "Body", "Engineer": "Ann", "Operation": "U"},
            {"PN": "1002", "ESO Plan Date": "2026-05-10", "ESO Actual Date": "", "Function Group": "Body", "Engineer": "Bob", "Operation": "U"},
            {"PN": "1003", "ESO Plan Date": "2026-05-10", "ESO Actual Date": "", "Function Group": "EE", "Engineer": "Cara", "Operation": "D"},
        ]
    )
    archive = pd.DataFrame(
        [
            {"Part Number": "1001", "Archive Date": "2026-05-11"},
        ]
    )

    primary_mapping = map_fields(primary, ESO_MAIN_FIELDS, "eso", "primary")
    archive_mapping = map_fields(archive, ESO_ARCHIVE_FIELDS, "eso", "archive")

    assert missing_required_fields(primary_mapping) == []
    assert missing_required_fields(archive_mapping) == []

    result = analyze_eso(primary, primary_mapping, archive, archive_mapping, date(2026, 5, 15))

    assert result["summary"]["未完成数量"] == 1
    assert result["summary"]["本次按归档清单可回填数量"] == 1
    assert result["summary"]["D行排除未完成数量"] == 1
    assert result["rows"][0]["零件号"] == "1002"


if __name__ == "__main__":
    test_eso_mapping_and_unfinished_count()
    print("ok")
