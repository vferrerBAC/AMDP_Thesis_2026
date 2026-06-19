from __future__ import annotations

import pandas as pd


def default_load_combinations(design_method: str = "LRFD") -> pd.DataFrame:
    method = str(design_method).upper()
    if method == "ASD":
        rows = [
            ("ASD-01", "Dead only", "D", "Baseline gravity case"),
            ("ASD-02", "Wind +X", "D + 0.6Wx+", "Wind in positive X direction"),
            ("ASD-03", "Wind -X", "D + 0.6Wx-", "Wind in negative X direction"),
            ("ASD-04", "Wind +Y", "D + 0.6Wy+", "Wind in positive Y direction"),
            ("ASD-05", "Wind -Y", "D + 0.6Wy-", "Wind in negative Y direction"),
            ("ASD-06", "Seismic +X", "D + 0.7Ex+", "Seismic in positive X direction"),
            ("ASD-07", "Seismic -X", "D + 0.7Ex-", "Seismic in negative X direction"),
        ]
    else:
        rows = [
            ("LRFD-01", "Dead only", "1.4D", "Baseline strength case"),
            ("LRFD-02", "Wind +X", "1.2D + 1.0Wx+", "Wind in positive X direction"),
            ("LRFD-03", "Wind -X", "1.2D + 1.0Wx-", "Wind in negative X direction"),
            ("LRFD-04", "Wind +Y", "1.2D + 1.0Wy+", "Wind in positive Y direction"),
            ("LRFD-05", "Wind -Y", "1.2D + 1.0Wy-", "Wind in negative Y direction"),
            ("LRFD-06", "Seismic +X", "1.2D + 1.0Ex+", "Seismic in positive X direction"),
            ("LRFD-07", "Seismic -X", "1.2D + 1.0Ex-", "Seismic in negative X direction"),
            ("LRFD-08", "Uplift wind +X", "0.9D + 1.0Wx+", "Stabilizing dead load reduced"),
            ("LRFD-09", "Uplift seismic +X", "0.9D + 1.0Ex+", "Stabilizing dead load reduced"),
        ]
    return pd.DataFrame(rows, columns=["combo_id", "combo_name", "expression", "plain_language_notes"])


def load_combination_template() -> pd.DataFrame:
    return default_load_combinations("LRFD")
