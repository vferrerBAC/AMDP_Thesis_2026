# BAC Joint Check Assistant — Streamlit Prototype v2

This version is designed for design engineers with limited structural/civil background. It uses a guided workflow, Simple/Advanced modes, connection templates, validation checks, traffic-light results, suggested fixes, and professional exports.

## Run locally

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

Or on Windows, double-click `run_local.bat`.

## Main workflow

1. Project setup
2. Upload or load sample CAD-derived joint geometry
3. Select the environment
4. Assign connection templates
5. Upload or edit joint loads
6. Validate inputs
7. Run joint screening checks
8. Export Excel, PDF, CSV templates, and JSON config

## Engineering limitation

The current joint capacities are **demo screening placeholders**. They are not final AISC/AISI/AWS code equations. Replace the placeholder formulas in `engine/joint_checks.py` with approved BAC/company/code equations before design release.

## Privacy recommendation

Start local-only. Uploaded files stay in the running Streamlit session unless the code is modified to save them. Do not deploy real BAC/customer CAD data to a cloud service without an approved security, authentication, storage, and deletion plan.
