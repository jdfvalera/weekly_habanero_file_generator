# Weekly Habanero Report Generator

This is a Streamlit web app used to generate a formatted weekly Excel report from two uploaded data files.

Users upload:
1. A **Weekly Data Pull** Excel file
2. A **Frequency File** Excel file

The app processes the data, previews the result, and generates a downloadable Excel report with the correct formatting applied.

No local setup or Python knowledge is required to use the app.

---

## How to Use (For Teammates)

1. Open the app using the shared link
2. Enter:
   - **Client name** (e.g. `Wray's`)
   - **Report number** (e.g. `8`)
3. Upload:
   - Weekly Data Pull (`.xlsx`)
   - Frequency File (`.xlsx`)
4. Click **Generate Report**
5. Review the preview table
6. Click **Download Excel File**

The downloaded file will be named automatically using the client name, report number, and date range.

---

## Input File Requirements

### Weekly Data Pull
- Excel file (`.xlsx`)
- Must contain a sheet named **`Data`**
- Must include at least:
  - `Date`
  - `Impressions`
  - `Click Rate (CTR)`

### Frequency File
- Excel file (`.xlsx`)
- Must contain a sheet named **`Data`**
- Must include:
  - `Date`
  - `Unique Reach: Average Impression Frequency`

Dates are parsed automatically and rows with invalid dates are ignored.

---

## Output

The generated Excel file:
- Contains merged weekly + frequency data
- Calculates **Reach = Impressions / Frequency**
- Applies formatting:
  - Date columns formatted as dates
  - CTR formatted as percentages
  - Reach formatted as whole numbers with separators

---

## For Developers

### Running Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app will be available at:
```
http://localhost:8501
```

---

## Deployment

This app is deployed using **Streamlit Community Cloud** (free tier).

Requirements:
- Public GitHub repository
- `app.py` at repo root
- `requirements.txt` at repo root

No additional configuration is required.

---

## Notes

- Uploaded files are processed in-memory only
- No data is stored after the session ends
- The app may briefly sleep when inactive (free tier behavior)

---

## Maintainer

Internal analytics utility maintained by the data team.