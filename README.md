# Y&S Tickets — Installment Projections
## Streamlit Web App

Upload a General Ledger `.xlsx` export and instantly download the formatted
Installment Projections report (Summary By Account, Summary By Team, Transactions).

---

## Local Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the app
streamlit run app.py
```

Then open http://localhost:8501 in your browser.

---

## Deploy to Streamlit Cloud (Free, Shareable URL)

1. **Create a free account** at https://share.streamlit.io

2. **Push these files to a GitHub repo:**
   ```
   app.py
   requirements.txt
   README.md
   ```

3. **In Streamlit Cloud:**
   - Click **New app**
   - Connect your GitHub repo
   - Set **Main file path** to `app.py`
   - Click **Deploy**

4. You'll get a public URL like:
   `https://your-app-name.streamlit.app`

   Share that link with anyone — no login required to use the app.

---

## What the Report Contains

| Tab | Contents |
|-----|----------|
| **Summary By Account** | Divvy/Slash/Wex grouped separately; other accounts below; subtotals + grand total |
| **Summary By Team** | All teams ranked high to low by spend |
| **Transactions** | All transactions sorted high to low, Date formatted MM/DD/YYYY |

### Account Relabeling Rules
| Original | Label |
|----------|-------|
| SP2, SP3, etc. | Slash |
| Divvy CR* | Divvy Credit |
| Divvy PF* | Divvy Prefund |
| Wex CR*, Wex (Credit*) | Wex Credit |
| Wex (Prefund) | Wex Prefund |

### Filters Applied
- Removes transactions with blank **Type**
- Removes **Clearing Account** transactions

### Filename Format
Auto-generated from the ledger file:
- **Company name** → Row 2
- **Date** → Row 3 date range → 15th of the following month

Example: `Y&S Tickets - Installment Projections - May 15th.xlsx`
