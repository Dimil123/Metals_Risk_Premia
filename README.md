# Metals Risk Premia Dashboard

Interactive Streamlit dashboard for exploring LME & CME metals market data.

## 📂 Required Data Files

Upload these via the sidebar when running the app:

1. **Metals Cash and 3M.xlsx** — LME Cash/3M/Spread data (one sheet per metal)
2. **Metals Futures Curve.xlsx** — Futures strip F1–F27 (one sheet per metal)

## 🚀 Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## ☁️ Deploy to Streamlit Cloud (Public URL)

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub account
4. Select this repo → `app.py` as the main file
5. Deploy — you'll get a public URL like `https://your-app.streamlit.app`

## 📊 Dashboard Tabs

| Tab | Description |
|-----|-------------|
| Market Overview | Latest prices, spreads, backwardation/contango status |
| Term Structure | Interactive futures curve with date slider |
| Cash vs 3M | Carry analysis, annualized carry, spread distribution |
| Volume & OI | Trading activity, liquidity heatmap |
| Cross-Metal | LME Copper vs COMEX Copper location arbitrage |
| Statistics | Summary stats, rolling vol, return distribution, regime analysis |

## 📝 Notes

- LME metals quoted in $/MT
- Precious metals (Gold, Silver, Platinum, Palladium) quoted in $/oz
- COMEX Copper quoted in $/lb (auto-converted to $/MT for comparison)
- NYU Financial Engineering — Metals Risk Premia Project
