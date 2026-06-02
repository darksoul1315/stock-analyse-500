# 🚀 Nifty 500 Quantitative Rotational Scanner & Telegram Alert Dispatcher

An advanced, institutional-grade sector rotation and market breadth scanner for Nifty 500 stocks. 

This engine automatically synthesizes capital-weighted industry indices, implements **Julius de Kempenaer's Relative Rotation Graphs (RRG)** model, computes multi-timeframe trend strengths, pulls real-time financial news sentiment polarity scores, and generates a beautifully styled 5-sheet corporate Excel dashboard. 

At the end of every run (including the scheduled **8:45 PM daily automation**), a formatted market alert and the compiled `.xlsx` report are sent straight to your **Telegram chat**!

---

## 📊 Key Mathematical & Quantitative Architecture

### 1. Relative Rotation Graph (RRG) Engine
Computes standard deviation normalized **RS-Ratio** (representing relative trend direction) and **RS-Momentum** (representing relative trend strength) baselined around a neutral threshold of `100.0` against the benchmark index (`^NSEI`):
* **RS-Ratio**: Measures relative outperformance over a 60-day trend horizon:
  $$\text{RS-Ratio}(t) = 100 + 10 \times \left( \frac{\text{RS}(t) - \text{EMA}_{60}(\text{RS})}{\sigma_{60}(\text{RS})} \right)$$
* **RS-Momentum**: Measures rate of change of relative trend outperformance:
  $$\text{RS-Momentum}(t) = 100 + 10 \times \left( \frac{\text{EMA}_{10}(d\_\text{RS-Ratio}) - \text{EMA}_{60}(d\_\text{RS-Ratio})}{\sigma_{60}(d\_\text{RS-Ratio})} \right)$$

### 2. Rotational Compass Direction & Heading Vectors
* **Rotational Velocity**: Measures the rate of rotational speed on the 2D plane:
  $$\text{Velocity} = \sqrt{(\text{RS-Ratio}_t - \text{RS-Ratio}_{t-1})^2 + (\text{RS-Momentum}_t - \text{RS-Momentum}_{t-1})^2}$$
* **Compass Heading Angle**: Compass direction of movement ($0^\circ$ to $360^\circ$). An angle of $0^\circ$ to $90^\circ$ represents acceleration pointing North-East towards the **Leading** quadrant:
  $$\text{Heading} = \text{atan2}(\text{RS-Momentum}_t - \text{RS-Momentum}_{t-1},\, \text{RS-Ratio}_t - \text{RS-Ratio}_{t-1}) \times \frac{180}{\pi}$$
  *(Heading angles are mapped compass-style where North is $90^\circ$, East is $0^\circ$, South is $270^\circ$, and West is $180^\circ$)*.

### 3. Quantitative Rotational Score (Sector Rankings)
Ranks all 20 Nifty 500 sectors using a weighted multi-factor scoring model (0 to 100):
* **Sector RSI** (25% weight)
* **EMA Breadth** (30% weight) - percentage of constituent stocks trading above their 50 EMA.
* **Heading Direction** (25% weight) - rewards sectors heading North-East towards the Leading quadrant using a smooth cosine wave: $\max(0, \cos(\text{heading} - 45^\circ)) \times 100$.
* **5D Average Delivery Volume %** (20% weight).

### 4. Advanced Monte Carlo Returns Simulation
Vectorized non-parametric joint bootstrapping Monte Carlo simulation for predicting the probability of sector outperformance:
* **Drift (Momentum Persistence)**: Computes the 30-day average relative return spread between the sector and the benchmark.
* **Volatility (Risk)**: Evaluates the historical standard deviation of the sector's daily relative spread.
* **5000-Path Vectorization**: Employs `np.random.normal(loc=drift, scale=std_dev, size=(5000, 30))` to simulate 5000 future market paths over 15 and 30 trading days.
* **Outperformance Probability (Win Rate)**: Computes the ratio of simulations where the cumulative relative return is strictly positive, translating theoretical risk into a percentage win-rate.

---

## 📂 Project Directory Structure

```text
├── sector_rotation_multi_scanner.py  # Core master execution scanner & compiler
├── verify_modules.py                 # Pipeline integrity verification utility
├── detect_chat_id.py                 # Automated Telegram chat ID configuration utility
├── requirements.txt                  # Python dependencies
├── .env                              # Private environment variables (ignored by git)
└── .gitignore                        # Git ignore rules for secrets and binary reports
```

---

## 💎 Spreadsheet Visualization & Style Sheets

The scanner generates `sector_rotation_multi_report.xlsx` featuring:
1. **Grid Lines**: Explicitly forced active on all worksheets.
2. **Corporate Headers**: Dark Navy Blue header fills (`#1B365D`) with bold white text.
3. **Price/EMA Relation Highlights**: Green fills (`#C6EFCE`) when price is above EMA; Red fills (`#FFC7CE`) when price is below.
4. **Extreme Proximity Bounds**: Green fills when a stock close is within 10% of its multi-year highs (`52W / 2Y / 5Y / 10Y`); Red fills when close is within 10% of its lows.
5. **RRG Quadrant Soft Highlights**: Soft Green for `LEADING`, Soft Blue for `IMPROVING`, Soft Yellow for `WEAKENING`, and Soft Red for `LAGGING`.

---

## 🛠️ Step-by-Step Setup Guide (Quick Start / Kaise Use Kare)

Follow these simple, step-by-step instructions to get the scanner fully running on your system and receiving daily Telegram alerts!

---

### 📋 Prerequisites (System Requirements)
Before you start, make sure you have:
1. **Python 3.8+** installed on your computer.
2. **Terminal** (macOS/Linux) or **Command Prompt/PowerShell** (Windows) open.
3. A **Telegram** account.

---

### 💻 Step 1: Clone the Repository & Go to Folder
Clone this repository from GitHub and navigate into the project directory:
```bash
# Clone the repository
git clone https://github.com/darksoul1315/stock-analyse-500.git

# Enter the project folder
cd stock-analyse-500
```

---

### 📦 Step 2: Install Python Libraries (Dependencies)
Install all required libraries (`pandas`, `numpy`, `yfinance`, and `openpyxl` for Excel generation) in one simple command:
```bash
pip install -r requirements.txt
```
*(If you run into permission errors, try `pip install -r requirements.txt --user` or use a python virtual environment)*.

---

### 🤖 Step 3: Set Up Your Telegram Bot (Alerts Activation)

To get automated alerts sent directly to your phone, you need a Telegram Bot token and your personal Chat ID. Follow these steps:

#### Part A: Create your Bot in 1 Minute
1. Open your Telegram app and search for **`@BotFather`** (the official, verified Telegram bot creator).
2. Click **Start** and send the command:
   ```text
   /newbot
   ```
3. Enter a friendly Name for your bot (e.g., `My Nifty Scanner`).
4. Enter a unique Username ending in `_bot` (e.g., `nifty_rotation_69_bot`).
5. **@BotFather** will reply with a long HTTP API **Token** (looks like `8760476239:AAEhHYH4AM6fZR6wbmCMrWD3Xaewn2fli-U`). **Copy this Token!**

#### Part B: Link your Chat & Get Chat ID Automatically
1. In Telegram, search for your newly created bot username (e.g., `@nifty_rotation_69_bot`) and click **START** (or send any test message to it). *This is important so the bot has permission to message you!*
2. Go back to your terminal, open [`.env`](file:///Users/rajeevkumar/Library/CloudStorage/GoogleDrive-kituraj22@gmail.com/My%20Drive/BACKTESTING/open%20intrest/.env) file (if it exists, or create a new file named `.env`) and paste your Bot Token:
   ```env
   TELEGRAM_BOT_TOKEN=your_copied_token_here
   TELEGRAM_CHAT_ID=
   ```
3. Run our automated Chat ID detector script in the terminal:
   ```bash
   python3 detect_chat_id.py
   ```
4. **Boom!** The script will listen to your bot, extract your private **Chat ID**, print it on the screen, and **automatically save it into your `.env` file!** Your `.env` will now look like this:
   ```env
   TELEGRAM_BOT_TOKEN=1234567890:ABCDefghIJKLmnopQRSTuvwxYZ123456789
   TELEGRAM_CHAT_ID=123456789
   ```

*(Note: Your `.env` file contains your private keys. It is automatically ignored by Git and will never be pushed to your public GitHub repo, keeping your bot 100% secure).*

---

### 🧪 Step 4: Verify Your Setup & Test Connection
Before running the full production scanner (which downloads 500 stocks), run this fast verification script to test if the data downloads, technical indicators, and news sentiment systems are fully operational:
```bash
python3 verify_modules.py
```
*If everything is correct, you will see a green success message:*
`🎉 ALL PIPELINES ARE 100% OPERATIONAL, INTEGRATED, AND CORRECT!`

---

### 🚀 Step 5: Run the Scanner Manually (with Smart Caching)

Run the production script to execute the scanner, build the custom Excel sheets, and send the results to your Telegram chat instantly:

```bash
# Normal Run (Uses local Parquet cache if available, runs in ~25 seconds!)
python3 sector_rotation_multi_scanner.py

# Force Refresh (Ignores cache, downloads fresh 10-year historical price data from Yahoo Finance)
python3 sector_rotation_multi_scanner.py --force-refresh
```

* **⚡ Smart Cache Mechanics**:
  1. **First-time Run**: Downloads the full historical 10-year data for all 500 stocks and caches it securely in a local `price_cache_<date>.parquet` file.
  2. **Subsequent Daily Runs**: Loads today's parquet file directly. Over **4x performance speedup** (running in ~24 seconds vs 98 seconds!).
  3. **Next-day Incremental Fetch**: Only downloads the missing days' data, automatically merges it with the existing cache, and cleans up old cache files.
* **What happens now?** 
  * The script will scan all Nifty 500 stocks, calculate technical breadths, parse financial news headlines, and compile the styled `sector_rotation_multi_report.xlsx` sheet.
  * You will immediately receive a **Telegram message** summarizing the **Top 3 RRG Sectors** and the **Top 5 Breakout Stock Candidates**, along with the **Excel Spreadsheet attached** to view on your mobile or PC!


---

## 📅 Step 6: Set Up Daily Automatic Runs (8:45 PM Daily)

To automate the script so it runs every single day at **8:45 PM (20:45)** without opening your terminal, configure a daily schedule:

### On macOS / Linux (using `cron`):
1. Find your system's exact Python 3 path by running:
   ```bash
   which python3
   ```
   *(Usually it is `/usr/bin/python3` or `/Library/Frameworks/Python.framework/...`)*.
2. Get the absolute path to your folder by running:
   ```bash
   pwd
   ```
3. Open your system's cron scheduler configuration:
   ```bash
   crontab -e
   ```
4. Press `i` to enter edit mode, and paste the following line at the very bottom (replace `/usr/bin/python3` with your python path and `/path/to/folder` with your absolute folder path):
   ```text
   45 20 * * * cd "/path/to/folder" && /usr/bin/python3 sector_rotation_multi_scanner.py >> daily_cron_run.log 2>&1
   ```
5. Press `Esc` then type `:wq` and press `Enter` to save and exit.
6. **Done!** The system will now execute the scanner daily at 8:45 PM and save all outputs/warnings to `daily_cron_run.log`.

---

## 📜 Disclaimer
This quantitative scanner is created solely for research, educational, and backtesting purposes. It does not constitute financial advice, buy/sell recommendations, or market solicitation. Please consult a registered investment advisor before committing capital.

