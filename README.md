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

## 🛠️ Installation & Setup (Step-by-Step)

### Step 1: Clone the Repository
Clone the repository to your local system:
```bash
git clone https://github.com/darksoul1315/stock-analyse-500.git
cd stock-analyse-500
```

### Step 2: Install Python Dependencies
Install the required packages using `requirements.txt`:
```bash
pip install -r requirements.txt
```

### Step 3: Configure Telegram Bot Credentials
To enable automated alerts, you must create a Telegram bot and configure your credentials:

1. **Create your Bot**:
   * Open Telegram, search for **`@BotFather`** and start a chat.
   * Send the `/newbot` command and follow the prompts to choose a name and username.
   * Copy the **HTTP API Token** provided (e.g., `8760476239:AAEhHYH4...`).

2. **Configure credentials automatically**:
   * Open Telegram and search for your newly created bot username (e.g., `@darksouls69_bot`).
   * Press **START** (or send any text message to it).
   * Run the interactive auto-detector script:
     ```bash
     python3 detect_chat_id.py
     ```
   * The script will listen for the start message, extract your private **Chat ID**, and automatically create and configure your local `.env` file!

3. **Verify Configuration**:
   Verify your `.env` contains:
   ```env
   TELEGRAM_BOT_TOKEN=your_token_here
   TELEGRAM_CHAT_ID=your_chat_id_here
   ```

---

## ⚡ Running & Verifying the Pipeline

### 1. Verify Pipeline Integrity
Run the verification check script to ensure all components (data fetcher, delivery downloader, technical calculators, news sentiment parser, and sector synthesis) execute cleanly:
```bash
python3 verify_modules.py
```

### 2. Run the Scanner manually
Execute the full production pipeline to compile the workbook and dispatch the Telegram alert:
```bash
python3 sector_rotation_multi_scanner.py
```

---

## 📅 Daily Cron Scheduling (macOS Automation)

To automatically run this program daily at **8:45 PM**, you can add it to your macOS `crontab`:

1. Open your terminal and run:
   ```bash
   crontab -e
   ```
2. Add the following entry (adjusting the paths to match your local python installation and workspace path):
   ```text
   45 20 * * * cd "/Users/rajeevkumar/Library/CloudStorage/GoogleDrive-kituraj22@gmail.com/My Drive/BACKTESTING/open intrest" && /usr/bin/python3 sector_rotation_multi_scanner.py >> daily_cron_run.log 2>&1
   ```
3. Save and close. Output logs and warnings will be saved daily to `daily_cron_run.log`.

---

## 📜 Disclaimer
This scanner is for research and educational purposes only. It is not financial advice.
