# Daily Report Bot

This macOS automation opens Chrome, copies a live weather summary from a public
web page, creates a report in Microsoft Excel, saves a date-stamped `.xlsx`, and
captures a screenshot of the completed sheet. The visible browser and Excel
interactions are controlled by PyAutoGUI.

## Setup

```bash
cd daily-report-bot
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

Before the first run, allow Terminal (or your IDE) under **System Settings >
Privacy & Security > Accessibility** and **Screen & System Audio Recording**.
Chrome and Microsoft Excel must be installed.

## Run

Close any open Save As dialog, then run:

```bash
python daily_report_bot.py
```

The bot waits five seconds before taking control. Do not type or move the mouse
until it finishes. To stop it immediately, move the pointer to the top-left
corner of the screen.

To choose another city:

```bash
python daily_report_bot.py --city Bengaluru
```

Generated files are placed in `output/`:

- `daily_report_YYYY-MM-DD.xlsx`
- `daily_report_YYYY-MM-DD.png`

For the assignment recording, start macOS screen recording first, run the bot,
and stop recording after the final Excel sheet and success message appear.
