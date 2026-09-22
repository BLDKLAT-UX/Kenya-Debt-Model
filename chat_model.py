"""
chat_model.py  —  Kenya Debt Model AI Chat Assistant
=====================================================
Ask questions about Kenya's debt sustainability in plain English.

NO EXTRA PACKAGES NEEDED — uses only Python's built-in libraries.

SETUP (one time only):
  1. Open api_key.txt in this folder
  2. Replace PASTE_YOUR_GEMINI_KEY_HERE with your Gemini API key
  3. Save api_key.txt
  4. Run:  python chat_model.py

GET A FREE GEMINI KEY:
  → https://aistudio.google.com/app/apikey
  → Sign in with Google → Create API key → Copy it → paste in api_key.txt

EXAMPLE QUESTIONS:
  • Is Kenya's debt sustainable by 2030?
  • Which stress scenario is most dangerous?
  • What is debt/GDP under FX depreciation in FY26/27?
  • How does debt service compare to the 30% threshold?
  • What fiscal reforms would reduce the debt ratio?
  • Summarise all key risks in two sentences
  • Compare the revenue shortfall and lower growth scenarios
"""

import urllib.request
import urllib.error
import json
import os
import sys
import textwrap
import time

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────
KEY_FILE    = "api_key.txt"
MIN_DELAY   = 5       # minimum seconds between requests (avoids rate limits)
RETRY_WAIT  = 40      # seconds to wait after a rate limit error
MAX_RETRIES = 3       # how many times to retry before giving up

# Current available Gemini models (as of Sep 2026) — tries each in order
GEMINI_MODELS = [
    "gemini-3.6-flash",
    "gemini-2.5-flash-lite",
    "gemini-2.5-pro",
]
GEMINI_URL = (
    "https://generativelanguage.googleapis.com/v1beta/models/"
    "{model}:generateContent?key={key}"
)

# ─────────────────────────────────────────────────────────────────────────────
# KENYA DSA DATA — loaded into every conversation as analyst context
# ─────────────────────────────────────────────────────────────────────────────
SYSTEM_CONTEXT = """You are an expert debt sustainability analyst for Kenya's official DSA model.
Answer all questions using ONLY the data below. Quote specific numbers, interpret policy implications, flag risks clearly. Be concise but thorough. If something is not in the data, say so.

=== KENYA DSA MODEL DATA ===

FISCAL YEARS: FY22/23 FY23/24 FY24/25 FY25/26 FY26/27 FY27/28 FY28/29 FY29/30
(FY22/23-FY23/24 = actuals | FY24/25 onwards = projections)

MACRO:
Real GDP Growth(%): 5.0 5.2 5.4 5.5 5.6 5.7 5.8 5.9
Inflation(%): 6.8 5.4 4.8 4.5 4.2 4.0 4.0 4.0
KES/USD: 125 130 138 142 145 148 151 154
Nominal GDP(KES Bn): 14200 15400 16800 18100 19500 21000 22600 24300
Revenue/GDP(%): 14.8 15.2 15.8 16.2 16.5 16.8 17.0 17.2
Primary Balance/GDP(%): -3.8 -3.4 -3.2 -2.8 -2.4 -2.0 -1.6 -1.2
Overall Fiscal Balance(%): -5.8 -5.4 -5.2 -4.8 -4.4 -4.0 -3.6 -3.2

DEBT STOCK (KES Bn):
Domestic: 1300 1480 1640 1750 1830 1880 1910 1950
External: 4160 4510 4475 4495 4485 4475 4440 4390
Guaranteed: 220 209 202 189 180 173 169 165
Total: 5680 6199 6317 6434 6495 6528 6519 6505

DEBT/GDP SCENARIOS (%) — 55% = IMF/WB anchor threshold:
Baseline:            52.3 57.1 63.7 62.5 61.0 59.8 58.5 57.2
Lower Growth:        52.3 57.1 63.7 66.4 66.9 67.3 67.5 67.4
FX Depreciation:     52.3 57.1 63.7 69.2 67.8 65.5 63.0 61.0
Interest Rate Shock: 52.3 57.1 63.7 64.0 63.2 62.1 61.0 60.0
Revenue Shortfall:   52.3 57.1 63.7 66.8 68.4 70.1 71.5 72.8
Optimistic:          52.3 57.1 63.7 61.0 58.8 56.5 54.2 52.0

DEBT SERVICE / REVENUE (%) — 30% = IMF threshold (Kenya is well above):
Baseline:           47.0 50.2 71.2 66.5 62.0 58.8 55.2 52.0
Lower Growth:       47.0 50.2 71.2 72.0 73.5 74.8 75.2 74.9
FX Depreciation:    47.0 50.2 71.2 78.4 75.1 71.0 67.5 64.0
Revenue Shortfall:  47.0 50.2 71.2 74.5 77.8 80.2 82.5 83.0
Optimistic:         47.0 50.2 71.2 63.0 59.5 56.0 52.5 49.5

PV EXTERNAL DEBT / GDP (%) — 40% = threshold (Kenya is below in all scenarios):
Baseline:     28.4 30.1 29.8 29.2 28.5 27.9 27.3 26.8
Lower Growth: 28.4 30.1 29.8 31.0 31.5 31.9 32.1 32.0
FX Shock:     28.4 30.1 29.8 33.5 32.8 31.5 30.2 29.0

SCENARIO DETAILS:
1. LOWER GROWTH: GDP 2pp below baseline every year → debt peaks 67.5% (FY28/29), never returns to anchor; DS/Rev hits 75.2%
2. FX DEPRECIATION: 15% extra KES depreciation FY25/26 → debt spikes 69.2% (highest single year); DS/Rev hits 78.4%
3. INTEREST RATE SHOCK: +200bps on external debt → moderate, debt stays 2-3pp above baseline, converges slowly
4. REVENUE SHORTFALL: Revenue/GDP 1.5pp below baseline → MOST SEVERE; debt reaches 72.8% by FY29/30 (still rising); DS/Rev hits 83.0%
5. OPTIMISTIC: +1pp growth + revenue gains + consolidation → ONLY scenario below 55% anchor (52.0% FY29/30)

OVERALL DSA RATING: MODERATE-TO-HIGH RISK
- Debt/GDP at 63.7% in FY24/25 breaches 55% anchor by 8.7pp
- Debt service at 71.2% of revenue is 2.4x the 30% IMF threshold — severe liquidity pressure
- Revenue shortfall = most dangerous long-run scenario
- FX depreciation = largest near-term spike risk
- PV external debt stays below 40% in all scenarios — only positive signal
- Baseline improves to 57.2% by FY29/30 but never reaches anchor
- Critical lever: reduce primary deficit toward zero via revenue mobilisation
=== END DATA ==="""


# ─────────────────────────────────────────────────────────────────────────────
# API KEY
# ─────────────────────────────────────────────────────────────────────────────
def load_api_key():
    key_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), KEY_FILE)

    # Try reading the file
    if os.path.exists(key_path):
        with open(key_path, "r", encoding="utf-8") as f:
            lines = f.read().strip().splitlines()
        # First non-comment, non-empty line is the key
        for line in lines:
            line = line.strip()
            if line and not line.startswith("#") and line != "PASTE_YOUR_GEMINI_KEY_HERE":
                return line

    # Key file missing or still has placeholder — guide user
    if not os.path.exists(key_path):
        with open(key_path, "w", encoding="utf-8") as f:
            f.write("PASTE_YOUR_GEMINI_KEY_HERE\n")

    print()
    print("  ┌──────────────────────────────────────────────────────────┐")
    print("  │  API KEY NEEDED — please follow the 3 steps below       │")
    print("  └──────────────────────────────────────────────────────────┘")
    print()
    print("  Step 1: Get your FREE Gemini key")
    print("          → https://aistudio.google.com/app/apikey")
    print("          → Sign in with Google")
    print("          → Click 'Create API key'")
    print("          → Copy the key (starts with AIza... or AQ...)")
    print()
    print("  Step 2: Open api_key.txt in this folder:")
    print(f"          {key_path}")
    print("          Delete PASTE_YOUR_GEMINI_KEY_HERE")
    print("          Paste your key on line 1. Save the file.")
    print()
    print("  Step 3: Run  python chat_model.py  again")
    print()
    print("  ── OR paste your key right now ──")
    key = input("  Key (or press Enter to exit): ").strip()
    if len(key) > 10:
        with open(key_path, "w", encoding="utf-8") as f:
            f.write(key + "\n")
        print(f"\n  ✅  Saved to {KEY_FILE}. Starting chat...\n")
        return key
    print("\n  Exiting. Edit api_key.txt and try again.\n")
    sys.exit(0)


# ─────────────────────────────────────────────────────────────────────────────
# GEMINI API CALL  (no external packages — pure stdlib urllib)
# ─────────────────────────────────────────────────────────────────────────────
_active_model_idx = 0   # tracks which model we're using

def _post_gemini(api_key, model, contents):
    """Single raw HTTP POST to Gemini. Returns (text, error)."""
    url = GEMINI_URL.format(model=model, key=api_key)
    payload = json.dumps({
        "contents": contents,
        "generationConfig": {"temperature": 0.3, "maxOutputTokens": 900},
    }).encode("utf-8")
    req = urllib.request.Request(
        url, data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            body = json.loads(resp.read().decode("utf-8"))
            return body["candidates"][0]["content"]["parts"][0]["text"].strip(), None
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", errors="replace")
        try:
            msg = json.loads(raw)["error"]["message"]
        except Exception:
            msg = raw[:300]
        return None, msg
    except urllib.error.URLError as e:
        return None, f"Network error: {e.reason}"
    except Exception as e:
        return None, str(e)


def call_gemini(api_key, history):
    """
    Call Gemini with:
    - System context prepended as first turn
    - Only last 6 history messages (keeps payload small)
    - Auto-retry on rate limits with wait
    - Fallback to alternate model if first fails
    """
    global _active_model_idx

    # Keep only recent history to reduce payload size
    recent = history[-6:] if len(history) > 6 else history

    contents = [
        {"role": "user",  "parts": [{"text": SYSTEM_CONTEXT}]},
        {"role": "model", "parts": [{"text": "Understood. Kenya DSA data loaded. Ready to answer as expert analyst."}]},
    ] + recent

    for attempt in range(1, MAX_RETRIES + 1):
        model = GEMINI_MODELS[_active_model_idx % len(GEMINI_MODELS)]
        reply, error = _post_gemini(api_key, model, contents)

        if reply:
            return reply, None

        err_low = (error or "").lower()
        is_rate  = "429" in (error or "") or "quota" in err_low or "exhausted" in err_low
        is_model = "not found" in err_low or "404" in (error or "") or "not supported" in err_low

        if is_model:
            # This model doesn't exist — try next one immediately
            _active_model_idx += 1
            if _active_model_idx < len(GEMINI_MODELS):
                continue
            return None, f"No working Gemini model found. Error: {error}"

        if is_rate and attempt < MAX_RETRIES:
            for remaining in range(RETRY_WAIT, 0, -1):
                print(f"  ⏳  Rate limit — retrying in {remaining}s...   ", end="\r")
                time.sleep(1)
            print(" " * 45, end="\r")
            continue

        return None, error

    return None, "Still rate limited after retrying. Wait 60 seconds and try again."


# ─────────────────────────────────────────────────────────────────────────────
# DISPLAY
# ─────────────────────────────────────────────────────────────────────────────
def wrap_print(text, width=74, indent="    "):
    """Print reply with clean word-wrapping."""
    for para in text.split("\n"):
        s = para.strip()
        if not s:
            print()
            continue
        # Preserve bullet points
        if s[0] in ("•", "-", "*", "→", "–"):
            prefix = indent + s[0] + " "
            body   = s[1:].strip()
            for i, line in enumerate(textwrap.wrap(body, width - len(prefix))):
                print((prefix if i == 0 else indent + "  ") + line)
        else:
            for line in textwrap.wrap(s, width - len(indent)):
                print(indent + line)


def banner():
    print()
    print("  ╔════════════════════════════════════════════════════════════╗")
    print("  ║        KENYA DEBT MODEL — AI ANALYST ASSISTANT            ║")
    print("  ║      Powered by Google Gemini 3.6  |  Free tier           ║")
    print("  ╚════════════════════════════════════════════════════════════╝")
    print()
    print("  Type your question and press Enter.")
    print("  Commands:  data · clear · help · quit")
    print()
    print("  Suggested questions:")
    print("    • Is Kenya's debt sustainable by 2030?")
    print("    • Which scenario is most dangerous and why?")
    print("    • Summarise the key risks in two sentences")
    print("    • What reforms would bring debt below 55%?")
    print("    • Compare FX depreciation vs revenue shortfall scenarios")
    print("    • What does debt service look like in the worst case?")
    print()


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    api_key = load_api_key()
    history = []   # list of {"role": "user"/"model", "parts": [{"text": "..."}]}
    last_request_time = 0.0

    banner()

    while True:
        try:
            user_input = input("  You: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n\n  Goodbye.\n")
            break

        if not user_input:
            continue

        cmd = user_input.lower()

        if cmd in ("quit", "exit", "q"):
            print("\n  Goodbye.\n")
            break

        if cmd == "clear":
            history = []
            print("\n  ✅  Conversation cleared. Fresh start.\n")
            continue

        if cmd == "data":
            print()
            print("  Data loaded into the AI:")
            print("    • Macro: GDP growth, inflation, exchange rate, revenue/GDP FY22–FY30")
            print("    • Debt stock: domestic, external, guaranteed (KES Billions)")
            print("    • Debt/GDP: 6 scenarios + 55% IMF anchor")
            print("    • Debt service/revenue: 5 scenarios + 30% IMF threshold")
            print("    • PV external debt/GDP: 3 scenarios + 40% threshold")
            print("    • Stress test descriptions and risk interpretations")
            print("    • Overall DSA rating: MODERATE-TO-HIGH RISK")
            print()
            continue

        if cmd == "help":
            print()
            print("  Commands:")
            print("    data   — show what data the AI knows")
            print("    clear  — reset the conversation")
            print("    quit   — exit")
            print("  Anything else is sent to the AI as a question.")
            print()
            continue

        # Enforce minimum gap between requests to avoid rate limits
        elapsed = time.time() - last_request_time
        if elapsed < MIN_DELAY and last_request_time > 0:
            wait = MIN_DELAY - elapsed
            time.sleep(wait)

        history.append({"role": "user", "parts": [{"text": user_input}]})

        print()
        print("  Analyst: thinking...", end="\r", flush=True)

        reply, error = call_gemini(api_key, history)
        last_request_time = time.time()

        if error:
            print("  " + " " * 30)   # clear the thinking line
            print()
            err_low = (error or "").lower()
            if "api_key" in err_low or "invalid" in err_low or "forbidden" in err_low or "403" in str(error):
                print("  ❌  API key is invalid or not recognised.")
                print()
                print("  Fix:")
                print("    1. Open api_key.txt in this folder")
                print("    2. Make sure it contains ONLY your key — nothing else")
                print("    3. The key should look like:  AIzaSy... or AQ.Ab...")
                print("    4. Save and run  python chat_model.py  again")
            elif "still rate limited" in err_low:
                print("  ❌  Gemini rate limit reached even after retrying.")
                print("  Wait 60 seconds, then ask your next question.")
                print("  Tip: space out questions — don't type several quickly.")
            elif "network" in err_low:
                print("  ❌  No internet connection. Check your network.")
            else:
                print(f"  ❌  Gemini error: {error}")
            print()
            history.pop()   # remove failed question from history
            continue

        print("  " + " " * 30)   # clear the thinking line
        print()
        wrap_print(reply)
        print()

        # Add model reply to history
        history.append({
            "role": "model",
            "parts": [{"text": reply}]
        })


if __name__ == "__main__":
    main()
