# Agreement Address Quality Checker (4-layer pipeline)

Flags improper addresses in an Excel sheet of customer agreements using four
layers, from cheapest/fastest to smartest:

1. **Structural rules** — mechanical corruption (free, offline, instant)
2. **Pincode master-data validation** — geographic plausibility (free public API)
3. **Placeholder/gibberish dictionary** — junk/fake entries (free, offline)
4. **AI semantic review** — catch-all for everything else (optional, your own free Gemini key)

## 1. Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

Open the local URL Streamlit prints. Click **"Try demo data"** to see all
four layers work on real examples (including "Main Road 1, Dubai" and "NA"),
or upload your `.xlsx` and pick your Agreement No / Address columns.

## 2. What each layer catches

### Layer 1 — Structural rules
| Flag | Meaning |
|---|---|
| `EMPTY_ADDRESS` | Address field is blank |
| `DOUBLE_COMMA_EMPTY_FIELD` | Contains `,,` — empty field between commas |
| `PINCODE_DUPLICATED` | Pincode repeated back-to-back, e.g. `400054400054` |
| `PINCODE_GLUED_TO_TEXT` | Pincode stuck to a word, e.g. `MUMBAI400206` |
| `STATE_NOT_FOUND` | No recognizable Indian state name present |
| `REPEATED_PHRASE(...)` | A phrase repeats, e.g. "MAIN ROAD ... MAIN ROAD" |
| `ADDRESS_TOO_SHORT` | Fewer words than the minimum (adjustable, default 5) |
| `POSSIBLE_MERGED_WORDS(...)` | A long word may be two+ words glued together |
| `HOUSE_NO_ZERO_OR_PLACEHOLDER` | House number looks like a placeholder, e.g. `# 0` |

### Layer 2 — Pincode master-data validation (needs internet)
Every pincode is checked against the free, keyless **All India Pincode API**
(static JSON on GitHub Pages, sourced from the official Dept. of Posts /
data.gov.in dataset — no server, no rate limit, no signup):
`https://aniket-thapa.github.io/india-pincode-api`

| Flag | Meaning |
|---|---|
| `PINCODE_NOT_FOUND_IN_INDIA` | Pincode doesn't exist in the official directory — catches fake numbers and non-Indian addresses like "Main Road 1, Dubai" |
| `PINCODE_STATE_MISMATCH(...)` | The pincode belongs to a different state than what's written in the address |

**License note:** this API's data is CC BY-NC 4.0 (non-commercial use with
attribution). If this tool will be used for a commercial product rather than
internal data-quality review, either get written permission from the API
author or swap in the official data.gov.in "All India Pincode Directory"
download (link in the API repo) as a local file instead — the app's design
makes that a small change confined to `pincode_lookup.py`.

If there's no internet or the API is briefly down, this layer is skipped
gracefully — you'll see a warning banner, and the row still gets Layers 1
and 3.

### Layer 3 — Placeholder / gibberish dictionary (always on, free, offline)
| Flag | Meaning |
|---|---|
| `PLACEHOLDER_ADDRESS` | Entire field is a placeholder value: NA, TEST, XXX, TBD, "same as above", etc. |
| `PLACEHOLDER_WORD(...)` | Contains a junk word like TEST, DUMMY, ASDF |
| `FOREIGN_LOCATION_MENTIONED(...)` | Mentions a non-Indian city/country (Dubai, Singapore, London, etc.) |
| `REPEATED_CHARACTER_RUN` | Same character repeated 4+ times, e.g. `aaaa`, `9999` |
| `POSSIBLE_GIBBERISH_TEXT` | Long consonant run suggests random typing |

Extend `PLACEHOLDER_PHRASES`, `PLACEHOLDER_WORDS`, and
`FOREIGN_LOCATION_HINTS` at the top of `app.py` as you discover new junk
patterns in your real data — this list is meant to grow over time.

### Layer 4 — Optional AI semantic review (Gemini free tier)
Catches everything the first three layers can't anticipate: nonsensical but
well-formatted addresses, subtly wrong details, anything genuinely novel.

1. Get a free key: https://aistudio.google.com/app/apikey
2. In the sidebar, enable "AI semantic review" and paste the key
3. Choose scope:
   - **Flagged rows only** (default) — cheapest, sanity-checks what's already caught
   - **Random sample** — quick spot-check across the whole file
   - **All rows** — most thorough, slowest, uses the most free-tier quota

Requests are paced (~1/second) to stay within free-tier rate limits. Only
the address text is sent — not other columns from your file.

## 3. Severity levels
- **Critical** — undeliverable: empty, no pincode, pincode doesn't exist,
  placeholder value, or far too short
- **Warning** — needs a human look: structural glitches, mismatches,
  possible merged words, gibberish signals
- **Clean** — passed all active layers

## 4. Review queue (human-in-the-loop)
After running a check, flagged rows appear in an editable table. Mark each
as **Confirmed Issue** or **False Positive** and click **Save reviewer
decisions**. This is written to `reviewer_feedback.csv` next to the app and
is automatically re-applied on every future run — an address marked "False
Positive" once won't be flagged again, so the tool gets quieter over time as
you use it.

Back up `reviewer_feedback.csv` if you redeploy the app somewhere new.

## 5. Exporting results
Two download buttons after each run:
- **Flagged only** — for sending back to whoever owns the data
- **Full results** — every row with status, for audit/record-keeping

## 6. Deploying for free (so your team can use it without you running it locally)
1. Push this folder to a GitHub repo.
2. Go to https://share.streamlit.io (Streamlit Community Cloud — free tier).
3. Sign in with GitHub, "New app", point it at your repo and `app.py`.
4. Deploy — you get a shareable `*.streamlit.app` link.

Don't hardcode a Gemini API key in the repo; the app already asks for it in
the sidebar at runtime so each user supplies their own.

## Files
- `app.py` — main Streamlit app, Layers 1 and 3, orchestration, UI
- `pincode_lookup.py` — Layer 2, pincode API client with caching
- `ai_review.py` — Layer 4, Gemini API client
- `feedback_store.py` — review-queue persistence
- `reviewer_feedback.csv` — created automatically after your first saved review
