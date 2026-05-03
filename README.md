# user-data-generator

Generates realistic fake educational-platform data and writes it to CSV files in the `data/` directory. Useful for prototyping dashboards, testing data pipelines, or seeding a database.

---

## Requirements

- Python 3.8+
- [`openpyxl`](https://pypi.org/project/openpyxl/) (for Excel export)

---

## Setup

Clone or download the repo, then navigate to the project folder:

```bash
cd user-data-generator
```

Create and activate a virtual environment, then install dependencies:

```bash
python3 -m venv venv
source venv/bin/activate
pip install openpyxl
```

To deactivate the venv when you're done:

```bash
deactivate
```

---

## Usage

```bash
python generate_data.py
```

The script will create (or overwrite) four CSV files inside the `data/` folder:

| File | Description |
|---|---|
| `data/user_sessions.csv` | Login/logout sessions with duration |
| `data/video_progress.csv` | Video watch progress, completion status, and ratings |
| `data/problem_attempts.csv` | Per-problem attempt results with difficulty, hints, and time |
| `data/quiz_results.csv` | Quiz scores by subject and subsection |

---

## Configuration

Edit the constants at the top of `generate_data.py` to control the output:

| Variable | Default | Description |
|---|---|---|
| `NUM_USERS` | `50` | Number of unique simulated users |
| `BASE_DATE` | `2025-11-01` | Start of the date range for generated timestamps |
| `END_DATE` | `2026-05-02` | End of the date range for generated timestamps |
| `random.seed` | `42` | Seed for reproducibility — change to get different data |

Each generator function (`gen_sessions`, `gen_video_progress`, `gen_problem_attempts`, `gen_quiz_results`) accepts an `n` parameter (default `220`) to control the number of rows produced.

---

## Subjects & Data

Data spans three subjects: **Math**, **Physics**, and **Computer Science**, each with 3 courses, multiple lessons, 15 practice problems, and 7 quiz subsections.
