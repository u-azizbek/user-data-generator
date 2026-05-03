import csv
import random
from datetime import datetime, timedelta
from pathlib import Path
import openpyxl

random.seed(42)

NUM_USERS = 50
BASE_DATE = datetime(2025, 11, 1)
END_DATE = datetime(2026, 5, 2)

def rand_dt(start=BASE_DATE, end=END_DATE):
    delta = int((end - start).total_seconds())
    return start + timedelta(seconds=random.randint(0, delta))

# Fixed IDs so FK references are consistent
user_ids    = [f"U{i+1:03d}" for i in range(NUM_USERS)]
lesson_ids  = [f"L{i+1:03d}" for i in range(80)]
problem_ids = [f"P{i+1:03d}" for i in range(100)]
quiz_ids    = [f"Q{i+1:03d}" for i in range(40)]

subjects = ["Math", "Physics", "Computer Science"]

courses = {
    "Math": [
        ("calculus-101",        ["limits-intro","derivatives-basics","chain-rule","integration-by-parts","definite-integrals","taylors-series"]),
        ("linear-algebra",      ["vectors-intro","matrix-operations","eigenvalues","dot-product","linear-transformations","determinants"]),
        ("probability-stats",   ["basic-probability","conditional-probability","distributions","hypothesis-testing","regression","bayes-theorem"]),
    ],
    "Physics": [
        ("classical-mechanics", ["newtons-laws","kinematics-1d","kinematics-2d","energy-work","momentum","rotational-motion"]),
        ("electromagnetism",    ["electric-fields","coulombs-law","magnetic-fields","faradays-law","circuits","capacitors"]),
        ("thermodynamics",      ["heat-transfer","first-law","second-law","entropy","ideal-gas","carnot-cycle"]),
    ],
    "Computer Science": [
        ("algorithms-101",      ["big-o-notation","sorting-basics","binary-search","recursion","dynamic-programming","greedy-algorithms"]),
        ("data-structures",     ["arrays-lists","linked-lists","trees-intro","hash-tables","graphs","heaps"]),
        ("web-development",     ["html-basics","css-layouts","javascript-intro","api-design","databases","authentication"]),
    ],
}

problems_data = {
    "Math": [
        ("Find the derivative of sin(x²)",               "Derivative",        "Easy"),
        ("Evaluate ∫₀¹ x² dx",                           "Integral",          "Easy"),
        ("Find eigenvalues of [[2,1],[1,2]]",             "Eigenvalue",        "Medium"),
        ("Solve system: 3x+2y=12, x-y=1",                "Linear System",     "Easy"),
        ("Compute lim(x→0) sin(x)/x",                    "Limit",             "Medium"),
        ("P(A∩B) if P(A)=0.4, P(B)=0.5, independent",   "Probability",       "Easy"),
        ("Find variance of [2,4,4,4,5,5,7,9]",           "Statistics",        "Medium"),
        ("Integrate x·eˣ by parts",                      "Integral",          "Medium"),
        ("Find critical points of f(x)=x³-3x",           "Derivative",        "Medium"),
        ("Matrix multiply [[1,2],[3,4]]·[[5,6],[7,8]]",  "Matrix",            "Easy"),
        ("Prove that √2 is irrational",                  "Proof",             "Hard"),
        ("Find the radius of convergence of Σ xⁿ/n",    "Series",            "Hard"),
        ("Diagonalise A=[[4,1],[2,3]]",                  "Eigenvalue",        "Hard"),
        ("Solve ∫ x²·ln(x) dx",                         "Integral",          "Hard"),
        ("Find gradient of f(x,y)=x²y+y³",              "Multivariable",     "Medium"),
    ],
    "Physics": [
        ("Ball thrown at 20 m/s upward. Find max height.",       "Kinematics",      "Easy"),
        ("Force on m=5 kg with a=3 m/s²",                        "Newton's Law",    "Easy"),
        ("Find KE: m=2 kg, v=10 m/s",                           "Energy",          "Easy"),
        ("Two charges q=1µC, r=0.1 m. Find E field.",           "Electric Field",  "Medium"),
        ("Spring k=200 N/m stretched 0.05 m. Find PE.",         "Potential Energy","Easy"),
        ("Find momentum: m=3 kg, v=4 m/s",                      "Momentum",        "Easy"),
        ("Ohm's law: V=12 V, R=4 Ω. Find I.",                   "Circuit",         "Easy"),
        ("Object on 30° incline. Find acceleration.",            "Kinematics",      "Medium"),
        ("Heat to raise 1 kg water by 10°C (c=4186 J/kg·K)",   "Thermodynamics",  "Medium"),
        ("Find wavelength: v=340 m/s, f=440 Hz",                "Wave",            "Easy"),
        ("Satellite orbit radius r. Find orbital speed.",        "Gravitation",     "Hard"),
        ("Derive Bernoulli's equation from energy conservation", "Fluid Mechanics", "Hard"),
        ("RLC circuit: find resonant frequency",                 "Circuit",         "Hard"),
        ("Relativistic momentum at v=0.8c",                     "Relativity",      "Hard"),
        ("Entropy change in isothermal expansion",               "Thermodynamics",  "Medium"),
    ],
    "Computer Science": [
        ("What is time complexity of merge sort?",              "Big-O",           "Easy"),
        ("Implement binary search iteratively",                 "Binary Search",   "Easy"),
        ("Find height of a BST",                               "Tree",            "Medium"),
        ("Reverse a singly linked list in-place",              "Linked List",     "Medium"),
        ("Fibonacci with memoization (top-down DP)",           "DP",              "Medium"),
        ("Generate all subsets of [1,2,3]",                    "Recursion",       "Medium"),
        ("Design a hash function for strings",                 "Hash Table",      "Hard"),
        ("BFS vs DFS — when to prefer each?",                  "Graph",           "Easy"),
        ("Longest common subsequence (LCS)",                   "DP",              "Hard"),
        ("Bubble sort vs selection sort complexity",            "Big-O",           "Easy"),
        ("Detect cycle in a directed graph",                   "Graph",           "Hard"),
        ("Implement a min-heap insert operation",              "Heap",            "Medium"),
        ("Prove that P⊆NP",                                    "Complexity",      "Hard"),
        ("Trie: insert and search implementation",             "Trie",            "Hard"),
        ("Find median of a stream of integers",                "Heap",            "Hard"),
    ],
}

quiz_subsections = {
    "Math":             ["Derivatives","Integrals","Linear Algebra","Probability","Statistics","Limits","Series"],
    "Physics":          ["Kinematics","Dynamics","Energy & Work","Electrostatics","Thermodynamics","Waves","Gravitation"],
    "Computer Science": ["Big-O Notation","Sorting Algorithms","Trees & Graphs","Dynamic Programming","Recursion","Hash Tables","System Design"],
}

# ---------- user profiles ----------
# Stable per-user cognitive traits that drive realistic cross-table correlations.
# Every generator uses a user's profile so that a "strong" student is consistently
# better across sessions, quizzes, problems, and video engagement.

def gen_user_profiles():
    profiles = {}
    for uid in user_ids:
        # Overall academic ability: drives correct rates, quiz scores, hint usage
        skill      = round(max(0.25, min(1.00, random.gauss(0.62, 0.20))), 3)
        # Engagement: drives video completion, session length, willingness to retry
        engagement = round(max(0.20, min(1.00, random.gauss(0.62, 0.20))), 3)
        # One subject where this student has an edge (+score, +correct rate)
        preferred_subject = random.choice(subjects)
        # Consistency: high = low variance; low = "hot/cold" student
        consistency = round(random.uniform(0.35, 1.00), 3)
        profiles[uid] = {
            "skill":             skill,
            "engagement":        engagement,
            "preferred_subject": preferred_subject,
            "consistency":       consistency,
        }
    return profiles

user_profiles = gen_user_profiles()

# ---------- fatigue & attention helpers ----------

def time_of_day_label(hour):
    if 5 <= hour < 11:
        return "Morning"
    elif 11 <= hour < 17:
        return "Afternoon"
    elif 17 <= hour < 22:
        return "Evening"
    else:
        return "Night"

def compute_fatigue(hour, duration_minutes):
    """
    Returns fatigue_score 1–10.
    Neuroscience basis: alertness follows a ~24 h circadian cycle with a strong dip
    between 2–4 AM and a smaller post-lunch dip (14:00–16:00).
    Each 30 min of continuous study adds ~1 extra fatigue point.
    """
    if 5 <= hour < 9:
        tod = random.uniform(1.0, 3.0)    # fresh morning start
    elif 9 <= hour < 14:
        tod = random.uniform(2.0, 5.0)    # productive mid-morning window
    elif 14 <= hour < 16:
        tod = random.uniform(4.0, 6.5)    # post-lunch circadian dip
    elif 16 <= hour < 20:
        tod = random.uniform(2.5, 5.5)    # afternoon second wind
    elif 20 <= hour < 23:
        tod = random.uniform(4.0, 7.0)    # evening wind-down
    else:
        tod = random.uniform(6.0, 9.5)    # late night / very early morning
    raw = tod + duration_minutes / 30.0 + random.gauss(0, 0.5)
    return round(max(1.0, min(10.0, raw)), 1)

def compute_attention(fatigue_score, skill):
    """
    Attention span 1–10.
    Strongly inversely correlated with fatigue (key neuroscience correlation).
    Higher-skill students sustain attention better under the same fatigue load.
    """
    base      = 11.0 - fatigue_score           # direct inverse of fatigue
    skill_adj = (skill - 0.62) * 2.2           # ±~0.8 across skill range 0.25–1.0
    return round(max(1.0, min(10.0, base + skill_adj + random.gauss(0, 0.8))), 1)

# ---------- generators ----------

def gen_sessions(n=220):
    rows = []
    for i in range(n):
        uid = random.choice(user_ids)
        p   = user_profiles[uid]
        login = rand_dt()
        hour  = login.hour
        # Engaged users study longer on average
        duration = max(5, min(240, int(random.gauss(30 + p["engagement"] * 90, 35))))
        logout = login + timedelta(minutes=duration)

        fatigue   = compute_fatigue(hour, duration)
        attention = compute_attention(fatigue, p["skill"])

        # KEY CORRELATION: high fatigue → shorter uninterrupted spans (more breaks)
        raw_unint   = duration * (attention / 10.0) * random.gauss(1.0, 0.10)
        uninterrupted = max(1, min(duration, int(raw_unint)))

        # primary_subject: weighted toward user's preferred subject
        other_subjects = [s for s in subjects if s != p["preferred_subject"]]
        primary_subject = random.choices(
            [p["preferred_subject"]] + other_subjects,
            weights=[0.55, 0.225, 0.225]
        )[0]

        # task_switching_count: low attention → more bouncing between modules
        switch_base = (10 - attention) * 1.5 + random.gauss(0, 2.0)
        task_switching_count = max(0, round(switch_base))

        # rage_click_count: high fatigue + low attention → frustration clicks
        rage_base = fatigue * 0.6 + (10 - attention) * 0.5 + random.gauss(0, 1.5)
        rage_click_count = max(0, round(rage_base))

        # idle_time_minutes: high fatigue + low attention → more dead time
        idle_base = (fatigue / 10) * duration * 0.25 + (10 - attention) / 10 * duration * 0.10 + random.gauss(0, 2)
        idle_time_minutes = max(0, min(duration - uninterrupted, round(idle_base)))

        # tab_out_count: low attention + high fatigue → leaving the platform tab
        tab_out_base = (10 - attention) * 0.9 + fatigue * 0.4 + random.gauss(0, 1.5)
        tab_out_count = max(0, round(tab_out_base))

        # lessons_completed: session length + engagement + attention (~20 min/lesson)
        lessons_base = (duration / 20) * (0.35 + p["engagement"] * 0.65) * (attention / 10)
        lessons_completed = max(0, round(random.gauss(lessons_base, 0.8)))

        # problems_attempted: engagement + attention drive attempts (~10 min/problem)
        problems_base = (duration / 10) * (0.25 + p["engagement"] * 0.5) * (attention / 10)
        problems_attempted = max(0, round(random.gauss(problems_base, 1.5)))

        # quizzes_taken: rare, ~1 per 50 min for engaged users
        quiz_base = (duration / 50) * (0.2 + p["engagement"] * 0.8)
        quizzes_taken = max(0, round(random.gauss(quiz_base, 0.4)))

        # device_type: mobile skews toward night hours
        if hour >= 22 or hour < 6:
            device_type = random.choices(["Desktop", "Mobile", "Tablet"], weights=[0.35, 0.50, 0.15])[0]
        else:
            device_type = random.choices(["Desktop", "Mobile", "Tablet"], weights=[0.55, 0.30, 0.15])[0]

        rows.append([
            f"S{i+1:04d}",
            uid,
            login.strftime("%Y-%m-%d %H:%M:%S"),
            logout.strftime("%Y-%m-%d %H:%M:%S"),
            duration,
            uninterrupted,
            time_of_day_label(hour),
            fatigue,
            attention,
            primary_subject,
            task_switching_count,
            rage_click_count,
            idle_time_minutes,
            tab_out_count,
            lessons_completed,
            problems_attempted,
            quizzes_taken,
            device_type,
            None,  # streak_day — filled in post-processing
            None,  # days_since_last_session — filled in post-processing
        ])

    # Sort by user then login time to compute per-user streak and gap
    rows.sort(key=lambda r: (r[1], r[2]))
    last_date = {}   # uid -> date of last session
    streak    = {}   # uid -> current streak_day
    for row in rows:
        uid        = row[1]
        login_date = datetime.strptime(row[2], "%Y-%m-%d %H:%M:%S").date()
        if uid not in last_date:
            days_since = ""   # first recorded session — no prior data
            streak_day = 1
        else:
            diff = (login_date - last_date[uid]).days
            days_since = diff
            if diff == 0:
                streak_day = streak[uid]          # multiple sessions same day
            elif diff == 1:
                streak_day = streak[uid] + 1      # consecutive day — streak continues
            else:
                streak_day = 1                    # gap — streak resets
        last_date[uid] = login_date
        streak[uid]    = streak_day
        row[-2] = streak_day
        row[-1]  = days_since

    # Re-number session IDs in chronological order
    for i, row in enumerate(rows):
        row[0] = f"S{i+1:04d}"

    return rows

def gen_video_progress(n=220):
    rows = []
    for i in range(n):
        uid = random.choice(user_ids)
        p   = user_profiles[uid]
        subject = random.choice(subjects)
        course_slug, lesson_slugs = random.choice(courses[subject])
        lesson_slug = random.choice(lesson_slugs)
        video_len = random.randint(300, 3600)

        # Engagement + subject affinity drive how far the user watches
        completion_prob = 0.30 + p["engagement"] * 0.58
        if subject == p["preferred_subject"]:
            completion_prob = min(0.97, completion_prob + 0.10)

        if random.random() < completion_prob:
            watched = random.randint(int(video_len * 0.85), video_len)
        else:
            # Low-engagement users drop off proportionally earlier
            max_frac = 0.20 + p["engagement"] * 0.55
            watched = random.randint(0, int(video_len * max_frac))

        is_completed = watched >= video_len * 0.9

        # playback_speed_avg: high-skill / preferred-subject users skim faster;
        # low-engagement users who struggle tend to stay at 1.0x
        speed_weights = [0.45, 0.35, 0.20]  # 1.0x, 1.5x, 2.0x baseline
        if p["skill"] >= 0.70 or subject == p["preferred_subject"]:
            speed_weights = [0.20, 0.40, 0.40]  # skew faster
        elif p["skill"] <= 0.40:
            speed_weights = [0.65, 0.28, 0.07]  # skew slower
        playback_speed_avg = random.choices([1.0, 1.5, 2.0], weights=speed_weights)[0]

        # video_duration_seconds: this IS video_len (expose it so AI can calc %)
        video_duration_seconds = video_len

        # watched_minutes proxy for per-minute event rates
        watched_minutes = max(1, watched / 60)

        # rewind_count: high skill → fewer rewinds; dense material (slow speed) → more
        rewind_base = (1.1 - p["skill"]) * 3.0 + (2.0 - playback_speed_avg) * 2.0
        rewind_count = max(0, round(rewind_base * watched_minutes / 10 + random.gauss(0, 1.5)))

        # pause_count: engagement drives note-taking pauses; low attention adds fatigue pauses
        pause_base = p["engagement"] * 2.5 + random.gauss(0, 1.2)
        pause_count = max(0, round(pause_base * watched_minutes / 10))

        # skip_forward_count: boredom / too-easy → more skips; preferred subject + low skill → fewer
        skip_base = (1.0 - p["engagement"]) * 3.0
        if subject == p["preferred_subject"] and p["skill"] < 0.55:
            skip_base *= 0.4   # struggling in favourite subject → watches carefully
        skip_forward_count = max(0, round(skip_base * watched_minutes / 10 + random.gauss(0, 1.0)))

        # tab_out_count: low engagement + long video → more tab-outs
        tab_base = (1.0 - p["engagement"]) * 2.5 + (video_len / 3600) * 2.0
        tab_out_count = max(0, round(tab_base + random.gauss(0, 1.0)))

        # suggested_problems_solved: engaged users who finish act on recommended problems
        if is_completed and random.random() < 0.55 + p["engagement"] * 0.35:
            suggested_problems_solved = random.randint(1, max(1, round(p["skill"] * 5)))
        else:
            suggested_problems_solved = 0

        rating = None
        if is_completed and random.random() < 0.7:
            # More engaged users skew ratings higher
            if p["engagement"] >= 0.72:
                weights = [1, 2, 10, 37, 50]
            elif p["engagement"] <= 0.40:
                weights = [5, 13, 28, 37, 17]
            else:
                weights = [2, 5, 15, 40, 38]
            rating = random.choices([1, 2, 3, 4, 5], weights=weights)[0]

        rows.append([
            f"VP{i+1:04d}",
            uid,
            random.choice(lesson_ids),
            subject,
            course_slug,
            lesson_slug,
            watched,
            is_completed,
            rating if rating is not None else "",
            rand_dt().strftime("%Y-%m-%d %H:%M:%S"),
            playback_speed_avg,
            rewind_count,
            pause_count,
            skip_forward_count,
            tab_out_count,
            video_duration_seconds,
            suggested_problems_solved,
        ])
    return rows

def gen_problem_attempts(n=220):
    base_correct = {"Easy": 0.78, "Medium": 0.55, "Hard": 0.30}
    base_time    = {"Easy": 120,  "Medium": 300,  "Hard": 600}
    rows = []
    for i in range(n):
        uid = random.choice(user_ids)
        p   = user_profiles[uid]
        subject = random.choice(subjects)
        title, topic_tag, difficulty = random.choice(problems_data[subject])
        ts   = rand_dt()
        hour = ts.hour
        # Estimate how far into a study session this attempt falls
        est_elapsed = random.randint(5, 120)
        fatigue   = compute_fatigue(hour, est_elapsed)
        attention = compute_attention(fatigue, p["skill"])

        # KEY CORRELATIONS:
        #   skill ↑  → correct rate ↑
        #   fatigue ↑ → correct rate ↓  (mental resource depletion)
        #   preferred subject → small boost
        rate  = base_correct[difficulty]
        rate += (p["skill"] - 0.62) * 0.45          # ±~0.17
        rate += 0.08 if subject == p["preferred_subject"] else 0.0
        rate -= (fatigue - 5) * 0.025               # fatigue > 5 starts hurting
        rate  = max(0.04, min(0.97, rate))
        is_correct = random.random() < rate

        # Time: low skill + high fatigue → slower (confusion, re-reading)
        t_base  = base_time[difficulty]
        t_scale = (1.5 - p["skill"]) * (1.0 + (fatigue - 5) * 0.05)
        time_taken = int(random.gauss(t_base * t_scale, t_base * 0.30))
        time_taken = max(15, min(3600, time_taken))

        # Hints: low skill + high fatigue → more hints requested
        hint_p = max(0.03, min(0.85, 0.45 - p["skill"] * 0.35 + (fatigue - 5) * 0.03))
        hints  = sum(1 for _ in range(3) if random.random() < hint_p)

        rows.append([
            f"PA{i+1:04d}",
            uid,
            random.choice(problem_ids),
            subject,
            title,
            difficulty,
            topic_tag,
            is_correct,
            time_taken,
            hints,
            fatigue,
            attention,
            ts.strftime("%Y-%m-%d %H:%M:%S"),
        ])
    return rows

def gen_quiz_results(n=220):
    rows = []
    for i in range(n):
        uid = random.choice(user_ids)
        p   = user_profiles[uid]
        subject    = random.choice(subjects)
        subsection = random.choice(quiz_subsections[subject])
        ts   = rand_dt()
        hour = ts.hour
        est_elapsed = random.randint(10, 180)
        fatigue   = compute_fatigue(hour, est_elapsed)
        attention = compute_attention(fatigue, p["skill"])

        # KEY CORRELATIONS:
        #   skill ↑     → score ↑
        #   fatigue ↑   → score ↓   (core neuroscience correlation)
        #   attention ↑ → score ↑   (mediates the fatigue → score link)
        #   preferred subject → score boost
        #   consistency → controls score variance (reliable vs. hot/cold)
        score  = 67.0
        score += (p["skill"] - 0.62) * 44           # ±~17 pts across full skill range
        score += 9.0 if subject == p["preferred_subject"] else -4.0
        score -= (fatigue - 5) * 2.2                # fatigue 10 → −11 pts; fatigue 1 → +8.8 pts
        score += (attention - 5) * 1.5              # attention 10 → +7.5 pts
        score += random.gauss(0, (1.1 - p["consistency"]) * 12)
        score  = round(max(0.0, min(100.0, score)), 1)

        # Time: lower skill + higher fatigue → slower completion
        t_scale    = (1.5 - p["skill"]) * (1.0 + (fatigue - 5) * 0.04)
        time_taken = int(random.gauss(900 * t_scale, 220))
        time_taken = max(120, min(3600, time_taken))

        rows.append([
            f"QR{i+1:04d}",
            uid,
            random.choice(quiz_ids),
            subject,
            subsection,
            score,
            time_taken,
            fatigue,
            attention,
            ts.strftime("%Y-%m-%d %H:%M:%S"),
        ])
    return rows

# ---------- write ----------

def write_csv(path, headers, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerows(rows)
    print(f"  {path}  →  {len(rows)} rows")

base = Path(__file__).parent / "data"
base.mkdir(exist_ok=True)

write_csv(base / "user_sessions.csv",
    ["session_id","user_id","login_timestamp","logout_timestamp",
     "session_duration_minutes","uninterrupted_minutes","time_of_day",
     "fatigue_score","attention_span_score",
     "primary_subject","task_switching_count","rage_click_count",
     "idle_time_minutes","tab_out_count",
     "lessons_completed","problems_attempted","quizzes_taken",
     "device_type","streak_day","days_since_last_session"],
    gen_sessions())

write_csv(base / "video_progress.csv",
    ["progress_id","user_id","lesson_id","subject_category","course_slug","lesson_slug",
     "seconds_watched","is_completed","user_rating","timestamp",
     "playback_speed_avg","rewind_count","pause_count","skip_forward_count",
     "tab_out_count","video_duration_seconds","suggested_problems_solved"],
    gen_video_progress())

write_csv(base / "problem_attempts.csv",
    ["attempt_id","user_id","problem_id","subject_category","title","difficulty",
     "topic_tag","is_correct","time_taken_seconds","hints_used",
     "fatigue_score","attention_span_score","timestamp"],
    gen_problem_attempts())

write_csv(base / "quiz_results.csv",
    ["result_id","user_id","quiz_id","subject_category","subsection",
     "score_percentage","time_taken_seconds","fatigue_score","attention_span_score",
     "completed_at"],
    gen_quiz_results())

print("\nDone.")

# ---------- combine CSVs into one Excel workbook ----------

sheets = [
    ("user_sessions",     "user_sessions.csv"),
    ("video_progress",    "video_progress.csv"),
    ("problem_attempts",  "problem_attempts.csv"),
    ("quiz_results",      "quiz_results.csv"),
]

wb = openpyxl.Workbook()
wb.remove(wb.active)  # remove default empty sheet

for sheet_name, csv_file in sheets:
    ws = wb.create_sheet(title=sheet_name)
    with open(base / csv_file, newline="", encoding="utf-8") as f:
        for row in csv.reader(f):
            ws.append(row)

excel_path = base / "all_data.xlsx"
wb.save(excel_path)
print(f"  {excel_path}  →  {len(sheets)} sheets")

