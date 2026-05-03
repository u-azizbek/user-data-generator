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

# ---------- generators ----------

def gen_sessions(n=220):
    rows = []
    for i in range(n):
        login = rand_dt()
        duration = random.randint(5, 240)
        logout = login + timedelta(minutes=duration)
        uninterrupted = random.randint(max(0, duration - random.randint(0, 60)), duration)
        rows.append([
            f"S{i+1:04d}",
            random.choice(user_ids),
            login.strftime("%Y-%m-%d %H:%M:%S"),
            logout.strftime("%Y-%m-%d %H:%M:%S"),
            uninterrupted,
        ])
    return rows

def gen_video_progress(n=220):
    rows = []
    for i in range(n):
        subject = random.choice(subjects)
        course_slug, lesson_slugs = random.choice(courses[subject])
        lesson_slug = random.choice(lesson_slugs)
        video_len = random.randint(300, 3600)
        # 60% chance user watched most of it
        if random.random() < 0.6:
            watched = random.randint(int(video_len * 0.8), video_len)
        else:
            watched = random.randint(0, int(video_len * 0.8))
        is_completed = watched >= video_len * 0.9
        rating = None
        if is_completed and random.random() < 0.7:
            rating = random.choices([1,2,3,4,5], weights=[2,5,15,40,38])[0]
        rows.append([
            f"VP{i+1:04d}",
            random.choice(user_ids),
            random.choice(lesson_ids),
            subject,
            course_slug,
            lesson_slug,
            watched,
            is_completed,
            rating if rating is not None else "",
            rand_dt().strftime("%Y-%m-%d %H:%M:%S"),
        ])
    return rows

def gen_problem_attempts(n=220):
    correct_rate = {"Easy": 0.78, "Medium": 0.55, "Hard": 0.30}
    rows = []
    for i in range(n):
        subject = random.choice(subjects)
        title, topic_tag, difficulty = random.choice(problems_data[subject])
        is_correct = random.random() < correct_rate[difficulty]
        # harder problems take longer
        base = {"Easy": 120, "Medium": 300, "Hard": 600}[difficulty]
        time_taken = int(random.gauss(base, base * 0.4))
        time_taken = max(15, min(time_taken, 3600))
        hints = random.choices([0,1,2,3], weights=[55,25,13,7])[0]
        rows.append([
            f"PA{i+1:04d}",
            random.choice(user_ids),
            random.choice(problem_ids),
            subject,
            title,
            difficulty,
            topic_tag,
            is_correct,
            time_taken,
            hints,
            rand_dt().strftime("%Y-%m-%d %H:%M:%S"),
        ])
    return rows

def gen_quiz_results(n=220):
    rows = []
    for i in range(n):
        subject = random.choice(subjects)
        subsection = random.choice(quiz_subsections[subject])
        score = round(min(100.0, max(0.0, random.gauss(67, 18))), 1)
        time_taken = random.randint(120, 3600)
        rows.append([
            f"QR{i+1:04d}",
            random.choice(user_ids),
            random.choice(quiz_ids),
            subject,
            subsection,
            score,
            time_taken,
            rand_dt().strftime("%Y-%m-%d %H:%M:%S"),
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
    ["session_id","user_id","login_timestamp","logout_timestamp","uninterrupted_minutes"],
    gen_sessions())

write_csv(base / "video_progress.csv",
    ["progress_id","user_id","lesson_id","subject_category","course_slug","lesson_slug",
     "seconds_watched","is_completed","user_rating","timestamp"],
    gen_video_progress())

write_csv(base / "problem_attempts.csv",
    ["attempt_id","user_id","problem_id","subject_category","title","difficulty",
     "topic_tag","is_correct","time_taken_seconds","hints_used","timestamp"],
    gen_problem_attempts())

write_csv(base / "quiz_results.csv",
    ["result_id","user_id","quiz_id","subject_category","subsection",
     "score_percentage","time_taken_seconds","completed_at"],
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

