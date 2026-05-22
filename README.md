# Workout Coach

A single-file HTML/JS workout tracker built around **antagonist supersets** and **muscle-group recovery tracking**. Picks today's workout based on which muscles are most recovered, logs your sets/reps/weight, and stores everything in browser localStorage with JSON export/import.

## Features

- **4 splits**: Push, Pull, Legs, Full Body
- **5 supersets per workout** (4 for Full Body), each paired A1/A2 with minimum-overlap antagonist logic
- **Recovery-aware recommendation**: tracks days-since-last-trained per muscle group and suggests whichever split is most recovered
- **Per-set logging**: reps + weight, with prefill from defaults
- **History tab**: view and delete past workouts
- **Recovery tab**: color-coded muscle map (fresh / recovering / hit today)
- **Export / Import JSON**: your real source of truth is a JSON file you control

## Splits & supersets

### Push (5 supersets)
- A. Flat Barbell Bench Press / Lateral Raises
- B. Incline Dumbbell Press / Tricep Pushdowns
- C. Shoulder Press / Cable Fly
- D. Dips / Overhead Cable Fly
- E. Push-ups / Tricep Extensions

### Pull (5 supersets)
- A. Pull-ups / Hammer Curls
- B. Barbell Row / Preacher Curls
- C. Lat Pulldown / Face Pulls
- D. Seated Cable Row / Single Arm Bicep Curls
- E. T-Bar Row / Reverse Dumbbell Fly

### Legs (5 supersets)
- A. Squats / Romanian Deadlift
- B. Leg Press / Hip Thrust
- C. Bulgarian Split Squats / Leg Curls
- D. Leg Extensions / Nordic Curls
- E. Standing Calf Raises / Seated Calf Raises

### Full Body (4 supersets)
- A. Squats / Pull-ups
- B. Deadlifts / Planks
- C. Cable Crunch / Pallof Press
- D. Decline Situps / Leg Raises

## Run it

Just open `index.html` in a browser. No build step, no dependencies.

```bash
# or serve locally
python3 -m http.server 8123
# then open http://localhost:8123
```

## Data

State lives in `localStorage` under the key `workout-coach-v1`. Use the **Export** button periodically to download a JSON backup. **Import** to load one back.

## How the recommendation works

Each muscle group tracks days-since-last-trained. Each split is scored by the average days-since across its muscles, with a penalty if any muscle in the split was hit < 1.5 days ago. The highest-scoring split is suggested.
