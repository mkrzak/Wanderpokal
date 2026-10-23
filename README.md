# Wanderpokal

Stats and live scoring for Wanderpokal, our Tuesday-night three-team volleyball session.
Teams A, B and C play nine 15-minute sets from 19:20 to 21:51, after a 20-minute warm-up.

**🏐 Live site (shared scoring): [claude.ai/artifact/7gg65naa5e6mFXPTKdBjkK](https://claude.ai/artifact/7gg65naa5e6mFXPTKdBjkK)**
Every team can enter scores here and everyone sees them live. You need to be signed in to Claude and given Contributor access to enter scores.

**Stats site (GitHub Pages):** `https://<your-account>.github.io/Wanderpokal/` (live once Pages is switched on, see below)

The site has four sections:

| Section | What it shows |
|---|---|
| **Live** | Tonight's schedule. Type in set scores and the winner, difference, total points, total wins and 1st/2nd/3rd update as you go. Each team card shows how many points it needs to top the points total. |
| **Players** | All-time leaderboard: nights played, titles, win %, 2nd/3rd places, sets won/lost, points per set, +/−. Click a player for their profile. |
| **Results** | Every night since November 2024, with set scores and rosters. |
| **Records** | Highest set score, biggest margin, most points in a night, one-point sets, perfect 6–0 nights and titles by team letter. |

Night placings follow the spreadsheet: most set wins first, then most total points.

## Project layout

```
Wanderpokal/
├── index.html            # the built site (open it directly or serve it with GitHub Pages)
├── src/template.html     # page layout, styles and scripts; edit this, not index.html
├── data/wanderpokal.json # every night: rosters and set scores, generated from the spreadsheet
├── scripts/build.py      # reads Wanderpokal.xlsx and rebuilds data/ and index.html
├── .nojekyll             # tells GitHub Pages to serve files as they are
└── .gitignore
```

## Updating the stats after a Tuesday

1. Fill in the new week's sheet in `Wanderpokal.xlsx` as usual.
2. Rebuild:

   ```bash
   pip install openpyxl          # first time only
   python3 scripts/build.py ../Wanderpokal.xlsx
   ```

3. Commit and push. GitHub Pages republishes within a minute or two.

```bash
git add index.html data/wanderpokal.json
git commit -m "Week 90 results"
git push
```

The script reads every sheet named like `89 - 13102026` and every special night (`Bryan Cup - 2810205`, `5-1 Special - 20102026`, …).
Weeks 1–16 and 27 come from the **WinnersList & Stats** sheet. Those weeks only record the winning team, so they count towards titles but not win %.

To merge two spellings of the same player, add them to `ALIASES` at the top of `scripts/build.py` and rebuild.

## Putting it on GitHub

```bash
cd Wanderpokal
git init
git add .
git commit -m "Wanderpokal site"
git branch -M main
git remote add origin https://github.com/<your-account>/Wanderpokal.git
git push -u origin main
```

Then on GitHub: **Settings → Pages → Build and deployment → Deploy from a branch → `main` / `(root)`**.
The site will appear at `https://<your-account>.github.io/Wanderpokal/`.

**Privacy:** the site lists players' first names and results. A public repository makes them visible to anyone.
Ask the group first, or keep the repository private. GitHub Pages on a private repository needs a paid GitHub plan.
`.gitignore` keeps `*.xlsx` out of the repository because the registration sheets hold more personal details.

## Live scoring: two versions

- **This GitHub version** saves scores in the browser of the device you type on. Use one phone or laptop as the scorekeeper.
  Other people's devices won't see those scores. Finishing a night adds it to that device's stats until you add it to the spreadsheet and rebuild.
- **The Claude version** of this page is shared. Anyone given Contributor access can enter scores, and everyone sees them live.

To get shared live scoring on GitHub Pages as well, connect the `makeLocalDb()` adapter in `src/template.html` to a hosted database (for example Firebase or Supabase).
The rest of the page already uses that adapter's `collection()/doc()/onSnapshot()` interface.
