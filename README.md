# Wanderpokal

Results, stats and live scoring for Wanderpokal, the Tuesday-night three-team volleyball session, compiled by Monika Krzak.

Live site: https://mkrzak.github.io/Wanderpokal/

One page (`index.html`, with team photos in `photos/`) with six sections:

* **Live:** tonight's schedule of nine 15-minute sets (19:20–21:51, after a 20-minute warm-up), with score entry. Winner, difference, total points, total wins, final standings and points needed to top the total update as you type.
* **Players:** all-time leaderboard for regular Tuesdays (nights, titles, win %, 2nd/3rd places, sets won/lost, points per set, +/−), with a profile for each player.
* **Results:** every regular Tuesday since November 2024, with set scores and team rosters.
* **Records:** highest set score, biggest margin, most points in a night, one-point sets, perfect 6–0 nights and titles by team letter.
* **Specials:** birthdays, cups and other one-off nights (Bryan Cup, Banana vs Wanderpokal, 5-1 Specials, …), each on its own card with winners, standings and set scores, plus a separate specials leaderboard. Specials are kept out of the weekly stats in the other sections.
* **Photos:** the trophy photo from each Tuesday and each special, newest first, in two separate galleries.

## Adding photos

**From the website (any phone or laptop):** every night has its own **+ Add photo** button: under each week in **Results**, on each card in **Specials** and on each card on the **Photos** page. Press it or drag photos onto that night.

1. The first time, the site asks for a GitHub fine-grained token (Repository access: only `mkrzak/Wanderpokal`; Permissions: Contents, Read and write). It is saved in that browser only; **turn off** at the top of the Photos page removes it.
2. Photos are resized (1600 px, plus a thumbnail) and committed straight to the repo, with `data/photos.json` updated in the same commit. Everyone sees them within a couple of minutes; no rebuild needed.
3. On a night with several photos the first one is the main photo (used as the cover): ★ makes another one main. Photos can't be removed from the website; to remove one, delete it from `photos/` (and `photos/thumbs/`) and from `data/photos.json`, then push.

**After uploading photos from the website**, each upload is a new commit on GitHub that your computer doesn't have yet. Before you push anything else (or rebuild from the spreadsheet), bring them in first:

```bash
git pull --rebase origin main
git push origin main
```

Commit your own changes before `git pull --rebase`; otherwise git refuses to pull.

**From your computer:** put the photo in `photos/` named `<date>_week<N>.jpg` (e.g. `2026-10-27_week90.jpg`) or `<date>_special.jpg`; add `_2`, `_3` for more photos of the same night. The next `python3 scripts/build.py ../Wanderpokal.xlsx` resizes it, makes the thumbnail and adds it to `data/photos.json` (needs `pip install pillow` once).

Night placings follow the match sheet: most set wins first, then most total points.

Live scores are shared through this repo: nights are saved in `data/live.json`, and every visitor's page reads it every 30 seconds (GitHub Pages publishes a save in about 1–2 minutes).

* **Scorekeepers:** in Live, unlock with the scorekeeper password, create the night or pick it, type the teams and press **Save teams**, enter scores and press **Save scores** (as often as you like), then **Finish night**. Finishing saves straight away and the night counts in Players, Results, Records and Stats for everyone.
* **One-time setup (Monika):** on a browser where photo uploads are on, unlock Live and press **Create scorekeeper key**. It stores your GitHub token in `data/scorekey.json`, locked with the scorekeeper password, so any device can save with just the password. Anyone with the password can write to this repo; keep the token limited to this repository, and if the password leaks, revoke the token on GitHub and create a new key.
* A finished live night is counted only until the same night (same date and week) appears in the spreadsheet, so rebuilding never counts it twice.

## Updating after a Tuesday

1. Fill in the week's sheet in `Wanderpokal.xlsx` as usual.
2. Rebuild the site from the spreadsheet (needs `pip install openpyxl` once):

   ```bash
   python3 scripts/build.py ../Wanderpokal.xlsx
   ```

3. Commit and push; the live site updates within a couple of minutes.

The script reads every weekly sheet (`89 - 13102026`) and special night (`Bryan Cup - 2810205`, `5-1 Special - 20102026`, …). Weeks 1–16 and 27 come from the **WinnersList & Stats** sheet and only record the winning team, so they count towards titles but not win %. To merge two spellings of one player, add them to `ALIASES` in `scripts/build.py` and rebuild.

## Files

* `index.html`: the built site (generated; edit `src/template.html` instead)
* `src/template.html`: page layout, styles and scripts
* `data/wanderpokal.json`: every night's rosters and set scores
* `scripts/build.py`: rebuilds `data/` and `index.html` from the spreadsheet
* `photos/` and `data/photos.json`: web-sized team photos (metadata removed) and which night each belongs to
