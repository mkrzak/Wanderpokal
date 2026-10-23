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

Night placings follow the match sheet: most set wins first, then most total points.

Shared live scoring, where every team enters scores and everyone sees them at once, runs on the Claude version: https://claude.ai/artifact/7gg65naa5e6mFXPTKdBjkK (Contributor access needed to enter scores). On the GitHub site, scores are saved in the browser of the device used to type them.

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
