# Study Café Analyzer — DESIGN.md

Direction: **Reading room.** The tool is a feasibility report you can operate, so it
should feel like a quiet library desk, not a neon dashboard: pale paper, dark ink,
one lamp-green for anything you can press, and colour reserved for the charts.

This direction is based on the product and its audience (course assessors and
portfolio visitors reading an analysis), not on a survey of other sites.

## What the screen is for

- **Job:** let a reader move through a five-part analysis and trust the numbers.
- **Reader:** someone assessing the work, reading on a laptop, sometimes a phone.
- **Must be able to:** always see where they are and get to any section in one click.

## Colour

| Token | Value | Role |
|---|---|---|
| `--bg` paper | `#F5F7F4` | Page background |
| `--panel` / `--card` sheet | `#FFFFFF` | Cards, tables, inputs, chart surface |
| `--text` ink | `#17231F` | Body text and all key figures |
| `--muted2` slate | `#55635D` | Secondary text, axis labels |
| `--muted` | `#5F6C66` | Tertiary text, captions |
| `--border` rule | `#DCE3DD` | Hairlines and card borders |
| `--border2` | `#C5D0C8` | Input borders, stronger dividers |
| `--accent` lamp green | `#1F5C4A` | Primary: sidebar, buttons, links, active tab, focus ring |
| `--accent3` brass | `#8F6212` | The one accent: the current or highlighted item |
| `--accent2` good | `#1F7A4D` | Positive status only |
| `--warn` | `#8A5A00` | Caution status only |
| `--danger` | `#B3261E` | Negative status and errors only |

Rules: key figures are ink, never rainbow. A chart with one series uses one colour (blue); colour only separates different series. Status colours appear only when the
value is a status (profit or loss, pass or fail). No gradients, no glows.

## Chart colour

Series colours are separate from the interface colours and used in this fixed
order, never cycled or reassigned:

`#2A78D6` blue · `#EB6834` orange · `#1BAF7A` aqua · `#EDA100` yellow ·
`#E87BA4` magenta · `#008300` green · `#6250D6` violet · `#E34948` red

Validated for colour-blind separation on the white chart surface. Aqua, yellow
and magenta are below 3:1 against white, so every chart with them keeps a
legend and hover values. Positive and negative values use aqua and red.
Axis text is slate, grid lines are a 7% ink tint.

## Typography

- **Headings and key figures:** Bricolage Grotesque, 600.
- **Everything else:** Public Sans, 400 and 600.
- **Scale:** 12 · 13 · 14 · 16 · 20 · 26 · 34 px. Body 14–16 px, line height 1.5.
- Numbers use tabular figures. Sentence case everywhere; no all-caps labels.
- Fallback when fonts cannot load: system UI sans.

## Layout and spacing

- **Shell:** fixed left sidebar (232 px, lamp green) with the five sections in
  reading order and the two utilities at the bottom. Content sits to its right,
  max width 1240 px, left aligned.
- **Phone (under 860 px):** the sidebar becomes a top bar that scrolls sideways.
- **Spacing scale:** 4 · 8 · 12 · 16 · 24 · 32 · 48 px.
- **Key figures:** a ledger row on a single sheet, figures divided by hairlines,
  not a row of separate coloured boxes. This is the one bold element.

## Shape and depth

- Radius: 6 px on cards and inputs, 4 px on small chips. No pills except tabs.
- No shadows on cards. Modals get one soft shadow.
- No entrance animation and no hover lift. Hover changes border or background only.

## Components

- **Buttons:** primary is solid lamp green with white text; secondary is white
  with a rule border; destructive is white with red text. Focus shows a 2 px
  lamp-green ring.
- **Tabs:** underlined text tabs; the active tab is ink with a lamp-green underline.
- **Tables:** white sheet, hairline row dividers, slate header text, numbers right
  of labels in tabular figures. No zebra stripes.
- **Cards:** white sheet, 1 px rule, 20 px padding, title 14 px semibold.
- **Inputs:** white, 1 px `--border2`, 6 px radius, label above in slate.
- **Icons:** none. Emoji are removed; status uses words and the ✅ ❌ ⚠ marks only.

## Copy

Sentence case. Buttons say what happens ("Run scenario", "Download report").
No decorative dashes or emoji in headings.

## Themes

Two themes share every rule above; only the token values change. The reader
switches with the Light / Dark control (sidebar, or top right on the upload
screen). The choice is remembered in the browser; with no saved choice the
page follows the device setting.

| Token | Light | Dark |
|---|---|---|
| `--bg` | `#F5F7F4` | `#101614` |
| `--panel` / `--card` | `#FFFFFF` | `#18211E` |
| `--text` | `#17231F` | `#E9EFEA` |
| `--muted2` | `#55635D` | `#A9B5AE` |
| `--muted` | `#5F6C66` | `#8C9992` |
| `--border` | `#DCE3DD` | `#2A3631` |
| `--border2` | `#C5D0C8` | `#3B4B44` |
| `--accent` | `#1F5C4A` | `#6FCFAE` |
| `--on-accent` (text on accent) | `#FFFFFF` | `#0F1513` |
| `--accent3` brass | `#8F6212` | `#D9A441` |
| `--accent2` good | `#1F7A4D` | `#5CC08A` |
| `--warn` | `#8A5A00` | `#E0B25A` |
| `--danger` | `#B3261E` | `#F0766C` |
| `--nav-bg` sidebar | `#1F5C4A` | `#13302A` |

Dark chart series, same fixed order, validated on the dark surface (all at
least 3:1): `#3987E5` · `#D95926` · `#199E70` · `#C98500` · `#D55181` ·
`#008300` · `#9085E9` · `#E66767`.

Never write a raw colour in a component: use a token, so both themes follow.

## Reports (v45)

Downloaded reports use the light theme only, whatever the dashboard is showing.

- **Analysis report** (`study_cafe_report.pdf` / `.html`): Data summary, Exploratory analysis, Predictive analytics. No what-if content.
- **What-if scenario report** (`study_cafe_whatif_report.pdf` / `.html`): downloaded from the What-if section after running a scenario. Lists the inputs used, then every tab.
- Cover band in lamp green `#1F5C4A`, ink `#17231F` text on white, hairlines `#DCE3DD`, brass `#8F6212` rule on notes. Key figures are ledger rows, tables are real text, charts sit two to a row unless wide.
- PDF uses Helvetica (built into every PDF reader); the rupee sign is written "Rs" because that font has no glyph for it. Tables wider than nine columns are left to the HTML report.
- HTML report redraws charts live with Chart.js when online and falls back to still images offline. The 3D cluster plot is rotatable there (drag, arrow keys, +/-, double-click to reset).

## Interactions (v48)

- **Card tools** sit in the top-right corner of a chart card (Reset, Reset zoom, Compare). Same button style everywhere: 1px `--border2` outline, `--panel` fill, 12.5px semibold.
- **Drill-down chips** (pill buttons) list a chart's x-axis values; the selected chip fills with `--accent`. The breakdown is a horizontal bar chart with its value axis on top and a label at the end of each bar. Colours keep each group's identity from the original chart.
- **Rule callout**: a panel beside the selected point (below it on narrow screens) with an arrow to the point; everything else fades to 12% opacity.
- **Radar callouts**: a line from each corner to its value, in the group's own colour, with a surface-coloured halo behind the number.
- **Heatmaps** use stepped colours only: correlation has three shades per direction plus grey; the rule heatmap has four greens cut at the quartiles of support. A key sits under each.
- **Lever notes** (amber, `--warn` rule) explain when a chart cannot or did not change.
- Motion: drill-downs cross-fade (180ms) and bars grow (650ms); all of it is skipped when the reader prefers reduced motion. Every interaction is cleared before a report is exported.

## Views (v53)
- Two levels of detail, switched in the sidebar like the theme: **Surface** (default) and **Deep**. The choice is remembered.
- Surface shows what an owner acts on: what students want, layout and hours patterns, customer groups (radar), "Rules worth acting on", the levers and forecasts, membership drivers, the membership predictor and What-if.
- Deep adds the method detail: correlation, regression, PCA and elbow, model comparison, confusion matrix, all rule views and the data summary.
- "Rules worth acting on": significant rules with lift 1.2 or more, ranked by students beyond chance, at most two per outcome, each with a plain action for the café.
- Reports follow the current view and say which one in the header.

## Start screen (v54)
- The app ships with the survey file. A fresh start shows it as an Excel icon above the drop box; dragging it in (or focusing it and pressing Enter) starts the run.
- "Analyse a new dataset" returns to the plain drop box with no icon; click it to choose a file.

## Chart colour (v55)
- Categorical palette, light: `#2b6ea3` steel blue, `#d2692c` burnt orange, `#2f8f6a` green, `#cfae2a` mustard, `#a8508e` plum, `#6a86d8` periwinkle, `#c4534a` brick, `#7f9a24` olive. Dark: `#5495cf #c6610a #47a67d #967400 #a34991 #5a7bd8 #d96d67 #657905`. Both pass the colour-blindness validator (adjacent pairs) on their card colours.
- A category keeps one colour on every chart: Morning mustard, Afternoon orange, Evening plum, Late Night steel blue; seat types fixed; Yes green, No steel blue.
- Ordered scales (spend tiers, visit frequency) use one green ramp, light to dark.
- Colour that carries meaning: green for positive and income, brick for negative and cost.
- Single-series charts use steel blue. Numbers are written in the text colour, never the series colour.

## Words and layout (v55)
- One product name everywhere: "Study café feasibility".
- Headings, chart titles and form labels in sentence case; words with inner capitals (PCA, K-Means, Wi-Fi) are kept.
- What-if settings in four groups: Market, Capacity and hours, Prices and membership, Costs.
- Buttons and menus at least 36 px tall, 44 px on touch screens.

## Pricing (v56)
- Exploratory analysis has a "Pricing and willingness to pay" section in both views, between Customer groups and Association rules.
- Demand curves come from stated maximums: everyone whose band starts at or above a price counts as still willing at that price. Only students who want a membership are used for the membership curve.
- Every share carries a 95% Wilson range; elasticities are midpoint (arc) elasticities with 95% ranges from 1,000 bootstrap resamples.
- Elasticity words: above −0.8 "few drop out, revenue rises"; −0.8 to −1.2 "revenue stays about the same"; below −1.2 "many drop out, revenue falls".
- Room to raise prices compares budget band tops with spend band tops; students with both answers in the open top band are shown separately as "can't tell".
- Willingness to pay by group uses the fee nearest the 50% point, with chi-square tests; groups under 30 students are left out.
- What-if Pricing tab: profit for the year, price change against catchment, seat occupancy or seats, from the same model as the rest of What-if. Shading is by rank: green profit, brick loss.
- The section carries a fixed note that these are stated answers, not observed purchases.

## Reports in plain words (v57)
- The analysis report is offered at two levels in the download dialog, whichever view is on screen: Surface (what matters for running the café) and Deep (the full analysis). Files: study_cafe_report_surface / study_cafe_report_deep.
- The Surface report leaves out the Data summary and the technical summary tiles.
- Both open with "Key findings in plain words", worked out from the data each time: who the results describe, when students come, seats, food and drink, membership interest and fee, the price of a visit, room to raise prices, who pays more, customer groups, a pattern worth acting on, visits and takings. The Deep report adds how far to trust the forecasts and how trustworthy the answers are.
- Every section and sub-section starts with one plain sentence on what it tells you.
- Data-reliability checks get plain titles and explanations; the model scorecard gets an "In plain words" column; rule and profile tables use plain headings.
- Both end with "Words used in this report": only the technical words that actually appear in that report.
- PDFs print "R-squared" and a plain hyphen for minus signs, which the PDF font lacks.
