# Radial selection menu (spec, 2026-09-07) - replaces the quarter-circle fan

Dashboard-only; no workbook writes; the four actions are unchanged in behaviour.
Reference: "Animated Circular Navigation Menu" (Online Tutorials, Simple Radial Menu 02):
a round central toggle with a plus that rotates 315° into an ×, two concentric ring
shadows rippling out behind it on open, and round icon buttons fading/scaling in around a
FULL circle with staggered delays (reverse order on close), each icon kept upright.

## Behaviour
- Appears only while jobs are ticked (`state.picked`), as now; the toggle shows the count.
- **Closed:** the 56 px round toggle sits at the bottom-right corner (18 px in), as now.
- **Open:** the toggle glides inwards to the ring centre (about 150 px from the right and
  bottom edges; on screens under 400 px wide use 130 px) and the four options appear on a
  full circle of radius 100 px around it, evenly spaced (top, right, bottom, left, starting
  at the top): **Alert to…** (bell icon; admin only - with three options use 120° spacing),
  **Export** (download icon), **Move to…** (arrows icon), **Untick all** (tick-square icon;
  this replaces the "Clear" wording). Each option: 52 px round button with an inline SVG
  icon and a small text label under it (never icon alone). Icons stay upright.
- Animation: toggle plus rotates 315° on open; box-shadow rings (inner dark ring 12 px,
  outer accent ring to ~40 px) expand on open; options fade + scale from 0 with
  `transition-delay: calc(var(--i) * 60ms)`, reversed on close; the glide is a 250 ms
  transform transition. `prefers-reduced-motion`: no transitions.
- Tap outside, Escape, or the toggle closes it (toggle glides back). Choosing an option
  runs exactly what it runs today (renderAlertMenu / Export window with scope "ticked" /
  renderMoveMenu / clear selection) and closes the ring.
- Hidden while a window or the drawer is open at phone width (existing `over` rule).
- Theme-aware with existing variables; both themes; the ring must stay fully on screen at
  320 px wide (compute the centre offset from the viewport).

## Tests
Update `test_phases.js`: the four (or three) options sit on a full circle at radius 100
around the open centre; the open centre keeps every option ≥ 8 px inside the viewport at
1280, 390 and 320 px; "Untick all" clears the selection; the rest of the existing wheel
tests keep passing. Keep every suite green.
