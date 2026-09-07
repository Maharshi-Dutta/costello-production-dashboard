# The selection wheel: how it moves

Nothing about the wheel pops on or off. It arrives by travelling up from under
the screen edge, opens by gliding its own button in to the middle of the ring
and sending the options out along their radii one after another, and closes by
playing all of that backwards. Every stage is a transform/opacity transition or
keyframe on its own node, so none of it touches layout or paint.

The whole feel is tunable from the stylesheet alone. Every duration, delay,
step and easing is a custom property declared on `:root` in `index.html`, next
to the `.fab*` rules:

| prefix | means | example |
|---|---|---|
| `--fab-t-*` | a duration | `--fab-t-glide: 450ms` |
| `--fab-d-*` | a delay | `--fab-d-home: 200ms` |
| `--fab-s-*` | a per-option step in a stagger | `--fab-s-opt: 70ms` |
| `--fab-e-*` | an easing | `--fab-e-spring: cubic-bezier(.34,1.4,.64,1)` |

`app.js` never writes a duration of its own: `fabMs()` reads the two it has to
wait for (`--fab-t-leave`, and `--fab-d-home` + `--fab-t-glide`) straight back
off `:root` with `getComputedStyle`, and only falls back to a literal when
there is no browser to ask. Change a number below and both the animation and
the JavaScript that waits for it move together.

## Every transition

| stage | node | property | from → to | duration | easing | delay |
|---|---|---|---|---|---|---|
| **Arrive** (first job ticked) | `.fabwrap` (`@keyframes fabrise`, fill `backwards`) | transform | `translateY(120px)` → none | `--fab-t-arrive` 420 ms | `--fab-e-arrive` `cubic-bezier(.34,1.56,.64,1)` | 0 |
| | | opacity | 0 → 1 | 420 ms | same | 0 |
| **Leave** (last job unticked) | `.fabwrap.leaving` | transform | none → `translateY(120px)` | `--fab-t-leave` 300 ms | `--fab-e-leave` `cubic-bezier(.4,0,1,1)` | 0 |
| | | opacity | 1 → 0 | 300 ms | linear | 0 |
| **Open** — toggle glides in | `.fabwrap.open .fab` | transform | `translate(0,0)` → `translate(--fdx,--fdy)` | `--fab-t-glide` 450 ms | `--fab-e-spring` `cubic-bezier(.34,1.4,.64,1)` | 0 |
| **Open** — × turns | `.fabwrap.open .fabx` | transform | `rotate(0)` → `rotate(315deg)` | `--fab-t-x` 500 ms | `--fab-e-spring` | `--fab-d-x` 80 ms |
| | | opacity | 0 → 1 | `--fab-t-fade` 200 ms | ease | 80 ms |
| **Open** — count fades out | `.fabwrap.open .fabn`, `.fabk` | opacity | 1 → 0 | 200 ms | ease | 0 |
| **Open** — inner ring grows | `.fabwrap.open .fabring.in` | transform | `translate(0,0) scale(.2)` → `translate(--fdx,--fdy) scale(1.45)` | `--fab-t-ring-in` 450 ms | `--fab-e-out` `cubic-bezier(.16,.84,.44,1)` | 0 |
| | | opacity | 0 → 1 | 450 ms | ease | 0 |
| **Open** — outer ring grows | `.fabwrap.open .fabring.out` | transform | `translate(0,0) scale(.2)` → `translate(--fdx,--fdy) scale(2.45)` | `--fab-t-ring-out` 550 ms | `--fab-e-out` | `--fab-d-ring-out` 80 ms |
| | | opacity | 0 → 1 | 550 ms | ease | 80 ms |
| **Open** — each option travels out | `.fabwrap.open .fabopt` | transform | `translate(--fdx,--fdy) scale(.3)` → `translate(--fx,--fy) scale(1)` | `--fab-t-opt` 450 ms | `--fab-e-spring` | `--fab-d-opt` 160 ms + `--i` × `--fab-s-opt` 70 ms |
| | | opacity | 0 → 1 | 450 ms | ease-out | same |
| **Close** — options draw back in | `.fabopt` | transform | `translate(--fx,--fy) scale(1)` → `translate(--fdx,--fdy) scale(.3)` | `--fab-t-optback` 320 ms | `--fab-e-in` `cubic-bezier(.4,0,1,1)` | `--rev` × `--fab-s-optback` 45 ms |
| | | opacity | 1 → 0 | 320 ms | ease-in | same |
| **Close** — rings shrink | `.fabring.in` / `.fabring.out` | transform, opacity | back to `scale(.2)`, 0 | 450 / 550 ms | `--fab-e-out` | 0 |
| **Close** — × turns back | `.fabx` | transform | `rotate(315deg)` → `rotate(0)` | 500 ms | `--fab-e-spring` | 0 |
| | | opacity | 1 → 0 | 200 ms | ease | 0 |
| **Close** — toggle glides home | `.fab` | transform | `translate(--fdx,--fdy)` → `translate(0,0)` | 450 ms | `--fab-e-spring` | `--fab-d-home` 200 ms |
| **Close** — count fades back | `.fabn`, `.fabk` | opacity | 0 → 1 | 200 ms | ease | 0 |
| **Count change** | `.fab.pulse .fabn` (`@keyframes fabpulse`) | transform | `scale(1)` → `scale(1.15)` at 45% → `scale(1)` | `--fab-t-pulse` 250 ms | `--fab-e-out` | 0 |

`--fdx` / `--fdy` (how far the toggle glides) and `--fx` / `--fy` (where each
option lands) are the only things JavaScript works out, in `fabGeom()`; they
are px offsets from the toggle's home in the corner, written onto the host and
the buttons as custom properties. `--i` is the option's index and `--rev` is
its index counted from the other end, so opening and closing stagger opposite
ways without a second set of rules.

An option's resting place is `translate(--fdx,--fdy)` — the middle of the ring,
which is exactly where the toggle will be by the time the first option is let
go, 160 ms in. So an option travels *along its own radius* out from under the
button, and on the way back draws in to the button rather than off to an empty
corner the button has not reached yet.

## How the stages line up

Opening, with four options (admin) — everything measured from the tap:

```
   0 ms  toggle starts gliding to the ring's middle      ┐ 450 ms
   0 ms  inner ring starts growing                       ┘ 450 ms
  80 ms  outer ring starts growing                         550 ms → 630 ms
  80 ms  the plus starts turning into a cross              500 ms → 580 ms
 160 ms  option 0 leaves the toggle                        450 ms → 610 ms
 230 ms  option 1 leaves                                             → 680 ms
 300 ms  option 2 leaves                                             → 750 ms
 370 ms  option 3 leaves                                             → 820 ms
```

Closing is the same in reverse — the options go first and the toggle follows
them home, so it never arrives at an empty corner ahead of its own ring:

```
   0 ms  option 3 draws back in (--rev 0)                  320 ms → 320 ms
   0 ms  rings shrink, plus starts turning back            450/550/500 ms
  45 ms  option 2 draws in                                          → 365 ms
  90 ms  option 1                                                   → 410 ms
 135 ms  option 0                                                   → 455 ms
 200 ms  toggle starts gliding home                        450 ms → 650 ms
```

## What the JavaScript waits for

| wait | length | why |
|---|---|---|
| host removed after leaving | `--fab-t-leave` + 60 ms = 360 ms | `fabDrop()` listens for `transitionend` on the host and removes it then; this timer stands behind it for the cases where `transitionend` never comes (the wheel hidden under a window, a background tab, no browser at all). Whichever fires first cancels the other, so a host is never orphaned. |

The `transitionend` listener sits on the host, so **everything inside the
wheel bubbles up through it**. Unticking the last job while the ring is open
means the toggle's glide home, four options drawing in, two rings shrinking and
the cross turning back all finish *before* the host has sunk 120 px, and any
one of them taken for the end of the leave whips the whole wheel off the screen
mid-motion. `fabDrop()` therefore only accepts `e.target === host` with a
`propertyName` of `transform` or `opacity` — the only two the host itself
animates (`.fabwrap.leaving`), and opacity alone under reduced motion.
| geometry re-laid after a close | `--fab-d-home` + `--fab-t-glide` + 40 ms = 690 ms | The ring is never rebuilt under a finger. If the viewport changes shape while the wheel is open, `renderFab()` closes it with its full motion and leaves `data-sig` alone; the new geometry is laid out when everything has come to rest. |
| resize settle | 120 ms | Plain debounce on the `resize` listener, so a drag of a desktop window is one re-lay rather than fifty. |

Ticking a job again while the wheel is on its way out cancels the leave
outright (`fabStay()`): the same host stays on the page, drops its `leaving`
class and rises back, and the removal that was queued behind it never fires.

## Reduced motion

Under `prefers-reduced-motion: reduce` every `.fab*` node keeps exactly where
it goes and stops travelling there: `animation: none`, and a single
`transition: opacity 120ms linear` with no delay. `.fabwrap.leaving` also drops
its `translateY(120px)`, so nothing is left mid-slide. The wheel still fades in
and out and the options still appear at their places on the circle.

## Where it is checked

`test_phases.js` (section 4b) reads `index.html` back and asserts that every
`--fab-*` property used is declared, that each duration and easing above is the
value written here, that no transition in the block names a property other than
transform or opacity, that `will-change` is on every animated node, and — with
the DOM stub — that the rings are built, that opening is one class, that a
changed count adds `pulse`, that the host survives its own leaving animation
and goes on `transitionend` *or* on the timer, that a fast re-tick calls it
back, and that a viewport change while open closes it before re-laying the
ring.

The bubbling case has its own test: the wheel is unticked with the ring open,
then a bubbled `transitionend` is delivered to the host's listener for an
option, for the toggle, for a ring, and for a property the sink does not
animate. The host has to survive all four, and go only on its own
`transform`.
