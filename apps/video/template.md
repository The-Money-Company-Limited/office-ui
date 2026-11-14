# Video

A page that plays like a video. You write an ordinary page and a list of what to draw on it; the player plays it with
a clock people can pause and scrub, and people comment on a moment the way they comment on any page. Nothing is
recorded and there is no build step: edit the files and publish.

## The files

| File | What it is | Change it? |
| --- | --- | --- |
| `index.html` | The stage: your content, inside `<div class="camera">`, drawn at 1280 by 720 | Yes: this is your video |
| `video.json` | What to draw, and where the camera goes, in order | Yes |
| `player.js` | The clock, the camera and the controls | Only to change how the player works |
| `player.css` | How the player looks | Only to change how the player looks |
| `artifact.json` | Declares the site a video | No |

## video.json

```json
{
  "by": "Claude",
  "steps": [
    { "marks": [{ "el": "#total", "shape": "ring", "color": "red" }] },
    { "marks": [{ "el": "#total", "write": "up 12% since June", "color": "red" }] },
    { "camera": { "el": "#chart", "zoom": 1.6 } },
    { "clear": true, "camera": { "zoom": 1 } }
  ]
}
```

- `by` is the name on the pen.
- Each step is done in turn. A step without `at` starts when the one before it is done, so most videos need no times.
- `marks` are drawn by the pen, in the same grammar as `cool comments leave --marks`. Each mark is on one element `el`
  (a selector that matches exactly one element) and is one of `"shape": "ring" | "box" | "underline" | "strike" | "cross"`,
  `"write": "<a few words>"` or `"path"`. `"color"` is `amber` (the default), `blue`, `red`, `green`, `purple` or `ink`.
- `camera` zooms in on an element (`"el"`, `"zoom"`), or back out with `{ "zoom": 1 }`. Drawings in the same step wait
  for the camera to land.
- `clear` wipes every drawing made so far, so later drawings are not drawn over old ones.
- `at` (seconds) starts a step at that moment instead. `wait` (seconds) holds before the next step.
- `duration` (seconds, at the top level) sets the length; without it the video ends shortly after its last step.

## Motion in the page

Anything that moves on the page, such as text appearing or a chart growing, must be a CSS animation or a Web
Animation. Its delay is its moment in the video: `animation: reveal 0.6s ease 9.8s both` appears at 9.8 seconds. The
player holds every animation and sets its time, so scrubbing and replaying always show the same picture.

For the same reason, don't use timers, `requestAnimationFrame`, the clock or unseeded randomness to move things, and
don't put `will-change: transform` on the camera. Video and canvas content does not follow the clock.

## How people watch it

- Space or K plays and pauses; arrows or J and L skip 5 seconds (with Shift, one frame); F is full screen.
- While it plays, the controls fade when the pointer rests.
- Comments: press C or the Comment button. The video pauses, and the comment is left at that moment. Its pin shows near
  that moment, and opening it takes the video back there. Your agent reads the moment as `anchor.moment_ms` and
  `cool comments` prints it as `at 0:22 in the video`.

The drawings are drawn by the Cool Computers comment script. `artifact.json` declares the site a video, so they show
even when the site's comments are switched off. Keep it, and keep `video.json` at the site's root.
