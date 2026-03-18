# SharpCap Collimation Overlay

An IronPython script for SharpCap Pro that automatically detects a defocused star's donut pattern and draws overlay circles to aid collimation of reflector telescopes. It measures the offset between the outer ring (primary mirror) and inner shadow (secondary mirror) in real time, showing you exactly how far off collimation is and which direction to correct.

## Requirements

- SharpCap Pro 4.1+ (scripting requires Pro license)
- A connected camera with live preview running
- A bright, defocused star visible in the field of view

## Quick Start

1. Connect your camera in SharpCap and start live preview
2. Point at a bright star and defocus until you see a clear donut shape
3. Open the scripting console: **Tools > Scripting Console**
4. In the console: **File > Run Script** and select `collimation_overlay.py`
5. The overlay appears immediately. Click the **Coll Overlay** toolbar button to cycle display modes.

## Display Modes

Click the **Coll Overlay** toolbar button to cycle through these modes:

| # | Mode | What's Shown |
|---|------|-------------|
| 1 | **All overlays** | Info panel, bottom bar, crosshairs, offset line, correction arrows, circles |
| 2 | **Hide info panel** | Bottom bar, crosshairs, offset line, correction arrows, circles |
| 3 | **Hide info panel + bottom bar** | Crosshairs, offset line, correction arrows, circles |
| 4 | **Circles & arrows only** | Just the two detected circles and correction arrows |
| 5 | **Alignment crosshairs** | Both circles with full-diameter crosshairs in matching colors (red for outer, green for inner) for visual alignment |
| 6 | **Off** | No overlay drawn. Cycles back to mode 1 on next click. |

You can also type `cycle_display_mode()` in the scripting console.

## Reading the Overlay

| Element | Meaning |
|---------|---------|
| **Red circle** | Outer edge of the donut (primary mirror light cone) |
| **Green circle** | Inner edge of the donut (secondary mirror shadow) |
| **Yellow crosshair** | Center of the outer circle |
| **Green crosshair** | Center of the inner circle |
| **Cyan dashed line** | Offset vector between the two centers |
| **Cyan arrows** | Direction the inner circle needs to move to become concentric |
| **Bottom bar** | X and Y offset with correction direction and color-coded severity |
| **Info panel** | Full stats: offset in px/%, direction, radii, quality rating |

### Quality Ratings

| Rating | Offset | Meaning |
|--------|--------|---------|
| EXCELLENT | < 2% | Nearly concentric |
| GOOD | < 5% | Minor offset |
| FAIR | < 10% | Noticeable offset |
| POOR | < 20% | Significant offset |
| BAD | > 20% | Major collimation needed |

Offset percentage is relative to the outer circle radius.

## Console Commands

Type these in the SharpCap scripting console:

| Command | Action |
|---------|--------|
| `cycle_display_mode()` | Cycle to the next display mode |
| `reset_tracking()` | Force re-detection of the donut |
| `show_state()` | Print current detection state and errors |
| `diagnose()` | Run detailed pixel access diagnostics |
| `stop()` | Fully stop overlay, remove toolbar button |

Re-running the script automatically cleans up the previous instance (handler, buttons).

## Configuration

Edit the `USER CONFIGURATION` section at the top of `collimation_overlay.py`:

**Colors** (RGBA, 0-255):
- `OUTER_CIRCLE_COLOR` — outer circle, default red
- `INNER_CIRCLE_COLOR` — inner circle, default green
- `CROSSHAIR_COLOR` — outer center crosshair, default yellow
- `OFFSET_LINE_COLOR` — offset line and arrows, default cyan

**Text sizes** (points):
- `TEXT_FONT_SIZE` — info panel (default 9)
- `BOTTOM_BAR_FONT_SIZE` — bottom bar (default 10)
- `ARROW_LABEL_FONT_SIZE` — correction arrow labels (default 8)

**Line widths**:
- `CIRCLE_PEN_WIDTH` — circle outlines (default 2.0)
- `CROSSHAIR_PEN_WIDTH` — crosshairs (default 1.0)
- `OFFSET_PEN_WIDTH` — offset line and arrows (default 2.0)

**Detection tuning**:
- `BRIGHTNESS_THRESHOLD_PERCENT` — % of peak brightness that defines the donut edge (default 30)
- `NUM_RAYS` — radial rays cast for edge detection; more = accurate, slower (default 90)
- `MIN_DONUT_RADIUS` — ignore detected circles smaller than this in pixels (default 15)
- `CENTROID_DOWNSAMPLE` — sample every Nth pixel when finding the donut center (default 4)

**Performance**:
- `ANALYSIS_INTERVAL` — run detection every Nth frame (default 2)
- `SMOOTHING_FACTOR` — exponential smoothing, 0.0 = none, 0.9 = very smooth (default 0.6)

## Troubleshooting

### "Searching for donut..." never goes away

- **No star visible**: Ensure a bright star is in the field of view and the camera is capturing.
- **Star not defocused enough**: The star needs to show a clear donut shape with a visible dark center. A focused or slightly defocused star that appears as a bright dot will not be detected.
- **Star too dim**: Increase exposure or gain. The peak brightness must be above 25 (on a 0-255 scale) for detection to trigger.
- **Star too small**: The donut must be at least ~30px in diameter. If it's smaller, defocus more or increase focal length/magnification.
- **Donut fills the entire frame**: The script needs some dark background around the donut. If the bright ring extends to the frame edges, the ray casting can't find the outer boundary.
- **Threshold too high/low**: If the donut has very low contrast (faint ring against a noisy background), try lowering `BRIGHTNESS_THRESHOLD_PERCENT` to 15-20. If there's a lot of background glow, try raising it to 40-50.

### Detection is jittery or unstable

- Increase `SMOOTHING_FACTOR` to 0.7-0.85 for more stable tracking.
- Increase `NUM_RAYS` from 90 to 120+ for more edge points (at slight performance cost).
- Ensure the image has good signal-to-noise ratio; increase exposure if the donut is noisy.

### Detection locks onto the wrong feature

- If there are multiple bright objects in the frame, the script will find the brightest one. Center your target star and reduce the field of view or increase zoom if needed.
- Run `reset_tracking()` after repositioning.

### Circles appear but don't match the donut edges well

- **Diffraction rings confusing edges**: If the star shows prominent diffraction rings, the threshold may trip on a ring instead of the true edge. Adjust `BRIGHTNESS_THRESHOLD_PERCENT` up or down.
- **Donut is not circular**: If the donut is significantly elongated (astigmatism, tilt), the circle fit will be an approximation. The script assumes circular symmetry.

### Overlay is slow or laggy

- Increase `ANALYSIS_INTERVAL` to 3-5 (analyze fewer frames).
- Increase `CENTROID_DOWNSAMPLE` to 6-8.
- Reduce `NUM_RAYS` to 60.
- The drawing phase runs every frame regardless and is fast; the analysis phase is where computation happens.

### Script errors on startup

- Ensure you have SharpCap Pro (scripting is a Pro feature).
- Ensure a camera is connected and live preview is running before executing the script.
- Check the scripting console for error messages; `show_state()` and `diagnose()` can help identify the issue.

### Toolbar button doesn't appear

- The button API varies across SharpCap versions. If the button fails to create, the console will show a warning. Use `cycle_display_mode()` in the console as an alternative.
- Re-running the script automatically removes old buttons before adding new ones.

## How It Works

1. Hooks into SharpCap's `BeforeFrameDisplay` event to run on every displayed frame
2. Gets the frame bitmap via `GetFrameBitmap()` and locks pixel data with `LockBits`
3. Finds the donut center using a brightness-weighted centroid
4. Casts 90 radial rays outward from the center, detecting dark-to-bright transitions (inner edge) and bright-to-dark transitions (outer edge)
5. Fits circles to each set of edge points using the Kasa algebraic least-squares method
6. Rejects outlier points and refits for higher accuracy
7. Applies exponential smoothing across frames to stabilize the overlay
8. Draws the overlay onto the frame using GDI+ via SharpCap's `GetDrawableBitmap()` wrapper
