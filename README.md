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

## Tips for Best Results

### Star Selection and Brightness

- Choose a **bright, isolated star** — the script locks onto the brightest object in the frame, so avoid crowded fields or double stars.
- **Ideal brightness:** Use SharpCap's histogram to check. The peak should land somewhere in the **middle third of the histogram** (roughly 80-180 on a 0-255 scale). This gives the detector plenty of contrast to work with.
- If the histogram peak is bunched up at the far right, the star is saturated — reduce exposure or gain. If it barely rises above the left side, increase exposure or gain.
- Brighter stars (mag 1-3) are easier to work with since they produce a strong, well-defined donut at shorter exposures.

### Defocus Amount

- Defocus the star until the donut is clearly visible with a **dark center and a bright ring**. You should be able to see the secondary mirror shadow distinctly.
- **Ideal donut size:** Aim for roughly **100-300 pixels in diameter**. The overlay info panel shows the detected outer radius — multiply by 2 for diameter.
- Too small (under ~30px diameter) and the script can't resolve the inner/outer edges. Too large and you lose brightness and contrast.
- A good starting point: defocus until the donut is about **1/4 to 1/3 the width of your frame**.

### Using SharpCap Features to Improve Detection

- **Frame stacking / live stacking:** If your donut appears noisy (especially at lower gain or with a dim star), enabling SharpCap's frame stacking can smooth out the noise and give the detector cleaner edges.
- **Planet/disk stabilization:** If seeing conditions cause the donut to bounce around the frame, SharpCap's stabilization feature can help keep it steady, which improves tracking accuracy and reduces jitter in the overlay.
- **ROI (Region of Interest):** Setting a smaller ROI around the star in SharpCap speeds up both capture and analysis, letting the script process frames faster.

## Troubleshooting

### "Searching for donut..." never goes away

- **No star visible**: Ensure a bright star is in the field of view and the camera is capturing.
- **Star not defocused enough**: The star needs to show a clear donut shape with a visible dark center. A focused or slightly defocused star that appears as a bright dot will not be detected.
- **Star too dim**: Increase exposure or gain. The peak brightness must be above 25 (on a 0-255 scale) for detection to trigger. Check the histogram — the peak should be clearly separated from the background noise.
- **Star too small**: The donut must be at least ~30px in diameter. If it's smaller, defocus more or increase focal length/magnification.
- **Donut fills the entire frame**: The script needs some dark background around the donut. If the bright ring extends to the frame edges, the edge detection can't find the outer boundary.

### Star is too dim or too bright

- **Too dim:** The donut appears faint and noisy, or detection keeps dropping out. Increase exposure time or gain. You can also try enabling frame stacking in SharpCap to boost signal.
- **Too bright / saturated:** The histogram peak is pinned to the right edge. The script can't distinguish the donut edges when everything is clipped at max brightness. Reduce exposure or gain until the histogram peak moves away from the right edge.

### Detection is jittery or unstable

- Poor seeing or low signal-to-noise is the most common cause. Try increasing exposure or enabling stabilization.
- If the donut is noisy, frame stacking can help smooth things out.
- You can also increase `SMOOTHING_FACTOR` in the script configuration (see [Configuration](#configuration)) to 0.7-0.85 for more stable tracking.

### Star moves too far or overlay is lost

- If the star drifts out of the detection area or the overlay disappears, **restart the script** by running it again from the scripting console. Re-running automatically cleans up the previous instance and starts fresh detection.
- Consider enabling SharpCap's stabilization to keep the star centered, or use a tracking mount.

### Detection locks onto the wrong feature

- If there are multiple bright objects in the frame, the script will find the brightest one. Center your target star and reduce the field of view or increase zoom if needed.
- Run `reset_tracking()` in the scripting console after repositioning.

### Circles appear but don't match the donut edges well

- **Diffraction rings confusing edges**: If the star shows prominent diffraction rings, the threshold may trip on a ring instead of the true edge. Adjust `BRIGHTNESS_THRESHOLD_PERCENT` up or down in the configuration.
- **Donut is not circular**: If the donut is significantly elongated (astigmatism, tilt), the circle fit will be an approximation. The script assumes circular symmetry.

### Overlay is slow or laggy

- Set a smaller ROI in SharpCap to reduce the area the script needs to analyze.
- Increase `ANALYSIS_INTERVAL` to 3-5 in the script configuration (analyze fewer frames).
- See [Configuration](#configuration) for more performance tuning options.

### Script errors on startup

- Ensure you have SharpCap Pro (scripting is a Pro feature).
- Ensure a camera is connected and live preview is running before executing the script.
- Check the scripting console for error messages; `show_state()` and `diagnose()` can help identify the issue.

### Toolbar button doesn't appear

- The button API varies across SharpCap versions. If the button fails to create, the console will show a warning. Use `cycle_display_mode()` in the console as an alternative.
- Re-running the script automatically removes old buttons before adding new ones.

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

## How It Works

1. Hooks into SharpCap's `BeforeFrameDisplay` event to run on every displayed frame
2. Gets the frame bitmap via `GetFrameBitmap()` and locks pixel data with `LockBits`
3. Finds the donut center using a brightness-weighted centroid
4. Casts 90 radial rays outward from the center, detecting dark-to-bright transitions (inner edge) and bright-to-dark transitions (outer edge)
5. Fits circles to each set of edge points using the Kasa algebraic least-squares method
6. Rejects outlier points and refits for higher accuracy
7. Applies exponential smoothing across frames to stabilize the overlay
8. Draws the overlay onto the frame using GDI+ via SharpCap's `GetDrawableBitmap()` wrapper
