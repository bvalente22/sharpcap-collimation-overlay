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
5. The settings palette opens automatically and the overlay begins detecting

To stop the overlay, close the settings palette (X button or **Close** button). To restart, simply re-run the script.

## Optimizing the Donut for Accurate Detection

Getting the best results requires a well-sized, well-exposed donut. The overlay provides two real-time guides to help you dial in the optimal settings.

### Step 1: Get the Right Brightness

Adjust your camera's **exposure** and **gain** until the **Target Brightness** bar (horizontal, above the bottom bar) shows the marker in the **green zone** (26–90%). Aim for the middle of the green zone for best results.

| Zone | Range | Color | What to Do |
|------|-------|-------|------------|
| Too dim | 0–26% | Red | Increase exposure or gain — weak signal means noisy edges |
| OK | 26–90% | Green | Good range — aim for the middle |
| Too bright | 90–100% | Red | Reduce exposure or gain — bloom and clipping distort edges |

### Step 2: Get the Right Size

The **Target Size** bar (vertical, right edge) shows how much of the frame the donut fills. Adjust your **focuser** (defocus more or less) until the marker lands in the **green zone** (20–60% fill).

| Zone | Fill % | Color | What to Do |
|------|--------|-------|------------|
| Too small | 0–15% | Red | Defocus more — sub-pixel errors dominate at this size |
| Marginal | 15–20% | Yellow | Workable but not ideal |
| Ideal | 20–60% | Green | Best accuracy for detection |
| Marginal | 60–70% | Yellow | Getting large — risk of edge clipping |
| Too large | 70–100% | Red | Focus inward — outer edge may clip at frame boundary |

The fill percentage is the outer circle diameter divided by the shorter frame dimension. You can also set a smaller **ROI (Capture Area)** in SharpCap's camera settings to increase the relative fill without changing focus.

### Step 3: Verify Detection

Once both bars are green, check that the red (outer) and green (inner) overlay circles match the donut edges visually. If the outer circle doesn't look right, see [Choosing an Algorithm](#choosing-an-algorithm) below.

Both indicator bars can be toggled on/off from the **Display** tab in the settings palette.

## Choosing an Algorithm

The script offers three outer edge detection algorithms, selectable from the **Detection** tab. Each has strengths for different donut shapes. The inner edge always uses a gradient-based approach that works well across all conditions.

| Algorithm | Best For | How It Works |
|-----------|----------|--------------|
| **Threshold Crossing** (default) | Most images — reliable and stable | Finds where brightness crosses below the threshold. Works consistently across different brightness levels and ring patterns. Try this first. |
| **Edge to Dark** | Prominent diffraction rings | Finds the steepest brightness drop that transitions to darkness. Ignores dips between bright rings by requiring the "after" side to be below threshold. |
| **Steepest Gradient** | Clean donuts, dim stars | Finds the sharpest brightness drop along each ray. Fast and precise for clean edges, but can latch onto diffraction ring edges on complex donuts. |

### When to Switch Algorithms

- **Outer circle is wobbly or too large** → Try **Threshold Crossing** — it's the most stable for general use
- **Donut has visible diffraction rings** (concentric bright rings around the main ring) → Try **Edge to Dark** — it skips past ring-to-ring dips
- **Donut is very dim or has clean, sharp edges** → Try **Steepest Gradient** — it's most precise when there's a single clear edge
- **Outer circle includes black areas inside the ring** → The edge is being detected too far out. Try **Threshold Crossing** or reduce **Gradient Window**

### Gradient Window

The **Gradient Window** setting (Detection tab) controls how many pixels are averaged when computing brightness gradients. This only affects the gradient-based algorithms (Edge to Dark and Steepest Gradient); Threshold Crossing doesn't use it.

- **Larger window (8–15):** Smooths over diffraction ring structure. Better for complex donuts with multiple rings.
- **Smaller window (3–5):** More precise for clean, sharp edges. Better for dim stars without rings.

The script automatically adapts the window size for small donuts, so you generally don't need to change this unless you see detection issues.

## Settings Palette

The settings palette opens automatically when the script runs. You can also reopen it with `settings()` in the console. All changes apply live.

### Display Tab

- **Enable Overlay** — master on/off toggle
- **Detection Circles** / **Centering Crosshairs** — choose between fitted circles or full-diameter crosshairs
- **Overlay Details** — detailed stats panel (top-left)
- **Bottom Numerical Details** — X/Y offset readout bar
- **Correction Arrows** — directional arrows showing which way to adjust
- **Target Brightness** — horizontal brightness level indicator
- **Target Size** — vertical donut size indicator

### Colors Tab

- Color pickers for all overlay elements
- Line widths for circles, crosshairs, and offset lines (default 1.0px — suitable for viewing at 200–300% zoom)
- Font sizes for info panel, bottom bar, and arrow labels

### Detection Tab

- **Algorithm** — outer edge detection method (see above)
- **Gradient Window** — pixel averaging window for gradient calculations
- **Number of Rays** — radial rays for edge detection (more = accurate, slower)
- **Brightness Threshold** — % of peak brightness defining the donut edge
- **Smoothing Factor** — exponential smoothing across frames (0 = none, 0.9 = very smooth)
- Tracking validation parameters (radius change limits, center jump limits, stale frame limit)

### Performance Tab

- **Analysis Interval** — analyze every Nth frame (higher = faster but less responsive)

### Bottom Buttons

| Button | Action |
|--------|--------|
| **Reset** | Reset all settings to defaults |
| **Restart** | Full detection restart — clears tracking and re-searches for the donut |
| **Save** | Save settings to `collimation_settings.cfg` (persists across sessions) |
| **Close** | Stop the overlay and close the palette |

## Reading the Overlay

| Element | Meaning |
|---------|---------|
| **Red circle** | Outer edge of the donut (primary mirror light cone) |
| **Green circle** | Inner edge of the donut (secondary mirror shadow) |
| **Yellow crosshair** | Center of the outer circle |
| **Green crosshair** | Center of the inner circle |
| **Cyan dashed line** | Offset vector between the two centers |
| **Cyan arrows** | Direction to adjust collimation |
| **Bottom bar** | X and Y offset with correction direction and color-coded severity |
| **Info panel** | Full stats: offset in px/%, direction, radii, quality rating |
| **Brightness bar** | Horizontal bar — current brightness level with dim/OK/bright zones |
| **Size bar** | Vertical bar — current donut size with small/ideal/large zones |

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

### Star Selection

- Choose a **bright, isolated star** — the script locks onto the brightest object in the frame, so avoid crowded fields or double stars.
- Brighter stars (mag 1–3) produce a strong, well-defined donut at shorter exposures.
- The script supports both **8-bit and 16-bit** image formats automatically.

### Using SharpCap Features

- **Frame stacking:** If your donut appears noisy, enabling SharpCap's frame stacking smooths out noise and gives cleaner edges.
- **Planet/disk stabilization:** If seeing causes the donut to bounce around, stabilization keeps it steady for better tracking.
- **ROI (Region of Interest):** A smaller ROI speeds up capture and analysis. It also increases the donut's relative frame fill, which can move you into the ideal size zone.

## Troubleshooting

### "Searching for donut..." never goes away

- **No star visible**: Ensure a bright star is in the field of view and the camera is capturing.
- **Star not defocused enough**: The star needs a clear donut shape with a visible dark center.
- **Star too dim**: Increase exposure or gain. Check the brightness bar — it should be in the green zone.
- **Star too small**: The donut must be at least ~30px in diameter. Defocus more or check the size bar.
- **Donut fills the entire frame**: The script needs dark background around the donut. If the ring extends to the frame edges, reduce defocus or increase the capture area.

### Outer circle doesn't match the donut edge

- **Try a different algorithm** — see [Choosing an Algorithm](#choosing-an-algorithm) above.
- **Adjust Gradient Window** — increase for heavy diffraction rings, decrease for clean edges.
- **Adjust Brightness Threshold** — lower if outer circle is too small, raise if too large.
- **Donut is not circular** — if significantly elongated (astigmatism, tilt), the circle fit will be an approximation.

### Detection is jittery or unstable

- Increase **Smoothing Factor** to 0.7–0.85 in the Detection tab.
- Increase exposure or enable frame stacking to improve signal-to-noise.
- Enable SharpCap's planet/disk stabilization if seeing is poor.

### Star moves too far or overlay is lost

- Click **Restart** in the settings palette, or type `restart()` in the console.
- Re-running the script also works — it automatically cleans up the previous instance.

### Overlay is slow or laggy

- Set a smaller ROI in SharpCap to reduce the area analyzed.
- Increase **Analysis Interval** to 3–5 in the Performance tab.

### Still having issues?

1. Type `debug()` in the SharpCap scripting console and press Enter
2. Copy the entire output (everything between the `====` lines)
3. Post it to the [SharpCap forum](https://forums.sharpcap.co.uk/) along with a description of what you're seeing

The `debug()` output includes your script version, camera settings, detection state, a live analysis of the current frame, and all configuration — everything needed to diagnose the problem remotely.

## Console Commands

Type these in the SharpCap scripting console:

| Command | Action |
|---------|--------|
| `settings()` | Open/reopen the settings palette |
| `restart()` | Full detection restart (same as Restart button) |
| `reset_tracking()` | Clear tracking and re-detect the donut |
| `debug()` | Full diagnostic dump |
| `debug_on()` / `debug_off()` | Toggle per-frame debug logging |
| `stop()` | Fully stop the overlay |
| `save_config()` / `load_config()` | Manually save or reload settings |
| `reset_config()` | Reset all settings to defaults |

Re-running the script automatically cleans up the previous instance.

## How It Works

1. Hooks into SharpCap's `BeforeFrameDisplay` event to run on every displayed frame
2. Gets the frame bitmap via `GetFrameBitmap()` and locks pixel data with `LockBits`
3. Finds the donut center using a brightness-weighted centroid
4. Casts radial rays outward from the center, detecting inner edges (dark-to-bright gradient) and outer edges (using the selected algorithm with sub-pixel interpolation)
5. Fits circles to each set of edge points using the Kasa algebraic least-squares method
6. Rejects outlier points and refits iteratively for higher accuracy
7. Applies exponential smoothing across frames to stabilize the overlay
8. Draws the overlay using SharpCap's `GetAnnotationGraphics()` API (with `GetDrawableBitmap()` fallback)
