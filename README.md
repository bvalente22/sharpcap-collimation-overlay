# SharpCap Collimation Overlay

A SharpCap IronPython script that automatically detects and overlays concentric circles on a defocused star (donut shape) to aid collimation of reflector telescopes.

## What It Does

When collimating a reflector telescope, a defocused star appears as a donut — a bright ring with a dark center (the secondary mirror shadow). Perfect collimation means the inner and outer circles of this donut are concentric.

This script:
- Detects the **outer circle** (primary mirror edge) and **inner circle** (secondary shadow)
- Draws overlay circles on the live preview in real-time
- Shows the **offset** between circle centers in pixels, percentage, and direction
- Provides a **quality indicator** (Excellent / Good / Fair / Poor / Bad)
- Updates automatically as you adjust collimation screws

## Requirements

- **SharpCap Pro 4.1+** (scripting requires Pro license)
- A connected camera with live preview running
- A bright, defocused star visible in the field of view

## Installation

1. Copy `collimation_overlay.py` to your SharpCap scripts folder, or any location you prefer.
2. In SharpCap, open the scripting console: **Tools > Scripting Console**

## Usage

### Starting the Overlay

1. Connect your camera and start live preview
2. Point at a bright star and defocus until you see a clear donut shape
3. In the SharpCap scripting console, either:
   - Click **File > Run Script** and select `collimation_overlay.py`, or
   - Type: `exec(open(r"C:\path\to\collimation_overlay.py").read())`

### Controls

| Action | Method |
|--------|--------|
| Toggle on/off | Click the **"Collimation Overlay"** toolbar button |
| Toggle on/off | Run `toggle_overlay()` in the scripting console |
| Force re-detection | Run `reset_tracking()` in the scripting console |
| Remove completely | Run `detach_handler()` in the scripting console |

### Reading the Display

| Element | Meaning |
|---------|---------|
| **Red circle** | Outer edge of donut (primary mirror light cone) |
| **Green circle** | Inner edge of donut (secondary mirror shadow) |
| **Yellow crosshair** | Center of outer circle |
| **Green crosshair** | Center of inner circle |
| **Cyan dashed line** | Offset between centers (should be zero) |
| **Info panel** | Offset values, direction, and quality rating |

### Info Panel Fields

- **Offset**: Distance between centers in pixels and as % of outer radius
- **dX / dY**: Directional offset (+X = right, +Y = up)
- **Direction**: Angle of offset in degrees (0 = right, 90 = up)
- **Outer R / Inner R**: Detected circle radii in pixels
- **Quality**: Overall collimation quality rating

### Quality Ratings

| Rating | Offset % | Meaning |
|--------|----------|---------|
| EXCELLENT | < 2% | Circles are nearly concentric |
| GOOD | < 5% | Minor offset, good for most uses |
| FAIR | < 10% | Noticeable offset |
| POOR | < 20% | Significant offset |
| BAD | > 20% | Major collimation needed |

## Configuration

Edit the `USER CONFIGURATION` section at the top of the script to customize:

- **Colors**: Change any overlay color (RGBA format)
- **Line widths**: Adjust circle and crosshair thickness
- **Analysis parameters**: Tune detection sensitivity
  - `NUM_RAYS`: More rays = more accurate but slower (default: 90)
  - `BRIGHTNESS_THRESHOLD_PERCENT`: Edge detection sensitivity (default: 30%)
  - `CENTROID_DOWNSAMPLE`: Analysis speed vs. accuracy (default: 4)
- **Smoothing**: `SMOOTHING_FACTOR` controls overlay stability (0.0-0.9)
- **Analysis interval**: `ANALYSIS_INTERVAL` — analyze every Nth frame (default: 2)

## How It Works

1. **Centroid detection**: Finds the brightness-weighted center of the image to locate the donut
2. **Radial ray casting**: Casts rays outward from the center in all directions to find brightness transitions (inner and outer edges)
3. **Circle fitting**: Uses the Kasa algebraic method (least-squares) to fit circles to the detected edge points
4. **Outlier rejection**: Removes noisy edge detections and refits for higher accuracy
5. **Exponential smoothing**: Stabilizes the overlay across frames to reduce jitter
6. **GDI+ drawing**: Renders the overlay directly onto SharpCap's live preview via the `BeforeFrameDisplay` event

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Searching for donut..." | Ensure a bright defocused star is visible; adjust exposure |
| Circles are jittery | Increase `SMOOTHING_FACTOR` (e.g., 0.8) |
| Detection is wrong | Adjust `BRIGHTNESS_THRESHOLD_PERCENT`; ensure star is well-defocused |
| Script is slow | Increase `ANALYSIS_INTERVAL` or `CENTROID_DOWNSAMPLE` |
| Overlay doesn't appear | Ensure camera is capturing; check console for errors |
| "No camera selected" | Connect and select a camera before running the script |
