# =============================================================================
# SharpCap Collimation Overlay Script
# =============================================================================
#
# PURPOSE:
#   Draws concentric circle overlays on a defocused star (donut) to aid
#   collimation of reflector telescopes. A defocused star in a reflector
#   appears as a donut: a bright ring (primary mirror) with a dark center
#   (secondary mirror shadow). Perfect collimation = concentric circles.
#   This script auto-detects both circles, measures their offset, and
#   shows correction guidance in real time on SharpCap's live preview.
#
# ARCHITECTURE:
#   The script registers an event handler on SharpCap's BeforeFrameDisplay
#   event. This handler fires on every displayed frame, allowing us to:
#     1. Read pixel data via frame.GetFrameBitmap() + LockBits
#     2. Analyze the image for a donut pattern (radial ray casting)
#     3. Fit circles to detected edges (Kasa least-squares method)
#     4. Draw overlays via frame.GetAnnotationGraphics() (frame annotations)
#        with fallback to frame.GetDrawableBitmap().GetGraphics()
#
#   Performance optimizations (when APIs are available):
#     - frame.GetAnnotationGraphics(): avoids bitmap alloc/dispose per frame
#     - frame.CalculateHistogram(): gets peak brightness natively (skips pixel loop)
#     - frame.CutROI(): analyzes only the donut region when tracking (smaller buffer)
#     - frame.Index: native frame counter (avoids manual tracking)
#
#   The script itself runs once to set up the handler and toolbar button,
#   then exits. The handler remains attached and runs per-frame until
#   stop() is called or SharpCap closes.
#
# SHARPCAP API NOTES (4.1+):
#   - GetAnnotationGraphics() returns a graphics context for frame annotations
#     drawn over the frame without bitmap overhead. Same API as GDI+ wrapper.
#   - GetDrawableBitmap().GetGraphics() is the fallback — returns WinformGraphicsImpl.
#     Most GDI+ methods work but:
#       * FillRectangle requires (brush, RectangleF) not (brush, x, y, w, h)
#       * SmoothingMode/DashStyle may not be supported (wrapped in try/except)
#   - GetDrawableBitmap().GetBitmap() returns the OVERLAY layer (blank),
#     NOT the camera frame. Use frame.GetFrameBitmap() for actual pixels.
#   - CalculateHistogram() returns HistogramData with centile accessors.
#   - CutROI(x, y, w, h) returns a sub-frame for focused processing.
#     Must call .Release() on the returned copy.
#   - frame.Index provides a native frame sequence number.
#   - AddCustomButton takes 4 args but the arg order varies by version.
#     We try 3 orderings and use whichever works.
#
# USAGE:
#   1. Open SharpCap, connect camera, start live preview
#   2. Defocus a bright star until you see a donut shape
#   3. Tools > Scripting Console > File > Run Script > select this file
#   4. Click "Coll Overlay" toolbar button to cycle display modes
#   5. Click "Coll Settings" toolbar button to open the settings palette
#
# Requires: SharpCap Pro 4.1+ (scripting is a Pro feature)
# =============================================================================

import clr
import math
import time
import os as _os

# .NET assemblies for drawing and pixel access
clr.AddReference("System.Drawing")
from System.Drawing import (
    Color, Pen, SolidBrush, Font, FontFamily, FontStyle,
    Bitmap, Rectangle, RectangleF, PointF, StringFormat,
)
from System.Drawing.Imaging import ImageLockMode, PixelFormat
from System.Drawing.Drawing2D import SmoothingMode, DashStyle
from System.Runtime.InteropServices import Marshal
from System import Array, Byte
import System.Drawing

# .NET WinForms for settings dialog
clr.AddReference("System.Windows.Forms")
import System.Windows.Forms
from System.Windows.Forms import (
    Form, TabControl, TabPage, Label, NumericUpDown, TrackBar,
    Button, Panel, ColorDialog, CheckBox, ToolTip, DockStyle, Padding,
    FormBorderStyle, FormStartPosition, AnchorStyles,
    TickStyle, Orientation, MessageBox, MessageBoxButtons,
    MessageBoxIcon, DialogResult, AutoScaleMode as WinFormsAutoScaleMode,
)

# =============================================================================
# CONFIGURATION SYSTEM
# =============================================================================
# All settings are stored in _config dict, populated from _DEFAULTS on startup
# and optionally overridden by a saved config file. The settings dialog and
# console commands modify _config directly — changes take effect on the next frame.

# Default values for all configuration options.
# Colors stored as (A, R, G, B) tuples for serialization.
_DEFAULTS = {
    # Overlay display toggles
    "SHOW_CIRCLES": True,                # Inner + outer detection circles
    "SHOW_INFO_PANEL": True,             # Detailed stats panel (top-left)
    "SHOW_BOTTOM_BAR": True,             # X/Y offset readout bar (bottom)
    "SHOW_CORRECTION_ARROWS": True,      # Directional correction arrows + labels
    "SHOW_ALIGNMENT_CROSSHAIRS": False,  # Full-diameter crosshairs for visual alignment
    # Overlay colors (ARGB)
    "OUTER_CIRCLE_COLOR": (220, 255, 60, 60),      # Red - primary mirror edge
    "INNER_CIRCLE_COLOR": (220, 60, 255, 60),       # Green - secondary shadow edge
    "CROSSHAIR_COLOR":    (180, 255, 255, 0),       # Yellow - outer center marker
    "OFFSET_LINE_COLOR":  (220, 0, 220, 255),       # Cyan - offset vector & arrows
    "TEXT_COLOR":          (240, 255, 255, 255),     # White - info panel text
    "TEXT_BG_COLOR":       (160, 0, 0, 0),           # Semi-transparent black background
    # Line widths (pixels)
    "CIRCLE_PEN_WIDTH":    2.0,
    "CROSSHAIR_PEN_WIDTH": 1.0,
    "OFFSET_PEN_WIDTH":    2.0,
    # Crosshair size (pixels extending from center, used in modes 0-2)
    "CROSSHAIR_LENGTH": 20,
    # Text sizes (points)
    "TEXT_FONT_SIZE":        9,
    "BOTTOM_BAR_FONT_SIZE": 10,
    "ARROW_LABEL_FONT_SIZE": 8,
    # Detection parameters
    "NUM_RAYS":                    90,   # Radial rays for edge detection
    "BRIGHTNESS_THRESHOLD_PERCENT": 30,  # % of peak brightness defining donut edge
    "CENTROID_DOWNSAMPLE":          4,   # Sample every Nth pixel for centroid
    "MIN_DONUT_RADIUS":            15,   # Ignore circles smaller than this (px)
    "RAY_STEP":                     1,   # Pixel step along each ray
    "ROI_MARGIN_FACTOR":          2.0,   # CutROI margin as multiple of outer radius
    # Tracking validation
    "MAX_RADIUS_CHANGE_PCT": 15,         # Max % change in outer radius per detection
    "MAX_CENTER_JUMP_PCT":   20,         # Max center jump as % of outer radius
    "STALE_FRAME_LIMIT":    15,          # Consecutive rejections before re-detect
    # Smoothing
    "SMOOTHING_FACTOR": 0.6,             # EMA factor (0.0 = none, 0.9 = heavy)
    # Performance
    "ANALYSIS_INTERVAL": 2,              # Analyze every Nth frame
}

# Color keys are identified by suffix for serialization
_COLOR_KEYS = frozenset(k for k in _DEFAULTS if k.endswith("_COLOR"))

# Float keys (non-integer numeric)
_FLOAT_KEYS = frozenset(["CIRCLE_PEN_WIDTH", "CROSSHAIR_PEN_WIDTH", "OFFSET_PEN_WIDTH",
                          "ROI_MARGIN_FACTOR", "SMOOTHING_FACTOR"])

# Bool keys (on/off toggles)
_BOOL_KEYS = frozenset(k for k in _DEFAULTS if isinstance(_DEFAULTS[k], bool))

# Live config dict — read by the frame handler on every frame
_config = {}


def _tuple_to_color(t):
    """Convert (A, R, G, B) tuple to System.Drawing.Color."""
    return Color.FromArgb(int(t[0]), int(t[1]), int(t[2]), int(t[3]))


def _color_to_tuple(c):
    """Convert System.Drawing.Color to (A, R, G, B) tuple."""
    return (c.A, c.R, c.G, c.B)


def _init_config():
    """Populate _config from _DEFAULTS, converting color tuples to Color objects."""
    for key, val in _DEFAULTS.items():
        if key in _COLOR_KEYS:
            _config[key] = _tuple_to_color(val)
        else:
            _config[key] = val


def _get_config_path():
    """Return the path to the settings file (next to the script)."""
    try:
        script_dir = _os.path.dirname(_os.path.abspath(__file__))
    except NameError:
        script_dir = "."
    return _os.path.join(script_dir, "collimation_settings.cfg")


def save_config():
    """Save current _config to the settings file."""
    path = _get_config_path()
    try:
        lines = ["# Collimation Overlay Settings", "# Edit with care or use the Settings dialog", ""]
        for key in sorted(_DEFAULTS.keys()):
            val = _config.get(key)
            if val is None:
                continue
            if key in _COLOR_KEYS:
                t = _color_to_tuple(val)
                lines.append("%s = %d,%d,%d,%d" % (key, t[0], t[1], t[2], t[3]))
            elif key in _BOOL_KEYS:
                lines.append("%s = %s" % (key, "true" if val else "false"))
            elif key in _FLOAT_KEYS:
                lines.append("%s = %.2f" % (key, val))
            else:
                lines.append("%s = %s" % (key, val))
        with open(path, "w") as f:
            f.write("\n".join(lines) + "\n")
        print("[Collimation] Settings saved to: %s" % path)
    except Exception as ex:
        print("[Collimation] Error saving settings: %s" % str(ex))


def load_config():
    """Load settings from file, overlaying onto current _config."""
    path = _get_config_path()
    if not _os.path.exists(path):
        return
    try:
        with open(path, "r") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                if "=" not in line:
                    continue
                key, _, val = line.partition("=")
                key = key.strip()
                val = val.strip()
                if key not in _DEFAULTS:
                    continue
                try:
                    if key in _COLOR_KEYS:
                        parts = [int(x.strip()) for x in val.split(",")]
                        if len(parts) == 4:
                            _config[key] = Color.FromArgb(parts[0], parts[1], parts[2], parts[3])
                    elif key in _BOOL_KEYS:
                        _config[key] = val.lower() in ("true", "1", "yes")
                    elif key in _FLOAT_KEYS:
                        _config[key] = float(val)
                    else:
                        _config[key] = int(val)
                except:
                    pass  # Skip malformed values, keep default
        print("[Collimation] Settings loaded from: %s" % path)
    except Exception as ex:
        print("[Collimation] Error loading settings: %s" % str(ex))


def reset_config():
    """Reset all settings to defaults and delete the config file."""
    _init_config()
    path = _get_config_path()
    try:
        if _os.path.exists(path):
            _os.remove(path)
    except:
        pass
    print("[Collimation] Settings reset to defaults")


# Initialize config from defaults, then overlay with saved file
_init_config()
load_config()


# =============================================================================
# INTERNAL STATE
# =============================================================================

# If the script is re-run, clean up the previous instance's event handler,
# toolbar buttons, and settings form to prevent duplicates.
try:
    _prev_state = _state
    if _prev_state.get("handler_attached"):
        try:
            cam = SharpCap.SelectedCamera
            if cam is not None:
                cam.BeforeFrameDisplay -= on_before_frame_display
        except:
            pass
    for btn_name in _prev_state.get("buttons", []):
        try:
            SharpCap.RemoveCustomButton(btn_name)
        except:
            pass
    # Close settings form if open
    sf = _prev_state.get("settings_form")
    if sf is not None:
        try:
            sf.Close()
            sf.Dispose()
        except:
            pass
    print("[Collimation] Cleaned up previous instance")
except NameError:
    pass  # First run, no previous _state exists

# Global state dictionary. Using a dict (mutable) so closures/handlers can
# modify values without needing the 'global' keyword (IronPython quirk).
_state = {
    "enabled": True,
    "outer_cx": None, "outer_cy": None, "outer_r": None,   # Smoothed outer circle
    "inner_cx": None, "inner_cy": None, "inner_r": None,   # Smoothed inner circle
    "frame_count": 0,
    "handler_attached": False,
    "status": "Searching for donut...",
    "last_error": "",
    "buttons": [],          # Toolbar button names (for cleanup)
    "ref_peak": None,       # Reference peak brightness from initial detection
    "ref_outer_r": None,    # Reference outer radius from initial detection (for ROI sizing)
    "reject_count": 0,      # Consecutive rejected detection count
    "debug_log": False,     # Toggle with debug_on() / debug_off()
    "debug_frames": 0,      # Frames logged since debug enabled
    "settings_form": None,  # Reference to open SettingsForm
}



# =============================================================================
# CIRCLE FITTING - Kasa algebraic method (least squares)
# =============================================================================
# Given a set of (x, y) edge points, fits the best circle (cx, cy, r).
# This is a non-iterative algebraic method that solves a linear system.
# Reference: I. Kasa, "A curve fitting procedure and its error analysis",
# IEEE Trans. Inst. Meas., 1976.

def fit_circle(points):
    """
    Fit a circle to a list of (x, y) points using the Kasa algebraic method.
    Returns (center_x, center_y, radius) or None if fitting fails.

    Requires at least 5 points. Returns None if the system is degenerate
    (collinear points) or the fitted radius is too small.
    """
    n = len(points)
    if n < 5:
        return None

    # Accumulate sums for the normal equations
    sx = sy = sx2 = sy2 = sxy = sx3 = sy3 = sx2y = sxy2 = 0.0
    for px, py in points:
        x2 = px * px
        y2 = py * py
        sx   += px
        sy   += py
        sx2  += x2
        sy2  += y2
        sxy  += px * py
        sx3  += x2 * px
        sy3  += y2 * py
        sx2y += x2 * py
        sxy2 += px * y2

    # Build and solve the 2x2 linear system for circle center (cx, cy)
    a11 = n * sx2 - sx * sx
    a12 = n * sxy - sx * sy
    a22 = n * sy2 - sy * sy
    b1  = 0.5 * (n * (sx3 + sxy2) - sx * (sx2 + sy2))
    b2  = 0.5 * (n * (sx2y + sy3) - sy * (sx2 + sy2))

    det = a11 * a22 - a12 * a12
    if abs(det) < 1e-10:
        return None  # Degenerate (collinear points)

    cx = (b1 * a22 - a12 * b2) / det
    cy = (a11 * b2 - b1 * a12) / det

    # Radius = RMS distance from center to all points
    sum_r2 = 0.0
    for px, py in points:
        dx = px - cx
        dy = py - cy
        sum_r2 += dx * dx + dy * dy
    radius = math.sqrt(sum_r2 / n) if sum_r2 > 0 else 0

    if radius < _config["MIN_DONUT_RADIUS"] * 0.5:
        return None

    return (cx, cy, radius)


def reject_outliers_and_refit(points, cx, cy, radius):
    """
    Iteratively remove edge points far from the fitted circle, then refit.
    Runs up to 3 passes with progressively tighter tolerance to handle
    cases where the initial fit is skewed by many outliers (e.g., diffraction
    rings causing 70-80px spread on a 90px radius circle).
    """
    if len(points) < 10:
        return cx, cy, radius

    cur_cx, cur_cy, cur_r = cx, cy, radius
    # Progressively tighter: 2.0x, 1.5x, 1.2x of base tolerance
    for tol_mult in [2.0, 1.5, 1.2]:
        tol = (cur_r * 0.10 + 3) * tol_mult
        filtered = []
        for px, py in points:
            dist = math.sqrt((px - cur_cx) ** 2 + (py - cur_cy) ** 2)
            if abs(dist - cur_r) < tol:
                filtered.append((px, py))

        if len(filtered) < 8:
            break  # Too many rejected at this tolerance, stop tightening

        result = fit_circle(filtered)
        if result is None:
            break
        cur_cx, cur_cy, cur_r = result

    return cur_cx, cur_cy, cur_r


# =============================================================================
# IMAGE ANALYSIS
# =============================================================================
# The donut detection algorithm:
#   1. Find the approximate center using brightness-weighted centroid
#   2. Refine the center using only bright pixels near the initial centroid
#   3. Cast NUM_RAYS radial rays outward from center. Along each ray, detect:
#      - Inner edge: first dark-to-bright transition (secondary shadow boundary)
#      - Outer edge: first bright-to-dark transition after that (primary mirror boundary)
#   4. Fit circles to the inner and outer edge point sets
#   5. Reject outliers and refit for higher accuracy
#   6. Validate: inner radius < outer radius, centers not too far apart, etc.

def _get_brightness(pixels, offset, bpp):
    """Extract grayscale brightness from raw byte array at given byte offset.
    For RGB (bpp>=3), averages the three channels. For mono, returns the byte value."""
    if bpp >= 3:
        return (pixels[offset] + pixels[offset + 1] + pixels[offset + 2]) / 3.0
    else:
        return float(pixels[offset])


def analyze_donut_raw(raw, stride, bpp, width, height, peak_override=None, approx_center=None):
    """
    Analyze raw byte pixel data (from LockBits) for a donut pattern.

    Args:
        raw: .NET byte array containing pixel data
        stride: row stride in bytes (may differ from width*bpp due to padding)
        bpp: bytes per pixel (1 for mono, 3 for RGB, 4 for ARGB)
        width, height: image dimensions in pixels
        peak_override: pre-computed peak brightness (from CalculateHistogram)
        approx_center: (cx, cy) from previous tracking, skips coarse centroid scan

    Returns:
        Dict with inner/outer circle parameters, or None if no donut found.
    """
    # --- Step 1: Determine peak brightness and approximate center ---
    step = _config["CENTROID_DOWNSAMPLE"]
    brightness_thresh_pct = _config["BRIGHTNESS_THRESHOLD_PERCENT"]

    if peak_override is not None and approx_center is not None:
        # Fastest path: both peak and center provided (histogram + tracking).
        peak = peak_override
        ref = _state.get("ref_peak")
        if ref is not None:
            peak = max(peak, ref * 0.5)
        threshold = peak * brightness_thresh_pct / 100.0
        approx_cx, approx_cy = approx_center
    elif approx_center is not None:
        # ROI path: center known from tracking, but scan pixels for accurate peak.
        approx_cx, approx_cy = approx_center
        peak = 0.0
        for y in range(0, height, step):
            row_off = y * stride
            for x in range(0, width, step):
                b = _get_brightness(raw, row_off + x * bpp, bpp)
                if b > peak:
                    peak = b
        if peak < 25.0:
            return None
        ref = _state.get("ref_peak")
        if ref is not None:
            peak = max(peak, ref * 0.5)
        threshold = peak * brightness_thresh_pct / 100.0
    else:
        # Standard path: scan every Nth pixel for peak brightness and
        # brightness-weighted centroid. Used for initial detection.
        sum_bx = 0.0
        sum_by = 0.0
        sum_b  = 0.0
        peak   = 0.0

        for y in range(0, height, step):
            row_off = y * stride
            for x in range(0, width, step):
                b = _get_brightness(raw, row_off + x * bpp, bpp)
                if b > peak:
                    peak = b
                if b > 15.0:  # Ignore pixels below noise floor
                    sum_bx += x * b
                    sum_by += y * b
                    sum_b  += b

        if sum_b < 1.0 or peak < 25.0:
            return None  # No bright object found

        approx_cx = sum_bx / sum_b
        approx_cy = sum_by / sum_b
        threshold = peak * brightness_thresh_pct / 100.0

    # --- Step 2: Refine center using only bright pixels near the centroid ---
    search_r = int(min(width, height) * 0.3)
    x_min = max(0, int(approx_cx) - search_r)
    x_max = min(width, int(approx_cx) + search_r)
    y_min = max(0, int(approx_cy) - search_r)
    y_max = min(height, int(approx_cy) + search_r)
    step2 = max(1, step // 2)  # Finer sampling for refinement

    sum_bx2 = 0.0
    sum_by2 = 0.0
    sum_b2  = 0.0

    for y in range(y_min, y_max, step2):
        row_off = y * stride
        for x in range(x_min, x_max, step2):
            b = _get_brightness(raw, row_off + x * bpp, bpp)
            if b > threshold:
                sum_bx2 += x * b
                sum_by2 += y * b
                sum_b2  += b

    if sum_b2 > 0:
        approx_cx = sum_bx2 / sum_b2
        approx_cy = sum_by2 / sum_b2

    # --- Step 3: Cast rays outward from center to find inner/outer edges ---
    inner_points = []
    outer_points = []
    max_ray_len = int(min(width, height) * 0.45)
    grad_window = 5  # Pixels to average for gradient calculation
    num_rays = _config["NUM_RAYS"]
    ray_step = _config["RAY_STEP"]

    for i in range(num_rays):
        angle = 2.0 * math.pi * i / num_rays
        cos_a = math.cos(angle)
        sin_a = math.sin(angle)

        found_inner = False
        in_bright = False
        ray_samples = []  # (dist, px, py, brightness) collected after inner edge

        for dist in range(3, max_ray_len, ray_step):
            px = int(approx_cx + cos_a * dist)
            py = int(approx_cy + sin_a * dist)

            if px < 0 or px >= width or py < 0 or py >= height:
                break

            b = _get_brightness(raw, py * stride + px * bpp, bpp)
            is_bright = b > threshold

            if not in_bright and is_bright:
                # Transition: dark -> bright = inner edge of the donut ring
                if not found_inner:
                    inner_points.append((px, py))
                    found_inner = True
                in_bright = True

            if found_inner:
                ray_samples.append((dist, px, py, b))

            # Stop collecting well past the expected outer edge
            if found_inner and not is_bright and len(ray_samples) > grad_window * 2:
                pass
            if found_inner and not is_bright and len(ray_samples) > grad_window * 3:
                if b < 5.0:
                    break

        # Find outer edge as point of steepest RELATIVE brightness decline.
        if found_inner and len(ray_samples) > grad_window * 2:
            best_grad = 0.0
            best_idx = -1
            for j in range(grad_window, len(ray_samples) - grad_window):
                avg_before = 0.0
                avg_after = 0.0
                for k in range(grad_window):
                    avg_before += ray_samples[j - 1 - k][3]
                    avg_after += ray_samples[j + k][3]
                avg_before /= grad_window
                avg_after /= grad_window
                if avg_before > 1.0:
                    grad = (avg_before - avg_after) / avg_before
                else:
                    grad = 0.0
                if grad > best_grad:
                    best_grad = grad
                    best_idx = j
            # Accept if at least 30% relative brightness drop
            if best_idx >= 0 and best_grad > 0.3:
                _, opx, opy, _ = ray_samples[best_idx]
                outer_points.append((opx, opy))

    # Need enough edge points from enough rays for a reliable circle fit
    n_inner = len(inner_points)
    n_outer = len(outer_points)
    if n_inner < 12 or n_outer < 12:
        return None

    # --- Step 4: Fit circles to the inner and outer edge point sets ---
    inner_fit = fit_circle(inner_points)
    outer_fit = fit_circle(outer_points)

    if inner_fit is None or outer_fit is None:
        return None

    # Refine by removing outliers (e.g., rays that hit diffraction rings)
    inner_fit = reject_outliers_and_refit(inner_points, inner_fit[0], inner_fit[1], inner_fit[2])
    outer_fit = reject_outliers_and_refit(outer_points, outer_fit[0], outer_fit[1], outer_fit[2])

    icx, icy, ir = inner_fit
    ocx, ocy, ore = outer_fit

    # --- Sanity checks ---
    min_donut_r = _config["MIN_DONUT_RADIUS"]
    if ir >= ore:
        return None  # Inner can't be larger than outer
    if ir < min_donut_r * 0.3:
        return None  # Inner circle too small to be real
    center_dist = math.sqrt((icx - ocx) ** 2 + (icy - ocy) ** 2)
    if center_dist > ore * 0.5:
        return None  # Centers too far apart — probably a bad detection

    # Compute outer edge point spread (diagnostic: how scattered are the points)
    outer_dists = []
    for px, py in outer_points:
        d = math.sqrt((px - ocx) ** 2 + (py - ocy) ** 2)
        outer_dists.append(d)
    outer_dists.sort()
    outer_spread = outer_dists[-1] - outer_dists[0] if outer_dists else 0

    return {
        "inner_cx": icx, "inner_cy": icy, "inner_r": ir,
        "outer_cx": ocx, "outer_cy": ocy, "outer_r": ore,
        "peak": peak,
        "threshold": threshold,
        "n_inner": n_inner, "n_outer": n_outer,
        "outer_spread": outer_spread,
        "approx_cx": approx_cx, "approx_cy": approx_cy,
    }


def analyze_donut_floats(pixels, width, height):
    """
    Same algorithm as analyze_donut_raw but for a flat float array
    (one float per pixel). Used by test_synthetic() which generates
    float data. Not used in live mode (live uses raw bytes from LockBits).
    """
    step = _config["CENTROID_DOWNSAMPLE"]
    brightness_thresh_pct = _config["BRIGHTNESS_THRESHOLD_PERCENT"]
    num_rays = _config["NUM_RAYS"]
    ray_step = _config["RAY_STEP"]
    min_donut_r = _config["MIN_DONUT_RADIUS"]

    # --- Step 1: Find peak brightness and brightness-weighted centroid ---
    sum_bx = 0.0
    sum_by = 0.0
    sum_b  = 0.0
    peak   = 0.0

    for y in range(0, height, step):
        row_off = y * width
        for x in range(0, width, step):
            b = pixels[row_off + x]

            if b > peak:
                peak = b
            if b > 15.0:
                sum_bx += x * b
                sum_by += y * b
                sum_b  += b

    if sum_b < 1.0 or peak < 25.0:
        return None

    approx_cx = sum_bx / sum_b
    approx_cy = sum_by / sum_b

    threshold = peak * brightness_thresh_pct / 100.0

    # --- Step 2: Refine center using bright pixels near the centroid ---
    search_r = int(min(width, height) * 0.3)
    x_min = max(0, int(approx_cx) - search_r)
    x_max = min(width, int(approx_cx) + search_r)
    y_min = max(0, int(approx_cy) - search_r)
    y_max = min(height, int(approx_cy) + search_r)

    sum_bx2 = 0.0
    sum_by2 = 0.0
    sum_b2  = 0.0
    step2 = max(1, step // 2)

    for y in range(y_min, y_max, step2):
        row_off = y * width
        for x in range(x_min, x_max, step2):
            b = pixels[row_off + x]
            if b > threshold:
                sum_bx2 += x * b
                sum_by2 += y * b
                sum_b2  += b

    if sum_b2 > 0:
        approx_cx = sum_bx2 / sum_b2
        approx_cy = sum_by2 / sum_b2

    # --- Step 3: Cast rays outward from center to find inner/outer edges ---
    inner_points = []
    outer_points = []
    max_ray_len = int(min(width, height) * 0.45)
    grad_window = 5

    for i in range(num_rays):
        angle = 2.0 * math.pi * i / num_rays
        cos_a = math.cos(angle)
        sin_a = math.sin(angle)

        found_inner = False
        in_bright = False
        ray_samples = []

        for dist in range(3, max_ray_len, ray_step):
            px = int(approx_cx + cos_a * dist)
            py = int(approx_cy + sin_a * dist)

            if px < 0 or px >= width or py < 0 or py >= height:
                break

            b = pixels[py * width + px]
            is_bright = b > threshold

            if not in_bright and is_bright:
                if not found_inner:
                    inner_points.append((px, py))
                    found_inner = True
                in_bright = True

            if found_inner:
                ray_samples.append((dist, px, py, b))

            if found_inner and not is_bright and len(ray_samples) > grad_window * 3:
                if b < 5.0:
                    break

        # Find outer edge as point of steepest relative brightness decline
        if found_inner and len(ray_samples) > grad_window * 2:
            best_grad = 0.0
            best_idx = -1
            for j in range(grad_window, len(ray_samples) - grad_window):
                avg_before = 0.0
                avg_after = 0.0
                for k in range(grad_window):
                    avg_before += ray_samples[j - 1 - k][3]
                    avg_after += ray_samples[j + k][3]
                avg_before /= grad_window
                avg_after /= grad_window
                if avg_before > 1.0:
                    grad = (avg_before - avg_after) / avg_before
                else:
                    grad = 0.0
                if grad > best_grad:
                    best_grad = grad
                    best_idx = j
            if best_idx >= 0 and best_grad > 0.3:
                _, opx, opy, _ = ray_samples[best_idx]
                outer_points.append((opx, opy))

    if len(inner_points) < 12 or len(outer_points) < 12:
        return None

    # --- Step 4: Fit circles to edge points ---
    inner_fit = fit_circle(inner_points)
    outer_fit = fit_circle(outer_points)

    if inner_fit is None or outer_fit is None:
        return None

    inner_fit = reject_outliers_and_refit(inner_points, inner_fit[0], inner_fit[1], inner_fit[2])
    outer_fit = reject_outliers_and_refit(outer_points, outer_fit[0], outer_fit[1], outer_fit[2])

    icx, icy, ir = inner_fit
    ocx, ocy, ore = outer_fit

    if ir >= ore:
        return None
    if ir < min_donut_r * 0.3:
        return None
    center_dist = math.sqrt((icx - ocx) ** 2 + (icy - ocy) ** 2)
    if center_dist > ore * 0.5:
        return None

    return {
        "inner_cx": icx, "inner_cy": icy, "inner_r": ir,
        "outer_cx": ocx, "outer_cy": ocy, "outer_r": ore,
        "peak": peak,
    }


# =============================================================================
# SMOOTHING
# =============================================================================

def smooth_val(old, new, factor):
    """Apply exponential smoothing: result = old*factor + new*(1-factor)."""
    if old is None:
        return new
    return old * factor + new * (1.0 - factor)


def _validate_detection(result):
    """Check if a new detection is consistent with the current smoothed state."""
    old_r = _state["outer_r"]
    if old_r is None:
        return True  # No previous state, accept anything

    new_r = result["outer_r"]
    radius_change_pct = abs(new_r - old_r) / old_r * 100.0
    if radius_change_pct > _config["MAX_RADIUS_CHANGE_PCT"]:
        return False

    dx = result["outer_cx"] - _state["outer_cx"]
    dy = result["outer_cy"] - _state["outer_cy"]
    center_jump = math.sqrt(dx * dx + dy * dy)
    if center_jump > old_r * _config["MAX_CENTER_JUMP_PCT"] / 100.0:
        return False

    return True


def update_state_with_result(result):
    """Validate and apply smoothing to new analysis results."""
    is_debug = _state.get("debug_log", False)

    if not _validate_detection(result):
        _state["reject_count"] = _state.get("reject_count", 0) + 1
        if is_debug:
            print("[DBG] REJECTED frame %d (count=%d) raw_or=%.1f spread=%.1f raw_ocx=%.1f raw_ocy=%.1f | smooth_or=%.1f smooth_ocx=%.1f" % (
                _state["frame_count"], _state["reject_count"],
                result["outer_r"], result.get("outer_spread", 0),
                result["outer_cx"], result["outer_cy"],
                _state["outer_r"] or 0, _state["outer_cx"] or 0))
        if _state["reject_count"] >= _config["STALE_FRAME_LIMIT"]:
            _state["outer_cx"] = None
            _state["outer_cy"] = None
            _state["outer_r"] = None
            _state["inner_cx"] = None
            _state["inner_cy"] = None
            _state["inner_r"] = None
            _state["ref_peak"] = None
            _state["ref_outer_r"] = None
            _state["reject_count"] = 0
            _state["status"] = "Re-detecting..."
            if is_debug:
                print("[DBG] RESET - too many rejections, re-detecting from scratch")
        return

    _state["reject_count"] = 0
    f = _config["SMOOTHING_FACTOR"]

    if is_debug:
        old_or = _state["outer_r"]
        old_ir = _state["inner_r"]
        roi_used = result.get("_roi_used", False)
        print("[DBG] f=%d peak=%.0f thr=%.0f rays_i=%d rays_o=%d spread=%.1f %s" % (
            _state["frame_count"],
            result.get("peak", 0) or 0,
            result.get("threshold", 0) or 0,
            result.get("n_inner", 0),
            result.get("n_outer", 0),
            result.get("outer_spread", 0),
            "ROI" if roi_used else "FULL"))
        print("[DBG]   raw:  ocx=%.1f ocy=%.1f or=%.1f | icx=%.1f icy=%.1f ir=%.1f" % (
            result["outer_cx"], result["outer_cy"], result["outer_r"],
            result["inner_cx"], result["inner_cy"], result["inner_r"]))
        new_ocx = smooth_val(_state["outer_cx"], result["outer_cx"], f)
        new_ocy = smooth_val(_state["outer_cy"], result["outer_cy"], f)
        new_or = smooth_val(old_or, result["outer_r"], f)
        new_ir = smooth_val(old_ir, result["inner_r"], f)
        print("[DBG]   smooth: ocx=%.1f ocy=%.1f or=%.1f->%.1f | ir=%.1f->%.1f  ratio=%.2f" % (
            new_ocx, new_ocy,
            old_or or 0, new_or,
            old_ir or 0, new_ir,
            new_ir / new_or if new_or > 0 else 0))
        _state["debug_frames"] = _state.get("debug_frames", 0) + 1

    _state["outer_cx"] = smooth_val(_state["outer_cx"], result["outer_cx"], f)
    _state["outer_cy"] = smooth_val(_state["outer_cy"], result["outer_cy"], f)
    _state["outer_r"]  = smooth_val(_state["outer_r"],  result["outer_r"],  f)
    _state["inner_cx"] = smooth_val(_state["inner_cx"], result["inner_cx"], f)
    _state["inner_cy"] = smooth_val(_state["inner_cy"], result["inner_cy"], f)
    _state["inner_r"]  = smooth_val(_state["inner_r"],  result["inner_r"],  f)
    _state["status"] = "Tracking"


# =============================================================================
# HISTOGRAM AND ROI HELPERS
# =============================================================================

def _get_histogram_peak(frame):
    """Get approximate peak brightness from frame histogram (native, fast)."""
    try:
        hist = frame.CalculateHistogram()
        vals = hist.GetCentilePointsF(99.9)
        n = min(3, len(vals))
        total = 0.0
        for i in range(n):
            total += float(vals[i])
        peak = total / n
        return peak if peak > 0 else None
    except:
        return None


def _try_cut_roi(frame):
    """Cut a region of interest around the tracked donut position."""
    try:
        roi_r = _state.get("ref_outer_r") or _state["outer_r"]
        margin = int(roi_r * _config["ROI_MARGIN_FACTOR"])
        info = frame.Info
        roi_x = max(0, int(_state["outer_cx"]) - margin)
        roi_y = max(0, int(_state["outer_cy"]) - margin)
        roi_w = min(info.Width - roi_x, margin * 2)
        roi_h = min(info.Height - roi_y, margin * 2)

        if roi_w > 50 and roi_h > 50:
            roi = frame.CutROI(Rectangle(roi_x, roi_y, roi_w, roi_h))
            return (roi, roi_x, roi_y)
    except:
        pass
    return (None, 0, 0)


# =============================================================================
# OVERLAY DRAWING
# =============================================================================

def draw_overlay(gfx, width, height):
    """Main overlay drawing function. Respects overlay toggle settings."""
    try:
        gfx.SmoothingMode = SmoothingMode.AntiAlias
    except:
        pass

    ocx = _state["outer_cx"]
    ocy = _state["outer_cy"]
    ore = _state["outer_r"]
    icx = _state["inner_cx"]
    icy = _state["inner_cy"]
    ir  = _state["inner_r"]

    if ocx is None or icx is None:
        draw_status_text(gfx, "Searching for donut...")
        return

    # --- Circles ---
    if _config["SHOW_CIRCLES"]:
        pen = Pen(_config["OUTER_CIRCLE_COLOR"], _config["CIRCLE_PEN_WIDTH"])
        gfx.DrawEllipse(pen, float(ocx - ore), float(ocy - ore), float(ore * 2), float(ore * 2))
        pen.Dispose()

        pen = Pen(_config["INNER_CIRCLE_COLOR"], _config["CIRCLE_PEN_WIDTH"])
        gfx.DrawEllipse(pen, float(icx - ir), float(icy - ir), float(ir * 2), float(ir * 2))
        pen.Dispose()

        # Small center crosshairs (tied to circles)
        pen = Pen(_config["CROSSHAIR_COLOR"], _config["CROSSHAIR_PEN_WIDTH"])
        ch = float(_config["CROSSHAIR_LENGTH"])
        gfx.DrawLine(pen, float(ocx - ch), float(ocy), float(ocx + ch), float(ocy))
        gfx.DrawLine(pen, float(ocx), float(ocy - ch), float(ocx), float(ocy + ch))
        pen.Dispose()

        pen = Pen(_config["INNER_CIRCLE_COLOR"], _config["CROSSHAIR_PEN_WIDTH"])
        ch2 = float(_config["CROSSHAIR_LENGTH"] * 0.6)
        gfx.DrawLine(pen, float(icx - ch2), float(icy), float(icx + ch2), float(icy))
        gfx.DrawLine(pen, float(icx), float(icy - ch2), float(icx), float(icy + ch2))
        pen.Dispose()

        # Offset vector line between centers
        dx = icx - ocx
        dy = icy - ocy
        offset_dist = math.sqrt(dx * dx + dy * dy)
        if offset_dist > 0.5:
            pen = Pen(_config["OFFSET_LINE_COLOR"], _config["OFFSET_PEN_WIDTH"])
            try:
                pen.DashStyle = DashStyle.Dash
            except:
                pass
            gfx.DrawLine(pen, float(ocx), float(ocy), float(icx), float(icy))
            pen.Dispose()

            if offset_dist > 3:
                arrow_len = min(12.0, offset_dist * 0.4)
                angle = math.atan2(dy, dx)
                a1 = angle + math.pi * 0.8
                a2 = angle - math.pi * 0.8
                pen = Pen(_config["OFFSET_LINE_COLOR"], _config["OFFSET_PEN_WIDTH"])
                gfx.DrawLine(pen,
                    float(icx), float(icy),
                    float(icx + arrow_len * math.cos(a1)), float(icy + arrow_len * math.sin(a1)))
                gfx.DrawLine(pen,
                    float(icx), float(icy),
                    float(icx + arrow_len * math.cos(a2)), float(icy + arrow_len * math.sin(a2)))
                pen.Dispose()
    else:
        dx = icx - ocx
        dy = icy - ocy
        offset_dist = math.sqrt(dx * dx + dy * dy)

    # --- Alignment crosshairs (full-diameter, for visual overlap) ---
    if _config["SHOW_ALIGNMENT_CROSSHAIRS"]:
        pen = Pen(_config["OUTER_CIRCLE_COLOR"], _config["CIRCLE_PEN_WIDTH"])
        gfx.DrawEllipse(pen, float(ocx - ore), float(ocy - ore), float(ore * 2), float(ore * 2))
        gfx.DrawLine(pen, float(ocx - ore), float(ocy), float(ocx + ore), float(ocy))
        gfx.DrawLine(pen, float(ocx), float(ocy - ore), float(ocx), float(ocy + ore))
        pen.Dispose()

        pen = Pen(_config["INNER_CIRCLE_COLOR"], _config["CIRCLE_PEN_WIDTH"])
        gfx.DrawEllipse(pen, float(icx - ir), float(icy - ir), float(ir * 2), float(ir * 2))
        gfx.DrawLine(pen, float(icx - ir), float(icy), float(icx + ir), float(icy))
        gfx.DrawLine(pen, float(icx), float(icy - ir), float(icx), float(icy + ir))
        pen.Dispose()

    # --- Correction arrows ---
    if _config["SHOW_CORRECTION_ARROWS"]:
        min_arrow_threshold = 1.5
        arrow_gap = 15.0
        arrow_shaft = 30.0
        arrow_head = 10.0
        arrow_pen_w = 2.5
        corr_x = -dx
        corr_y = -dy

        if abs(corr_x) > min_arrow_threshold:
            ax_sign = 1.0 if corr_x > 0 else -1.0
            ax_y = float(ocy - ore - arrow_gap - 8)
            ax_start = float(ocx)
            ax_end = float(ocx + ax_sign * arrow_shaft)

            pen = Pen(_config["OFFSET_LINE_COLOR"], arrow_pen_w)
            gfx.DrawLine(pen, ax_start, ax_y, ax_end, ax_y)
            gfx.DrawLine(pen, ax_end, ax_y,
                float(ax_end - ax_sign * arrow_head * 0.7), float(ax_y - arrow_head * 0.5))
            gfx.DrawLine(pen, ax_end, ax_y,
                float(ax_end - ax_sign * arrow_head * 0.7), float(ax_y + arrow_head * 0.5))
            pen.Dispose()

            font = Font(FontFamily.GenericMonospace, _config["ARROW_LABEL_FONT_SIZE"], FontStyle.Bold)
            brush = SolidBrush(_config["OFFSET_LINE_COLOR"])
            if corr_x > 0:
                gfx.DrawString("MOVE X >>>", font, brush, float(ocx + 5), float(ax_y - 16))
            else:
                gfx.DrawString("<<< MOVE X", font, brush, float(ocx - arrow_shaft - 20), float(ax_y - 16))
            font.Dispose()
            brush.Dispose()

        if abs(corr_y) > min_arrow_threshold:
            ay_sign = 1.0 if corr_y > 0 else -1.0
            ay_x = float(ocx + ore + arrow_gap + 8)
            ay_start = float(ocy)
            ay_end = float(ocy + ay_sign * arrow_shaft)

            pen = Pen(_config["OFFSET_LINE_COLOR"], arrow_pen_w)
            gfx.DrawLine(pen, ay_x, ay_start, ay_x, ay_end)
            gfx.DrawLine(pen, ay_x, ay_end,
                float(ay_x - arrow_head * 0.5), float(ay_end - ay_sign * arrow_head * 0.7))
            gfx.DrawLine(pen, ay_x, ay_end,
                float(ay_x + arrow_head * 0.5), float(ay_end - ay_sign * arrow_head * 0.7))
            pen.Dispose()

            font = Font(FontFamily.GenericMonospace, _config["ARROW_LABEL_FONT_SIZE"], FontStyle.Bold)
            brush = SolidBrush(_config["OFFSET_LINE_COLOR"])
            if corr_y < 0:
                gfx.DrawString("MOVE Y UP", font, brush, float(ay_x + 5), float(ocy - arrow_shaft - 5))
            else:
                gfx.DrawString("MOVE Y DN", font, brush, float(ay_x + 5), float(ocy + arrow_shaft - 5))
            font.Dispose()
            brush.Dispose()

    # --- Info panel ---
    if _config["SHOW_INFO_PANEL"]:
        draw_info_panel(gfx, width, height, ocx, ocy, ore, icx, icy, ir, offset_dist)

    # --- Bottom bar ---
    if _config["SHOW_BOTTOM_BAR"]:
        draw_bottom_bar(gfx, width, height, dx, dy, offset_dist, ore)


def draw_bottom_bar(gfx, width, height, dx, dy, offset_dist, ore):
    """Draw the X/Y offset readout bar at the bottom of the frame."""
    bar_h = float(_config["BOTTOM_BAR_FONT_SIZE"] + 16)
    bar_y = float(height) - bar_h
    offset_pct = (offset_dist / ore * 100.0) if ore > 0 else 0.0

    bg = SolidBrush(Color.FromArgb(200, 0, 0, 0))
    gfx.FillRectangle(bg, RectangleF(0.0, bar_y, float(width), bar_h))
    bg.Dispose()

    font_lg = Font(FontFamily.GenericMonospace, _config["BOTTOM_BAR_FONT_SIZE"], FontStyle.Bold)
    font_sm = Font(FontFamily.GenericMonospace, _config["BOTTOM_BAR_FONT_SIZE"] - 2.0, FontStyle.Regular)

    def offset_color(val):
        pct = abs(val) / ore * 100.0 if ore > 0 else 0
        if pct < 2:
            return Color.FromArgb(255, 0, 255, 0)
        elif pct < 5:
            return Color.FromArgb(255, 150, 255, 0)
        elif pct < 10:
            return Color.FromArgb(255, 255, 255, 0)
        elif pct < 20:
            return Color.FromArgb(255, 255, 150, 0)
        else:
            return Color.FromArgb(255, 255, 50, 50)

    x_color = offset_color(dx)
    x_brush = SolidBrush(x_color)
    x_corr = "<<<" if dx > 0 else ">>>"
    if abs(dx) < 1.5:
        x_corr = "OK"
    x_text = "X: %+.1f px  %s" % (-dx, x_corr)
    gfx.DrawString(x_text, font_lg, x_brush, float(width * 0.1), bar_y + 6.0)
    x_brush.Dispose()

    y_color = offset_color(dy)
    y_brush = SolidBrush(y_color)
    y_corr = "UP" if dy > 0 else "DN"
    if abs(dy) < 1.5:
        y_corr = "OK"
    y_text = "Y: %+.1f px  %s" % (dy, y_corr)
    gfx.DrawString(y_text, font_lg, y_brush, float(width * 0.45), bar_y + 6.0)
    y_brush.Dispose()

    tot_color = offset_color(offset_dist)
    tot_brush = SolidBrush(tot_color)
    tot_text = "Total: %.1f px (%.1f%%)" % (offset_dist, offset_pct)
    gfx.DrawString(tot_text, font_sm, tot_brush, float(width * 0.75), bar_y + 9.0)
    tot_brush.Dispose()

    font_lg.Dispose()
    font_sm.Dispose()


def draw_info_panel(gfx, width, height, ocx, ocy, ore, icx, icy, ir, offset_dist):
    """Draw the detailed info panel at top-left with offset stats and quality rating."""
    dx = icx - ocx
    dy = icy - ocy
    offset_pct = (offset_dist / ore * 100.0) if ore > 0 else 0.0

    if offset_dist > 0.5:
        angle_deg = math.degrees(math.atan2(-dy, dx))
        if angle_deg < 0:
            angle_deg += 360.0
        dir_str = "%.0f deg" % angle_deg
    else:
        dir_str = "--"

    if offset_pct < 2:
        quality = "EXCELLENT"
        qcolor = Color.FromArgb(255, 0, 255, 0)
    elif offset_pct < 5:
        quality = "GOOD"
        qcolor = Color.FromArgb(255, 150, 255, 0)
    elif offset_pct < 10:
        quality = "FAIR"
        qcolor = Color.FromArgb(255, 255, 255, 0)
    elif offset_pct < 20:
        quality = "POOR"
        qcolor = Color.FromArgb(255, 255, 150, 0)
    else:
        quality = "BAD"
        qcolor = Color.FromArgb(255, 255, 50, 50)

    lines = [
        "COLLIMATION OVERLAY",
        "-------------------",
        "Offset: %.1f px (%.1f%%)" % (offset_dist, offset_pct),
        "dX: %+.1f  dY: %+.1f" % (dx, -dy),
        "Direction: %s" % dir_str,
        "Outer R: %.0f  Inner R: %.0f" % (ore, ir),
        "Quality: %s" % quality,
    ]

    txt_font_size = _config["TEXT_FONT_SIZE"]
    font = Font(FontFamily.GenericMonospace, txt_font_size, FontStyle.Regular)
    txt_brush = SolidBrush(_config["TEXT_COLOR"])
    bg_brush = SolidBrush(_config["TEXT_BG_COLOR"])
    q_brush = SolidBrush(qcolor)

    pad = 6
    line_h = txt_font_size + 4
    panel_w = float(txt_font_size * 22)
    panel_h = float(len(lines) * line_h + pad * 2)

    gfx.FillRectangle(bg_brush, RectangleF(float(pad), float(pad), panel_w, panel_h))

    y = float(pad + 4)
    for i, line in enumerate(lines):
        if i == 6:
            gfx.DrawString("Quality: ", font, txt_brush, float(pad + 6), y)
            gfx.DrawString("         %s" % quality, font, q_brush, float(pad + 6), y)
        else:
            gfx.DrawString(line, font, txt_brush, float(pad + 6), y)
        y += line_h

    font.Dispose()
    txt_brush.Dispose()
    bg_brush.Dispose()
    q_brush.Dispose()


def draw_status_text(gfx, text):
    """Draw a status message (e.g., 'Searching for donut...') at top-left."""
    txt_font_size = _config["TEXT_FONT_SIZE"]
    font = Font(FontFamily.GenericMonospace, txt_font_size, FontStyle.Bold)
    brush = SolidBrush(Color.Orange)
    bg = SolidBrush(_config["TEXT_BG_COLOR"])
    w = float(len(text) * (txt_font_size * 0.65) + 16)
    h = float(txt_font_size + 10)
    gfx.FillRectangle(bg, RectangleF(6.0, 6.0, w, h))
    gfx.DrawString(text, font, brush, 10.0, 8.0)
    font.Dispose()
    brush.Dispose()
    bg.Dispose()


def draw_error_text(gfx, text):
    """Draw an error message at top-left (red text on dark background)."""
    font = Font(FontFamily.GenericMonospace, 10.0, FontStyle.Regular)
    brush = SolidBrush(Color.Red)
    bg = SolidBrush(_config["TEXT_BG_COLOR"])
    gfx.FillRectangle(bg, RectangleF(10.0, 10.0, 500.0, 20.0))
    gfx.DrawString(text, font, brush, 14.0, 12.0)
    font.Dispose()
    brush.Dispose()
    bg.Dispose()


# =============================================================================
# FRAME HANDLER
# =============================================================================

def on_before_frame_display(*args):
    """BeforeFrameDisplay event handler. Called by SharpCap on every frame."""
    if not _state["enabled"]:
        return

    dbitmap = None
    gfx = None

    try:
        frame = args[1].Frame

        try:
            gfx = frame.GetAnnotationGraphics("collimation")
        except:
            dbitmap = frame.GetDrawableBitmap()
            if dbitmap is None:
                return
            gfx = dbitmap.GetGraphics()

        if gfx is None:
            if dbitmap is not None:
                dbitmap.Dispose()
            return

        _state["frame_count"] += 1

        # --- Analysis phase (runs every ANALYSIS_INTERVAL frames) ---
        if _state["frame_count"] % _config["ANALYSIS_INTERVAL"] == 0:
            roi_frame = None
            try:
                approx_center = None
                roi_ox = roi_oy = 0
                peak = None
                if _state["outer_cx"] is not None and _state["outer_r"] is not None:
                    roi_frame, roi_ox, roi_oy = _try_cut_roi(frame)
                    if roi_frame is not None:
                        approx_center = (
                            _state["outer_cx"] - roi_ox,
                            _state["outer_cy"] - roi_oy,
                        )
                    else:
                        peak = _get_histogram_peak(frame)
                        approx_center = (
                            _state["outer_cx"],
                            _state["outer_cy"],
                        )

                target = roi_frame if roi_frame is not None else frame
                bmp = target.GetFrameBitmap()
                if bmp is not None:
                    result = analyze_bitmap(bmp, peak, approx_center)
                    if result is not None:
                        result["_roi_used"] = roi_frame is not None
                        if roi_frame is not None:
                            result["inner_cx"] += roi_ox
                            result["inner_cy"] += roi_oy
                            result["outer_cx"] += roi_ox
                            result["outer_cy"] += roi_oy
                        if _state["ref_peak"] is None and result.get("peak"):
                            _state["ref_peak"] = result["peak"]
                        if _state["ref_outer_r"] is None:
                            _state["ref_outer_r"] = result["outer_r"]
                        update_state_with_result(result)
                        _state["last_error"] = ""
                    elif _state["outer_cx"] is None:
                        _state["status"] = "Searching for donut..."
                else:
                    _state["last_error"] = "GetFrameBitmap returned None"
            except Exception as ex:
                _state["last_error"] = "Analysis: %s" % str(ex)
                if _state["frame_count"] <= 5:
                    print("[Collimation] Analysis error: %s" % str(ex))
            finally:
                if roi_frame is not None:
                    try:
                        roi_frame.Release()
                    except:
                        pass

        # --- Drawing phase (runs every frame for smooth overlay) ---
        info = frame.Info
        draw_overlay(gfx, info.Width, info.Height)

    except Exception as ex:
        if gfx is not None:
            try:
                draw_error_text(gfx, "Error: %s" % str(ex))
            except:
                pass
        if _state.get("frame_count", 0) < 5:
            print("[Collimation] Handler error: %s" % str(ex))

    finally:
        if gfx is not None:
            try:
                gfx.Dispose()
            except:
                pass
        if dbitmap is not None:
            try:
                dbitmap.Dispose()
            except:
                pass


def analyze_bitmap(bmp, peak_override=None, approx_center=None):
    """Extract raw pixel data from a System.Drawing.Bitmap and run donut analysis."""
    width = bmp.Width
    height = bmp.Height
    rect = Rectangle(0, 0, width, height)

    bmpdata = None
    for fmt in [PixelFormat.Format24bppRgb, PixelFormat.Format32bppArgb,
                PixelFormat.Format32bppRgb, PixelFormat.Format8bppIndexed]:
        try:
            bmpdata = bmp.LockBits(rect, ImageLockMode.ReadOnly, fmt)
            break
        except:
            continue

    if bmpdata is None:
        try:
            bmpdata = bmp.LockBits(rect, ImageLockMode.ReadOnly, bmp.PixelFormat)
        except Exception as ex:
            _state["last_error"] = "LockBits failed: %s" % str(ex)
            return None

    stride = bmpdata.Stride
    total = abs(stride) * height
    raw = Array.CreateInstance(Byte, total)
    Marshal.Copy(bmpdata.Scan0, raw, 0, total)
    bmp.UnlockBits(bmpdata)

    bpp = abs(stride) // width
    if bpp < 1:
        bpp = 1

    return analyze_donut_raw(raw, abs(stride), bpp, width, height, peak_override, approx_center)


# =============================================================================
# TOGGLE AND CONTROL
# =============================================================================

def diagnose():
    """Diagnostic tool: captures one frame and tests all pixel access methods."""
    print("[Diag] Capturing frame - testing all pixel access methods...")

    _diag = {"done": False, "info": []}

    def diag_handler(*args):
        if _diag["done"]:
            return
        _diag["done"] = True
        info = _diag["info"]
        try:
            frame = args[1].Frame
            fi = frame.Info
            w = fi.Width
            h = fi.Height
            info.append("Frame: %dx%d, BitDepth=%s, ColorPlanes=%s, BPP=%s" % (
                w, h, fi.BitDepth, fi.ColorPlanes, fi.BytesPerPixel))
            info.append("Stride=%s" % fi.Stride)

            cx = w // 2
            cy = h // 2

            info.append("")
            info.append("=== Method 1: GetPixelValue ===")
            try:
                val_center = frame.GetPixelValue(cx, cy)
                val_corner = frame.GetPixelValue(10, 10)
                val_qr = frame.GetPixelValue(cx + w // 4, cy)
                info.append("  center(%d,%d) = %s (type: %s)" % (cx, cy, val_center, type(val_center)))
                info.append("  corner(10,10) = %s" % val_corner)
                info.append("  quarter-right(%d,%d) = %s" % (cx + w // 4, cy, val_qr))
            except Exception as ex:
                info.append("  FAILED: %s" % ex)

            info.append("")
            info.append("=== Method 2: frame.ToBytes() ===")
            try:
                raw = frame.ToBytes()
                info.append("  Length: %d (expected mono=%d, rgb=%d)" % (
                    len(raw), w * h, w * h * 3))
                bpp_guess = len(raw) // (w * h)
                info.append("  Implied BPP: %d" % bpp_guess)
                if bpp_guess >= 1:
                    stride_guess = len(raw) // h
                    off = cy * stride_guess + cx * bpp_guess
                    if bpp_guess >= 3:
                        val = (raw[off] + raw[off+1] + raw[off+2]) / 3.0
                    elif bpp_guess == 2:
                        val = raw[off] + raw[off+1] * 256.0
                    else:
                        val = float(raw[off])
                    info.append("  center pixel: %.1f" % val)
                    off2 = 10 * stride_guess + 10 * bpp_guess
                    if bpp_guess >= 3:
                        val2 = (raw[off2] + raw[off2+1] + raw[off2+2]) / 3.0
                    elif bpp_guess == 2:
                        val2 = raw[off2] + raw[off2+1] * 256.0
                    else:
                        val2 = float(raw[off2])
                    info.append("  corner(10,10): %.1f" % val2)
            except Exception as ex:
                info.append("  FAILED: %s" % ex)

            info.append("")
            info.append("=== Method 3: frame.GetFrameBitmap() ===")
            try:
                fbmp = frame.GetFrameBitmap()
                info.append("  Result: %s (type: %s)" % (fbmp, type(fbmp)))
                if fbmp is not None:
                    info.append("  Size: %dx%d, Format: %s" % (fbmp.Width, fbmp.Height, fbmp.PixelFormat))
                    px = fbmp.GetPixel(cx, cy)
                    info.append("  center pixel: R=%d G=%d B=%d A=%d" % (px.R, px.G, px.B, px.A))
                    px2 = fbmp.GetPixel(cx + w // 4, cy)
                    info.append("  quarter-right: R=%d G=%d B=%d" % (px2.R, px2.G, px2.B))
            except Exception as ex:
                info.append("  FAILED: %s" % ex)

            info.append("")
            info.append("=== Method 4: frame.CreateBitmap() ===")
            try:
                cbmp = frame.CreateBitmap()
                info.append("  Result: %s (type: %s)" % (cbmp, type(cbmp)))
                if cbmp is not None:
                    info.append("  Size: %dx%d, Format: %s" % (cbmp.Width, cbmp.Height, cbmp.PixelFormat))
                    px = cbmp.GetPixel(cx, cy)
                    info.append("  center pixel: R=%d G=%d B=%d" % (px.R, px.G, px.B))
            except Exception as ex:
                info.append("  FAILED: %s" % ex)

            info.append("")
            info.append("=== Method 5: GetDrawableBitmap().GetBitmap() ===")
            try:
                db = frame.GetDrawableBitmap()
                bmp = db.GetBitmap()
                px = bmp.GetPixel(cx, cy)
                info.append("  center pixel: R=%d G=%d B=%d A=%d (expect 0 = overlay layer)" % (px.R, px.G, px.B, px.A))
                db.Dispose()
            except Exception as ex:
                info.append("  FAILED: %s" % ex)

            info.append("")
            info.append("=== Method 6: frame.GetBufferLease() ===")
            try:
                lease = frame.GetBufferLease()
                info.append("  Result: %s (type: %s)" % (lease, type(lease)))
                info.append("  Members: %s" % [x for x in dir(lease) if not x.startswith("_")])
                lease.Dispose()
            except Exception as ex:
                info.append("  FAILED: %s" % ex)

        except Exception as ex:
            info.append("TOP-LEVEL EXCEPTION: %s" % str(ex))

    cam = SharpCap.SelectedCamera
    cam.BeforeFrameDisplay += diag_handler

    import time
    time.sleep(2)

    cam.BeforeFrameDisplay -= diag_handler

    print("[Diag] Results:")
    for line in _diag["info"]:
        print("  %s" % line)
    if not _diag["info"]:
        print("  No frames received!")


def probe_apis():
    """Probe the new SharpCap APIs and print what's available."""
    print("[Probe] Capturing one frame to test APIs...")

    _probe = {"done": False, "info": []}

    def probe_handler(*args):
        if _probe["done"]:
            return
        _probe["done"] = True
        info = _probe["info"]
        try:
            frame = args[1].Frame

            info.append("=== frame.Index ===")
            try:
                idx = frame.Index
                info.append("  Value: %s (type: %s)" % (idx, type(idx)))
            except Exception as ex:
                info.append("  FAILED: %s" % ex)

            info.append("")
            info.append("=== frame.GetAnnotationGraphics() ===")
            for test_args, desc in [
                ((), "no args"),
                (("collimation",), "string key"),
                ((0,), "int 0"),
                ((True,), "bool True"),
            ]:
                try:
                    ag = frame.GetAnnotationGraphics(*test_args)
                    info.append("  (%s) Result: %s (type: %s)" % (desc, ag, type(ag)))
                    info.append("  (%s) Members: %s" % (desc, [x for x in dir(ag) if not x.startswith("_")]))
                    ag.Dispose()
                    break
                except Exception as ex:
                    info.append("  (%s) FAILED: %s" % (desc, ex))

            info.append("")
            info.append("=== frame.CalculateHistogram() ===")
            try:
                hist = frame.CalculateHistogram()
                info.append("  Result: %s (type: %s)" % (hist, type(hist)))
                info.append("  Members: %s" % [x for x in dir(hist) if not x.startswith("_")])

                info.append("")
                info.append("  --- GetCentilePointsF ---")
                try:
                    val = hist.GetCentilePointsF(99.9)
                    info.append("  GetCentilePointsF(99.9) = %s (type: %s)" % (val, type(val)))
                except Exception as ex2:
                    info.append("  GetCentilePointsF(99.9) FAILED: %s" % ex2)
                try:
                    from System import Array, Double, Single
                    for arr_type in [Double, Single]:
                        try:
                            arr = Array[arr_type]([50.0, 90.0, 99.0, 99.9])
                            val = hist.GetCentilePointsF(arr)
                            info.append("  GetCentilePointsF(Array[%s]([50,90,99,99.9])) = %s" % (arr_type.__name__, val))
                            if val is not None:
                                info.append("    Length: %s, Values: %s" % (len(val), list(val)))
                            break
                        except Exception as ex3:
                            info.append("  GetCentilePointsF(Array[%s]) FAILED: %s" % (arr_type.__name__, ex3))
                except Exception as ex2:
                    info.append("  Array attempt FAILED: %s" % ex2)

                info.append("")
                info.append("  --- GetCentilePoints ---")
                try:
                    val = hist.GetCentilePoints(99.9)
                    info.append("  GetCentilePoints(99.9) = %s (type: %s)" % (val, type(val)))
                except Exception as ex2:
                    info.append("  GetCentilePoints(99.9) FAILED: %s" % ex2)
                try:
                    from System import Array, Double, Single
                    for arr_type in [Double, Single]:
                        try:
                            arr = Array[arr_type]([50.0, 90.0, 99.0, 99.9])
                            val = hist.GetCentilePoints(arr)
                            info.append("  GetCentilePoints(Array[%s]([50,90,99,99.9])) = %s" % (arr_type.__name__, val))
                            if val is not None:
                                info.append("    Length: %s, Values: %s" % (len(val), list(val)))
                            break
                        except Exception as ex3:
                            info.append("  GetCentilePoints(Array[%s]) FAILED: %s" % (arr_type.__name__, ex3))
                except Exception as ex2:
                    info.append("  Array attempt FAILED: %s" % ex2)

                info.append("")
                info.append("  --- GetMeansBelowCentile ---")
                try:
                    val = hist.GetMeansBelowCentile(99.9)
                    info.append("  GetMeansBelowCentile(99.9) = %s (type: %s)" % (val, type(val)))
                except Exception as ex2:
                    info.append("  GetMeansBelowCentile(99.9) FAILED: %s" % ex2)

                info.append("")
                info.append("  --- Values ---")
                try:
                    vals = hist.Values
                    info.append("  Type: %s" % type(vals))
                    info.append("  Channels: %d" % len(vals))
                    for ch in range(len(vals)):
                        ch_data = vals[ch]
                        info.append("  Channel %d: %d bins, max=%d, last_nonzero_idx=%s" % (
                            ch, len(ch_data), max(ch_data),
                            max([i for i in range(len(ch_data)) if ch_data[i] > 0]) if any(ch_data[i] > 0 for i in range(len(ch_data))) else "none"
                        ))
                except Exception as ex2:
                    info.append("  Values FAILED: %s" % ex2)

            except Exception as ex:
                info.append("  FAILED: %s" % ex)

            info.append("")
            info.append("=== frame.CutROI() ===")
            fi = frame.Info
            rx = fi.Width // 2 - 50
            ry = fi.Height // 2 - 50
            for desc, roi_args in [
                ("Rectangle", (Rectangle(rx, ry, 100, 100),)),
                ("4 ints", (rx, ry, 100, 100)),
                ("2 Points", (System.Drawing.Point(rx, ry), System.Drawing.Size(100, 100))),
            ]:
                try:
                    roi = frame.CutROI(*roi_args)
                    info.append("  (%s) Result: %s (type: %s)" % (desc, roi, type(roi)))
                    ri = roi.Info
                    info.append("  (%s) ROI size: %dx%d" % (desc, ri.Width, ri.Height))
                    try:
                        rbmp = roi.GetFrameBitmap()
                        info.append("  (%s) GetFrameBitmap: %dx%d (works!)" % (desc, rbmp.Width, rbmp.Height))
                    except Exception as ex2:
                        info.append("  (%s) GetFrameBitmap FAILED: %s" % (desc, ex2))
                    roi.Release()
                    info.append("  (%s) Release: OK" % desc)
                    break
                except Exception as ex:
                    info.append("  (%s) FAILED: %s" % (desc, ex))

        except Exception as ex:
            info.append("TOP-LEVEL EXCEPTION: %s" % str(ex))

    cam = SharpCap.SelectedCamera
    if cam is None:
        print("[Probe] ERROR: No camera selected!")
        return

    cam.BeforeFrameDisplay += probe_handler

    import time
    time.sleep(2)

    cam.BeforeFrameDisplay -= probe_handler

    print("[Probe] Results:")
    for line in _probe["info"]:
        print("  %s" % line)
    if not _probe["info"]:
        print("  No frames received!")
    print("")
    print("Copy everything above and paste it back for analysis.")


def reset_tracking():
    """Clear detection state and force re-acquisition of the donut."""
    _state["outer_cx"] = None
    _state["outer_cy"] = None
    _state["outer_r"] = None
    _state["inner_cx"] = None
    _state["inner_cy"] = None
    _state["inner_r"] = None
    _state["ref_peak"] = None
    _state["ref_outer_r"] = None
    _state["reject_count"] = 0
    _state["status"] = "Searching for donut..."
    print("[Collimation] Tracking reset")


def show_state():
    """Print current detection state and last error to the console."""
    print("[Collimation] State:")
    print("  enabled: %s" % _state["enabled"])
    print("  status: %s" % _state["status"])
    print("  overlays: circles=%s info=%s bar=%s arrows=%s crosshairs=%s" % (
        _config["SHOW_CIRCLES"], _config["SHOW_INFO_PANEL"],
        _config["SHOW_BOTTOM_BAR"], _config["SHOW_CORRECTION_ARROWS"],
        _config["SHOW_ALIGNMENT_CROSSHAIRS"]))
    print("  frames: %d" % _state["frame_count"])
    print("  outer: cx=%.1f cy=%.1f r=%.1f" % (
        _state["outer_cx"] or 0, _state["outer_cy"] or 0, _state["outer_r"] or 0))
    print("  inner: cx=%.1f cy=%.1f r=%.1f" % (
        _state["inner_cx"] or 0, _state["inner_cy"] or 0, _state["inner_r"] or 0))
    print("  ref_peak: %s" % _state.get("ref_peak"))
    print("  ref_outer_r: %s" % _state.get("ref_outer_r"))
    print("  reject_count: %d" % _state.get("reject_count", 0))
    print("  debug_log: %s" % _state.get("debug_log", False))
    if _state["outer_r"] and _state["inner_r"]:
        print("  inner/outer ratio: %.2f" % (_state["inner_r"] / _state["outer_r"]))
    print("  last_error: %s" % _state["last_error"])


def debug_on():
    """Enable per-frame debug logging to console."""
    _state["debug_log"] = True
    _state["debug_frames"] = 0
    print("[Collimation] Debug logging ON - each analysis frame will print details")
    print("[Collimation] Use debug_off() to stop")


def debug_off():
    """Disable per-frame debug logging."""
    _state["debug_log"] = False
    print("[Collimation] Debug logging OFF (logged %d frames)" % _state.get("debug_frames", 0))


# =============================================================================
# TEST MODE - synthetic donut for verifying the pipeline
# =============================================================================

def test_synthetic(width=800, height=600, outer_r=220, inner_r=90,
                   offset_x=15, offset_y=-10):
    """Generate a synthetic donut image and run analysis on it."""
    print("[Test] Generating synthetic donut: %dx%d" % (width, height))
    print("[Test] Outer R=%d, Inner R=%d, Offset=(%d, %d)" % (outer_r, inner_r, offset_x, offset_y))

    cx = width / 2.0
    cy = height / 2.0
    icx = cx + offset_x
    icy = cy + offset_y

    pixels = []
    for y in range(height):
        for x in range(width):
            dx = x - cx
            dy = y - cy
            dist_outer = math.sqrt(dx * dx + dy * dy)

            dxi = x - icx
            dyi = y - icy
            dist_inner = math.sqrt(dxi * dxi + dyi * dyi)

            if dist_inner < inner_r:
                val = 5.0
            elif dist_outer < outer_r:
                inner_fade = min(1.0, (dist_inner - inner_r) / 15.0) if dist_inner > inner_r else 0.0
                outer_fade = min(1.0, (outer_r - dist_outer) / 20.0) if dist_outer < outer_r else 0.0
                val = 200.0 * inner_fade * outer_fade
                val += 30.0 * max(0, math.cos(dist_outer * 0.3)) * inner_fade * outer_fade
            else:
                val = 2.0

            pixels.append(max(0.0, val))

    print("[Test] Running analysis...")
    result = analyze_donut_floats(pixels, width, height)

    if result is None:
        print("[Test] FAILED - analysis returned None")
        return

    print("[Test] SUCCESS!")
    print("[Test] Detected outer: cx=%.1f cy=%.1f r=%.1f (expected cx=%.1f cy=%.1f r=%d)" % (
        result["outer_cx"], result["outer_cy"], result["outer_r"], cx, cy, outer_r))
    print("[Test] Detected inner: cx=%.1f cy=%.1f r=%.1f (expected cx=%.1f cy=%.1f r=%d)" % (
        result["inner_cx"], result["inner_cy"], result["inner_r"], icx, icy, inner_r))

    dx = result["inner_cx"] - result["outer_cx"]
    dy = result["inner_cy"] - result["outer_cy"]
    offset = math.sqrt(dx * dx + dy * dy)
    expected_offset = math.sqrt(offset_x * offset_x + offset_y * offset_y)
    print("[Test] Detected offset: %.1f px (expected: %.1f px)" % (offset, expected_offset))

    update_state_with_result(result)
    _state["status"] = "Test mode"
    print("[Test] State updated - overlay should now draw on the live view!")


def test_image(path):
    """Load an image file and run donut analysis on it."""
    print("[Test] Loading image: %s" % path)
    try:
        bmp = Bitmap(path)
        print("[Test] Image size: %dx%d" % (bmp.Width, bmp.Height))
        result = analyze_bitmap(bmp)
        bmp.Dispose()

        if result is None:
            print("[Test] FAILED - no donut detected in image")
            return

        print("[Test] SUCCESS!")
        print("[Test] Outer: cx=%.1f cy=%.1f r=%.1f" % (
            result["outer_cx"], result["outer_cy"], result["outer_r"]))
        print("[Test] Inner: cx=%.1f cy=%.1f r=%.1f" % (
            result["inner_cx"], result["inner_cy"], result["inner_r"]))

        dx = result["inner_cx"] - result["outer_cx"]
        dy = result["inner_cy"] - result["outer_cy"]
        offset = math.sqrt(dx * dx + dy * dy)
        pct = (offset / result["outer_r"] * 100.0) if result["outer_r"] > 0 else 0
        print("[Test] Offset: %.1f px (%.1f%%)" % (offset, pct))

        update_state_with_result(result)
        _state["status"] = "Test image"
        print("[Test] State updated - overlay should now draw!")
    except Exception as ex:
        print("[Test] Error: %s" % str(ex))


# =============================================================================
# SETTINGS DIALOG
# =============================================================================
# WinForms-based non-modal settings palette. Opens via toolbar button or
# settings() console command. Changes apply live to the running overlay.

# Metadata for building the settings UI: (key, label, tab, control_type, min, max, step)
_SETTINGS_META = [
    # --- Overlay Display tab ---
    ("SHOW_CIRCLES",              "Detection Circles",          "Display", "bool", None, None, None),
    ("SHOW_ALIGNMENT_CROSSHAIRS", "Centering Crosshairs",       "Display", "bool", None, None, None),
    ("SHOW_INFO_PANEL",           "Overlay Details",            "Display", "bool", None, None, None),
    ("SHOW_BOTTOM_BAR",           "Bottom Numerical Details",   "Display", "bool", None, None, None),
    ("SHOW_CORRECTION_ARROWS",    "Correction Arrows",          "Display", "bool", None, None, None),
    # --- Appearance tab ---
    ("OUTER_CIRCLE_COLOR",    "Outer Circle Color",    "Colors", "color", None, None, None),
    ("INNER_CIRCLE_COLOR",    "Inner Circle Color",    "Colors", "color", None, None, None),
    ("CROSSHAIR_COLOR",       "Crosshair Color",       "Colors", "color", None, None, None),
    ("OFFSET_LINE_COLOR",     "Offset/Arrow Color",    "Colors", "color", None, None, None),
    ("TEXT_COLOR",             "Text Color",            "Colors", "color", None, None, None),
    ("TEXT_BG_COLOR",          "Text Background Color", "Colors", "color", None, None, None),
    ("CIRCLE_PEN_WIDTH",      "Circle Line Width",     "Colors", "float", 0.5, 10.0, 0.5),
    ("CROSSHAIR_PEN_WIDTH",   "Crosshair Line Width",  "Colors", "float", 0.5, 10.0, 0.5),
    ("OFFSET_PEN_WIDTH",      "Offset Line Width",     "Colors", "float", 0.5, 10.0, 0.5),
    ("CROSSHAIR_LENGTH",      "Crosshair Length (px)", "Colors", "int",   5, 100, 1),
    ("TEXT_FONT_SIZE",         "Info Panel Font Size",  "Colors", "int",   6, 24, 1),
    ("BOTTOM_BAR_FONT_SIZE",  "Bottom Bar Font Size",  "Colors", "int",   6, 24, 1),
    ("ARROW_LABEL_FONT_SIZE", "Arrow Label Font Size", "Colors", "int",   6, 24, 1),
    # --- Detection tab ---
    ("NUM_RAYS",                    "Number of Rays",          "Detection", "int",   12, 360, 1),
    ("BRIGHTNESS_THRESHOLD_PERCENT","Brightness Threshold (%)", "Detection", "int",   1, 80, 1),
    ("CENTROID_DOWNSAMPLE",         "Centroid Downsample",     "Detection", "int",   1, 16, 1),
    ("MIN_DONUT_RADIUS",            "Min Donut Radius (px)",   "Detection", "int",   5, 200, 1),
    ("RAY_STEP",                    "Ray Step (px)",           "Detection", "int",   1, 5, 1),
    ("ROI_MARGIN_FACTOR",           "ROI Margin Factor",       "Detection", "float", 1.0, 5.0, 0.1),
    ("SMOOTHING_FACTOR",            "Smoothing Factor",        "Detection", "float", 0.0, 0.95, 0.05),
    ("MAX_RADIUS_CHANGE_PCT",       "Max Radius Change (%)",   "Detection", "int",   1, 50, 1),
    ("MAX_CENTER_JUMP_PCT",         "Max Center Jump (%)",     "Detection", "int",   1, 50, 1),
    ("STALE_FRAME_LIMIT",           "Stale Frame Limit",       "Detection", "int",   3, 100, 1),
    # --- Performance tab ---
    ("ANALYSIS_INTERVAL",           "Analysis Interval",       "Performance", "int", 1, 10, 1),
]


def _build_settings_form():
    """Create and return the settings Form with all controls."""
    form = Form()
    form.Text = "Live Collimation Aid"
    form.Width = 270
    form.Height = 400
    form.FormBorderStyle = FormBorderStyle.FixedToolWindow
    form.MaximizeBox = False
    form.MinimizeBox = False
    form.StartPosition = FormStartPosition.CenterScreen
    form.TopMost = True
    try:
        form.AutoScaleMode = WinFormsAutoScaleMode.Dpi
    except:
        pass

    # Tab control
    tabs = TabControl()
    tabs.Dock = DockStyle.Fill
    tabs.Padding = System.Drawing.Point(3, 3)

    # Create tab pages
    tab_pages = {}
    for tab_name in ["Display", "Colors", "Detection", "Performance"]:
        tp = TabPage()
        tp.Text = tab_name
        tp.AutoScroll = True
        tp.Padding = Padding(4, 4, 4, 4)
        tab_pages[tab_name] = tp
        tabs.TabPages.Add(tp)

    # Bottom panel with buttons
    bottom = Panel()
    bottom.Height = 36
    bottom.Dock = DockStyle.Bottom

    btn_reset = Button()
    btn_reset.Text = "Reset"
    btn_reset.Width = 60
    btn_reset.Height = 24
    btn_reset.Left = 6
    btn_reset.Top = 5

    btn_save = Button()
    btn_save.Text = "Save"
    btn_save.Width = 60
    btn_save.Height = 24
    btn_save.Left = form.Width - 82
    btn_save.Top = 5

    bottom.Controls.Add(btn_reset)
    bottom.Controls.Add(btn_save)

    form.Controls.Add(tabs)
    form.Controls.Add(bottom)

    # Track all controls for refresh on reset
    control_map = {}  # key -> (control, extra1, extra2)

    # Layout: stack controls vertically per tab
    tab_y = {}  # current y position per tab
    lbl_w = 120  # label width
    ctrl_left = 125  # control x position

    # Master enable/disable checkbox at top of Display tab
    display_controls = []  # track Display tab controls to enable/disable

    cb_enable = CheckBox()
    cb_enable.Text = "Enable Overlay"
    cb_enable.Left = 4
    cb_enable.Top = 4
    cb_enable.Width = 220
    cb_enable.Height = 22
    cb_enable.Checked = _state["enabled"]
    font_bold = Font(cb_enable.Font, FontStyle.Bold)
    cb_enable.Font = font_bold
    tab_pages["Display"].Controls.Add(cb_enable)
    tab_y["Display"] = 32

    def on_enable_change(sender, e):
        _state["enabled"] = sender.Checked
        for c in display_controls:
            c.Enabled = sender.Checked
    cb_enable.CheckedChanged += on_enable_change

    for key, label_text, tab_name, ctrl_type, c_min, c_max, c_step in _SETTINGS_META:
        tp = tab_pages[tab_name]
        y = tab_y.get(tab_name, 4)

        lbl = Label()
        lbl.Text = label_text
        lbl.Left = 4
        lbl.Top = y + 3
        lbl.Width = lbl_w
        lbl.Height = 18
        tp.Controls.Add(lbl)

        if ctrl_type == "bool":
            tp.Controls.Remove(lbl)

            # Circles and crosshairs are mutually exclusive
            if key in ("SHOW_CIRCLES", "SHOW_ALIGNMENT_CROSSHAIRS"):
                from System.Windows.Forms import RadioButton
                rb = RadioButton()
                rb.Text = label_text
                rb.Left = 4
                rb.Top = y
                rb.Width = 220
                rb.Height = 22
                rb.Checked = bool(_config[key])
                tp.Controls.Add(rb)

                def make_radio_change(k, other_k):
                    def on_change(sender, e):
                        if sender.Checked:
                            _config[k] = True
                            _config[other_k] = False
                            # Update the other radio button's control
                            other_entry = control_map.get(other_k)
                            if other_entry and other_entry[0] is not None:
                                other_entry[0].Checked = False
                    return on_change

                other_key = "SHOW_ALIGNMENT_CROSSHAIRS" if key == "SHOW_CIRCLES" else "SHOW_CIRCLES"
                rb.CheckedChanged += make_radio_change(key, other_key)

                control_map[key] = (rb, None, None)
                if tab_name == "Display":
                    rb.Enabled = _state["enabled"]
                    display_controls.append(rb)
                tab_y[tab_name] = y + 24
            else:
                cb = CheckBox()
                cb.Text = label_text
                cb.Left = 4
                cb.Top = y
                cb.Width = 220
                cb.Height = 22
                cb.Checked = bool(_config[key])
                tp.Controls.Add(cb)

                def make_bool_change(k):
                    def on_change(sender, e):
                        _config[k] = sender.Checked
                    return on_change
                cb.CheckedChanged += make_bool_change(key)

                control_map[key] = (cb, None, None)
                if tab_name == "Display":
                    cb.Enabled = _state["enabled"]
                    display_controls.append(cb)
                tab_y[tab_name] = y + 24

        elif ctrl_type == "color":
            color_val = _config[key]

            pnl = Panel()
            pnl.Left = ctrl_left
            pnl.Top = y
            pnl.Width = 90
            pnl.Height = 20
            pnl.BackColor = Color.FromArgb(255, color_val.R, color_val.G, color_val.B)
            pnl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            tp.Controls.Add(pnl)

            def make_color_click(k, p):
                def on_click(sender, e):
                    dlg = ColorDialog()
                    cur = _config[k]
                    dlg.Color = Color.FromArgb(255, cur.R, cur.G, cur.B)
                    dlg.FullOpen = True
                    if dlg.ShowDialog() == DialogResult.OK:
                        c = dlg.Color
                        _config[k] = Color.FromArgb(cur.A, c.R, c.G, c.B)
                        p.BackColor = Color.FromArgb(255, c.R, c.G, c.B)
                return on_click
            pnl.Click += make_color_click(key, pnl)

            control_map[key] = (pnl, None, None)
            tab_y[tab_name] = y + 24

        elif ctrl_type == "float":
            nud = NumericUpDown()
            nud.Left = ctrl_left
            nud.Top = y
            nud.Width = 65
            nud.Height = 22
            nud.DecimalPlaces = 2
            nud.Minimum = System.Decimal(c_min)
            nud.Maximum = System.Decimal(c_max)
            nud.Increment = System.Decimal(c_step)
            nud.Value = System.Decimal(float(_config[key]))
            tp.Controls.Add(nud)

            def make_float_change(k):
                def on_change(sender, e):
                    _config[k] = float(sender.Value)
                return on_change
            nud.ValueChanged += make_float_change(key)

            control_map[key] = (nud, None, None)
            tab_y[tab_name] = y + 26

        elif ctrl_type == "int":
            nud = NumericUpDown()
            nud.Left = ctrl_left
            nud.Top = y
            nud.Width = 65
            nud.Height = 22
            nud.DecimalPlaces = 0
            nud.Minimum = System.Decimal(c_min)
            nud.Maximum = System.Decimal(c_max)
            nud.Increment = System.Decimal(c_step)
            nud.Value = System.Decimal(int(_config[key]))
            tp.Controls.Add(nud)

            def make_int_change(k):
                def on_change(sender, e):
                    _config[k] = int(sender.Value)
                return on_change
            nud.ValueChanged += make_int_change(key)

            control_map[key] = (nud, None, None)
            tab_y[tab_name] = y + 26

    # --- Button handlers ---
    def on_reset(sender, e):
        reset_config()
        # Refresh all controls
        for key, label_text, tab_name, ctrl_type, c_min, c_max, c_step in _SETTINGS_META:
            entry = control_map.get(key)
            if entry is None:
                continue
            ctrl, extra1, extra2 = entry
            if ctrl_type == "bool":
                ctrl.Checked = bool(_config[key])
            elif ctrl_type == "color":
                color_val = _config[key]
                ctrl.BackColor = Color.FromArgb(255, color_val.R, color_val.G, color_val.B)
            elif ctrl_type == "float":
                ctrl.Value = System.Decimal(float(_config[key]))
            elif ctrl_type == "int":
                ctrl.Value = System.Decimal(int(_config[key]))
        print("[Collimation] Settings reset to defaults in dialog")

    def on_save(sender, e):
        save_config()

    btn_reset.Click += on_reset
    btn_save.Click += on_save

    return form


def open_settings():
    """Open the settings dialog (or bring to front if already open)."""
    sf = _state.get("settings_form")
    if sf is not None:
        try:
            if not sf.IsDisposed:
                sf.BringToFront()
                return
        except:
            pass

    form = _build_settings_form()

    def on_closed(sender, e):
        _state["settings_form"] = None

    form.FormClosed += on_closed
    _state["settings_form"] = form
    form.Show()


def settings():
    """Console command: open the settings dialog."""
    open_settings()


# =============================================================================
# SETUP AND TEARDOWN
# =============================================================================

def _make_icon(icon_type):
    """Create a 24x24 toolbar icon bitmap using GDI+ drawing."""
    size = 24
    bmp = Bitmap(size, size)
    g = System.Drawing.Graphics.FromImage(bmp)
    g.Clear(Color.Transparent)

    if icon_type == "toggle":
        p = Pen(Color.FromArgb(255, 220, 60, 60), 2.0)
        g.DrawEllipse(p, 2, 2, 19, 19)
        p.Dispose()
        p = Pen(Color.FromArgb(255, 60, 220, 60), 2.0)
        g.DrawEllipse(p, 7, 7, 9, 9)
        p.Dispose()
        b = SolidBrush(Color.Yellow)
        g.FillRectangle(b, RectangleF(10.0, 10.0, 3.0, 3.0))
        b.Dispose()

    elif icon_type == "settings":
        # Wrench/gear icon
        p = Pen(Color.FromArgb(255, 200, 200, 200), 2.0)
        # Outer gear ring
        g.DrawEllipse(p, 4, 4, 15, 15)
        p.Dispose()
        # Inner dot
        b = SolidBrush(Color.FromArgb(255, 200, 200, 200))
        g.FillRectangle(b, RectangleF(9.0, 9.0, 5.0, 5.0))
        b.Dispose()
        # Gear teeth (4 lines crossing the circle)
        p = Pen(Color.FromArgb(255, 200, 200, 200), 2.0)
        g.DrawLine(p, 11.0, 1.0, 11.0, 5.0)   # top
        g.DrawLine(p, 11.0, 18.0, 11.0, 22.0)  # bottom
        g.DrawLine(p, 1.0, 11.0, 5.0, 11.0)    # left
        g.DrawLine(p, 18.0, 11.0, 22.0, 11.0)  # right
        p.Dispose()

    elif icon_type == "reset":
        p = Pen(Color.FromArgb(255, 100, 180, 255), 2.0)
        g.DrawArc(p, 4, 4, 15, 15, 0, 270)
        p.Dispose()
        p = Pen(Color.FromArgb(255, 100, 180, 255), 2.0)
        g.DrawLine(p, 11.0, 4.0, 19.0, 4.0)
        g.DrawLine(p, 19.0, 4.0, 19.0, 10.0)
        p.Dispose()

    g.Dispose()
    return bmp


def _add_toolbar_button(btn_name, icon, tooltip, callback):
    """Try to add a toolbar button using the 3 known arg orderings."""
    for arg_order in [
        (btn_name, icon, tooltip, callback),
        (btn_name, callback, icon, tooltip),
        (btn_name, callback, tooltip, icon),
    ]:
        try:
            SharpCap.AddCustomButton(*arg_order)
            return True
        except:
            pass
    return False


def attach_handler():
    """Register the overlay event handler on the selected camera."""
    if _state["handler_attached"]:
        print("[Collimation] Handler already attached")
        return

    cam = SharpCap.SelectedCamera
    if cam is None:
        print("[Collimation] ERROR: No camera selected!")
        return

    cam.BeforeFrameDisplay += on_before_frame_display
    _state["handler_attached"] = True
    print("[Collimation] Handler attached to: %s" % cam.DeviceName)


def detach_handler():
    """Unregister the overlay event handler from the camera."""
    if not _state["handler_attached"]:
        print("[Collimation] No handler to detach")
        return
    try:
        cam = SharpCap.SelectedCamera
        if cam is not None:
            cam.BeforeFrameDisplay -= on_before_frame_display
    except:
        pass
    _state["handler_attached"] = False
    print("[Collimation] Handler detached")


def stop():
    """Fully stop the overlay: detach handler, remove toolbar button, disable."""
    detach_handler()
    for btn_name in _state.get("buttons", []):
        try:
            SharpCap.RemoveCustomButton(btn_name)
            print("[Collimation] Removed button: %s" % btn_name)
        except:
            pass
    _state["buttons"] = []
    _state["enabled"] = False
    sf = _state.get("settings_form")
    if sf is not None:
        try:
            sf.Close()
            sf.Dispose()
        except:
            pass
        _state["settings_form"] = None
    print("[Collimation] Stopped. Run the script again to restart.")


def setup():
    """
    Initialize the collimation overlay:
      1. Add toolbar buttons (Overlay + Settings)
      2. Attach the BeforeFrameDisplay handler to the camera
      3. Print usage instructions to the console
    """
    print("=" * 50)
    print("  COLLIMATION OVERLAY v2.0")
    print("=" * 50)

    buttons = []

    # Settings button
    btn_settings = "Coll Settings"
    icon_settings = _make_icon("settings")
    if _add_toolbar_button(btn_settings, icon_settings, "Open collimation settings", open_settings):
        buttons.append(btn_settings)
        print("[OK] Toolbar button: '%s'" % btn_settings)
    else:
        print("[WARN] Could not add settings button")
        print("  Use settings() in console")

    _state["buttons"] = buttons

    attach_handler()

    print("")
    print("Console commands:")
    print("  settings()             - open settings dialog")
    print("  reset_tracking()       - re-detect donut")
    print("  show_state()           - debug info")
    print("  debug_on() / debug_off() - per-frame logging")
    print("  stop()                 - fully stop + remove button")
    print("=" * 50)


# =============================================================================
# AUTO-START
# =============================================================================
# When this script is executed (via File > Run Script or exec()), setup()
# runs automatically, registering the handler and creating the toolbar buttons.
# The script then exits, but the handler remains active on every frame.

setup()
