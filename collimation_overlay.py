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
#
# Requires: SharpCap Pro 4.1+ (scripting is a Pro feature)
# =============================================================================

import clr
import math
import time

# .NET assemblies for drawing and pixel access
clr.AddReference("System.Drawing")
from System.Drawing import (
    Color, Pen, SolidBrush, Font, FontFamily, FontStyle,
    Bitmap, Rectangle, RectangleF, PointF,
)
from System.Drawing.Imaging import ImageLockMode, PixelFormat
from System.Drawing.Drawing2D import SmoothingMode, DashStyle
from System.Runtime.InteropServices import Marshal
from System import Array, Byte
import System.Drawing

# =============================================================================
# USER CONFIGURATION
# =============================================================================
# Edit these values to customize the overlay appearance and behavior.
# All colors use ARGB format: Color.FromArgb(alpha, red, green, blue)

# Overlay colors
OUTER_CIRCLE_COLOR = Color.FromArgb(220, 255, 60, 60)     # Red - primary mirror edge
INNER_CIRCLE_COLOR = Color.FromArgb(220, 60, 255, 60)     # Green - secondary shadow edge
CROSSHAIR_COLOR    = Color.FromArgb(180, 255, 255, 0)     # Yellow - outer center marker
OFFSET_LINE_COLOR  = Color.FromArgb(220, 0, 220, 255)     # Cyan - offset vector & arrows
TEXT_COLOR          = Color.FromArgb(240, 255, 255, 255)   # White - info panel text
TEXT_BG_COLOR       = Color.FromArgb(160, 0, 0, 0)         # Semi-transparent black background

# Line widths (in pixels)
CIRCLE_PEN_WIDTH    = 2.0
CROSSHAIR_PEN_WIDTH = 1.0
OFFSET_PEN_WIDTH    = 2.0

# Crosshair size (pixels extending from center, used in modes 0-2)
CROSSHAIR_LENGTH = 20

# Text sizes (in points)
TEXT_FONT_SIZE = 9              # Info panel and status text
BOTTOM_BAR_FONT_SIZE = 10      # Bottom bar offset readout
ARROW_LABEL_FONT_SIZE = 8      # Correction arrow labels (MOVE X/Y)

# Detection parameters
NUM_RAYS = 90                       # Radial rays for edge detection. More = accurate, slower.
BRIGHTNESS_THRESHOLD_PERCENT = 30   # % of peak brightness that defines the donut edge.
                                    # Lower = more sensitive (picks up faint edges).
                                    # Higher = stricter (only bright edges).
CENTROID_DOWNSAMPLE = 4             # Sample every Nth pixel when computing centroid.
                                    # Higher = faster but less precise initial center.
MIN_DONUT_RADIUS = 15               # Ignore detected circles smaller than this (pixels).
RAY_STEP = 1                        # Pixel step along each ray. 1 = every pixel.
ROI_MARGIN_FACTOR = 2.0             # CutROI margin as multiple of outer radius.
                                    # 2.0 = 100% padding around the donut for tracking.

# Tracking validation: reject new detections that differ too much from
# the current smoothed state. Prevents drift from bad frames.
MAX_RADIUS_CHANGE_PCT = 15          # Max % change in outer radius per detection.
MAX_CENTER_JUMP_PCT = 20            # Max center jump as % of outer radius.
STALE_FRAME_LIMIT = 15              # After this many consecutive rejected frames,
                                    # force a full re-detection from scratch.

# Smoothing: exponential moving average applied to circle positions/radii.
# 0.0 = no smoothing (instant response, may jitter)
# 0.9 = heavy smoothing (very stable, slow to respond to changes)
SMOOTHING_FACTOR = 0.6

# Analyze every Nth frame. Higher = less CPU but slower response.
# 1 = every frame, 2 = every other frame, etc.
ANALYSIS_INTERVAL = 2

# =============================================================================
# INTERNAL STATE
# =============================================================================

# If the script is re-run, clean up the previous instance's event handler
# and toolbar buttons to prevent duplicates. This works because IronPython
# keeps global variables alive across script runs in the same console session.
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
    "display_mode": 0,      # Index into _DISPLAY_MODE_NAMES
    "ref_peak": None,       # Reference peak brightness from initial detection
    "ref_outer_r": None,    # Reference outer radius from initial detection (for ROI sizing)
    "reject_count": 0,      # Consecutive rejected detection count
    "debug_log": False,     # Toggle with debug_on() / debug_off()
    "debug_frames": 0,      # Frames logged since debug enabled
}

# Display modes cycled by the toolbar button. The last mode ("Off") disables
# the overlay entirely. Cycling past Off wraps back to the first mode.
_DISPLAY_MODE_NAMES = [
    "All overlays",                     # 0: info panel + bottom bar + crosshairs + arrows + circles
    "Hide info panel",                  # 1: bottom bar + crosshairs + arrows + circles
    "Hide info panel + bottom bar",     # 2: crosshairs + arrows + circles
    "Circles & arrows only",            # 3: just circles + correction arrows
    "Alignment crosshairs",             # 4: circles with full-diameter crosshairs for visual alignment
    "Off",                              # 5: no overlay, handler returns immediately
]


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

    if radius < MIN_DONUT_RADIUS * 0.5:
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
    step = CENTROID_DOWNSAMPLE
    if peak_override is not None and approx_center is not None:
        # Fastest path: both peak and center provided (histogram + tracking).
        # Skip the pixel scan entirely.
        peak = peak_override
        # Use reference peak as floor to prevent threshold collapse when dimming
        ref = _state.get("ref_peak")
        if ref is not None:
            peak = max(peak, ref * 0.5)
        threshold = peak * BRIGHTNESS_THRESHOLD_PERCENT / 100.0
        approx_cx, approx_cy = approx_center
    elif approx_center is not None:
        # ROI path: center known from tracking, but scan pixels for accurate
        # peak brightness. Used with CutROI where the scan is on a small region.
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
        # Use reference peak as floor to prevent threshold collapse when dimming
        ref = _state.get("ref_peak")
        if ref is not None:
            peak = max(peak, ref * 0.5)
        threshold = peak * BRIGHTNESS_THRESHOLD_PERCENT / 100.0
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
        threshold = peak * BRIGHTNESS_THRESHOLD_PERCENT / 100.0

    # --- Step 2: Refine center using only bright pixels near the centroid ---
    # The coarse centroid can be pulled off by background gradients. Here we
    # restrict to a region around the initial estimate and only use bright pixels.
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
    # Each ray walks pixel-by-pixel from the center outward, looking for
    # brightness transitions. The first dark->bright transition is the inner
    # edge (boundary of the secondary shadow).
    #
    # For the outer edge, instead of a single-pixel threshold crossing (which
    # is very noisy on dim images), we find the point of steepest brightness
    # decline. This is robust to absolute brightness variations because it
    # detects the edge shape rather than an absolute level.
    inner_points = []
    outer_points = []
    max_ray_len = int(min(width, height) * 0.45)
    grad_window = 5  # Pixels to average for gradient calculation

    for i in range(NUM_RAYS):
        angle = 2.0 * math.pi * i / NUM_RAYS
        cos_a = math.cos(angle)
        sin_a = math.sin(angle)

        found_inner = False
        in_bright = False
        ray_samples = []  # (dist, px, py, brightness) collected after inner edge

        for dist in range(3, max_ray_len, RAY_STEP):
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
                # Continue a bit past the first dark pixel to capture the full gradient
                pass
            if found_inner and not is_bright and len(ray_samples) > grad_window * 3:
                # We're past the bright region and have enough samples for
                # gradient calculation. Stop once brightness is very low.
                if b < 5.0:
                    break

        # Find outer edge as point of steepest RELATIVE brightness decline.
        # Using relative gradient (pct drop) instead of absolute ensures that
        # dim parts of the donut find the edge just as reliably as bright parts.
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
                # Relative gradient: fraction of brightness that drops
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
    if ir >= ore:
        return None  # Inner can't be larger than outer
    if ir < MIN_DONUT_RADIUS * 0.3:
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

    pixels: Python list or .NET float array, length = width * height
    """
    # --- Step 1: Find peak brightness and brightness-weighted centroid ---
    sum_bx = 0.0
    sum_by = 0.0
    sum_b  = 0.0
    peak   = 0.0
    step   = CENTROID_DOWNSAMPLE

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

    threshold = peak * BRIGHTNESS_THRESHOLD_PERCENT / 100.0

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

    for i in range(NUM_RAYS):
        angle = 2.0 * math.pi * i / NUM_RAYS
        cos_a = math.cos(angle)
        sin_a = math.sin(angle)

        found_inner = False
        in_bright = False
        ray_samples = []

        for dist in range(3, max_ray_len, RAY_STEP):
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
    if ir < MIN_DONUT_RADIUS * 0.3:
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
# Exponential moving average to stabilize circle positions across frames.
# Without smoothing, the overlay can jitter due to noise in edge detection.

def smooth_val(old, new, factor):
    """Apply exponential smoothing: result = old*factor + new*(1-factor)."""
    if old is None:
        return new
    return old * factor + new * (1.0 - factor)


def _validate_detection(result):
    """Check if a new detection is consistent with the current smoothed state.
    Returns True if the result should be accepted, False if it should be rejected.
    Always accepts when there is no previous state (initial detection)."""
    old_r = _state["outer_r"]
    if old_r is None:
        return True  # No previous state, accept anything

    # Check outer radius didn't change too much
    new_r = result["outer_r"]
    radius_change_pct = abs(new_r - old_r) / old_r * 100.0
    if radius_change_pct > MAX_RADIUS_CHANGE_PCT:
        return False

    # Check outer center didn't jump too far
    dx = result["outer_cx"] - _state["outer_cx"]
    dy = result["outer_cy"] - _state["outer_cy"]
    center_jump = math.sqrt(dx * dx + dy * dy)
    if center_jump > old_r * MAX_CENTER_JUMP_PCT / 100.0:
        return False

    return True


def update_state_with_result(result):
    """Validate and apply smoothing to new analysis results.
    Rejects detections that differ too much from the current state.
    After too many consecutive rejections, forces a full re-detection."""
    is_debug = _state.get("debug_log", False)

    if not _validate_detection(result):
        _state["reject_count"] = _state.get("reject_count", 0) + 1
        if is_debug:
            print("[DBG] REJECTED frame %d (count=%d) raw_or=%.1f spread=%.1f raw_ocx=%.1f raw_ocy=%.1f | smooth_or=%.1f smooth_ocx=%.1f" % (
                _state["frame_count"], _state["reject_count"],
                result["outer_r"], result.get("outer_spread", 0),
                result["outer_cx"], result["outer_cy"],
                _state["outer_r"] or 0, _state["outer_cx"] or 0))
        if _state["reject_count"] >= STALE_FRAME_LIMIT:
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
    f = SMOOTHING_FACTOR

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
# These use SharpCap APIs revealed by the author to avoid expensive IronPython
# pixel loops and full-frame buffer copies when possible.

def _get_histogram_peak(frame):
    """Get approximate peak brightness from frame histogram (native, fast).
    Uses GetCentilePointsF(99.9) which returns a per-channel array of floats.
    Averages the first 3 channels (RGB) to match _get_brightness() behavior,
    which computes (R+G+B)/3. For mono cameras (1 channel), returns that value."""
    try:
        hist = frame.CalculateHistogram()
        vals = hist.GetCentilePointsF(99.9)
        # vals is Array[Single] with one entry per channel (e.g., RGBA=4 values).
        # Average the first 3 (RGB), skipping channel 4 (alpha/luminance) which
        # can be much higher and would inflate the threshold.
        n = min(3, len(vals))
        total = 0.0
        for i in range(n):
            total += float(vals[i])
        peak = total / n
        return peak if peak > 0 else None
    except:
        return None


def _try_cut_roi(frame):
    """Cut a region of interest around the tracked donut position.
    Returns (roi_frame, offset_x, offset_y) or (None, 0, 0) on failure.
    Caller MUST call roi_frame.Release() when done."""
    try:
        # Use reference radius (from initial detection) for ROI sizing to prevent
        # feedback loop where shrinking radius -> smaller ROI -> fewer rays -> smaller radius
        roi_r = _state.get("ref_outer_r") or _state["outer_r"]
        margin = int(roi_r * ROI_MARGIN_FACTOR)
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
# All drawing uses SharpCap's WinformGraphicsImpl wrapper, obtained via
# frame.GetDrawableBitmap().GetGraphics(). This wrapper supports standard
# GDI+ drawing methods (DrawEllipse, DrawLine, DrawString, FillRectangle)
# but FillRectangle requires a RectangleF object instead of separate x/y/w/h.
# SmoothingMode and DashStyle are wrapped in try/except as they may not be
# supported on all wrapper versions.

def draw_overlay(gfx, width, height):
    """Main overlay drawing function. Respects current display_mode."""
    try:
        gfx.SmoothingMode = SmoothingMode.AntiAlias
    except:
        pass

    mode = _state["display_mode"]
    ocx = _state["outer_cx"]
    ocy = _state["outer_cy"]
    ore = _state["outer_r"]
    icx = _state["inner_cx"]
    icy = _state["inner_cy"]
    ir  = _state["inner_r"]

    if ocx is None or icx is None:
        if mode < 3:
            draw_status_text(gfx, "Searching for donut...")
        return

    # --- Mode 4: Alignment crosshairs ---
    # Draws both circles with full-diameter crosshairs in their respective
    # colors. When collimation is perfect, the red and green crosshairs
    # overlap exactly. This is a minimal mode for precise visual alignment.
    if mode == 4:
        # Outer circle + full-diameter crosshair in red
        pen = Pen(OUTER_CIRCLE_COLOR, CIRCLE_PEN_WIDTH)
        gfx.DrawEllipse(pen, float(ocx - ore), float(ocy - ore), float(ore * 2), float(ore * 2))
        gfx.DrawLine(pen, float(ocx - ore), float(ocy), float(ocx + ore), float(ocy))
        gfx.DrawLine(pen, float(ocx), float(ocy - ore), float(ocx), float(ocy + ore))
        pen.Dispose()

        # Inner circle + full-diameter crosshair in green
        pen = Pen(INNER_CIRCLE_COLOR, CIRCLE_PEN_WIDTH)
        gfx.DrawEllipse(pen, float(icx - ir), float(icy - ir), float(ir * 2), float(ir * 2))
        gfx.DrawLine(pen, float(icx - ir), float(icy), float(icx + ir), float(icy))
        gfx.DrawLine(pen, float(icx), float(icy - ir), float(icx), float(icy + ir))
        pen.Dispose()
        return  # Nothing else drawn in alignment mode

    # --- Circles (all modes 0-3) ---
    pen = Pen(OUTER_CIRCLE_COLOR, CIRCLE_PEN_WIDTH)
    gfx.DrawEllipse(pen, float(ocx - ore), float(ocy - ore), float(ore * 2), float(ore * 2))
    pen.Dispose()

    pen = Pen(INNER_CIRCLE_COLOR, CIRCLE_PEN_WIDTH)
    gfx.DrawEllipse(pen, float(icx - ir), float(icy - ir), float(ir * 2), float(ir * 2))
    pen.Dispose()

    # --- Center crosshairs (modes 0-2) ---
    # Yellow crosshair at outer center, green crosshair at inner center.
    # These are small fixed-size markers, unlike mode 4's full-diameter ones.
    if mode <= 2:
        pen = Pen(CROSSHAIR_COLOR, CROSSHAIR_PEN_WIDTH)
        ch = float(CROSSHAIR_LENGTH)
        gfx.DrawLine(pen, float(ocx - ch), float(ocy), float(ocx + ch), float(ocy))
        gfx.DrawLine(pen, float(ocx), float(ocy - ch), float(ocx), float(ocy + ch))
        pen.Dispose()

        pen = Pen(INNER_CIRCLE_COLOR, CROSSHAIR_PEN_WIDTH)
        ch2 = float(CROSSHAIR_LENGTH * 0.6)
        gfx.DrawLine(pen, float(icx - ch2), float(icy), float(icx + ch2), float(icy))
        gfx.DrawLine(pen, float(icx), float(icy - ch2), float(icx), float(icy + ch2))
        pen.Dispose()

    # --- Offset vector line between centers (modes 0-2) ---
    # Dashed cyan line from outer center to inner center with arrowhead.
    dx = icx - ocx
    dy = icy - ocy
    offset_dist = math.sqrt(dx * dx + dy * dy)

    if mode <= 2 and offset_dist > 0.5:
        pen = Pen(OFFSET_LINE_COLOR, OFFSET_PEN_WIDTH)
        try:
            pen.DashStyle = DashStyle.Dash
        except:
            pass
        gfx.DrawLine(pen, float(ocx), float(ocy), float(icx), float(icy))
        pen.Dispose()

        # Arrowhead at the inner-center end of the offset line
        if offset_dist > 3:
            arrow_len = min(12.0, offset_dist * 0.4)
            angle = math.atan2(dy, dx)
            a1 = angle + math.pi * 0.8
            a2 = angle - math.pi * 0.8
            pen = Pen(OFFSET_LINE_COLOR, OFFSET_PEN_WIDTH)
            gfx.DrawLine(pen,
                float(icx), float(icy),
                float(icx + arrow_len * math.cos(a1)), float(icy + arrow_len * math.sin(a1)))
            gfx.DrawLine(pen,
                float(icx), float(icy),
                float(icx + arrow_len * math.cos(a2)), float(icy + arrow_len * math.sin(a2)))
            pen.Dispose()

    # --- Correction arrows (all modes 0-3) ---
    # These arrows point in the direction the inner circle needs to MOVE to
    # become concentric with the outer circle (opposite of the offset vector).
    # X arrow is placed above the donut, Y arrow to the right.
    min_arrow_threshold = 1.5  # Don't show arrows for sub-pixel offsets
    arrow_gap = 15.0           # Gap between outer circle edge and arrow start
    arrow_shaft = 30.0
    arrow_head = 10.0
    arrow_pen_w = 2.5
    corr_x = -dx  # Correction = opposite of offset
    corr_y = -dy

    # X-axis correction arrow (horizontal, above the donut)
    if abs(corr_x) > min_arrow_threshold:
        ax_sign = 1.0 if corr_x > 0 else -1.0
        ax_y = float(ocy - ore - arrow_gap - 8)
        ax_start = float(ocx)
        ax_end = float(ocx + ax_sign * arrow_shaft)

        pen = Pen(OFFSET_LINE_COLOR, arrow_pen_w)
        gfx.DrawLine(pen, ax_start, ax_y, ax_end, ax_y)
        gfx.DrawLine(pen, ax_end, ax_y,
            float(ax_end - ax_sign * arrow_head * 0.7), float(ax_y - arrow_head * 0.5))
        gfx.DrawLine(pen, ax_end, ax_y,
            float(ax_end - ax_sign * arrow_head * 0.7), float(ax_y + arrow_head * 0.5))
        pen.Dispose()

        if mode <= 2:
            font = Font(FontFamily.GenericMonospace, ARROW_LABEL_FONT_SIZE, FontStyle.Bold)
            brush = SolidBrush(OFFSET_LINE_COLOR)
            if corr_x > 0:
                gfx.DrawString("MOVE X >>>", font, brush, float(ocx + 5), float(ax_y - 16))
            else:
                gfx.DrawString("<<< MOVE X", font, brush, float(ocx - arrow_shaft - 20), float(ax_y - 16))
            font.Dispose()
            brush.Dispose()

    # Y-axis correction arrow (vertical, right of the donut)
    if abs(corr_y) > min_arrow_threshold:
        ay_sign = 1.0 if corr_y > 0 else -1.0
        ay_x = float(ocx + ore + arrow_gap + 8)
        ay_start = float(ocy)
        ay_end = float(ocy + ay_sign * arrow_shaft)

        pen = Pen(OFFSET_LINE_COLOR, arrow_pen_w)
        gfx.DrawLine(pen, ay_x, ay_start, ay_x, ay_end)
        gfx.DrawLine(pen, ay_x, ay_end,
            float(ay_x - arrow_head * 0.5), float(ay_end - ay_sign * arrow_head * 0.7))
        gfx.DrawLine(pen, ay_x, ay_end,
            float(ay_x + arrow_head * 0.5), float(ay_end - ay_sign * arrow_head * 0.7))
        pen.Dispose()

        if mode <= 2:
            font = Font(FontFamily.GenericMonospace, ARROW_LABEL_FONT_SIZE, FontStyle.Bold)
            brush = SolidBrush(OFFSET_LINE_COLOR)
            # In screen coords +Y is down, but for user display +Y means "up"
            if corr_y < 0:
                gfx.DrawString("MOVE Y UP", font, brush, float(ay_x + 5), float(ocy - arrow_shaft - 5))
            else:
                gfx.DrawString("MOVE Y DN", font, brush, float(ay_x + 5), float(ocy + arrow_shaft - 5))
            font.Dispose()
            brush.Dispose()

    # --- Info panel (mode 0 only) ---
    if mode == 0:
        draw_info_panel(gfx, width, height, ocx, ocy, ore, icx, icy, ir, offset_dist)

    # --- Bottom bar (modes 0-1) ---
    if mode <= 1:
        draw_bottom_bar(gfx, width, height, dx, dy, offset_dist, ore)


def draw_bottom_bar(gfx, width, height, dx, dy, offset_dist, ore):
    """Draw the X/Y offset readout bar at the bottom of the frame.
    Shows correction direction and color-codes severity (green=good, red=bad)."""
    bar_h = float(BOTTOM_BAR_FONT_SIZE + 16)
    bar_y = float(height) - bar_h
    offset_pct = (offset_dist / ore * 100.0) if ore > 0 else 0.0

    bg = SolidBrush(Color.FromArgb(200, 0, 0, 0))
    gfx.FillRectangle(bg, RectangleF(0.0, bar_y, float(width), bar_h))
    bg.Dispose()

    font_lg = Font(FontFamily.GenericMonospace, BOTTOM_BAR_FONT_SIZE, FontStyle.Bold)
    font_sm = Font(FontFamily.GenericMonospace, BOTTOM_BAR_FONT_SIZE - 2.0, FontStyle.Regular)

    def offset_color(val):
        """Color-code an offset value: green (< 2%) through red (> 20%)."""
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

    # X offset with correction direction indicator
    x_color = offset_color(dx)
    x_brush = SolidBrush(x_color)
    x_corr = "<<<" if dx > 0 else ">>>"  # Correction = opposite of offset
    if abs(dx) < 1.5:
        x_corr = "OK"
    x_text = "X: %+.1f px  %s" % (-dx, x_corr)
    gfx.DrawString(x_text, font_lg, x_brush, float(width * 0.1), bar_y + 6.0)
    x_brush.Dispose()

    # Y offset with correction direction (negate so +Y = up for user)
    y_color = offset_color(dy)
    y_brush = SolidBrush(y_color)
    y_corr = "UP" if dy > 0 else "DN"
    if abs(dy) < 1.5:
        y_corr = "OK"
    y_text = "Y: %+.1f px  %s" % (dy, y_corr)
    gfx.DrawString(y_text, font_lg, y_brush, float(width * 0.45), bar_y + 6.0)
    y_brush.Dispose()

    # Total offset magnitude
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

    # Direction of offset in degrees (0=right, 90=up)
    if offset_dist > 0.5:
        angle_deg = math.degrees(math.atan2(-dy, dx))  # Negate Y for screen coords
        if angle_deg < 0:
            angle_deg += 360.0
        dir_str = "%.0f deg" % angle_deg
    else:
        dir_str = "--"

    # Quality rating based on offset as % of outer radius
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
        "dX: %+.1f  dY: %+.1f" % (dx, -dy),  # -dy so +Y = up for user
        "Direction: %s" % dir_str,
        "Outer R: %.0f  Inner R: %.0f" % (ore, ir),
        "Quality: %s" % quality,
    ]

    font = Font(FontFamily.GenericMonospace, TEXT_FONT_SIZE, FontStyle.Regular)
    txt_brush = SolidBrush(TEXT_COLOR)
    bg_brush = SolidBrush(TEXT_BG_COLOR)
    q_brush = SolidBrush(qcolor)

    pad = 6
    line_h = TEXT_FONT_SIZE + 4
    panel_w = float(TEXT_FONT_SIZE * 22)
    panel_h = float(len(lines) * line_h + pad * 2)

    gfx.FillRectangle(bg_brush, RectangleF(float(pad), float(pad), panel_w, panel_h))

    y = float(pad + 4)
    for i, line in enumerate(lines):
        if i == 6:
            # Quality line: "Quality: " in white, value in color
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
    font = Font(FontFamily.GenericMonospace, TEXT_FONT_SIZE, FontStyle.Bold)
    brush = SolidBrush(Color.Orange)
    bg = SolidBrush(TEXT_BG_COLOR)
    w = float(len(text) * (TEXT_FONT_SIZE * 0.65) + 16)
    h = float(TEXT_FONT_SIZE + 10)
    gfx.FillRectangle(bg, RectangleF(6.0, 6.0, w, h))
    gfx.DrawString(text, font, brush, 10.0, 8.0)
    font.Dispose()
    brush.Dispose()
    bg.Dispose()


def draw_error_text(gfx, text):
    """Draw an error message at top-left (red text on dark background)."""
    font = Font(FontFamily.GenericMonospace, 10.0, FontStyle.Regular)
    brush = SolidBrush(Color.Red)
    bg = SolidBrush(TEXT_BG_COLOR)
    gfx.FillRectangle(bg, RectangleF(10.0, 10.0, 500.0, 20.0))
    gfx.DrawString(text, font, brush, 14.0, 12.0)
    font.Dispose()
    brush.Dispose()
    bg.Dispose()


# =============================================================================
# FRAME HANDLER
# =============================================================================
# This is the core event handler registered on camera.BeforeFrameDisplay.
# It runs on every frame SharpCap is about to display. The handler:
#   1. Gets the drawable bitmap (for drawing overlays)
#   2. Periodically gets the frame bitmap (for pixel analysis)
#   3. Runs donut analysis on the frame pixels
#   4. Draws the overlay using current (smoothed) detection state
#
# IMPORTANT: All GDI+ resources (Pen, Brush, Font, Graphics, Bitmap) must
# be Disposed to avoid memory leaks. The finally block ensures cleanup.

def on_before_frame_display(*args):
    """BeforeFrameDisplay event handler. Called by SharpCap on every frame."""
    # Skip if disabled or in "Off" display mode
    if not _state["enabled"] or _state["display_mode"] >= len(_DISPLAY_MODE_NAMES) - 1:
        return

    dbitmap = None
    gfx = None

    try:
        frame = args[1].Frame

        # Try the frame annotation API first (avoids bitmap alloc/dispose overhead).
        # Requires a string key to identify the annotation layer.
        # Falls back to the older GetDrawableBitmap approach if unavailable.
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
        if _state["frame_count"] % ANALYSIS_INTERVAL == 0:
            roi_frame = None
            try:
                # When tracking, use CutROI for smaller analysis region
                approx_center = None
                roi_ox = roi_oy = 0
                peak = None
                if _state["outer_cx"] is not None and _state["outer_r"] is not None:
                    roi_frame, roi_ox, roi_oy = _try_cut_roi(frame)
                    if roi_frame is not None:
                        # CutROI active: pixel scan on small region is fast,
                        # so let it compute the real peak (more accurate than histogram).
                        # Only pass approx_center to skip the coarse centroid pass.
                        approx_center = (
                            _state["outer_cx"] - roi_ox,
                            _state["outer_cy"] - roi_oy,
                        )
                    else:
                        # CutROI failed but tracking: use histogram peak to
                        # skip the expensive full-frame coarse pixel scan.
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
                        # Offset ROI coordinates back to full frame
                        if roi_frame is not None:
                            result["inner_cx"] += roi_ox
                            result["inner_cy"] += roi_oy
                            result["outer_cx"] += roi_ox
                            result["outer_cy"] += roi_oy
                        # Save reference values on initial detection
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
        # Try to show the error on-screen so user sees it without the console
        if gfx is not None:
            try:
                draw_error_text(gfx, "Error: %s" % str(ex))
            except:
                pass
        if _state.get("frame_count", 0) < 5:
            print("[Collimation] Handler error: %s" % str(ex))

    finally:
        # Always dispose GDI+ resources to prevent memory leaks.
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
    """
    Extract raw pixel data from a System.Drawing.Bitmap and run donut analysis.

    Uses LockBits to copy pixel data into a .NET byte array, then calls
    analyze_donut_raw(). Tries multiple pixel formats for compatibility
    with different camera types (RGB, ARGB, indexed).
    """
    width = bmp.Width
    height = bmp.Height
    rect = Rectangle(0, 0, width, height)

    # Try to lock as 24bpp RGB first (most compatible and efficient).
    # Fall back to other formats if the bitmap doesn't support conversion.
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

    # Copy pixel data from unmanaged memory to a .NET byte array
    stride = bmpdata.Stride
    total = abs(stride) * height
    raw = Array.CreateInstance(Byte, total)
    Marshal.Copy(bmpdata.Scan0, raw, 0, total)
    bmp.UnlockBits(bmpdata)

    # Compute bytes per pixel from stride (stride may include padding bytes)
    bpp = abs(stride) // width
    if bpp < 1:
        bpp = 1

    return analyze_donut_raw(raw, abs(stride), bpp, width, height, peak_override, approx_center)


# =============================================================================
# TOGGLE AND CONTROL
# =============================================================================

def cycle_display_mode():
    """Cycle to the next display mode. Called by toolbar button or console."""
    _state["display_mode"] = (_state["display_mode"] + 1) % len(_DISPLAY_MODE_NAMES)
    mode = _state["display_mode"]
    name = _DISPLAY_MODE_NAMES[mode]
    print("[Collimation] Display: %s" % name)
    # When cycling back from Off to All, reset detection state so it
    # re-acquires the donut from scratch (position may have changed).
    if mode == 0:
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


def diagnose():
    """
    Diagnostic tool: captures one frame and tests all pixel access methods.
    Prints detailed results to the console. Use this to debug detection
    issues — it shows what pixel data is actually available and what values
    the camera is producing.
    """
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

            # Method 1: Direct pixel value access (SharpCap native)
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

            # Method 2: Raw byte buffer
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

            # Method 3: Frame as System.Drawing.Bitmap (THIS IS WHAT WE USE)
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

            # Method 4: CreateBitmap (requires args in some versions)
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

            # Method 5: Overlay layer (expected to be blank - included for reference)
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

            # Method 6: Raw memory buffer
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
    """
    Probe the new SharpCap APIs and print what's available.
    Run this in the console and paste the output for debugging.
    Usage: probe_apis()
    """
    print("[Probe] Capturing one frame to test APIs...")

    _probe = {"done": False, "info": []}

    def probe_handler(*args):
        if _probe["done"]:
            return
        _probe["done"] = True
        info = _probe["info"]
        try:
            frame = args[1].Frame

            # --- frame.Index ---
            info.append("=== frame.Index ===")
            try:
                idx = frame.Index
                info.append("  Value: %s (type: %s)" % (idx, type(idx)))
            except Exception as ex:
                info.append("  FAILED: %s" % ex)

            # --- frame.GetAnnotationGraphics() ---
            info.append("")
            info.append("=== frame.GetAnnotationGraphics() ===")
            # Try with various argument patterns
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

            # --- frame.CalculateHistogram() ---
            info.append("")
            info.append("=== frame.CalculateHistogram() ===")
            try:
                hist = frame.CalculateHistogram()
                info.append("  Result: %s (type: %s)" % (hist, type(hist)))
                info.append("  Members: %s" % [x for x in dir(hist) if not x.startswith("_")])

                # Probe GetCentilePointsF (returns float points for centile values)
                info.append("")
                info.append("  --- GetCentilePointsF ---")
                try:
                    # Try with a single float (99.9th percentile)
                    val = hist.GetCentilePointsF(99.9)
                    info.append("  GetCentilePointsF(99.9) = %s (type: %s)" % (val, type(val)))
                except Exception as ex2:
                    info.append("  GetCentilePointsF(99.9) FAILED: %s" % ex2)
                try:
                    # Try with a list/array of centile values
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

                # Probe GetMeansBelowCentile
                info.append("")
                info.append("  --- GetMeansBelowCentile ---")
                try:
                    val = hist.GetMeansBelowCentile(99.9)
                    info.append("  GetMeansBelowCentile(99.9) = %s (type: %s)" % (val, type(val)))
                except Exception as ex2:
                    info.append("  GetMeansBelowCentile(99.9) FAILED: %s" % ex2)

                # Probe Values (histogram bin counts)
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

            # --- frame.CutROI() ---
            info.append("")
            info.append("=== frame.CutROI() ===")
            fi = frame.Info
            rx = fi.Width // 2 - 50
            ry = fi.Height // 2 - 50
            # Try Rectangle argument
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
    print("  display_mode: %d (%s)" % (_state["display_mode"], _DISPLAY_MODE_NAMES[_state["display_mode"]]))
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
    """Enable per-frame debug logging to console. Use debug_off() to stop."""
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
# These functions let you test the detection algorithm without a real camera.
# test_synthetic() generates a fake donut image; test_image() loads a file.
# Both update _state so the overlay draws on the next live frame.

def test_synthetic(width=800, height=600, outer_r=220, inner_r=90,
                   offset_x=15, offset_y=-10):
    """
    Generate a synthetic donut image and run analysis on it.
    Sets the state so the overlay draws on the next frame.

    Args:
        width, height: image dimensions (should match your camera)
        outer_r: outer donut radius in pixels
        inner_r: inner donut radius in pixels
        offset_x, offset_y: offset of inner circle from outer (simulates miscollimation)
    """
    print("[Test] Generating synthetic donut: %dx%d" % (width, height))
    print("[Test] Outer R=%d, Inner R=%d, Offset=(%d, %d)" % (outer_r, inner_r, offset_x, offset_y))

    cx = width / 2.0
    cy = height / 2.0
    icx = cx + offset_x
    icy = cy + offset_y

    # Build float array with donut pattern (soft edges + diffraction rings)
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
                val = 5.0  # Dark center (secondary shadow)
            elif dist_outer < outer_r:
                inner_fade = min(1.0, (dist_inner - inner_r) / 15.0) if dist_inner > inner_r else 0.0
                outer_fade = min(1.0, (outer_r - dist_outer) / 20.0) if dist_outer < outer_r else 0.0
                val = 200.0 * inner_fade * outer_fade
                val += 30.0 * max(0, math.cos(dist_outer * 0.3)) * inner_fade * outer_fade
            else:
                val = 2.0  # Dark background

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
    """
    Load an image file and run donut analysis on it.
    Sets the state so the overlay draws on the next frame.

    Usage: test_image(r"C:\\path\\to\\donut.png")
    """
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
# SETUP AND TEARDOWN
# =============================================================================

def _make_icon(icon_type):
    """Create a 24x24 toolbar icon bitmap using GDI+ drawing.
    Uses System.Drawing.Graphics directly (not the SharpCap wrapper)."""
    size = 24
    bmp = Bitmap(size, size)
    g = System.Drawing.Graphics.FromImage(bmp)
    g.Clear(Color.Transparent)

    if icon_type == "toggle":
        # Concentric circles icon: red outer, green inner, yellow center dot
        p = Pen(Color.FromArgb(255, 220, 60, 60), 2.0)
        g.DrawEllipse(p, 2, 2, 19, 19)
        p.Dispose()
        p = Pen(Color.FromArgb(255, 60, 220, 60), 2.0)
        g.DrawEllipse(p, 7, 7, 9, 9)
        p.Dispose()
        b = SolidBrush(Color.Yellow)
        g.FillRectangle(b, RectangleF(10.0, 10.0, 3.0, 3.0))
        b.Dispose()

    elif icon_type == "reset":
        # Circular arrow icon (refresh symbol)
        p = Pen(Color.FromArgb(255, 100, 180, 255), 2.0)
        g.DrawArc(p, 4, 4, 15, 15, 0, 270)
        p.Dispose()
        p = Pen(Color.FromArgb(255, 100, 180, 255), 2.0)
        g.DrawLine(p, 11.0, 4.0, 19.0, 4.0)
        g.DrawLine(p, 19.0, 4.0, 19.0, 10.0)
        p.Dispose()

    g.Dispose()
    return bmp


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
    print("[Collimation] Stopped. Run the script again to restart.")


def setup():
    """
    Initialize the collimation overlay:
      1. Add the "Coll Overlay" toolbar button
      2. Attach the BeforeFrameDisplay handler to the camera
      3. Print usage instructions to the console
    """
    print("=" * 50)
    print("  COLLIMATION OVERLAY v4")
    print("=" * 50)

    # Add a single toolbar button that cycles through display modes.
    # SharpCap's AddCustomButton takes 4 args but the ordering varies by
    # version, so we try 3 possible orderings and use whichever works.
    btn_name = "Coll Overlay"
    icon = _make_icon("toggle")
    btn_added = False

    for arg_order in [
        (btn_name, icon, "Cycle collimation overlay mode", cycle_display_mode),
        (btn_name, cycle_display_mode, icon, "Cycle collimation overlay mode"),
        (btn_name, cycle_display_mode, "Cycle collimation overlay mode", icon),
    ]:
        if not btn_added:
            try:
                SharpCap.AddCustomButton(*arg_order)
                btn_added = True
            except:
                pass

    if btn_added:
        _state["buttons"] = [btn_name]
        print("[OK] Toolbar button: '%s'" % btn_name)
    else:
        print("[WARN] Could not add toolbar button")
        print("  Use cycle_display_mode() in console")

    attach_handler()

    print("")
    print("'%s' button cycles:" % btn_name)
    for i, name in enumerate(_DISPLAY_MODE_NAMES):
        marker = ">>>" if i == 0 else "   "
        print("  %s %d. %s" % (marker, i + 1, name))
    print("")
    print("Console commands:")
    print("  cycle_display_mode()   - same as button")
    print("  reset_tracking()       - re-detect donut")
    print("  show_state()           - debug info")
    print("  debug_on() / debug_off() - per-frame logging")
    print("  probe_apis()           - test new SharpCap APIs")
    print("  stop()                 - fully stop + remove button")
    print("=" * 50)


# =============================================================================
# AUTO-START
# =============================================================================
# When this script is executed (via File > Run Script or exec()), setup()
# runs automatically, registering the handler and creating the toolbar button.
# The script then exits, but the handler remains active on every frame.

setup()
