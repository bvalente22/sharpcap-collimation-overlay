# =============================================================================
# DIAGNOSTIC SCRIPT - Run this in SharpCap's Scripting Console
# =============================================================================
# This script tests each API to figure out what works in your SharpCap version.
#
# HOW TO RUN:
#   1. In SharpCap, go to Tools > Scripting Console
#   2. In the console, paste the contents of this script (or use File > Run Script)
#   3. Read the output in the console window
# =============================================================================

import clr
clr.AddReference("System.Drawing")
from System.Drawing import Color, Pen, Rectangle
from System.Drawing.Imaging import ImageLockMode, PixelFormat
from System.Runtime.InteropServices import Marshal
from System import Array, Byte

print("=" * 60)
print("  SHARPCAP COLLIMATION OVERLAY - DIAGNOSTICS")
print("=" * 60)

# --- Test 1: SharpCap object ---
print("")
print("[TEST 1] SharpCap object")
try:
    print("  SharpCap type: %s" % type(SharpCap))
    print("  SharpCap dir: %s" % [x for x in dir(SharpCap) if not x.startswith("_")])
except Exception as ex:
    print("  FAIL: %s" % ex)

# --- Test 2: Camera ---
print("")
print("[TEST 2] Camera")
try:
    cam = SharpCap.SelectedCamera
    if cam is None:
        print("  FAIL: No camera selected!")
        print("  Please connect a camera and start preview first.")
    else:
        print("  Camera: %s" % cam.DeviceName)
        print("  Camera type: %s" % type(cam))
        cam_members = [x for x in dir(cam) if not x.startswith("_")]
        print("  Camera members: %s" % cam_members)
except Exception as ex:
    print("  FAIL: %s" % ex)

# --- Test 3: Check for BeforeFrameDisplay event ---
print("")
print("[TEST 3] BeforeFrameDisplay event")
try:
    cam = SharpCap.SelectedCamera
    if cam is not None:
        has_bfd = hasattr(cam, "BeforeFrameDisplay")
        print("  HasAttribute 'BeforeFrameDisplay': %s" % has_bfd)

        # Also check for other possible event names
        for name in ["BeforeFrameDisplay", "FrameCaptured", "FrameReady",
                      "OnFrameDisplay", "PreFrameDisplay", "FrameArrived",
                      "NewFrame", "ImageReady"]:
            if hasattr(cam, name):
                print("  Found event: %s" % name)
except Exception as ex:
    print("  FAIL: %s" % ex)

# --- Test 4: Try attaching a simple handler ---
print("")
print("[TEST 4] Attach simple BeforeFrameDisplay handler")
_diag_count = [0]
_diag_errors = []
_diag_info = []

def diag_handler(*args):
    _diag_count[0] += 1
    if _diag_count[0] <= 3:  # Only log first 3 calls
        try:
            _diag_info.append("Call #%d: num args=%d" % (_diag_count[0], len(args)))
            for i, a in enumerate(args):
                _diag_info.append("  arg[%d] type=%s" % (i, type(a)))
                if i == 1:
                    try:
                        members = [x for x in dir(a) if not x.startswith("_")]
                        _diag_info.append("  arg[1] members: %s" % members)
                    except:
                        pass
                    try:
                        _diag_info.append("  arg[1].Frame type: %s" % type(a.Frame))
                        frame = a.Frame
                        frame_members = [x for x in dir(frame) if not x.startswith("_")]
                        _diag_info.append("  Frame members: %s" % frame_members)
                    except Exception as ex:
                        _diag_info.append("  arg[1].Frame access failed: %s" % ex)
        except Exception as ex:
            _diag_errors.append("Handler error: %s" % ex)

try:
    cam = SharpCap.SelectedCamera
    if cam is not None:
        cam.BeforeFrameDisplay += diag_handler
        print("  Handler attached OK. Waiting 3 seconds for frames...")

        import time
        time.sleep(3)

        cam.BeforeFrameDisplay -= diag_handler
        print("  Handler detached.")
        print("  Frames received: %d" % _diag_count[0])
        if _diag_count[0] == 0:
            print("  WARNING: No frames received! Is the camera capturing?")
            print("  Make sure live preview is running.")
        for line in _diag_info:
            print("  %s" % line)
        for err in _diag_errors:
            print("  ERROR: %s" % err)
except Exception as ex:
    print("  FAIL: %s" % ex)

# --- Test 5: Try drawing on a frame ---
print("")
print("[TEST 5] Drawing test")
_draw_ok = [False]
_draw_info = []

def draw_test_handler(*args):
    if _draw_ok[0]:
        return
    try:
        frame = args[1].Frame
        _draw_info.append("Got frame")

        # Try GetDrawableBitmap
        try:
            dbitmap = frame.GetDrawableBitmap()
            _draw_info.append("GetDrawableBitmap: %s (type: %s)" % (dbitmap, type(dbitmap)))
            _draw_info.append("DrawableBitmap members: %s" % [x for x in dir(dbitmap) if not x.startswith("_")])

            gfx = dbitmap.GetGraphics()
            _draw_info.append("GetGraphics: %s (type: %s)" % (gfx, type(gfx)))

            # Try to draw a test circle
            pen = Pen(Color.Red, 3.0)
            gfx.DrawEllipse(pen, 50.0, 50.0, 100.0, 100.0)
            gfx.DrawString("OVERLAY TEST",
                          System.Drawing.Font(System.Drawing.FontFamily.GenericMonospace, 16.0),
                          System.Drawing.SolidBrush(Color.Yellow),
                          60.0, 90.0)
            pen.Dispose()
            gfx.Dispose()
            dbitmap.Dispose()
            _draw_info.append("Drawing completed - check top-left of image for red circle!")
            _draw_ok[0] = True
        except Exception as ex:
            _draw_info.append("GetDrawableBitmap FAILED: %s" % ex)

        # Try frame.Info
        try:
            info = frame.Info
            _draw_info.append("Frame.Info: Width=%s, Height=%s" % (info.Width, info.Height))
            _draw_info.append("Frame.Info members: %s" % [x for x in dir(info) if not x.startswith("_")])
        except Exception as ex:
            _draw_info.append("Frame.Info FAILED: %s" % ex)

        # Try CutROI
        try:
            info = frame.Info
            roi = frame.CutROI(Rectangle(0, 0, info.Width, info.Height))
            _draw_info.append("CutROI: %s (type: %s)" % (roi, type(roi)))
            _draw_info.append("CutROI members: %s" % [x for x in dir(roi) if not x.startswith("_")])
            try:
                raw = roi.ToBytes()
                _draw_info.append("ToBytes: length=%d" % len(raw))
            except Exception as ex2:
                _draw_info.append("ToBytes FAILED: %s" % ex2)
            try:
                floats = roi.ToFloats(0)
                _draw_info.append("ToFloats(0): length=%d" % len(floats))
            except Exception as ex2:
                _draw_info.append("ToFloats FAILED: %s" % ex2)
        except Exception as ex:
            _draw_info.append("CutROI FAILED: %s" % ex)

    except Exception as ex:
        _draw_info.append("Frame access failed: %s" % ex)

# Need System.Drawing imports for the drawing test
import System.Drawing

try:
    cam = SharpCap.SelectedCamera
    if cam is not None:
        cam.BeforeFrameDisplay += draw_test_handler
        print("  Draw test handler attached. Waiting 3 seconds...")

        import time
        time.sleep(3)

        cam.BeforeFrameDisplay -= draw_test_handler
        print("  Handler detached.")
        print("  Draw succeeded: %s" % _draw_ok[0])
        for line in _draw_info:
            print("  %s" % line)
except Exception as ex:
    print("  FAIL: %s" % ex)

# --- Test 6: Check AddCustomButton ---
print("")
print("[TEST 6] Custom button")
try:
    def test_btn():
        print("[Collimation] Button clicked!")
    SharpCap.AddCustomButton("Diag Test", "Test button", test_btn)
    print("  Button added OK. Check toolbar for 'Diag Test' button and click it.")
except Exception as ex:
    print("  FAIL: %s" % ex)

print("")
print("=" * 60)
print("  DIAGNOSTICS COMPLETE")
print("  Copy ALL output above and share it.")
print("=" * 60)
