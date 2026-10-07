#!/usr/bin/env python3
"""Build qc_automation_tracker.xlsx from the post-wall QC export.

Usage:  python3 -I build_workbook.py [data/qc_post_wall_source.csv] [qc_automation_tracker.xlsx]

The tracker keeps the useful columns of the export, adds a proposed automation
design per check, and leaves tracking columns for the team to fill. Every
figure on the Dashboard is a formula over the QC Tracker tab.
"""
import csv
import datetime as dt
import math
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

SRC = Path(sys.argv[1] if len(sys.argv) > 1 else "data/qc_post_wall_source.csv")
OUT = Path(sys.argv[2] if len(sys.argv) > 2 else "qc_automation_tracker.xlsx")
SOURCE_NOTE = "Source: Tool QC Automation - Sheet2.csv (post-wall QC export), uploaded 2026-10-07"

# ----------------------------------------------------------------------------
# Method catalogue (codes used in the tracker's "Primary method" column)
# ----------------------------------------------------------------------------
METHODS = [
    # code, name, what the SE does differently, extra seconds per use, accuracy, build effort, limits
    ("G-VIS", "Gyro frames to vision model",
     "Nothing new. The existing gyro capture is used.",
     0, "Qualitative. Reliable for presence, finish, gross defects and matching an item to a written spec or reference image.",
     "M", "Does not measure. Give it the drawing or spec text in the prompt and ask for pass, flag or needs-a-photo."),
    ("G-GEO", "Gyro to metric geometry",
     "Nothing new, provided the tripod height is fixed or the floor-to-ceiling height from the drawing is available.",
     0, "About ±20–30 mm at 3 m for heights, sizes and positions on a wall; about ±5 mm when a tile or block in the scene is used as the ruler.",
     "L", "Needs gravity-aligned frames, a room-layout model and a known camera height. Lengths across the view (widths) are less accurate than heights."),
    ("G-PLUMB", "Gyro plumb, level and straightness screening",
     "Nothing new.",
     0, "About ±0.2°, which is ±7 mm over 2 m. A screening tool for gross errors, not a plumb-bob replacement.",
     "M", "Only in-plane deviation is visible. A wall bowing toward or away from the camera does not show as a slanted line."),
    ("G-FLOOR", "Gyro to top-down floor image",
     "Optionally one extra band at about −70° so the floor near the tripod is covered.",
     0, "About 1 mm for in-plane alignment when the tile size is used as the ruler.",
     "M", "Lippage (height steps between tiles) does not show from above."),
    ("TAP", "Requested close-up photo",
     "Takes one photo when the app asks, with a tile, block, coin or scale sticker in frame.",
     30, "About 0.3–0.5 mm with a reference object in frame.",
     "S", "Keep requests rare: only checks that always need one, or when the vision model reports it is unsure."),
    ("IMU", "Phone as a level",
     "Lays the phone flat on the surface for two seconds, at two or three points.",
     20, "About ±0.1° (1 mm per metre) after a one-time calibration.",
     "S", "Needs a flat phone back or a flat case. Good for slopes of 1 % and above."),
    ("BT-TOOL", "Bluetooth laser meter or digital level",
     "Takes the reading with the tool; the app receives it over Bluetooth.",
     30, "Millimetre grade.",
     "M", "Hardware cost per SE. Best used only on items the gyro screening has flagged."),
    ("EXT", "Exterior capture",
     "One walking video along the compound wall, or gyro captures at two or three points in the plot. Once per site.",
     120, "Same as G-GEO and G-VIS, applied outdoors.",
     "M", "New action in the app. Strong sun and shadows need a prompt that tolerates them."),
    ("ATTEST", "One-tap confirmation with evidence",
     "Taps yes and attaches a photo, a short video or a voice note; or the client approves in the app.",
     10, "A process record, not a measurement.",
     "S", "Use for process checks and approvals where no image after the fact can prove the point."),
    ("MANUAL", "Stays with the SE and a physical tool",
     "Takes the reading as today and types or dictates it into the app.",
     60, "Tool grade.",
     "S", "The app still gets a structured record and a photo, so the check is not lost."),
]

STATUSES = ["Not started", "Exploring", "Prototype", "Validating", "Live", "Parked"]
LEVELS = ["Full", "Partial", "Evidence", "Manual"]
CONFIDENCES = ["High", "Medium", "Low"]

# ----------------------------------------------------------------------------
# Per-check design, keyed by the export's sno
# ----------------------------------------------------------------------------
def d(family, measure, method, level, conf, how, fallback, inputs, tol, same=""):
    return dict(family=family, measure=measure, method=method, level=level, conf=conf,
                how=how, fallback=fallback, inputs=inputs, tol=tol, same=same)

DESIGN = {
    30: d("Squareness and plumb", "Jamb lean against gravity; jamb-to-soffit angle, per opening.",
          "G-PLUMB", "Partial", "Medium",
          "Gravity-align the frames with the gyro, rectify each wall to a flat elevation view and detect the opening edges. Jambs should be vertical and the soffit horizontal; report the lean in mm over the opening height and the corner angle. Flags go to the SE.",
          "BT-TOOL digital level on flagged openings only. The vertical vanishing point of the room's own edges can refine the gyro's gravity estimate.",
          "Plumb tolerance.", "Plumb within 5 mm over the opening height; 90° ± 0.5°"),
    31: d("Alignment and straightness", "Straightness of courses (horizontal) and perpends or wall ends (vertical) in the wall plane.",
          "G-PLUMB", "Partial", "Medium",
          "On the rectified wall view, detect the mortar lines, fit straight lines and report the largest deviation; the block height gives the mm scale. This checks in-plane alignment. A bow toward or away from the camera is not visible from one viewpoint.",
          "Straightedge by the SE on flagged walls; a BT laser meter read at three points along the wall gives the bow.",
          "Block size (mm); straightness tolerance.", "About 6 mm over 3 m"),
    33: d("Plumb and level", "Jamb lean against gravity.",
          "G-PLUMB", "Partial", "Medium",
          "Same detector as ID 30: this is the vertical component of the opening check.",
          "BT-TOOL on flagged openings.", "Plumb tolerance.", "Within 5 mm over the opening height", same=30),
    34: d("Height and position", "Height of the band's top and bottom edges above floor level, per wall.",
          "G-GEO", "Full", "High",
          "The band is a distinct concrete strip in the blockwork. Segment it on each rectified wall and convert its edges to heights using the camera height and the wall distance from the room layout. Compare with 1200 mm.",
          "TAP photo with a tape held against the band.",
          "Band height standard (1200 mm); band depth from the structural drawing.", "1200 mm ± 25 mm"),
    36: d("Thickness and joints", "Average and local mortar joint thickness.",
          "G-GEO", "Partial", "Medium",
          "Average joint = course pitch minus block height, measured over several courses on the rectified wall (the block is the ruler). Local joints wider than about 15 mm show at 1 mm per pixel; finer differences do not.",
          "TAP close-up with a block edge in frame for exact joints. Thin-bed AAC joints (3–4 mm) always need the close-up.",
          "Block height (mm); mortar type (conventional or thin-bed).", "10 mm ± 3 mm conventional; 3–4 mm thin-bed"),
    37: d("Plumb and level", "Soffit height of every opening; soffit line horizontal.",
          "G-GEO", "Partial", "Medium",
          "Soffit heights from the rectified walls (about ±20 mm) compared across all openings in the room, and each soffit line checked against gravity. Catches an opening set 30 mm high; cannot match a tube level's millimetres.",
          "BT laser level or digital level on flagged openings.", "Lintel height from the drawing.", "Within 10 mm between openings"),
    38: d("Plumb and level", "Column edge lean against gravity above lintel level.",
          "G-PLUMB", "Partial", "Medium",
          "Columns read as concrete against blockwork before plaster. Detect both edges on the rectified wall view and measure the lean over the visible height.",
          "BT-TOOL or plumb bob on flagged columns.", "Column plumb tolerance from the structural spec.", "Per structural spec"),
    40: d("Height and position", "Lintel soffit height above floor, per opening.",
          "G-GEO", "Full", "High",
          "Lintel soffit height from the rectified wall view against the height given in the drawing for that opening.",
          "TAP with a tape.", "Lintel heights from the drawing (for example 2100 mm).", "± 25 mm"),
    63: d("Size against drawing", "Lintel depth and bearing length each side; width equals the wall thickness.",
          "G-GEO", "Partial", "Medium",
          "Depth and bearing are visible as the concrete band around the opening before plaster; measure them on the rectified wall (about ±10–20 mm). Width cannot be seen from inside the room.",
          "TAP with a tape across the lintel.", "Structural drawing (depth, bearing).", "± 10 mm"),
    69: d("Process and approval", "Vibrator used during the pour; no honeycombing after deshuttering.",
          "ATTEST", "Evidence", "Medium",
          "During the pour, a mandatory ten-second video or photo of the vibrator in use. After deshuttering, the staircase gyro frames go through the vision model for honeycombing, voids and segregation on the waist slab, risers and soffit.",
          "TAP close-ups of suspect patches.", "None.", "No honeycombing visible"),
    72: d("Alignment and straightness", "Straightness of the base course in both directions.",
          "G-PLUMB", "Partial", "Medium",
          "Same detector as ID 31, run at the base-course stage.",
          "Straightedge on flagged walls.", "Block size (mm).", "About 6 mm over 3 m", same=31),
    76: d("Thickness and joints", "Average and local mortar joint thickness at the base course.",
          "G-GEO", "Partial", "Medium",
          "Same detector as ID 36, run at the base-course stage.",
          "TAP close-up with a block edge in frame.", "Block height (mm); mortar type.", "10 mm ± 3 mm conventional", same=36),
    77: d("Squareness and plumb", "Jamb lean and jamb-to-soffit angle, per opening.",
          "G-PLUMB", "Partial", "Medium",
          "Same detector as ID 30, run at the lintel-level stage.",
          "BT-TOOL on flagged openings.", "Plumb tolerance.", "Within 5 mm; 90° ± 0.5°", same=30),
    79: d("Plumb and level", "Jamb lean against gravity.",
          "G-PLUMB", "Partial", "Medium",
          "Same detector as ID 30 (vertical component), run at the lintel-level stage.",
          "BT-TOOL on flagged openings.", "Plumb tolerance.", "Within 5 mm over the opening height", same=30),
    83: d("Height and position", "Sill top height above floor, per window.",
          "G-GEO", "Full", "High",
          "The sill band is visible before plaster. Height of its top edge on the rectified wall against the drawing.",
          "TAP with a tape.", "Sill heights from the drawing (for example 900 mm).", "± 25 mm"),
    84: d("Plumb and level", "Wall lean.",
          "G-PLUMB", "Partial", "Medium",
          "Corner lines checked against gravity. A lean toward or away from the camera does not show as a slanted line, so also compare each wall's floor-line and ceiling-line distances from the layout; the difference is the lean over the full height (resolves about 30 mm).",
          "BT-TOOL digital level or plumb bob on flagged walls.", "Plumb tolerance.", "About 6 mm per 3 m"),
    107: d("Presence and spec match", "Buttress count and spacing along the compound wall.",
           "EXT", "Partial", "Medium",
           "A walking video along the compound wall, or gyro captures at two or three points in the plot. Detect the buttresses, count them and get the spacing from the wall length in the plot drawing.",
           "TAP photo per stretch with a count.", "Plot drawing (wall lengths); buttress spacing standard.", "Spacing at or under the spec"),
    108: d("Size against drawing", "Wall height; wall thickness.",
           "EXT", "Partial", "Medium",
           "Height from the exterior capture with the camera height known (about ±30 mm). Thickness needs a view of the top or an end: one TAP photo with a block as the scale.",
           "BT laser meter.", "Spec height and width.", "± 25 mm"),
    109: d("Surface finish", "Gross undulations, patches and cracks in the plaster.",
           "EXT", "Partial", "Low",
           "The vision model reads the exterior frames for gross undulations, patches and cracks, best in low sun or with a torch at a grazing angle. Millimetre flatness is not visible.",
           "MANUAL 2 m straightedge, reading recorded in the app.", "None.", "About 3 mm under a 2 m straightedge"),
    110: d("Height and position", "Top level of the SSM against the datum; width.",
           "MANUAL", "Manual", "Low",
           "Levels need an instrument. The SE records the auto-level or laser readings in the app with a photo. Width from a TAP photo with a tape.",
           "A BT laser level with a receiver could feed the app directly.", "Drawing (level, width).", "Per drawing"),
    111: d("Size against drawing", "Footing positions against the plan.",
           "MANUAL", "Evidence", "Low",
           "Photo of the laid-out lines with the tape and the drawing, kept as the record.",
           "A photo from an elevated point (first floor or terrace) with four marked reference points can be warped onto the plan for a direct overlay.",
           "Foundation plan.", "Per drawing"),
    123: d("Surface finish", "Even gaps, clean sealant line, no visible fasteners or damage.",
           "G-VIS", "Full", "Medium",
           "The vision model reads a crop of each door, window and grill from the gyro frames with a short rubric: gap evenness, sealant line, fastener heads, scratches, damage.",
           "TAP close-up where the model is unsure.", "Finish rubric (what uniform means for your packages).", "Rubric pass"),
    125: d("Size against drawing", "Frame width and height; design against spec.",
           "G-GEO", "Full", "Medium",
           "Width and height of each frame from the rectified wall (about ±30 mm separates catalogue sizes such as 900 and 1000 mm). Design matched by the vision model against the BOQ line or a reference image for the package (type, material, track count, mesh).",
           "TAP with a tape for borderline sizes.", "Door and window schedule; BOQ or package spec with reference images.", "Size ± 25 mm; design must match"),
    127: d("Functional", "Shutter alignment and gaps; latching.",
           "G-VIS", "Partial", "Medium",
           "Ask the SE to close all shutters before the gyro capture; the vision model then checks even gaps and alignment. Latching is a one-tap confirmation per door and window.",
           "IMU: phone laid on the shutter's top edge for level. A short video of the latch.", "None.", "Even gap; latch engages"),
    128: d("Height and position", "Opening position along the wall; sill height.",
           "G-GEO", "Full", "High",
           "Each opening is located on its wall from the layout (distance from the corner) and its sill height is read, both compared with the drawing.",
           "TAP with a tape.", "Floor plan and window schedule.", "Position ± 50 mm; sill ± 25 mm"),
    171: d("Presence and spec match", "Chamber cover size, type and material.",
           "TAP", "Partial", "Medium",
           "One photo per chamber with a tape or a known tile in frame. The vision model checks type and material against the spec (RCC, CI, SFRC, frame); size from the reference object.",
           "Include the chambers in the exterior capture (EXT).", "BOQ spec for chamber covers.", "As per spec"),
    172: d("Size against drawing", "Railing height; baluster spacing; post plumb.",
           "G-GEO", "Partial", "Medium",
           "Needs the staircase and balcony as capture points. Height from the floor line (about ±20 mm), spacing from the rectified view scaled by the railing height, posts checked against gravity.",
           "TAP with a tape; BT-TOOL on flagged posts.", "Railing spec (height, maximum gap).", "Height ± 20 mm; gap at or under spec"),
    176: d("Workmanship defect", "Weld surface quality; throat thickness.",
           "TAP", "Partial", "Low",
           "Close-up photos of the joints; the vision model screens for porosity, spatter, undercut, incomplete fusion and grinding finish. Throat thickness needs a weld gauge.",
           "MANUAL gauge reading typed into the app.", "Weld spec.", "Per spec"),
    186: d("Size against drawing", "Count, size and type of doors and windows per room against the contract.",
           "G-GEO", "Full", "Medium",
           "Same detector as ID 125, with the room's door and window list from the contract as the prompt input; also counts the items per room.",
           "TAP.", "Contract door and window schedule.", "Must match the contract", same=125),
    199: d("Process and approval", "Client approval recorded.",
           "ATTEST", "Evidence", "High",
           "Make this a client approval step in the app: the gyro frames of the bull-marked floor (or a layout render) go to the client, who approves with an OTP or a signature. The approval is the QC record.",
           "Site-visit sign-off photo.", "Tile layout drawing.", "Approval present"),
    206: d("Slope", "Floor slope toward the drain.",
           "IMU", "Partial", "Medium",
           "The SE lays the phone flat on the bed at two or three points (or on a straightedge). The accelerometer reads tilt to about 0.1°, enough for a 1 % slope (0.57°). The app records the direction toward the drain.",
           "ATTEST: a short video of water or a marble running to the drain.", "Required slope (for example 1:100).", "Slope at or above spec, toward the drain"),
    207: d("Alignment and straightness", "Grout-line straightness and offset at intersections.",
           "G-FLOOR", "Full", "High",
           "Warp the downward gyro band into a top-down floor image, detect the grout lines, fit the grid and report the largest offset at intersections and the line waviness (tile size is the ruler, about 1 mm resolution). Lippage does not show.",
           "TAP photo with a torch at a grazing angle for lippage; MANUAL straightedge.", "Tile size; spacer width.", "Offset at or under 1 mm"),
    216: d("Height and position", "Skirting height and uniformity; material.",
           "G-GEO", "Full", "High",
           "Skirting top edge detected along each wall; height from the floor line with the adjacent tile as the ruler (about ±5 mm). Material and finish judged by the vision model.",
           "TAP with a tape.", "Skirting height and material from spec (for example 100 mm).", "± 5 mm"),
    219: d("Workmanship defect", "Cut shape and fit around gratings.",
           "G-VIS", "Full", "Medium",
           "The vision model reads the downward-band crop around each drain: shape matches the grating, cut edges clean, gap even.",
           "TAP when the drain is small or far from the tripod.", "Grating type per bathroom and balcony.", "Rubric pass"),
    222: d("Process and approval", "Client approval recorded.",
           "ATTEST", "Evidence", "High",
           "Same approval flow as ID 199, for wet areas.",
           "Site-visit sign-off photo.", "Tile layout drawing.", "Approval present", same=199),
    272: d("Thickness and joints", "Counter slab thickness at the exposed edge.",
           "TAP", "Partial", "Medium",
           "One close-up of the counter edge with a coin or a printed scale sticker in frame; thickness in pixels converted by the reference.",
           "MANUAL vernier or tape reading.", "Spec thickness (for example 18 or 20 mm).", "± 1 mm"),
    274: d("Presence and spec match", "Cut-out, brackets or clamps, and water and waste points at the sink location.",
           "G-VIS", "Full", "High",
           "The vision model reads the kitchen counter crop from the gyro frames: cut-out present and sized, support brackets, tap and waste points in place.",
           "TAP.", "Sink type from spec.", "All items present"),
    286: d("Presence and spec match", "Railing material type and profile against the package.",
           "G-VIS", "Partial", "Medium",
           "The vision model compares the railing with the package spec and reference image (stainless, mild steel or glass; profile sizes). The steel grade (304 against 202) is not visible.",
           "MANUAL: invoice or material test certificate check.", "Package spec with reference images.", "Must match the package"),
    294: d("Alignment and straightness", "Corner lines straight and vertical; finish.",
           "G-PLUMB", "Partial", "Medium",
           "Internal and external corner lines detected in the gravity-aligned view; straightness and lean reported. Finish (waviness, patching) judged by the vision model on the crop.",
           "MANUAL straightedge on flagged corners.", "Tolerance.", "About 3 mm over 2 m"),
    296: d("Plumb and level", "Frame edge lean against gravity.",
           "G-PLUMB", "Partial", "Medium",
           "Same detector as ID 30, applied to the installed frame edges.",
           "BT-TOOL.", "Tolerance.", "About 3 mm over the frame height", same=30),
    317: d("Height and position", "Box count and positions against the electrical layout; height; finish flush with plaster.",
           "G-GEO", "Full", "High",
           "Detect the boxes on each rectified wall; compare count and positions with the electrical layout and heights with the standards table. The vision model checks flush fit and clean edges.",
           "TAP for finish.", "Electrical layout drawing; standard heights table.", "Height ± 25 mm; position ± 50 mm"),
    318: d("Squareness and plumb", "Angle between adjacent walls.",
           "G-GEO", "Partial", "Medium",
           "The floor plan recovered from the floor or ceiling boundary gives each corner angle to about ±1°. That catches a 3° error, not the 3 mm over 600 mm a square checks.",
           "MANUAL square, or a BT laser meter with the 3-4-5 method.", "Tolerance.", "90° ± 0.3° (about 3 mm over 600 mm)"),
    319: d("Squareness and plumb", "Jamb lean and jamb-to-soffit angle, per opening.",
           "G-PLUMB", "Partial", "Medium",
           "Same detector as ID 30, run after plastering.",
           "BT-TOOL on flagged openings.", "Plumb tolerance.", "Within 5 mm; 90° ± 0.5°", same=30),
    325: d("Surface finish", "Undulations, patches and cracks; flatness.",
           "G-VIS", "Partial", "Low",
           "The vision model reads the rectified wall views for gross undulations, patches and cracks. Millimetre flatness is not visible from the room centre.",
           "MANUAL 2 m straightedge. A phone lidar scan (iPhone Pro) resolves about 5 mm. TAP with the torch at a grazing angle shows waviness.",
           "None.", "About 3 mm under a 2 m straightedge"),
    327: d("Workmanship defect", "Edge chipping, rounding and unevenness at junctions.",
           "G-VIS", "Partial", "Medium",
           "The vision model reads zoomed crops of each junction and asks for a TAP photo when the crop is too small to judge.",
           "TAP.", "Rubric.", "Rubric pass"),
    328: d("Plumb and level", "Grill bar angles against gravity.",
           "G-PLUMB", "Partial", "Medium",
           "Many parallel bars give a robust angle estimate on the gravity-aligned view; horizontal and vertical bars are both checked.",
           "BT-TOOL.", "Tolerance.", "About 3 mm over the grill size"),
    329: d("Squareness and plumb", "Jamb angle; edge sharpness.",
           "G-PLUMB", "Partial", "Medium",
           "Same detector as ID 30 for the angle, plus the edge rubric of ID 327.",
           "TAP.", "Tolerance; rubric.", "90° ± 0.5°; edges intact", same=30),
    432: d("Height and position", "Height of each box above FFL.",
           "G-GEO", "Full", "High",
           "Detect the boxes on each bathroom wall; height above the floor line (or above the FFL mark if flooring is not done) against the standard heights table.",
           "TAP with a tape.", "Standard heights table; FFL offset before flooring.", "± 25 mm"),
    433: d("Height and position", "Height of each electrical box and pipe stub above FFL.",
           "G-GEO", "Full", "High",
           "Same detector as ID 432, with pipe stubs and elbows added to the detection set.",
           "TAP with a tape.", "Standard heights table; FFL offset.", "± 25 mm", same=432),
    434: d("Height and position", "Mixer and basin point heights above FFL.",
           "G-GEO", "Full", "High",
           "Detect the pipe stubs, elbows and the mixer body; heights above the FFL mark compared with the standards table (mixer, basin, shower).",
           "TAP with a tape.", "Standard heights table; FFL offset.", "± 25 mm"),
    436: d("Plumb and level", "Wall lean before dado tiling.",
           "G-PLUMB", "Partial", "Medium",
           "Same detector as ID 84, run before dado tiling.",
           "BT-TOOL on flagged walls.", "Plumb tolerance.", "About 6 mm per 3 m", same=84),
    437: d("Height and position", "Ventilator position on the wall; sill height.",
           "G-GEO", "Full", "High",
           "Same detector as ID 128, applied to the ventilator opening.",
           "TAP with a tape.", "Floor plan and window schedule.", "Position ± 50 mm; sill ± 25 mm", same=128),
    438: d("Presence and spec match", "Bull marks, level lines and the starter row present.",
           "G-VIS", "Partial", "Low",
           "The vision model looks for level lines, bull marks and the starter row on each wall. Chalk lines are faint from the room centre, so expect a TAP request often.",
           "TAP photo of each marked wall.", "Dado layout (height, starting line).", "Marks present"),
    441: d("Thickness and joints", "Joint width between dado tiles.",
           "TAP", "Full", "High",
           "One close-up per bathroom wall with a tile edge in frame; the tile size gives the scale and joints measure to about 0.3 mm. Spacer presence can be checked in the same photo.",
           "G-GEO coarse check: course pitch minus tile size over several courses.", "Tile size; limit of 3 mm.", "Under 3 mm"),
    456: d("Height and position", "Top edge of the coat above floor on every wet-area wall.",
           "G-GEO", "Full", "High",
           "The coat is a distinct colour. Detect its top edge on each wall, convert to height above the floor line, compare with 300 mm and flag gaps.",
           "TAP with a tape.", "Required height (300 mm).", "At least 300 mm"),
}

# ----------------------------------------------------------------------------
# Styling
# ----------------------------------------------------------------------------
FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_GREY = Font(name=FONT, size=10, color="595959")
F_BOLD = Font(name=FONT, size=10, bold=True)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True)
F_BLUE = Font(name=FONT, size=10, color="0000FF")
F_GREEN = Font(name=FONT, size=10, color="008000")
FILL_HEAD = PatternFill("solid", fgColor="1F3864")
FILL_SRC = PatternFill("solid", fgColor="D9E1F2")
FILL_DESIGN = PatternFill("solid", fgColor="E2EFDA")
FILL_TRACK = PatternFill("solid", fgColor="FFF2CC")
FILL_REF = PatternFill("solid", fgColor="EDEDED")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_range(ws, ref, font=None, fill=None, align=None, border=None, number_format=None):
    for row in ws[ref]:
        for c in row:
            if font is not None:
                c.font = font
            if fill is not None:
                c.fill = fill
            if align is not None:
                c.alignment = align
            if border is not None:
                c.border = border
            if number_format is not None:
                c.number_format = number_format


def set_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def fit_row_heights(ws, first_row, last_row, widths, min_h=15, max_h=170):
    """Estimate wrapped line counts and set row heights (openpyxl cannot autofit)."""
    for r in range(first_row, last_row + 1):
        lines = 1
        for col, w in widths.items():
            v = ws[f"{col}{r}"].value
            if v is None or isinstance(v, (int, float, dt.date)):
                continue
            text = str(v)
            if text.startswith("="):
                continue
            per_line = max(1.0, w * 1.15)
            n = sum(max(1, math.ceil(len(part) / per_line)) for part in text.split("\n"))
            lines = max(lines, n)
        ws.row_dimensions[r].height = max(min_h, min(max_h, lines * 12.75 + 4))


# ----------------------------------------------------------------------------
# Read the export
# ----------------------------------------------------------------------------
with SRC.open(encoding="utf-8-sig", newline="") as fh:
    reader = csv.DictReader(fh)
    SRC_COLS = reader.fieldnames
    ROWS = list(reader)
ROWS.sort(key=lambda r: int(r["sno"]))
assert len(ROWS) == 55, len(ROWS)
missing = [r["sno"] for r in ROWS if int(r["sno"]) not in DESIGN]
assert not missing, f"no design for sno {missing}"

wb = Workbook()

# ============================================================================
# Tab: QC Tracker
# ============================================================================
ws = wb.active
ws.title = "QC Tracker"

HEADERS = [
    # (header, group, width)
    ("ID (sno)", "src", 7),
    ("Stage", "src", 22),
    ("Task", "src", 28),
    ("QC check", "src", 44),
    ("Sub-phase", "src", 16),
    ("Current SE method", "src", 10),
    ("Priority", "src", 10),
    ("No-go", "src", 7),
    ("Level (2.0)", "src", 9),
    ("Stage type", "src", 10),
    ("In 2.0", "src", 7),
    ("Live today", "src", 8),
    ("Check family", "design", 17),
    ("What to measure or judge", "design", 30),
    ("Primary method", "design", 11),
    ("Automation level", "design", 11),
    ("Confidence", "design", 10),
    ("How it would work", "design", 56),
    ("Fallback or upgrade", "design", 40),
    ("Inputs the prompt needs", "design", 30),
    ("Target tolerance (typical; confirm against your spec)", "design", 26),
    ("Same detector as (ID)", "design", 10),
    ("Extra SE time (s)", "design", 9),
    ("Status", "track", 12),
    ("Owner", "track", 12),
    ("Target sprint", "track", 10),
    ("Validation accuracy", "track", 10),
    ("Validation sample (n)", "track", 9),
    ("Last updated", "track", 11),
    ("Notes", "track", 32),
    ("qc_id", "ref", 34),
]
COL = {h: get_column_letter(i + 1) for i, (h, _, _) in enumerate(HEADERS)}
WIDTHS = {get_column_letter(i + 1): w for i, (_, _, w) in enumerate(HEADERS)}
LAST_COL = get_column_letter(len(HEADERS))
HEAD_ROW, FIRST, = 2, 3
LAST = FIRST + len(ROWS) - 1
RANGE_END = 200  # formulas look this far down so rows can be added

# Row 1: group bands
groups = [("src", "From your QC sheet (trimmed to the useful columns)", FILL_SRC),
          ("design", "Proposed automation design", FILL_DESIGN),
          ("track", "Your tracking (fill in)", FILL_TRACK),
          ("ref", "Ref", FILL_REF)]
for key, label, fill in groups:
    cols = [get_column_letter(i + 1) for i, (_, g, _) in enumerate(HEADERS) if g == key]
    ref = f"{cols[0]}1:{cols[-1]}1"
    if len(cols) > 1:
        ws.merge_cells(ref)
    ws[f"{cols[0]}1"] = label
    style_range(ws, ref, font=F_BOLD, fill=fill, align=CENTER, border=BORDER)
ws.row_dimensions[1].height = 20

# Row 2: headers
for i, (h, _, _) in enumerate(HEADERS):
    c = ws.cell(row=HEAD_ROW, column=i + 1, value=h)
    c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BORDER
ws.row_dimensions[HEAD_ROW].height = 42
ws[f"A{HEAD_ROW}"].comment = Comment(
    "Columns A-L: copied from the uploaded QC sheet (Sheet2.csv, 2026-10-07); text as exported, "
    "with TRUE/FALSE and YES/NO shown as Yes/No. Columns M-W: proposed automation design. "
    "Columns X-AD: for the team to fill in. " + SOURCE_NOTE, "Workbook notes")
ws[f"{COL['Extra SE time (s)']}{HEAD_ROW}"].comment = Comment(
    "Looked up from the Methods tab by the primary method code. Change the seconds there, not here.", "Workbook notes")
ws[f"{COL['Target tolerance (typical; confirm against your spec)']}{HEAD_ROW}"].comment = Comment(
    "Typical values for residential work, written from general practice, not from your standards. "
    "Replace with the numbers in your own spec and drawings before encoding any check.", "Workbook notes")
ws[f"{COL['Same detector as (ID)']}{HEAD_ROW}"].comment = Comment(
    "Blank means this row is the first (canonical) check using its detector. A number points at the row "
    "whose detector also serves this check, so the two are built once.", "Workbook notes")

level_map = {"Floor > Space": "Space", "Floor > Space > Wall": "Wall", "": ""}
stage_type_map = {"FLOORWISE_STAGE": "Floorwise", "INITIAL_ONE_TIME_STAGE": "One-time"}

for i, r in enumerate(ROWS):
    row = FIRST + i
    sno = int(r["sno"])
    dsg = DESIGN[sno]
    values = {
        "ID (sno)": sno,
        "Stage": r["stage"],
        "Task": r["task"],
        "QC check": r["qc"],
        "Sub-phase": r["Phase(More Divided)"],
        "Current SE method": r["QC_TYPE_MEASUREMENT_VISUAL_TOOl"].strip().title(),
        "Priority": r["qc_type"],
        "No-go": "Yes" if r["is_no_go"].upper() == "TRUE" else "No",
        "Level (2.0)": level_map.get(r["2_0_hierarchy"], r["2_0_hierarchy"]),
        "Stage type": stage_type_map.get(r["stage_type"], r["stage_type"]),
        "In 2.0": r["in_2_0"].title(),
        "Live today": "Yes" if r["CURRENT LIVE"] else "No",
        "Check family": dsg["family"],
        "What to measure or judge": dsg["measure"],
        "Primary method": dsg["method"],
        "Automation level": dsg["level"],
        "Confidence": dsg["conf"],
        "How it would work": dsg["how"],
        "Fallback or upgrade": dsg["fallback"],
        "Inputs the prompt needs": dsg["inputs"],
        "Target tolerance (typical; confirm against your spec)": dsg["tol"],
        "Same detector as (ID)": dsg["same"] if dsg["same"] != "" else None,
        "Extra SE time (s)": (f"=INDEX(Methods!$D$2:$D${len(METHODS) + 1},"
                              f"MATCH({COL['Primary method']}{row},Methods!$A$2:$A${len(METHODS) + 1},0))"),
        "Status": "Not started",
        "Owner": None, "Target sprint": None, "Validation accuracy": None,
        "Validation sample (n)": None, "Last updated": None, "Notes": None,
        "qc_id": r["qc_id"],
    }
    for h, _, _ in HEADERS:
        ws[f"{COL[h]}{row}"] = values[h]
    # formats per group
    for h, g, _ in HEADERS:
        c = ws[f"{COL[h]}{row}"]
        c.alignment = WRAP_TOP
        c.border = BORDER
        if g == "src" or g == "ref":
            c.font = F_GREY
        elif g == "track":
            c.font = F_BASE
            c.fill = FILL_TRACK
        else:
            c.font = F_BASE
    ws[f"{COL['Extra SE time (s)']}{row}"].font = F_GREEN  # link to another sheet
    ws[f"{COL['Extra SE time (s)']}{row}"].number_format = "0"
    ws[f"{COL['Validation accuracy']}{row}"].number_format = "0%"
    ws[f"{COL['Validation sample (n)']}{row}"].number_format = "0"
    ws[f"{COL['Last updated']}{row}"].number_format = "yyyy-mm-dd"
    for h in ("ID (sno)", "No-go", "In 2.0", "Live today", "Same detector as (ID)", "Extra SE time (s)",
              "Validation sample (n)", "Confidence", "Automation level", "Primary method"):
        ws[f"{COL[h]}{row}"].alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)

set_widths(ws, WIDTHS)
fit_row_heights(ws, FIRST, LAST, WIDTHS)
ws.freeze_panes = f"{COL['Sub-phase']}{FIRST}"

# Table with filter buttons and banding
tbl = Table(displayName="QCTracker", ref=f"A{HEAD_ROW}:{LAST_COL}{LAST}")
tbl.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True, showColumnStripes=False)
ws.add_table(tbl)

# Collapsible groups: Stage/Task and the flag columns, so the check text, design and tracking stay visible
for col in ("B", "C"):
    ws.column_dimensions[col].outline_level = 1
for col in [COL[h] for h in ("Sub-phase", "Current SE method", "Priority", "No-go", "Level (2.0)",
                              "Stage type", "In 2.0", "Live today")]:
    ws.column_dimensions[col].outline_level = 1
ws.sheet_format.outlineLevelCol = 1
ws.sheet_properties.outlinePr.summaryRight = False

# Dropdowns
def add_list_validation(ws, formula, ref):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.error, dv.errorTitle = "Pick a value from the list.", "Not in list"
    ws.add_data_validation(dv)
    dv.add(ref)

add_list_validation(ws, '"' + ",".join(STATUSES) + '"', f"{COL['Status']}{FIRST}:{COL['Status']}{RANGE_END}")
add_list_validation(ws, '"' + ",".join(LEVELS) + '"', f"{COL['Automation level']}{FIRST}:{COL['Automation level']}{RANGE_END}")
add_list_validation(ws, '"' + ",".join(CONFIDENCES) + '"', f"{COL['Confidence']}{FIRST}:{COL['Confidence']}{RANGE_END}")
add_list_validation(ws, f"=Methods!$A$2:$A${len(METHODS) + 1}", f"{COL['Primary method']}{FIRST}:{COL['Primary method']}{RANGE_END}")

# Conditional formatting on Status and Automation level
status_ref = f"{COL['Status']}{FIRST}:{COL['Status']}{RANGE_END}"
for label, color in (("Live", "C6EFCE"), ("Validating", "FFEB9C"), ("Prototype", "FFEB9C"),
                     ("Exploring", "DDEBF7"), ("Parked", "D9D9D9")):
    ws.conditional_formatting.add(status_ref, CellIsRule(operator="equal", formula=[f'"{label}"'],
                                                         fill=PatternFill("solid", fgColor=color)))
level_ref = f"{COL['Automation level']}{FIRST}:{COL['Automation level']}{RANGE_END}"
for label, color in (("Full", "C6EFCE"), ("Partial", "FFEB9C"), ("Evidence", "DDEBF7"), ("Manual", "F8CBAD")):
    ws.conditional_formatting.add(level_ref, CellIsRule(operator="equal", formula=[f'"{label}"'],
                                                        fill=PatternFill("solid", fgColor=color)))

# Named ranges for formulas elsewhere (plain A1 ranges are used; these are documented helpers)
TR = "'QC Tracker'!"
def tr(h):
    c = COL[h]
    return f"{TR}${c}${FIRST}:${c}${RANGE_END}"

# ============================================================================
# Tab: Methods
# ============================================================================
wm = wb.create_sheet("Methods")
m_heads = ["Code", "Method", "What the SE does differently", "Extra SE time per use (s)",
           "Accuracy you can expect", "Build effort (S/M/L)", "Checks using it", "Limits and notes"]
m_widths = {"A": 10, "B": 26, "C": 40, "D": 12, "E": 44, "F": 10, "G": 10, "H": 50}
for i, h in enumerate(m_heads):
    c = wm.cell(row=1, column=i + 1, value=h)
    c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BORDER
wm.row_dimensions[1].height = 32
wm["D1"].comment = Comment("Estimates of added site time per use, written as a starting point. "
                           "Blue cells: replace with timings measured on site. Source: estimate.", "Workbook notes")
wm["G1"].comment = Comment("Count of tracker rows whose primary method is this code.", "Workbook notes")
for i, (code, name, se, secs, acc, effort, limits) in enumerate(METHODS):
    r = i + 2
    vals = [code, name, se, secs, acc, effort, f'=COUNTIF({tr("Primary method")},A{r})', limits]
    for j, v in enumerate(vals):
        c = wm.cell(row=r, column=j + 1, value=v)
        c.font, c.alignment, c.border = F_BASE, WRAP_TOP, BORDER
    wm[f"A{r}"].font = F_BOLD
    wm[f"D{r}"].font = F_BLUE
    wm[f"D{r}"].fill = FILL_INPUT
    wm[f"D{r}"].number_format = "0"
    wm[f"F{r}"].alignment = Alignment(horizontal="center", vertical="top")
    wm[f"G{r}"].alignment = Alignment(horizontal="center", vertical="top")
    wm[f"G{r}"].font = F_GREEN
    wm[f"G{r}"].number_format = "0"
set_widths(wm, m_widths)
fit_row_heights(wm, 2, len(METHODS) + 1, m_widths)
wm.freeze_panes = "B2"
M_LAST = len(METHODS) + 1

# ============================================================================
# Tab: Dashboard
# ============================================================================
wd = wb.create_sheet("Dashboard")
wd["A1"] = "QC automation dashboard"
wd["A1"].font = F_TITLE
wd["A2"] = "Every figure is a formula over the QC Tracker tab. Edit there, not here."
wd["A2"].font = F_GREY

def head(ws, ref_cells):
    for ref, text in ref_cells:
        ws[ref] = text
        ws[ref].font, ws[ref].fill, ws[ref].alignment, ws[ref].border = F_HEAD, FILL_HEAD, CENTER, BORDER

def label(ws, ref, text, bold=False):
    ws[ref] = text
    ws[ref].font = F_BOLD if bold else F_BASE
    ws[ref].alignment = WRAP_TOP
    ws[ref].border = BORDER

def num(ws, ref, formula, fmt="0"):
    ws[ref] = formula
    ws[ref].font = F_GREEN
    ws[ref].number_format = fmt
    ws[ref].alignment = Alignment(horizontal="right", vertical="top")
    ws[ref].border = BORDER

ID, ST, PR, NG, FAM, MTH, LVL, CONF, SAME, SECS, STAT = (
    tr("ID (sno)"), tr("Stage"), tr("Priority"), tr("No-go"), tr("Check family"), tr("Primary method"),
    tr("Automation level"), tr("Confidence"), tr("Same detector as (ID)"), tr("Extra SE time (s)"), tr("Status"))

# Key figures (A4:B16)
head(wd, [("A4", "Key figures"), ("B4", "Value")])
key = [
    ("Total checks", f"=COUNTA({ID})", "0"),
    ("Distinct detectors to build (rows without a 'same detector as' pointer)", f"=COUNTA({ID})-COUNTA({SAME})", "0"),
    ("Checks needing no new site action (G- methods)", f"=SUMPRODUCT((LEFT({MTH},2)=\"G-\")*1)", "0"),
    ("Share needing no new site action", "=IF(B5=0,0,B7/B5)", "0%"),
    ("Critical checks", f"=COUNTIF({PR},\"CRITICAL\")", "0"),
    ("Critical checks with Full automation", f"=COUNTIFS({PR},\"CRITICAL\",{LVL},\"Full\")", "0"),
    ("No-go checks", f"=COUNTIF({NG},\"Yes\")", "0"),
    ("No-go checks covered (Full or Evidence)", f"=COUNTIFS({NG},\"Yes\",{LVL},\"Full\")+COUNTIFS({NG},\"Yes\",{LVL},\"Evidence\")", "0"),
    ("Added SE time if every check runs its primary method (s, upper bound)", f"=SUM({SECS})", "0"),
    ("Same, in minutes", "=B13/60", "0.0"),
    ("Checks Live", f"=COUNTIF({STAT},\"Live\")", "0"),
    ("Share Live", "=IF(B5=0,0,B15/B5)", "0%"),
]
for i, (lab, f, fmt) in enumerate(key):
    r = 5 + i
    label(wd, f"A{r}", lab)
    num(wd, f"B{r}", f, fmt)

# By status (D4:F11)
head(wd, [("D4", "Status"), ("E4", "Checks"), ("F4", "Share")])
for i, s in enumerate(STATUSES):
    r = 5 + i
    label(wd, f"D{r}", s)
    num(wd, f"E{r}", f"=COUNTIF({STAT},D{r})")
    num(wd, f"F{r}", f"=IF($B$5=0,0,E{r}/$B$5)", "0%")
r = 5 + len(STATUSES)
label(wd, f"D{r}", "Total", bold=True)
num(wd, f"E{r}", f"=SUM(E5:E{r - 1})")
num(wd, f"F{r}", f"=IF($B$5=0,0,E{r}/$B$5)", "0%")

# By automation level (D13:G18)
head(wd, [("D13", "Automation level"), ("E13", "Checks"), ("F13", "Of which critical"), ("G13", "Of which no-go")])
for i, lv in enumerate(LEVELS):
    r = 14 + i
    label(wd, f"D{r}", lv)
    num(wd, f"E{r}", f"=COUNTIF({LVL},D{r})")
    num(wd, f"F{r}", f"=COUNTIFS({LVL},D{r},{PR},\"CRITICAL\")")
    num(wd, f"G{r}", f"=COUNTIFS({LVL},D{r},{NG},\"Yes\")")
r = 14 + len(LEVELS)
label(wd, f"D{r}", "Total", bold=True)
num(wd, f"E{r}", f"=SUM(E14:E{r - 1})")
num(wd, f"F{r}", f"=SUM(F14:F{r - 1})")
num(wd, f"G{r}", f"=SUM(G14:G{r - 1})")

# By confidence (D21:E24)
head(wd, [("D21", "Confidence"), ("E21", "Checks")])
for i, cf in enumerate(CONFIDENCES):
    r = 22 + i
    label(wd, f"D{r}", cf)
    num(wd, f"E{r}", f"=COUNTIF({CONF},D{r})")

# By primary method (A19:E30)
head(wd, [("A19", "Primary method"), ("B19", "Checks"), ("C19", "Added SE time (s)")])
# (the method name is on the Methods tab; keep this block compact)
for i in range(len(METHODS)):
    r = 20 + i
    wd[f"A{r}"] = f"=Methods!A{i + 2}"
    wd[f"A{r}"].font, wd[f"A{r}"].border, wd[f"A{r}"].alignment = F_GREEN, BORDER, WRAP_TOP
    num(wd, f"B{r}", f"=COUNTIF({MTH},A{r})")
    num(wd, f"C{r}", f"=SUMIF({MTH},A{r},{SECS})")
r = 20 + len(METHODS)
label(wd, f"A{r}", "Total", bold=True)
num(wd, f"B{r}", f"=SUM(B20:B{r - 1})")
num(wd, f"C{r}", f"=SUM(C20:C{r - 1})")

# By check family (A33:C44)
families = sorted({v["family"] for v in DESIGN.values()})
head(wd, [("A33", "Check family"), ("B33", "Checks"), ("C33", "Distinct detectors")])
for i, fam in enumerate(families):
    r = 34 + i
    label(wd, f"A{r}", fam)
    num(wd, f"B{r}", f"=COUNTIF({FAM},A{r})")
    num(wd, f"C{r}", f"=SUMPRODUCT(({FAM}=A{r})*({SAME}=\"\"))")
r = 34 + len(families)
label(wd, f"A{r}", "Total", bold=True)
num(wd, f"B{r}", f"=SUM(B34:B{r - 1})")
num(wd, f"C{r}", f"=SUM(C34:C{r - 1})")

# By stage (E27:I40)
stages = sorted({row["stage"] for row in ROWS})
head(wd, [("E27", "Stage"), ("F27", "Checks"), ("G27", "Critical"), ("H27", "Full automation"), ("I27", "Live")])
for i, stg in enumerate(stages):
    r = 28 + i
    label(wd, f"E{r}", stg)
    num(wd, f"F{r}", f"=COUNTIF({ST},E{r})")
    num(wd, f"G{r}", f"=COUNTIFS({ST},E{r},{PR},\"CRITICAL\")")
    num(wd, f"H{r}", f"=COUNTIFS({ST},E{r},{LVL},\"Full\")")
    num(wd, f"I{r}", f"=COUNTIFS({ST},E{r},{STAT},\"Live\")")
r = 28 + len(stages)
label(wd, f"E{r}", "Total", bold=True)
for c in "FGHI":
    num(wd, f"{c}{r}", f"=SUM({c}28:{c}{r - 1})")

set_widths(wd, {"A": 46, "B": 10, "C": 16, "D": 20, "E": 40, "F": 11, "G": 11, "H": 12, "I": 9})
wd.freeze_panes = "A4"

# ============================================================================
# Tab: Experiments
# ============================================================================
we = wb.create_sheet("Experiments")
e_heads = ["Date", "QC ID (sno)", "Method", "What was tried", "Sample (n)", "Result",
           "Decision", "Next step", "Owner"]
e_widths = {"A": 11, "B": 9, "C": 10, "D": 44, "E": 9, "F": 44, "G": 11, "H": 36, "I": 14}
for i, h in enumerate(e_heads):
    c = we.cell(row=1, column=i + 1, value=h)
    c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BORDER
we.row_dimensions[1].height = 30
example = [dt.date(2026, 10, 7), 456, "G-GEO",
           "Waterproofing line height read from 12 bathroom captures, tripod fixed at 1.50 m. Compared with tape readings.",
           12, "Within 20 mm of the tape on 11 of 12; the miss was a dark wall where the coat edge was not found.",
           "Iterate", "Turn the torch on for wet-area captures and rerun.", "Example row, overwrite"]
for j, v in enumerate(example):
    c = we.cell(row=2, column=j + 1, value=v)
    c.font, c.alignment, c.border, c.fill = F_BASE, WRAP_TOP, BORDER, FILL_TRACK
we["A2"].number_format = "yyyy-mm-dd"
for r in range(3, 60):
    for j in range(len(e_heads)):
        c = we.cell(row=r, column=j + 1)
        c.fill, c.border, c.font, c.alignment = FILL_TRACK, BORDER, F_BASE, WRAP_TOP
    we[f"A{r}"].number_format = "yyyy-mm-dd"
add_list_validation(we, '"Adopt,Iterate,Drop"', "G2:G200")
add_list_validation(we, f"=Methods!$A$2:$A${M_LAST}", "C2:C200")
set_widths(we, e_widths)
fit_row_heights(we, 2, 2, e_widths)
we.freeze_panes = "A2"

# ============================================================================
# Tab: Source data (the export, unchanged)
# ============================================================================
wsrc = wb.create_sheet("Source data")
for j, h in enumerate(SRC_COLS):
    c = wsrc.cell(row=1, column=j + 1, value=h)
    c.font, c.fill, c.alignment, c.border = F_HEAD, FILL_HEAD, CENTER, BORDER
wsrc["A1"].comment = Comment(SOURCE_NOTE + ". 55 rows, all columns as exported; nothing changed.", "Workbook notes")
for i, r in enumerate(ROWS):
    for j, h in enumerate(SRC_COLS):
        v = r[h]
        if h == "sno":
            v = int(v)
        elif v == "":
            v = None
        c = wsrc.cell(row=i + 2, column=j + 1, value=v)
        c.font = F_GREY
        if isinstance(v, str) and (v[:1] in "=+-@" ):
            c.value = "'" + v  # keep text as text
for j, h in enumerate(SRC_COLS):
    wsrc.column_dimensions[get_column_letter(j + 1)].width = 14 if h not in ("stage", "task", "qc") else 40
wsrc.row_dimensions[1].height = 30
wsrc.freeze_panes = "B2"

# ============================================================================
# Tab: Read me (first tab)
# ============================================================================
wr = wb.create_sheet("Read me", 0)
set_widths(wr, {"A": 4, "B": 30, "C": 100})
R = [1]

def para(text, bold=False, col="B", size=None, merge=True):
    r = R[0]
    cell = wr[f"{col}{r}"]
    cell.value = text
    cell.font = Font(name=FONT, size=size or 10, bold=bold)
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    if merge and col == "B":
        wr.merge_cells(f"B{r}:C{r}")
        per_line = 130 * 1.15
        lines = sum(max(1, math.ceil(len(p) / per_line)) for p in str(text).split("\n"))
        wr.row_dimensions[r].height = max(15, lines * 12.75 + 4)
    R[0] += 1

def blank():
    R[0] += 1

def pair(k, v, kfont=None):
    r = R[0]
    wr[f"B{r}"] = k
    wr[f"B{r}"].font = kfont or F_BOLD
    wr[f"B{r}"].alignment = WRAP_TOP
    wr[f"C{r}"] = v
    wr[f"C{r}"].font = F_BASE
    wr[f"C{r}"].alignment = WRAP_TOP
    lines = max(math.ceil(len(str(v)) / (100 * 1.15)), math.ceil(len(str(k)) / 28), 1)  # bold keys are wider
    wr.row_dimensions[r].height = max(15, lines * 12.75 + 4)
    R[0] += 1

para("Post-wall QC automation tracker", bold=True, size=14)
para("55 post-wall quality checks from the QC sheet, each mapped to the cheapest capture method that could automate it, "
     "with space to track the work.")
blank()
para("How to use this workbook", bold=True, size=11)
pair("QC Tracker", "One row per check. Columns A to L are copied from the QC sheet. Columns M to W hold the proposed design "
     "for each check. Columns X to AD (yellow) are yours to fill as work progresses. Filter and sort with the header "
     "buttons. The + controls above the columns collapse Stage/Task and the flag columns, leaving the check text, the "
     "design and the tracking in view.")
pair("Methods", "The method codes used in the tracker: what each asks of the site engineer, the added seconds it costs "
     "(blue cells, estimates you can edit) and the accuracy to expect. Changing a time here updates the tracker and the dashboard.")
pair("Dashboard", "Counts by status, method, automation level, confidence, check family and stage. Formulas only; nothing to type.")
pair("Experiments", "A log of what was tried and decided. The first row is an example to overwrite.")
pair("Source data", "The QC sheet export as received, for reference. Columns left out of the tracker: those identical on every "
     "row (Phase, Removal Status, GYRO+AI, 2D PHONE+AI, AI_NOT_ENABLED, measurement_type, hierarchy, 2_0_measurement_type), "
     "empty ones (capture_type, 2.0 Current, measurement_config) and admin fields (config_updated_on, stage_id, task_id, "
     "2_0_measurement_config, qc_category, task_type).")
blank()
para("Legend", bold=True, size=11)
r = R[0]
wr[f"B{r}"] = "Yellow fill"; wr[f"B{r}"].fill = FILL_TRACK; wr[f"B{r}"].font = F_BASE
wr[f"C{r}"] = "Cells for the team to fill in."; wr[f"C{r}"].font = F_BASE; R[0] += 1
r = R[0]
wr[f"B{r}"] = "Blue text"; wr[f"B{r}"].font = F_BLUE
wr[f"C{r}"] = "An estimate you can change (Methods tab, column D)."; wr[f"C{r}"].font = F_BASE; R[0] += 1
r = R[0]
wr[f"B{r}"] = "Green text"; wr[f"B{r}"].font = F_GREEN
wr[f"C{r}"] = "A formula that reads another tab."; wr[f"C{r}"].font = F_BASE; R[0] += 1
r = R[0]
wr[f"B{r}"] = "Grey text"; wr[f"B{r}"].font = F_GREY
wr[f"C{r}"] = "Copied from the QC sheet; edit it there."; wr[f"C{r}"].font = F_BASE; R[0] += 1
blank()
para("Definitions used in the tracker", bold=True, size=11)
pair("Full", "The capture, plus at most one requested photo, decides pass or fail. The SE reviews flags only.")
pair("Partial", "The system screens and flags. The SE confirms borderline cases, or part of the check stays manual.")
pair("Evidence", "The system records proof or an approval. The judgement stays with a person.")
pair("Manual", "Stays with the SE and a physical tool. The app records the reading.")
pair("Confidence", "How likely the primary method reaches useful accuracy in production with the current gyro capture: High, Medium or Low.")
pair("Same detector as", "The ID of the check whose detector also serves this row. Build it once and run it at both stages.")
r = R[0]
wr[f"B{r}"] = "Distinct detectors"; wr[f"B{r}"].font = F_BOLD
wr[f"C{r}"] = '=Dashboard!B6&" detectors cover "&Dashboard!B5&" checks."'; wr[f"C{r}"].font = F_GREEN; R[0] += 1
blank()
para("Proposed pipeline from the existing gyro capture", bold=True, size=11)
steps = [
    "1. Gravity-align every frame with the gyro and accelerometer, then stitch the bands into one equirectangular panorama. "
    "The phone turns about a fixed point on the tripod, so stitching is a pure-rotation problem and is reliable.",
    "2. Run a room-layout model on the panorama to get the floor and ceiling boundaries and the corners. Set the metric scale "
    "from the camera height: either fix the tripod height (a mark on the stick) or let the app derive it from the "
    "floor-to-ceiling height in the drawing. No new site action is needed.",
    "3. Re-project the panorama into one flat elevation image per wall and one top-down image of the floor. Straight lines "
    "stay straight in these views, which the measurement checks rely on, and vision models read them better than a "
    "distorted panorama.",
    "4. Detect the elements each check needs (openings, boxes, pipe stubs, bands, skirting, tiles, grills, columns) and "
    "convert pixel positions to millimetres with the wall distance from step 2. Known objects in the scene (block height, "
    "tile size) give a second, finer scale.",
    "5. Send each check's crop, the drawing or specification values and the standard-heights table to the vision model "
    "with a check-specific prompt. It returns pass, flag or needs-a-photo.",
    "6. Only when a check needs it, the app asks the SE for one close-up photo (TAP) or a one-tap confirmation (ATTEST). "
    "The added site time stays at seconds per room.",
]
for s in steps:
    para(s)
blank()
para("Why not send the cylinder straight to the language model", bold=True, size=11)
para("A vision model judges well and measures badly. It cannot tell 10 mm from 14 mm or 0.3° from 1°, and it reads distorted "
     "projections worse than flat ones. Use geometry to measure, and the model to judge finish, presence and spec matches. "
     "The stitched panorama or cylinder is still worth building: it is the best view for human review and the input for the "
     "layout model.")
blank()
para("What one viewpoint can and cannot do", bold=True, size=11)
pair("Heights and positions on a wall (G-GEO)", "About ±20–30 mm at 3 m; about ±5 mm with a tile or block as the ruler. Enough for standard "
     "heights, lintel and sill heights, skirting, the waterproofing line and box positions.")
pair("Plumb and level of edges (G-PLUMB)", "About ±0.2°, which is ±7 mm over 2 m. Screens for gross errors; not a replacement for a plumb bob at ±3 mm.")
pair("Alignment on the floor (G-FLOOR)", "About 1 mm with the tile as the ruler. Grout-line alignment and pattern.")
pair("Close-up with a reference object (TAP)", "About 0.3–0.5 mm. Joint widths, granite thickness, weld surface.")
pair("Angles between walls (G-GEO)", "About ±1°. Gross out-of-square rooms only.")
pair("Not possible from one point", "Out-of-plane bow or flatness of a wall, tile lippage, anything hidden behind materials or outside the frames, "
     "functional tests such as latching, and process checks that happened earlier (vibration during a pour).")
blank()
para("What changes for the site engineer", bold=True, size=11)
for s in [
    "Close all shutters before the gyro capture (ID 127).",
    "Optional: one extra band at about −70° so the floor near the tripod is covered (G-FLOOR).",
    "A close-up photo only when the app asks. Four checks always need one (IDs 171, 176, 272, 441); a few others ask when the model is unsure.",
    "Phone laid on the mortar bed for the slope check (ID 206).",
    "One exterior capture per site for the compound wall (IDs 107 to 109).",
    "One-tap confirmations: vibrator evidence (69), latch test (127), client approval of the tile layout (199, 222).",
]:
    para(s)
blank()
para("Suggested order of work", bold=True, size=11)
for s in [
    "1. G-GEO heights and positions, no added site time and high confidence: IDs 34, 40, 83, 128, 216, 317, 432, 433, 434, 437, 456. One detector family covers eleven checks.",
    "2. G-VIS presence and finish checks with the spec in the prompt: IDs 123, 125, 186, 219, 274, 286.",
    "3. G-FLOOR grout alignment (207) and the client approval flow (199, 222).",
    "4. G-PLUMB screening for openings, corners, walls and grills. Validate against plumb-bob readings before trusting it.",
    "5. TAP close-ups (171, 176, 272, 441), the IMU slope check (206), then the exterior capture (107 to 109).",
]:
    para(s)

wb.save(OUT)
print(f"wrote {OUT} with {len(ROWS)} checks, {len(METHODS)} methods")
