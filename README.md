# Fusion 360 Timing Pulley + Belt Add-Ins

This repository contains **two Fusion 360 Python add-ins**:

1. `AdjustableDrivePulley`: builds a drive + driven timing pulley pair.
2. `AdjustableTimingBeltDrive`: builds a timing belt loop that automatically accounts for drive/driven pulley center distance.

## Add-In 1: Adjustable Timing Pulley Pair

Path:

- `AdjustableDrivePulley/AdjustableDrivePulley.py`
- `AdjustableDrivePulley/AdjustableDrivePulley.manifest`

What it does:

- Adjustable drive and driven tooth count.
- Optional auto-tooth mode from:
  - target ratio (driven:drive),
  - max pulley diameter,
  - min/max tooth range.
- Adjustable belt pitch, tooth height, pulley thickness, tip clearance, and bores.
- Manual center distance or auto center distance from belt pitch count.
- Creates separate drive and driven pulley components.
- Tags generated pulleys with custom Fusion attributes for robust downstream detection.
- Optional CSV summary export for basic BOM/reporting.

Core math:

- Pitch radius: `r = p / (2 sin(pi / z))`
- Belt-length approximation for auto center:
  - `L = 2m + (z1 + z2)/2 + ((z2 - z1)^2)/(4 pi^2 m)`
  - `m = C / p`

Validation and engineering checks:

- Minimum pulley tooth count: `9`.
- Tooth height must be less than belt pitch.
- Bore diameters are validated against computed root/hub envelope.
- Center distance warns when outside the common recommended range of `30-50` pitch lengths.
- Ratio mode warns/fails when no tooth pair can satisfy ratio + diameter + range constraints.

## Add-In 2: Adjustable Timing Belt Drive

Path:

- `AdjustableTimingBeltDrive/AdjustableTimingBeltDrive.py`
- `AdjustableTimingBeltDrive/AdjustableTimingBeltDrive.manifest`

What it does:

- Creates a timing belt loop around a drive and driven pulley pair.
- Uses drive/driven tooth counts and belt pitch to compute pitch radii.
- Automatically accounts for pulley center distance by:
  - reading selected drive/driven pulley occurrences, or
  - finding attribute-tagged pulleys created by `AdjustableDrivePulley`, with name-based fallback.
- Supports manual center distance when selection mode is off.
- Supports auto belt tooth count or manual tooth count.
- Includes live in-dialog summary text that updates with current inputs.
- Optional CSV summary export for belt/BOM-style reporting.

Notes:

- Auto tooth count chooses the nearest practical count and reports effective pitch deviation.
- Can force even tooth count and explicitly reports odd-tooth phase-indexing notes.
- Warns when center distance is outside `30-50` pitches and when computed wrap angle is low.
- Geometry is generated as an actual belt body with repeated inward-facing tooth profiles along the path.

## Install in Fusion 360

Install each add-in folder separately:

1. Open Fusion 360.
2. Go to `Utilities` -> `Add-Ins` -> `Scripts and Add-Ins`.
3. Open the `Add-Ins` tab.
4. Click `+` / `Add Existing`.
5. Select `AdjustableDrivePulley` and/or `AdjustableTimingBeltDrive`.
6. Run the add-in(s).

## Recommended workflow

1. Run `Adjustable Timing Pulley Pair` to generate drive and driven pulleys.
2. Run `Adjustable Timing Belt Drive` with `Use Selected Pulley Centers` enabled.
3. Select the two pulley occurrences to auto-place the belt at the correct distance.
