import adsk.core
import adsk.fusion
import traceback
import math
import csv
import datetime

APP_NAME = 'Adjustable Timing Belt Drive'
CMD_ID = 'com.lenicolas.adjustablebeltdrive'
CMD_NAME = 'Adjustable Timing Belt Drive'
CMD_DESC = 'Generate a timing belt loop around drive and driven pulleys with automatic center-distance handling.'
WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'SolidCreatePanel'

MIN_SPROCKET_TEETH = 9
RECOMMENDED_CENTER_MIN_PITCHES = 30.0
RECOMMENDED_CENTER_MAX_PITCHES = 50.0
ATTRIBUTE_GROUP = 'com.lenicolas.pulley'
ATTR_ROLE = 'role'
ATTR_PAIR_ID = 'pair_id'

handlers = []


def run(context):
    app = adsk.core.Application.get()
    ui = app.userInterface
    try:
        command_def = ui.commandDefinitions.itemById(CMD_ID)
        if not command_def:
            command_def = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_DESC)

        on_command_created = CommandCreatedHandler()
        command_def.commandCreated.add(on_command_created)
        handlers.append(on_command_created)

        workspace = ui.workspaces.itemById(WORKSPACE_ID)
        panel = workspace.toolbarPanels.itemById(PANEL_ID)
        control = panel.controls.itemById(CMD_ID)
        if not control:
            control = panel.controls.addCommand(command_def)

        control.isPromoted = True
        control.isPromotedByDefault = True
    except Exception:
        if ui:
            ui.messageBox('Add-in start failed:\n{}'.format(traceback.format_exc()))


def stop(context):
    app = adsk.core.Application.get()
    ui = app.userInterface
    try:
        workspace = ui.workspaces.itemById(WORKSPACE_ID)
        panel = workspace.toolbarPanels.itemById(PANEL_ID)
        control = panel.controls.itemById(CMD_ID)
        if control:
            control.deleteMe()

        command_def = ui.commandDefinitions.itemById(CMD_ID)
        if command_def:
            command_def.deleteMe()
    except Exception:
        if ui:
            ui.messageBox('Add-in stop failed:\n{}'.format(traceback.format_exc()))


def _distance_2d(a, b):
    dx = b[0] - a[0]
    dy = b[1] - a[1]
    return math.sqrt((dx * dx) + (dy * dy))


def _lerp(a, b, t):
    return (a[0] + ((b[0] - a[0]) * t), a[1] + ((b[1] - a[1]) * t))


def _point_from_angle(center_xy, radius, angle):
    return (center_xy[0] + (radius * math.cos(angle)), center_xy[1] + (radius * math.sin(angle)))


def _normalize_2d(vx, vy):
    mag = math.sqrt((vx * vx) + (vy * vy))
    if mag <= 1e-9:
        return (0.0, 0.0)
    return (vx / mag, vy / mag)


def _point3d_xy(xy):
    return adsk.core.Point3D.create(xy[0], xy[1], 0)


def _polygon_area_2d(points):
    if len(points) < 3:
        return 0.0

    area = 0.0
    for i in range(len(points)):
        p1 = points[i]
        p2 = points[(i + 1) % len(points)]
        area += (p1[0] * p2[1]) - (p2[0] * p1[1])

    return area * 0.5


def _pitch_radius(belt_pitch, tooth_count):
    return belt_pitch / (2.0 * math.sin(math.pi / float(tooth_count)))


def _get_occurrence_center(occurrence):
    transform = occurrence.transform
    offset = transform.translation
    return (offset.x, offset.y, offset.z)


def _first_occurrence_for_component(root_component, component):
    occs = root_component.allOccurrencesByComponent(component)
    if occs and occs.count > 0:
        return occs.item(0)
    return None


def _selection_to_occurrence(selection_input, root_component):
    if selection_input.selectionCount < 1:
        return None

    entity = selection_input.selection(0).entity
    occurrence = adsk.fusion.Occurrence.cast(entity)
    if occurrence:
        return occurrence

    component = adsk.fusion.Component.cast(entity)
    if component:
        return _first_occurrence_for_component(root_component, component)

    return None


def _get_attribute_value(entity, key):
    if not entity:
        return None
    attr = entity.attributes.itemByName(ATTRIBUTE_GROUP, key)
    return attr.value if attr else None


def _tagged_role_for_occurrence(occurrence):
    role = _get_attribute_value(occurrence, ATTR_ROLE)
    if not role:
        role = _get_attribute_value(occurrence.component, ATTR_ROLE)
    return role.lower() if role else None


def _tagged_pair_id_for_occurrence(occurrence):
    pair_id = _get_attribute_value(occurrence, ATTR_PAIR_ID)
    if not pair_id:
        pair_id = _get_attribute_value(occurrence.component, ATTR_PAIR_ID)
    return pair_id


def _find_tagged_pulley_occurrences(root_component):
    pair_map = {}
    fallback_drive = None
    fallback_driven = None

    all_occurrences = root_component.allOccurrences
    for i in range(all_occurrences.count):
        occurrence = all_occurrences.item(i)
        role = _tagged_role_for_occurrence(occurrence)
        if role not in ['drive', 'driven']:
            continue

        if role == 'drive' and fallback_drive is None:
            fallback_drive = occurrence
        if role == 'driven' and fallback_driven is None:
            fallback_driven = occurrence

        pair_id = _tagged_pair_id_for_occurrence(occurrence)
        if pair_id:
            pair_map.setdefault(pair_id, {})[role] = occurrence

    for pair in pair_map.values():
        if 'drive' in pair and 'driven' in pair:
            return pair['drive'], pair['driven']

    return fallback_drive, fallback_driven


def _find_named_pulley_occurrences(root_component):
    drive_occ = None
    driven_occ = None

    all_occurrences = root_component.allOccurrences
    for i in range(all_occurrences.count):
        occurrence = all_occurrences.item(i)
        comp_name = occurrence.component.name.lower()
        if 'pulley' not in comp_name:
            continue

        if drive_occ is None and ('drive' in comp_name) and ('driven' not in comp_name):
            drive_occ = occurrence

        if driven_occ is None and ('driven' in comp_name):
            driven_occ = occurrence

        if drive_occ and driven_occ:
            return drive_occ, driven_occ

    return drive_occ, driven_occ


def _resolve_occurrences(inputs, root_component):
    drive_selection = adsk.core.SelectionCommandInput.cast(inputs.itemById('driveOccurrence'))
    driven_selection = adsk.core.SelectionCommandInput.cast(inputs.itemById('drivenOccurrence'))

    drive_occ = _selection_to_occurrence(drive_selection, root_component)
    driven_occ = _selection_to_occurrence(driven_selection, root_component)
    if drive_occ and driven_occ:
        return drive_occ, driven_occ, 'explicitly selected pulley occurrences'

    tagged_drive = None
    tagged_driven = None
    if (not drive_occ) or (not driven_occ):
        tagged_drive, tagged_driven = _find_tagged_pulley_occurrences(root_component)
        if not drive_occ:
            drive_occ = tagged_drive
        if not driven_occ:
            driven_occ = tagged_driven

    if drive_occ and driven_occ:
        if drive_selection.selectionCount > 0 or driven_selection.selectionCount > 0:
            return drive_occ, driven_occ, 'selected + attribute-tagged pulley occurrences'
        return drive_occ, driven_occ, 'attribute-tagged pulley occurrences'

    if (not drive_occ) or (not driven_occ):
        named_drive, named_driven = _find_named_pulley_occurrences(root_component)
        if not drive_occ:
            drive_occ = named_drive
        if not driven_occ:
            driven_occ = named_driven

    if drive_occ and driven_occ:
        if tagged_drive or tagged_driven:
            return drive_occ, driven_occ, 'attribute/name-detected pulley occurrences'
        return drive_occ, driven_occ, 'name-detected pulley occurrences'

    return drive_occ, driven_occ, 'unresolved'

def _compute_belt_path(c1_xy, c2_xy, r1, r2):
    center_distance = _distance_2d(c1_xy, c2_xy)
    if center_distance <= 0:
        return None

    radius_delta_ratio = (r2 - r1) / center_distance
    if radius_delta_ratio <= -1.0 or radius_delta_ratio >= 1.0:
        return None

    theta = math.atan2(c2_xy[1] - c1_xy[1], c2_xy[0] - c1_xy[0])
    beta = math.acos(radius_delta_ratio)

    angle_1_upper = theta + beta
    angle_1_lower = theta - beta
    angle_2_upper = theta + beta
    angle_2_lower = theta - beta

    p1_upper = _point_from_angle(c1_xy, r1, angle_1_upper)
    p1_lower = _point_from_angle(c1_xy, r1, angle_1_lower)
    p2_upper = _point_from_angle(c2_xy, r2, angle_2_upper)
    p2_lower = _point_from_angle(c2_xy, r2, angle_2_lower)

    upper_length = _distance_2d(p1_upper, p2_upper)
    lower_length = _distance_2d(p2_lower, p1_lower)

    arc1_delta = 2.0 * beta
    arc2_delta = (2.0 * math.pi) - arc1_delta

    arc2_length = r2 * arc2_delta
    arc1_length = r1 * arc1_delta

    total_length = upper_length + arc2_length + lower_length + arc1_length

    return {
        'center_1': c1_xy,
        'center_2': c2_xy,
        'radius_1': r1,
        'radius_2': r2,
        'angle_1_lower': angle_1_lower,
        'angle_1_upper': angle_1_upper,
        'angle_2_upper': angle_2_upper,
        'p1_upper': p1_upper,
        'p2_upper': p2_upper,
        'p2_lower': p2_lower,
        'p1_lower': p1_lower,
        'upper_length': upper_length,
        'arc2_length': arc2_length,
        'lower_length': lower_length,
        'arc1_length': arc1_length,
        'arc1_delta': arc1_delta,
        'arc2_delta': arc2_delta,
        'total_length': total_length,
        'center_distance': center_distance,
    }


def _path_frame_at(path_data, s):
    segment_1_end = path_data['upper_length']
    segment_2_end = segment_1_end + path_data['arc2_length']
    segment_3_end = segment_2_end + path_data['lower_length']

    if s < segment_1_end:
        if path_data['upper_length'] <= 1e-9:
            point_xy = path_data['p1_upper']
        else:
            t = s / path_data['upper_length']
            point_xy = _lerp(path_data['p1_upper'], path_data['p2_upper'], t)

        tangent = _normalize_2d(
            path_data['p2_upper'][0] - path_data['p1_upper'][0],
            path_data['p2_upper'][1] - path_data['p1_upper'][1],
        )

        v1 = _normalize_2d(
            path_data['center_1'][0] - point_xy[0],
            path_data['center_1'][1] - point_xy[1],
        )
        v2 = _normalize_2d(
            path_data['center_2'][0] - point_xy[0],
            path_data['center_2'][1] - point_xy[1],
        )
        inward = _normalize_2d(v1[0] + v2[0], v1[1] + v2[1])

    elif s < segment_2_end:
        arc_s = s - segment_1_end
        angle = path_data['angle_2_upper'] + (arc_s / path_data['radius_2'])
        point_xy = _point_from_angle(path_data['center_2'], path_data['radius_2'], angle)
        tangent = _normalize_2d(-math.sin(angle), math.cos(angle))
        inward = _normalize_2d(
            path_data['center_2'][0] - point_xy[0],
            path_data['center_2'][1] - point_xy[1],
        )

    elif s < segment_3_end:
        line_s = s - segment_2_end
        if path_data['lower_length'] <= 1e-9:
            point_xy = path_data['p2_lower']
        else:
            t = line_s / path_data['lower_length']
            point_xy = _lerp(path_data['p2_lower'], path_data['p1_lower'], t)

        tangent = _normalize_2d(
            path_data['p1_lower'][0] - path_data['p2_lower'][0],
            path_data['p1_lower'][1] - path_data['p2_lower'][1],
        )

        v1 = _normalize_2d(
            path_data['center_1'][0] - point_xy[0],
            path_data['center_1'][1] - point_xy[1],
        )
        v2 = _normalize_2d(
            path_data['center_2'][0] - point_xy[0],
            path_data['center_2'][1] - point_xy[1],
        )
        inward = _normalize_2d(v1[0] + v2[0], v1[1] + v2[1])

    else:
        arc_s = s - segment_3_end
        angle = path_data['angle_1_lower'] + (arc_s / path_data['radius_1'])
        point_xy = _point_from_angle(path_data['center_1'], path_data['radius_1'], angle)
        tangent = _normalize_2d(-math.sin(angle), math.cos(angle))
        inward = _normalize_2d(
            path_data['center_1'][0] - point_xy[0],
            path_data['center_1'][1] - point_xy[1],
        )

    if inward == (0.0, 0.0):
        inward = _normalize_2d(-tangent[1], tangent[0])

    return {
        'point': point_xy,
        'tangent': tangent,
        'inward': inward,
    }


def _sample_belt_frames(path_data, tooth_count):
    frames = []
    step = path_data['total_length'] / float(tooth_count)

    for i in range(tooth_count):
        s = i * step
        frames.append(_path_frame_at(path_data, s))

    return frames, step


def _create_reference_sketch(component, path_data):
    sketch = component.sketches.add(component.xYConstructionPlane)

    center_1 = adsk.core.Point3D.create(path_data['center_1'][0], path_data['center_1'][1], 0)
    center_2 = adsk.core.Point3D.create(path_data['center_2'][0], path_data['center_2'][1], 0)

    curves = sketch.sketchCurves
    center_line = curves.sketchLines.addByTwoPoints(center_1, center_2)
    center_line.isConstruction = True

    circle_1 = curves.sketchCircles.addByCenterRadius(center_1, path_data['radius_1'])
    circle_1.isConstruction = True

    circle_2 = curves.sketchCircles.addByCenterRadius(center_2, path_data['radius_2'])
    circle_2.isConstruction = True

    upper_line = curves.sketchLines.addByTwoPoints(
        adsk.core.Point3D.create(path_data['p1_upper'][0], path_data['p1_upper'][1], 0),
        adsk.core.Point3D.create(path_data['p2_upper'][0], path_data['p2_upper'][1], 0),
    )
    upper_line.isConstruction = True

    lower_line = curves.sketchLines.addByTwoPoints(
        adsk.core.Point3D.create(path_data['p2_lower'][0], path_data['p2_lower'][1], 0),
        adsk.core.Point3D.create(path_data['p1_lower'][0], path_data['p1_lower'][1], 0),
    )
    lower_line.isConstruction = True


def _select_profile_by_target_area(sketch, target_area):
    best_profile = None
    best_error = None

    for i in range(sketch.profiles.count):
        profile = sketch.profiles.item(i)
        area = profile.areaProperties().area
        if area <= 0:
            continue

        error = abs(area - target_area)
        if (best_error is None) or (error < best_error):
            best_error = error
            best_profile = profile

    return best_profile


def _create_belt_base_body(component, path_data, belt_width, inner_offset, outer_offset, sample_count):
    samples = max(80, int(sample_count))
    outer_points = []
    inner_points = []

    step = path_data['total_length'] / float(samples)
    for i in range(samples):
        frame = _path_frame_at(path_data, i * step)
        p = frame['point']
        n = frame['inward']

        inner_points.append((p[0] + (n[0] * inner_offset), p[1] + (n[1] * inner_offset)))
        outer_points.append((p[0] - (n[0] * outer_offset), p[1] - (n[1] * outer_offset)))

    sketch = component.sketches.add(component.xYConstructionPlane)
    lines = sketch.sketchCurves.sketchLines

    for i in range(samples):
        lines.addByTwoPoints(
            _point3d_xy(outer_points[i]),
            _point3d_xy(outer_points[(i + 1) % samples]),
        )

    for i in range(samples):
        a = inner_points[(samples - i) % samples]
        b = inner_points[(samples - i - 1) % samples]
        lines.addByTwoPoints(_point3d_xy(a), _point3d_xy(b))

    target_area = abs(_polygon_area_2d(outer_points)) - abs(_polygon_area_2d(inner_points))
    if target_area <= 0:
        raise RuntimeError('Computed belt area is invalid for the current geometry.')

    belt_profile = _select_profile_by_target_area(sketch, target_area)
    if not belt_profile:
        raise RuntimeError('Could not determine timing belt base profile.')

    extrudes = component.features.extrudeFeatures
    belt_extrude_input = extrudes.createInput(
        belt_profile,
        adsk.fusion.FeatureOperations.NewBodyFeatureOperation,
    )
    belt_extrude_input.setDistanceExtent(False, adsk.core.ValueInput.createByReal(belt_width))
    belt_extrude = extrudes.add(belt_extrude_input)

    if belt_extrude.bodies.count > 0:
        belt_extrude.bodies.item(0).name = 'TimingBelt_Base'

    return belt_extrude


def _create_belt_teeth(component, belt_frames, tooth_pitch, tooth_height, belt_width, inner_offset):
    if not belt_frames:
        raise RuntimeError('No belt tooth frames were generated.')

    base_half = max(tooth_pitch * 0.26, tooth_height * 0.25)
    tip_half = max(tooth_pitch * 0.16, tooth_height * 0.14)
    tooth_depth = tooth_height * 0.78
    root_offset = max(inner_offset * 0.72, tooth_height * 0.05)

    tooth_sketch = component.sketches.add(component.xYConstructionPlane)
    lines = tooth_sketch.sketchCurves.sketchLines

    for frame in belt_frames:
        p = frame['point']
        t = frame['tangent']
        n = frame['inward']

        base_center = (p[0] + (n[0] * root_offset), p[1] + (n[1] * root_offset))
        tip_center = (
            base_center[0] + (n[0] * tooth_depth),
            base_center[1] + (n[1] * tooth_depth),
        )

        p1 = (base_center[0] + (t[0] * base_half), base_center[1] + (t[1] * base_half))
        p2 = (tip_center[0] + (t[0] * tip_half), tip_center[1] + (t[1] * tip_half))
        p3 = (tip_center[0] - (t[0] * tip_half), tip_center[1] - (t[1] * tip_half))
        p4 = (base_center[0] - (t[0] * base_half), base_center[1] - (t[1] * base_half))

        lines.addByTwoPoints(_point3d_xy(p1), _point3d_xy(p2))
        lines.addByTwoPoints(_point3d_xy(p2), _point3d_xy(p3))
        lines.addByTwoPoints(_point3d_xy(p3), _point3d_xy(p4))
        lines.addByTwoPoints(_point3d_xy(p4), _point3d_xy(p1))

    expected_area = (base_half + tip_half) * tooth_depth
    min_area = expected_area * 0.25
    max_area = expected_area * 4.0

    tooth_profiles = []
    for i in range(tooth_sketch.profiles.count):
        profile = tooth_sketch.profiles.item(i)
        area = profile.areaProperties().area
        if min_area <= area <= max_area:
            tooth_profiles.append(profile)

    if not tooth_profiles:
        raise RuntimeError('Could not determine timing belt tooth profiles.')

    extrudes = component.features.extrudeFeatures
    for profile in tooth_profiles:
        tooth_extrude_input = extrudes.createInput(
            profile,
            adsk.fusion.FeatureOperations.JoinFeatureOperation,
        )
        tooth_extrude_input.setDistanceExtent(False, adsk.core.ValueInput.createByReal(belt_width))
        extrudes.add(tooth_extrude_input)

def _validate_inputs(
    drive_teeth,
    driven_teeth,
    belt_pitch,
    roller_diameter,
    belt_width,
    use_selected,
    manual_center_distance,
    auto_link_count,
    requested_link_count,
):
    errors = []

    if drive_teeth < MIN_SPROCKET_TEETH or driven_teeth < MIN_SPROCKET_TEETH:
        errors.append('Both tooth counts must be at least {}.'.format(MIN_SPROCKET_TEETH))

    if belt_pitch <= 0 or roller_diameter <= 0 or belt_width <= 0:
        errors.append('Belt pitch, tooth height, and belt width must be positive.')

    if roller_diameter >= belt_pitch:
        errors.append('Tooth height must be smaller than belt pitch.')

    if (not use_selected) and (manual_center_distance <= 0):
        errors.append('Manual center distance must be positive when selection mode is off.')

    if (not auto_link_count) and (requested_link_count < 10):
        errors.append('Manual belt tooth count must be at least 10.')

    return errors


def _center_distance_warnings(center_distance, belt_pitch):
    warnings = []

    if belt_pitch <= 0:
        return warnings

    center_in_pitches = center_distance / belt_pitch
    if center_in_pitches < RECOMMENDED_CENTER_MIN_PITCHES or center_in_pitches > RECOMMENDED_CENTER_MAX_PITCHES:
        warnings.append(
            (
                'Center distance is {:.2f} pitches. Recommended range is '
                '{:.0f}-{:.0f} pitches (ISO 606 / ANSI B29.1 common practice).'
            ).format(center_in_pitches, RECOMMENDED_CENTER_MIN_PITCHES, RECOMMENDED_CENTER_MAX_PITCHES)
        )

    return warnings


def _determine_link_count(path_data, belt_pitch, auto_link_count, requested_link_count, enforce_even_links):
    raw_link_count = max(10, int(round(path_data['total_length'] / belt_pitch))) if auto_link_count else requested_link_count
    final_link_count = raw_link_count
    even_adjusted = False

    if enforce_even_links and (final_link_count % 2 != 0):
        final_link_count += 1
        even_adjusted = True

    return raw_link_count, final_link_count, even_adjusted


def _half_link_note(link_count, drive_teeth, driven_teeth):
    if link_count % 2 == 0:
        return None
    if drive_teeth == driven_teeth:
        return 'Odd belt tooth count with equal pulleys can require phase indexing during installation.'
    return 'Odd belt tooth count can require phase indexing during installation.'


def _format_issues(issues):
    return '' if not issues else '\n- ' + '\n- '.join(issues)


def _write_csv_rows(path, rows):
    with open(path, 'w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(['Field', 'Value'])
        for key, value in rows:
            writer.writerow([key, value])


def _export_csv_dialog(ui, filename_seed, rows):
    file_dialog = ui.createFileDialog()
    file_dialog.title = 'Export Timing Belt Drive CSV'
    file_dialog.filter = 'CSV Files (*.csv)'
    file_dialog.filterIndex = 0
    file_dialog.initialFilename = '{}_{}.csv'.format(
        filename_seed,
        datetime.datetime.now().strftime('%Y%m%d_%H%M%S'),
    )

    if file_dialog.showSave() != adsk.core.DialogResults.DialogOK:
        return None

    _write_csv_rows(file_dialog.filename, rows)
    return file_dialog.filename


def _set_input_state(inputs):
    use_selected = inputs.itemById('useSelectedPulleys').value
    auto_links = inputs.itemById('autoLinkCount').value

    adsk.core.SelectionCommandInput.cast(inputs.itemById('driveOccurrence')).isEnabled = use_selected
    adsk.core.SelectionCommandInput.cast(inputs.itemById('drivenOccurrence')).isEnabled = use_selected
    adsk.core.ValueCommandInput.cast(inputs.itemById('manualCenterDistance')).isEnabled = not use_selected
    adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById('linkCount')).isEnabled = not auto_links


def _build_preview_text(inputs):
    try:
        drive_teeth = inputs.itemById('driveToothCount').value
        driven_teeth = inputs.itemById('drivenToothCount').value
        belt_pitch = inputs.itemById('beltPitch').value
        roller_diameter = inputs.itemById('rollerDiameter').value
        belt_width = inputs.itemById('beltWidth').value
        use_selected = inputs.itemById('useSelectedPulleys').value
        manual_center_distance = inputs.itemById('manualCenterDistance').value
        auto_link_count = inputs.itemById('autoLinkCount').value
        requested_link_count = inputs.itemById('linkCount').value
        enforce_even_links = inputs.itemById('enforceEvenLinks').value

        errors = _validate_inputs(
            drive_teeth,
            driven_teeth,
            belt_pitch,
            roller_diameter,
            belt_width,
            use_selected,
            manual_center_distance,
            auto_link_count,
            requested_link_count,
        )
        if errors:
            return 'Input issues:{}\n\nFix values to preview derived belt data.'.format(_format_issues(errors))

        pitch_radius_1 = _pitch_radius(belt_pitch, drive_teeth)
        pitch_radius_2 = _pitch_radius(belt_pitch, driven_teeth)

        center_distance = None
        center_source = ''
        if use_selected:
            design = adsk.fusion.Design.cast(adsk.core.Application.get().activeProduct)
            if design:
                drive_occ, driven_occ, center_source = _resolve_occurrences(inputs, design.rootComponent)
                if drive_occ and driven_occ:
                    c1 = _get_occurrence_center(drive_occ)
                    c2 = _get_occurrence_center(driven_occ)
                    center_distance = _distance_2d((c1[0], c1[1]), (c2[0], c2[1]))
            if center_distance is None:
                return (
                    'Live summary\n\n'
                    'Center source unresolved. Select both pulleys or run tagged pulley generation first.'
                )
        else:
            center_distance = manual_center_distance
            center_source = 'manual center distance input'

        lines = [
            'Live summary',
            '',
            'Center source: {}'.format(center_source),
            'Center distance: {:.3f} mm ({:.2f} pitches)'.format(center_distance * 10.0, center_distance / belt_pitch),
        ]

        if center_distance <= (pitch_radius_1 + pitch_radius_2):
            lines.append('Warning: pulley centers are too close for valid wrap.')
            return '\n'.join(lines)

        path_data = _compute_belt_path((0.0, 0.0), (center_distance, 0.0), pitch_radius_1, pitch_radius_2)
        if not path_data:
            lines.append('Warning: could not compute valid belt tangency.')
            return '\n'.join(lines)

        raw_count, final_count, even_adjusted = _determine_link_count(
            path_data,
            belt_pitch,
            auto_link_count,
            requested_link_count,
            enforce_even_links,
        )
        lines.append('Estimated loop length: {:.3f} mm'.format(path_data['total_length'] * 10.0))
        lines.append('Preview belt tooth count: {}'.format(final_count))

        if even_adjusted:
            lines.append('Note: rounded from {} to {} to keep an even tooth count.'.format(raw_count, final_count))

        half_link = _half_link_note(final_count, drive_teeth, driven_teeth)
        if half_link:
            lines.append('Warning: {}'.format(half_link))

        for warning in _center_distance_warnings(center_distance, belt_pitch):
            lines.append('Warning: {}'.format(warning))

        return '\n'.join(lines)
    except Exception:
        return 'Live summary unavailable for current inputs.'


def _update_preview_text(inputs):
    preview = adsk.core.TextBoxCommandInput.cast(inputs.itemById('previewInfo'))
    if preview:
        preview.text = _build_preview_text(inputs)

class CommandCreatedHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        app = adsk.core.Application.get()
        ui = app.userInterface

        try:
            command = adsk.core.Command.cast(args.command)
            inputs = command.commandInputs

            inputs.addIntegerSpinnerCommandInput('driveToothCount', 'Drive Tooth Count', MIN_SPROCKET_TEETH, 240, 1, 24)
            inputs.addIntegerSpinnerCommandInput('drivenToothCount', 'Driven Tooth Count', MIN_SPROCKET_TEETH, 240, 1, 48)

            inputs.addValueInput('beltPitch', 'Belt Pitch', 'mm', adsk.core.ValueInput.createByString('12.7 mm'))
            inputs.addValueInput('rollerDiameter', 'Tooth Height', 'mm', adsk.core.ValueInput.createByString('7.9 mm'))
            inputs.addValueInput('beltWidth', 'Belt Width', 'mm', adsk.core.ValueInput.createByString('6 mm'))

            inputs.addBoolValueInput('useSelectedPulleys', 'Use Selected Pulley Centers', True, '', True)

            drive_occurrence_input = inputs.addSelectionInput(
                'driveOccurrence',
                'Drive Pulley Occurrence',
                'Select the drive pulley occurrence (optional when tagged pulleys exist).',
            )
            drive_occurrence_input.addSelectionFilter('Occurrences')
            drive_occurrence_input.setSelectionLimits(0, 1)

            driven_occurrence_input = inputs.addSelectionInput(
                'drivenOccurrence',
                'Driven Pulley Occurrence',
                'Select the driven pulley occurrence (optional when tagged pulleys exist).',
            )
            driven_occurrence_input.addSelectionFilter('Occurrences')
            driven_occurrence_input.setSelectionLimits(0, 1)

            inputs.addValueInput('manualCenterDistance', 'Manual Center Distance', 'mm', adsk.core.ValueInput.createByString('150 mm'))

            inputs.addBoolValueInput('autoLinkCount', 'Auto Belt Tooth Count', True, '', True)
            inputs.addIntegerSpinnerCommandInput('linkCount', 'Manual Belt Tooth Count', 10, 6000, 1, 120)
            inputs.addBoolValueInput('enforceEvenLinks', 'Force Even Tooth Count', True, '', True)
            inputs.addBoolValueInput('exportCsv', 'Export CSV Summary', True, '', False)

            inputs.addTextBoxCommandInput(
                'previewInfo',
                'Live Summary',
                'Adjust values to preview center distance, belt tooth count, and warnings.',
                8,
                True,
            )

            _set_input_state(inputs)
            _update_preview_text(inputs)

            on_input_changed = CommandInputChangedHandler()
            command.inputChanged.add(on_input_changed)
            handlers.append(on_input_changed)

            on_execute = CommandExecuteHandler()
            command.execute.add(on_execute)
            handlers.append(on_execute)

            on_destroy = CommandDestroyHandler()
            command.destroy.add(on_destroy)
            handlers.append(on_destroy)

        except Exception:
            if ui:
                ui.messageBox('Command creation failed:\n{}'.format(traceback.format_exc()))


class CommandInputChangedHandler(adsk.core.InputChangedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        app = adsk.core.Application.get()
        ui = app.userInterface

        try:
            event_args = adsk.core.InputChangedEventArgs.cast(args)
            changed_input = event_args.input
            if not changed_input:
                return

            command = adsk.core.Command.cast(event_args.firingEvent.sender)
            if changed_input.id in ['useSelectedPulleys', 'autoLinkCount']:
                _set_input_state(command.commandInputs)

            if changed_input.id != 'previewInfo':
                _update_preview_text(command.commandInputs)

        except Exception:
            if ui:
                ui.messageBox('Input change handling failed:\n{}'.format(traceback.format_exc()))


class CommandExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        app = adsk.core.Application.get()
        ui = app.userInterface

        try:
            design = adsk.fusion.Design.cast(app.activeProduct)
            if not design:
                ui.messageBox('An active Fusion 360 design is required.')
                return

            command = adsk.core.Command.cast(args.firingEvent.sender)
            inputs = command.commandInputs

            drive_teeth = inputs.itemById('driveToothCount').value
            driven_teeth = inputs.itemById('drivenToothCount').value

            belt_pitch = inputs.itemById('beltPitch').value
            roller_diameter = inputs.itemById('rollerDiameter').value
            belt_width = inputs.itemById('beltWidth').value

            use_selected = inputs.itemById('useSelectedPulleys').value
            manual_center_distance = inputs.itemById('manualCenterDistance').value

            auto_link_count = inputs.itemById('autoLinkCount').value
            requested_link_count = inputs.itemById('linkCount').value
            enforce_even_links = inputs.itemById('enforceEvenLinks').value
            export_csv = inputs.itemById('exportCsv').value

            input_errors = _validate_inputs(
                drive_teeth,
                driven_teeth,
                belt_pitch,
                roller_diameter,
                belt_width,
                use_selected,
                manual_center_distance,
                auto_link_count,
                requested_link_count,
            )
            if input_errors:
                ui.messageBox('Input validation failed:{}\n\nFix the values and run again.'.format(_format_issues(input_errors)))
                return

            pitch_radius_1 = _pitch_radius(belt_pitch, drive_teeth)
            pitch_radius_2 = _pitch_radius(belt_pitch, driven_teeth)

            root_component = design.rootComponent
            center_source = 'manual center distance input'

            if use_selected:
                drive_occ, driven_occ, center_source = _resolve_occurrences(inputs, root_component)
                if (not drive_occ) or (not driven_occ):
                    ui.messageBox(
                        'Could not find drive/driven pulley occurrences. '
                        'Select both occurrences, or generate tagged pulleys first, or disable "Use Selected Pulley Centers".'
                    )
                    return
                if drive_occ.entityToken == driven_occ.entityToken:
                    ui.messageBox('Drive and driven selections must be different occurrences.')
                    return
                center_1 = _get_occurrence_center(drive_occ)
                center_2 = _get_occurrence_center(driven_occ)
            else:
                center_1 = (0.0, 0.0, 0.0)
                center_2 = (manual_center_distance, 0.0, 0.0)

            center_1_xy = (center_1[0], center_1[1])
            center_2_xy = (center_2[0], center_2[1])
            center_distance = _distance_2d(center_1_xy, center_2_xy)

            if center_distance <= (pitch_radius_1 + pitch_radius_2):
                ui.messageBox(
                    'Pulley centers are too close for belt wrap. '
                    'Increase center distance or reduce pulley sizes.'
                )
                return

            path_data = _compute_belt_path(center_1_xy, center_2_xy, pitch_radius_1, pitch_radius_2)
            if not path_data:
                ui.messageBox('Could not compute valid tangency for the provided pulley geometry.')
                return

            raw_link_count, link_count, even_adjusted = _determine_link_count(
                path_data,
                belt_pitch,
                auto_link_count,
                requested_link_count,
                enforce_even_links,
            )
            if link_count < 10:
                ui.messageBox('Belt tooth count is too small to form a stable belt loop.')
                return

            belt_frames, actual_pitch = _sample_belt_frames(path_data, link_count)
            if len(belt_frames) < 4:
                ui.messageBox('Failed to generate sufficient belt tooth positions.')
                return

            engineering_warnings = _center_distance_warnings(center_distance, belt_pitch)

            drive_wrap_deg = math.degrees(path_data['arc1_delta'])
            driven_wrap_deg = math.degrees(path_data['arc2_delta'])
            if drive_wrap_deg < 120.0:
                engineering_warnings.append(
                    'Drive pulley wrap angle is {:.2f} deg; low wrap can reduce tooth engagement under load.'.format(
                        drive_wrap_deg
                    )
                )
            if driven_wrap_deg < 120.0:
                engineering_warnings.append(
                    'Driven pulley wrap angle is {:.2f} deg; low wrap can reduce tooth engagement under load.'.format(
                        driven_wrap_deg
                    )
                )

            if even_adjusted:
                engineering_warnings.append(
                    'Belt tooth count adjusted from {} to {} to keep an even count.'.format(
                        raw_link_count,
                        link_count,
                    )
                )

            half_link_note = _half_link_note(link_count, drive_teeth, driven_teeth)
            if half_link_note:
                engineering_warnings.append(half_link_note)

            plane_z = (center_1[2] + center_2[2]) * 0.5
            z_mismatch = abs(center_1[2] - center_2[2])

            belt_transform = adsk.core.Matrix3D.create()
            belt_transform.translation = adsk.core.Vector3D.create(0, 0, plane_z)
            belt_occurrence = root_component.occurrences.addNewComponent(belt_transform)
            belt_component = belt_occurrence.component
            belt_component.name = 'Timing Belt Drive {}T-{}T'.format(drive_teeth, driven_teeth)

            belt_tooth_height = roller_diameter
            belt_inner_offset = max(0.08 * belt_tooth_height, 0.03 * belt_pitch)
            belt_backing_thickness = max(0.40 * belt_tooth_height, 0.18 * belt_pitch)
            belt_outer_offset = belt_backing_thickness + (0.16 * belt_tooth_height)
            belt_profile_samples = max(180, link_count * 4)

            _create_reference_sketch(belt_component, path_data)
            _create_belt_base_body(
                belt_component,
                path_data,
                belt_width,
                belt_inner_offset,
                belt_outer_offset,
                belt_profile_samples,
            )
            _create_belt_teeth(
                belt_component,
                belt_frames,
                actual_pitch,
                belt_tooth_height,
                belt_width,
                belt_inner_offset,
            )

            center_mm = center_distance * 10.0
            center_pitches = center_distance / belt_pitch
            requested_pitch_mm = belt_pitch * 10.0
            actual_pitch_mm = actual_pitch * 10.0
            pitch_error_pct = abs(actual_pitch - belt_pitch) / belt_pitch * 100.0
            half_link_required = (link_count % 2 != 0)

            csv_export_note = ''
            if export_csv:
                rows = [
                    ('GeneratedAt', datetime.datetime.now().isoformat(timespec='seconds')),
                    ('DriveTeeth', str(drive_teeth)),
                    ('DrivenTeeth', str(driven_teeth)),
                    ('BeltPitch_mm', '{:.4f}'.format(requested_pitch_mm)),
                    ('ToothHeight_mm', '{:.4f}'.format(roller_diameter * 10.0)),
                    ('BeltWidth_mm', '{:.4f}'.format(belt_width * 10.0)),
                    ('CenterDistance_mm', '{:.4f}'.format(center_mm)),
                    ('CenterDistance_pitches', '{:.6f}'.format(center_pitches)),
                    ('CenterSource', center_source),
                    ('AutoBeltToothCount', str(auto_link_count)),
                    ('RawBeltToothCount', str(raw_link_count)),
                    ('FinalBeltToothCount', str(link_count)),
                    ('EvenToothAdjusted', str(even_adjusted)),
                    ('OddToothPhaseIndexingNote', str(half_link_required)),
                    ('EffectivePitch_mm', '{:.4f}'.format(actual_pitch_mm)),
                    ('PitchDeviation_pct', '{:.6f}'.format(pitch_error_pct)),
                    ('DriveWrap_deg', '{:.4f}'.format(drive_wrap_deg)),
                    ('DrivenWrap_deg', '{:.4f}'.format(driven_wrap_deg)),
                    ('ZMismatch_mm', '{:.4f}'.format(z_mismatch * 10.0)),
                    ('BeltInnerOffset_mm', '{:.4f}'.format(belt_inner_offset * 10.0)),
                    ('BeltBackingThickness_mm', '{:.4f}'.format(belt_backing_thickness * 10.0)),
                    ('BeltOuterOffset_mm', '{:.4f}'.format(belt_outer_offset * 10.0)),
                    ('EngineeringWarnings', ' | '.join(engineering_warnings) if engineering_warnings else ''),
                ]

                export_path = _export_csv_dialog(ui, 'belt_drive', rows)
                if export_path:
                    csv_export_note = '\nCSV summary: {}'.format(export_path)
                else:
                    csv_export_note = '\nCSV summary: skipped by user.'

            warning_block = ''
            if engineering_warnings:
                warning_block = '\n\nEngineering warnings:{}'.format(_format_issues(engineering_warnings))

            ui.messageBox(
                'Created timing belt drive.\n\n'
                'Drive teeth: {}\n'
                'Driven teeth: {}\n'
                'Center distance: {:.3f} mm ({:.3f} pitches)\n'
                'Belt tooth count: {}\n'
                'Odd-tooth phase indexing note: {}\n'
                'Requested pitch: {:.3f} mm\n'
                'Effective pitch: {:.3f} mm\n'
                'Pitch deviation: {:.3f}%\n'
                'Wrap angle (drive / driven): {:.2f} / {:.2f} deg\n'
                'Center source: {}\n'
                'Z mismatch between pulley centers: {:.3f} mm{}{}'.format(
                    drive_teeth,
                    driven_teeth,
                    center_mm,
                    center_pitches,
                    link_count,
                    'Yes' if half_link_required else 'No',
                    requested_pitch_mm,
                    actual_pitch_mm,
                    pitch_error_pct,
                    drive_wrap_deg,
                    driven_wrap_deg,
                    center_source,
                    z_mismatch * 10.0,
                    warning_block,
                    csv_export_note,
                )
            )

        except Exception:
            if ui:
                ui.messageBox('Timing belt drive generation failed:\n{}'.format(traceback.format_exc()))


class CommandDestroyHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        pass
