import adsk.core
import adsk.fusion
import traceback
import math
import csv
import datetime
import uuid

APP_NAME = 'Adjustable Timing Pulley Pair'
CMD_ID = 'com.lenicolas.adjustabledrivepulley'
CMD_NAME = 'Adjustable Timing Pulley Pair'
CMD_DESC = 'Create drive and driven timing pulleys with adjustable tooth counts and belt geometry.'
WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'SolidCreatePanel'

MIN_SPROCKET_TEETH = 9
RECOMMENDED_CENTER_MIN_PITCHES = 30.0
RECOMMENDED_CENTER_MAX_PITCHES = 50.0

ATTRIBUTE_GROUP = 'com.lenicolas.pulley'
ATTR_ROLE = 'role'
ATTR_PAIR_ID = 'pair_id'
ATTR_TOOTH_COUNT = 'tooth_count'
ATTR_CHAIN_PITCH = 'belt_pitch_cm'
ATTR_TOOTH_HEIGHT = 'tooth_height_cm'

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


def _largest_profile(sketch):
    largest = None
    largest_area = -1.0

    for i in range(sketch.profiles.count):
        profile = sketch.profiles.item(i)
        area = profile.areaProperties().area
        if area > largest_area:
            largest_area = area
            largest = profile

    return largest


def _polar_point(radius, angle_rad):
    return adsk.core.Point3D.create(radius * math.cos(angle_rad), radius * math.sin(angle_rad), 0)


def _pulley_radii(tooth_count, belt_pitch, roller_diameter, tip_clearance):
    pitch_radius = belt_pitch / (2.0 * math.sin(math.pi / tooth_count))
    root_radius = pitch_radius - (0.55 * roller_diameter)
    tip_radius = pitch_radius + (0.55 * roller_diameter) + tip_clearance
    return pitch_radius, root_radius, tip_radius


def _center_distance_from_belt_links(belt_links, drive_teeth, driven_teeth, belt_pitch):
    # Standard approximate belt-length relationship:
    # L = 2m + (z1+z2)/2 + ((z2-z1)^2)/(4*pi^2*m), where m = C/p
    avg_teeth = (drive_teeth + driven_teeth) / 2.0
    tooth_delta_term = ((driven_teeth - drive_teeth) ** 2) / (4.0 * math.pi * math.pi)
    discriminant = ((belt_links - avg_teeth) ** 2) - (8.0 * tooth_delta_term)

    if discriminant < 0:
        return None

    sqrt_disc = math.sqrt(discriminant)
    m1 = ((belt_links - avg_teeth) + sqrt_disc) / 4.0
    m2 = ((belt_links - avg_teeth) - sqrt_disc) / 4.0

    m = max(m1, m2)
    if m <= 0:
        return None

    return m * belt_pitch


def _solve_tooth_counts_for_ratio(
    target_ratio,
    min_teeth,
    max_teeth,
    max_pulley_diameter,
    belt_pitch,
    roller_diameter,
    tip_clearance,
):
    best = None

    for drive_teeth in range(min_teeth, max_teeth + 1):
        for driven_teeth in range(min_teeth, max_teeth + 1):
            actual_ratio = float(driven_teeth) / float(drive_teeth)
            _, _, drive_tip_radius = _pulley_radii(drive_teeth, belt_pitch, roller_diameter, tip_clearance)
            _, _, driven_tip_radius = _pulley_radii(driven_teeth, belt_pitch, roller_diameter, tip_clearance)

            drive_diameter = 2.0 * drive_tip_radius
            driven_diameter = 2.0 * driven_tip_radius

            if drive_diameter > max_pulley_diameter or driven_diameter > max_pulley_diameter:
                continue

            ratio_error = abs(actual_ratio - target_ratio) / target_ratio

            # Prefer ratio accuracy first, then larger teeth counts for smoother running.
            score = (
                ratio_error,
                -min(drive_teeth, driven_teeth),
                -max(drive_teeth, driven_teeth),
            )

            if (best is None) or (score < best['score']):
                best = {
                    'drive_teeth': drive_teeth,
                    'driven_teeth': driven_teeth,
                    'actual_ratio': actual_ratio,
                    'ratio_error': ratio_error,
                    'drive_diameter': drive_diameter,
                    'driven_diameter': driven_diameter,
                    'score': score,
                }

    return best


def _validate_inputs(
    drive_teeth,
    driven_teeth,
    belt_pitch,
    roller_diameter,
    thickness,
    drive_bore_diameter,
    driven_bore_diameter,
    tip_clearance,
    auto_center,
    belt_links,
    manual_center_distance,
):
    errors = []

    if drive_teeth < MIN_SPROCKET_TEETH or driven_teeth < MIN_SPROCKET_TEETH:
        errors.append('Both tooth counts must be at least {}.'.format(MIN_SPROCKET_TEETH))

    if belt_pitch <= 0 or roller_diameter <= 0 or thickness <= 0:
        errors.append('Belt pitch, tooth height, and pulley thickness must be positive.')

    if roller_diameter >= belt_pitch:
        errors.append('Tooth height must be smaller than belt pitch.')

    if tip_clearance < 0:
        errors.append('Tip clearance cannot be negative.')

    if drive_bore_diameter < 0 or driven_bore_diameter < 0:
        errors.append('Bore diameters cannot be negative.')

    if auto_center:
        if belt_links < 10:
            errors.append('Belt links must be at least 10.')
    else:
        if manual_center_distance <= 0:
            errors.append('Manual center distance must be positive.')

    return errors


def _validate_ratio_solver_inputs(
    auto_tooth_by_ratio,
    target_ratio,
    max_pulley_diameter,
    min_tooth_limit,
    max_tooth_limit,
):
    errors = []

    if not auto_tooth_by_ratio:
        return errors

    if target_ratio <= 0:
        errors.append('Target ratio must be positive.')

    if max_pulley_diameter <= 0:
        errors.append('Max pulley diameter must be positive.')

    if min_tooth_limit < MIN_SPROCKET_TEETH:
        errors.append('Minimum tooth limit must be at least {}.'.format(MIN_SPROCKET_TEETH))

    if max_tooth_limit < min_tooth_limit:
        errors.append('Maximum tooth limit must be greater than or equal to the minimum tooth limit.')

    return errors


def _center_distance_warnings(center_distance, belt_pitch):
    warnings = []

    if belt_pitch <= 0:
        return warnings

    center_in_pitches = center_distance / belt_pitch

    if center_in_pitches < RECOMMENDED_CENTER_MIN_PITCHES:
        warnings.append(
            (
                'Center distance is {:.2f} pitches. Recommended range is '
                '{:.0f}-{:.0f} pitches (ISO 606 / ANSI B29.1 common practice).'
            ).format(center_in_pitches, RECOMMENDED_CENTER_MIN_PITCHES, RECOMMENDED_CENTER_MAX_PITCHES)
        )
    elif center_in_pitches > RECOMMENDED_CENTER_MAX_PITCHES:
        warnings.append(
            (
                'Center distance is {:.2f} pitches. Recommended range is '
                '{:.0f}-{:.0f} pitches (ISO 606 / ANSI B29.1 common practice).'
            ).format(center_in_pitches, RECOMMENDED_CENTER_MIN_PITCHES, RECOMMENDED_CENTER_MAX_PITCHES)
        )

    return warnings


def _set_input_state(inputs):
    auto_center = inputs.itemById('autoCenter').value
    auto_tooth_by_ratio = inputs.itemById('autoToothByRatio').value

    belt_links_input = adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById('beltLinks'))
    center_distance_input = adsk.core.ValueCommandInput.cast(inputs.itemById('centerDistance'))
    drive_teeth_input = adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById('driveToothCount'))
    driven_teeth_input = adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById('drivenToothCount'))
    target_ratio_input = adsk.core.ValueCommandInput.cast(inputs.itemById('targetRatio'))
    max_pulley_diameter_input = adsk.core.ValueCommandInput.cast(inputs.itemById('maxPulleyDiameter'))
    min_tooth_limit_input = adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById('minToothLimit'))
    max_tooth_limit_input = adsk.core.IntegerSpinnerCommandInput.cast(inputs.itemById('maxToothLimit'))

    belt_links_input.isEnabled = auto_center
    center_distance_input.isEnabled = not auto_center
    drive_teeth_input.isEnabled = not auto_tooth_by_ratio
    driven_teeth_input.isEnabled = not auto_tooth_by_ratio
    target_ratio_input.isEnabled = auto_tooth_by_ratio
    max_pulley_diameter_input.isEnabled = auto_tooth_by_ratio
    min_tooth_limit_input.isEnabled = auto_tooth_by_ratio
    max_tooth_limit_input.isEnabled = auto_tooth_by_ratio


def _create_pulley_geometry(component, tooth_count, root_radius, tip_radius, thickness, bore_diameter, body_name):
    sketches = component.sketches
    extrudes = component.features.extrudeFeatures
    patterns = component.features.circularPatternFeatures

    base_sketch = sketches.add(component.xYConstructionPlane)
    center = adsk.core.Point3D.create(0, 0, 0)
    base_sketch.sketchCurves.sketchCircles.addByCenterRadius(center, root_radius)
    if bore_diameter > 0:
        base_sketch.sketchCurves.sketchCircles.addByCenterRadius(center, bore_diameter / 2.0)

    base_profile = _largest_profile(base_sketch)
    if not base_profile:
        raise RuntimeError('Could not find a valid base profile.')

    base_extrude_input = extrudes.createInput(base_profile, adsk.fusion.FeatureOperations.NewBodyFeatureOperation)
    base_extrude_input.setDistanceExtent(False, adsk.core.ValueInput.createByReal(thickness))
    base_extrude = extrudes.add(base_extrude_input)

    tooth_sketch = sketches.add(component.xYConstructionPlane)
    pitch_angle = (2.0 * math.pi) / tooth_count
    root_half_angle = pitch_angle * 0.22
    tip_half_angle = pitch_angle * 0.12

    p1 = _polar_point(root_radius, -root_half_angle)
    p2 = _polar_point(tip_radius, -tip_half_angle)
    p3 = _polar_point(tip_radius, tip_half_angle)
    p4 = _polar_point(root_radius, tip_half_angle)

    lines = tooth_sketch.sketchCurves.sketchLines
    lines.addByTwoPoints(p1, p2)
    lines.addByTwoPoints(p2, p3)
    lines.addByTwoPoints(p3, p4)
    lines.addByTwoPoints(p4, p1)

    tooth_profile = _largest_profile(tooth_sketch)
    if not tooth_profile:
        raise RuntimeError('Could not find a valid tooth profile.')

    tooth_extrude_input = extrudes.createInput(tooth_profile, adsk.fusion.FeatureOperations.JoinFeatureOperation)
    tooth_extrude_input.setDistanceExtent(False, adsk.core.ValueInput.createByReal(thickness))
    tooth_extrude = extrudes.add(tooth_extrude_input)

    feature_collection = adsk.core.ObjectCollection.create()
    feature_collection.add(tooth_extrude)

    pattern_input = patterns.createInput(feature_collection, component.zConstructionAxis)
    pattern_input.quantity = adsk.core.ValueInput.createByString(str(tooth_count))
    pattern_input.totalAngle = adsk.core.ValueInput.createByString('360 deg')
    patterns.add(pattern_input)

    if base_extrude.bodies.count > 0:
        base_extrude.bodies.item(0).name = body_name


def _tag_pulley_entity(entity, role, pair_id, tooth_count, belt_pitch, roller_diameter):
    if not entity:
        return

    attrs = entity.attributes
    attrs.add(ATTRIBUTE_GROUP, ATTR_ROLE, role)
    attrs.add(ATTRIBUTE_GROUP, ATTR_PAIR_ID, pair_id)
    attrs.add(ATTRIBUTE_GROUP, ATTR_TOOTH_COUNT, str(tooth_count))
    attrs.add(ATTRIBUTE_GROUP, ATTR_CHAIN_PITCH, '{:.10f}'.format(belt_pitch))
    attrs.add(ATTRIBUTE_GROUP, ATTR_TOOTH_HEIGHT, '{:.10f}'.format(roller_diameter))


def _write_csv_rows(path, rows):
    with open(path, 'w', newline='') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(['Field', 'Value'])
        for row in rows:
            writer.writerow([row[0], row[1]])


def _export_csv_dialog(ui, filename_seed, rows):
    file_dialog = ui.createFileDialog()
    file_dialog.title = 'Export Timing Pulley Pair CSV'
    file_dialog.filter = 'CSV Files (*.csv)'
    file_dialog.filterIndex = 0

    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    file_dialog.initialFilename = '{}_{}.csv'.format(filename_seed, timestamp)

    dialog_result = file_dialog.showSave()
    if dialog_result != adsk.core.DialogResults.DialogOK:
        return None

    _write_csv_rows(file_dialog.filename, rows)
    return file_dialog.filename


def _format_issues(issues):
    if not issues:
        return ''
    return '\n- ' + '\n- '.join(issues)


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
            inputs.addBoolValueInput('autoToothByRatio', 'Auto Tooth Count From Ratio', True, '', False)
            inputs.addValueInput('targetRatio', 'Target Ratio (driven:drive)', '', adsk.core.ValueInput.createByReal(6.5))
            inputs.addValueInput('maxPulleyDiameter', 'Max Pulley Diameter', 'mm', adsk.core.ValueInput.createByString('120 mm'))
            inputs.addIntegerSpinnerCommandInput('minToothLimit', 'Min Tooth Limit', MIN_SPROCKET_TEETH, 240, 1, 12)
            inputs.addIntegerSpinnerCommandInput('maxToothLimit', 'Max Tooth Limit', MIN_SPROCKET_TEETH, 240, 1, 80)

            inputs.addValueInput('beltPitch', 'Belt Pitch', 'mm', adsk.core.ValueInput.createByString('12.7 mm'))
            inputs.addValueInput('rollerDiameter', 'Tooth Height', 'mm', adsk.core.ValueInput.createByString('7.9 mm'))
            inputs.addValueInput('thickness', 'Pulley Thickness', 'mm', adsk.core.ValueInput.createByString('6 mm'))
            inputs.addValueInput('driveBoreDiameter', 'Drive Bore Diameter', 'mm', adsk.core.ValueInput.createByString('8 mm'))
            inputs.addValueInput('drivenBoreDiameter', 'Driven Bore Diameter', 'mm', adsk.core.ValueInput.createByString('8 mm'))
            inputs.addValueInput('tipClearance', 'Tip Clearance', 'mm', adsk.core.ValueInput.createByString('1.5 mm'))

            inputs.addBoolValueInput('autoCenter', 'Auto Center Distance From Belt Links', True, '', True)
            inputs.addIntegerSpinnerCommandInput('beltLinks', 'Belt Links (pitches)', 20, 2000, 1, 120)
            inputs.addValueInput('centerDistance', 'Manual Center Distance', 'mm', adsk.core.ValueInput.createByString('150 mm'))
            inputs.addBoolValueInput('exportCsv', 'Export CSV Summary', True, '', False)

            _set_input_state(inputs)

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
            if changed_input and changed_input.id in ['autoCenter', 'autoToothByRatio']:
                command = adsk.core.Command.cast(event_args.firingEvent.sender)
                _set_input_state(command.commandInputs)

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

            event_args = adsk.core.CommandEventArgs.cast(args)
            command = event_args.command
            inputs = command.commandInputs

            drive_teeth = inputs.itemById('driveToothCount').value
            driven_teeth = inputs.itemById('drivenToothCount').value
            auto_tooth_by_ratio = inputs.itemById('autoToothByRatio').value
            target_ratio = inputs.itemById('targetRatio').value
            max_pulley_diameter = inputs.itemById('maxPulleyDiameter').value
            min_tooth_limit = inputs.itemById('minToothLimit').value
            max_tooth_limit = inputs.itemById('maxToothLimit').value

            belt_pitch = inputs.itemById('beltPitch').value
            roller_diameter = inputs.itemById('rollerDiameter').value
            thickness = inputs.itemById('thickness').value
            drive_bore_diameter = inputs.itemById('driveBoreDiameter').value
            driven_bore_diameter = inputs.itemById('drivenBoreDiameter').value
            tip_clearance = inputs.itemById('tipClearance').value

            auto_center = inputs.itemById('autoCenter').value
            belt_links = inputs.itemById('beltLinks').value
            manual_center_distance = inputs.itemById('centerDistance').value
            export_csv = inputs.itemById('exportCsv').value

            ratio_mode_errors = _validate_ratio_solver_inputs(
                auto_tooth_by_ratio,
                target_ratio,
                max_pulley_diameter,
                min_tooth_limit,
                max_tooth_limit,
            )

            if ratio_mode_errors:
                ui.messageBox(
                    'Ratio mode validation failed:{}\n\nFix the values and run again.'.format(
                        _format_issues(ratio_mode_errors)
                    )
                )
                return

            solver_summary_lines = []
            if auto_tooth_by_ratio:
                solver_result = _solve_tooth_counts_for_ratio(
                    target_ratio,
                    min_tooth_limit,
                    max_tooth_limit,
                    max_pulley_diameter,
                    belt_pitch,
                    roller_diameter,
                    tip_clearance,
                )

                if not solver_result:
                    ui.messageBox(
                        'Could not find tooth counts that satisfy the ratio/diameter/range constraints.\n\n'
                        'Try increasing max diameter, widening the tooth range, or relaxing the target ratio.'
                    )
                    return

                drive_teeth = solver_result['drive_teeth']
                driven_teeth = solver_result['driven_teeth']
                solver_summary_lines = [
                    'Selection mode: auto from ratio constraints',
                    'Target ratio (driven:drive): {:.4f}'.format(target_ratio),
                    'Ratio error: {:.4f}%'.format(solver_result['ratio_error'] * 100.0),
                    'Max pulley diameter limit: {:.3f} mm'.format(max_pulley_diameter * 10.0),
                    'Resolved drive diameter: {:.3f} mm'.format(solver_result['drive_diameter'] * 10.0),
                    'Resolved driven diameter: {:.3f} mm'.format(solver_result['driven_diameter'] * 10.0),
                    'Tooth search range: {}-{}'.format(min_tooth_limit, max_tooth_limit),
                ]
            else:
                solver_summary_lines = ['Selection mode: manual tooth counts']

            input_errors = _validate_inputs(
                drive_teeth,
                driven_teeth,
                belt_pitch,
                roller_diameter,
                thickness,
                drive_bore_diameter,
                driven_bore_diameter,
                tip_clearance,
                auto_center,
                belt_links,
                manual_center_distance,
            )

            if input_errors:
                ui.messageBox('Input validation failed:{}\n\nFix the values and run again.'.format(_format_issues(input_errors)))
                return

            drive_pitch_radius, drive_root_radius, drive_tip_radius = _pulley_radii(
                drive_teeth, belt_pitch, roller_diameter, tip_clearance
            )
            _, driven_root_radius, driven_tip_radius = _pulley_radii(
                driven_teeth, belt_pitch, roller_diameter, tip_clearance
            )

            if drive_root_radius <= 0 or driven_root_radius <= 0:
                ui.messageBox('Computed root radius is invalid. Increase pitch or reduce tooth height.')
                return

            if drive_bore_diameter > (2.0 * drive_root_radius * 0.95):
                ui.messageBox('Drive bore diameter is too large for the drive pulley root/hub geometry.')
                return

            if driven_bore_diameter > (2.0 * driven_root_radius * 0.95):
                ui.messageBox('Driven bore diameter is too large for the driven pulley root/hub geometry.')
                return

            if auto_center:
                center_distance = _center_distance_from_belt_links(
                    belt_links, drive_teeth, driven_teeth, belt_pitch
                )
                if center_distance is None:
                    ui.messageBox(
                        'Belt links value is not feasible for this tooth pair. '
                        'Increase belt links or reduce tooth difference.'
                    )
                    return
                center_source = 'auto from belt links approximation'
            else:
                center_distance = manual_center_distance
                center_source = 'manual center distance input'

            if center_distance <= (drive_tip_radius + driven_tip_radius):
                ui.messageBox(
                    'Center distance is too small and causes pulley overlap. '
                    'Increase center distance or use more belt links.'
                )
                return

            engineering_warnings = _center_distance_warnings(center_distance, belt_pitch)

            root_comp = design.rootComponent
            pair_id = str(uuid.uuid4())

            drive_occ = root_comp.occurrences.addNewComponent(adsk.core.Matrix3D.create())
            drive_comp = drive_occ.component
            drive_comp.name = '{}T Drive Pulley'.format(drive_teeth)

            _create_pulley_geometry(
                drive_comp,
                drive_teeth,
                drive_root_radius,
                drive_tip_radius,
                thickness,
                drive_bore_diameter,
                '{}T_Drive_Pulley_Body'.format(drive_teeth),
            )

            driven_transform = adsk.core.Matrix3D.create()
            driven_transform.translation = adsk.core.Vector3D.create(center_distance, 0, 0)
            driven_occ = root_comp.occurrences.addNewComponent(driven_transform)
            driven_comp = driven_occ.component
            driven_comp.name = '{}T Driven Pulley'.format(driven_teeth)

            _create_pulley_geometry(
                driven_comp,
                driven_teeth,
                driven_root_radius,
                driven_tip_radius,
                thickness,
                driven_bore_diameter,
                '{}T_Driven_Pulley_Body'.format(driven_teeth),
            )

            _tag_pulley_entity(drive_comp, 'drive', pair_id, drive_teeth, belt_pitch, roller_diameter)
            _tag_pulley_entity(driven_comp, 'driven', pair_id, driven_teeth, belt_pitch, roller_diameter)
            _tag_pulley_entity(drive_occ, 'drive', pair_id, drive_teeth, belt_pitch, roller_diameter)
            _tag_pulley_entity(driven_occ, 'driven', pair_id, driven_teeth, belt_pitch, roller_diameter)

            ratio = float(driven_teeth) / float(drive_teeth)
            speed_factor = float(drive_teeth) / float(driven_teeth)
            ratio_error_pct = (abs(ratio - target_ratio) / target_ratio * 100.0) if auto_tooth_by_ratio else 0.0
            center_mm = center_distance * 10.0
            center_pitches = center_distance / belt_pitch

            csv_export_note = ''
            if export_csv:
                rows = [
                    ('GeneratedAt', datetime.datetime.now().isoformat(timespec='seconds')),
                    ('PairID', pair_id),
                    ('DriveTeeth', str(drive_teeth)),
                    ('DrivenTeeth', str(driven_teeth)),
                    ('SelectionMode', 'auto_ratio' if auto_tooth_by_ratio else 'manual'),
                    ('TargetRatioDrivenToDrive', '{:.6f}'.format(target_ratio) if auto_tooth_by_ratio else ''),
                    ('RatioError_pct', '{:.6f}'.format(ratio_error_pct) if auto_tooth_by_ratio else ''),
                    ('MaxPulleyDiameterLimit_mm', '{:.4f}'.format(max_pulley_diameter * 10.0) if auto_tooth_by_ratio else ''),
                    ('ToothRangeMin', str(min_tooth_limit) if auto_tooth_by_ratio else ''),
                    ('ToothRangeMax', str(max_tooth_limit) if auto_tooth_by_ratio else ''),
                    ('RatioDrivenToDrive', '{:.6f}'.format(ratio)),
                    ('DrivenSpeedFactor', '{:.6f}'.format(speed_factor)),
                    ('BeltPitch_mm', '{:.4f}'.format(belt_pitch * 10.0)),
                    ('ToothHeight_mm', '{:.4f}'.format(roller_diameter * 10.0)),
                    ('PulleyThickness_mm', '{:.4f}'.format(thickness * 10.0)),
                    ('DriveBoreDiameter_mm', '{:.4f}'.format(drive_bore_diameter * 10.0)),
                    ('DrivenBoreDiameter_mm', '{:.4f}'.format(driven_bore_diameter * 10.0)),
                    ('TipClearance_mm', '{:.4f}'.format(tip_clearance * 10.0)),
                    ('CenterDistance_mm', '{:.4f}'.format(center_mm)),
                    ('CenterDistance_pitches', '{:.6f}'.format(center_pitches)),
                    ('CenterSource', center_source),
                    ('EngineeringWarnings', ' | '.join(engineering_warnings) if engineering_warnings else ''),
                ]

                export_path = _export_csv_dialog(ui, 'pulley_pair', rows)
                if export_path:
                    csv_export_note = '\nCSV summary: {}'.format(export_path)
                else:
                    csv_export_note = '\nCSV summary: skipped by user.'

            warning_block = ''
            if engineering_warnings:
                warning_block = '\n\nEngineering warnings:{}'.format(_format_issues(engineering_warnings))

            selection_summary = '\n'.join(solver_summary_lines)
            if selection_summary:
                selection_summary = '{}\n'.format(selection_summary)

            ui.messageBox(
                'Created pulley pair.\n\n'
                '{}'
                'Drive teeth: {}\n'
                'Driven teeth: {}\n'
                'Ratio (driven:drive): {:.4f}\n'
                '{}'
                'Driven speed factor: {:.4f}x drive speed\n'
                'Center distance: {:.3f} mm\n'
                'Center distance: {:.3f} pitches\n'
                'Center source: {}{}{}'.format(
                    selection_summary,
                    drive_teeth,
                    driven_teeth,
                    ratio,
                    ('Ratio error vs target: {:.4f}%\n'.format(ratio_error_pct) if auto_tooth_by_ratio else ''),
                    speed_factor,
                    center_mm,
                    center_pitches,
                    center_source,
                    warning_block,
                    csv_export_note,
                )
            )

        except Exception:
            if ui:
                ui.messageBox('Pulley pair generation failed:\n{}'.format(traceback.format_exc()))


class CommandDestroyHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, args):
        # No explicit termination needed for add-ins.
        pass
