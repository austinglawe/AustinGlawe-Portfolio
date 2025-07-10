' -----------------------------------------
' UserForm layout and design best practices
' -----------------------------------------
'
' Size and proportions:
' - Keep forms compact and task-focused.
' - Avoid overcrowding; leave enough space between controls.
' - Minimize form height to avoid unnecessary scrolling.
'
' Alignment:
' - Align similar controls neatly using VBA editor alignment tools.
' - Labels should align with their associated input fields.
'
' Tab order (TabIndex property):
' - Set TabIndex to ensure a logical navigation sequence.
' - Follow reading order (typically left to right, top to bottom).
'
' Label association:
' - Every input control should have a clear, adjacent Label.
'
' Grouping controls:
' - Use Frames to group related controls logically.
' - OptionButtons inside a Frame act as an independent group.
'
' Default and Cancel buttons:
' - Set Default = True on primary action buttons (e.g., "Submit").
' - Set Cancel = True on secondary buttons (e.g., "Cancel" or "Close").
'
' Consistent fonts and colors:
' - Avoid unnecessary styling unless it serves a purpose.
' - Use colors sparingly and meaningfully (e.g., highlight editable fields).
'
' Size anchors and responsiveness:
' - UserForms do not resize automatically.
' - Manual coding is required to make forms and controls resize dynamically.
'
' Best practices summary:
' - Group related inputs with Frames.
' - Align controls neatly for readability.
' - Use Labels for all inputs.
' - Set TabIndex properly for good keyboard navigation.
' - Provide clear default and cancel buttons.
' - Maintain a clean and consistent visual style.
'
' -----------------------------------------

