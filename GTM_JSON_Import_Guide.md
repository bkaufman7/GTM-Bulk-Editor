# GTM JSON Import Guide

## Purpose
This guide explains the recommended import workflow for the GTM Bulk Trigger Assignment Editor.

The tool supports:

- Trigger assignment edits on existing tags.
- Floodlight tag creation from pasted DCM TagSheet rows.

For trigger assignment edits, it only edits these GTM tag properties:

- `tag.firingTriggerId`
- `tag.blockingTriggerId`

For Floodlight creation, it clones an existing Floodlight template tag (`type: flc`) and updates selected fields such as:

- `name`
- `activityTag`
- `groupTag`
- `ordinalType`
- `sessionId`
- numbered custom variables (`u1`, `u2`, etc.)
- transaction fields when provided (`value`, `transaction_id`, `quantity`)

It does not publish changes or call live GTM approval queue APIs.

## Workflow
1. Export the full GTM container JSON from GTM.
2. Paste the JSON into the Google Sheet using the custom `GTM Bulk Editor` menu.
3. Run `Build Editor Tabs`.
4. Choose one or both edit paths:
	- Trigger assignments: make row-based edits or rule-based edits.
	- Floodlight imports: paste DCM rows in the sheet, select them, then run `Import DCM Floodlights From Selection`.
5. In `Floodlight Import`, verify trigger IDs and any optional fields.
6. Run `Build Preview` for trigger assignment review.
7. Review warnings and errors in `Edit Preview`.
8. Run `Apply Edits & Create Export JSON`.
9. Download or open the generated JSON export.
10. Import the modified JSON into a new GTM workspace.
11. Choose `Merge`, then choose overwrite for conflicting tags/triggers/variables.
12. Review GTM’s detailed change summary before confirming.

## Floodlight Import Notes
- DCM TagSheet mapping expects rows with fields like `Activity Name`, `Group Name`, and `Event Snippet`.
- `activityTag` and `groupTag` are parsed from `send_to` or image tag values in the event snippet.
- Counting type is mapped to GTM ordinal type where available.
- Numbered custom variables are always written when provided (`u1`, `u2`, etc.).
- If no template tag is specified per row, the tool uses the first existing GTM Floodlight tag in the container.

## Safety Notes
- Treat the pasted container export as the source of truth.
- Import into a new workspace first so GTM can show a reviewable diff.
- Do not rely on trigger names alone when a Trigger ID is available.
- If GTM reports conflicts, review them before accepting the import.

## Known Limitations
- No live GTM API calls for approval queue actions.
- No publishing.
- No trigger deletion.
- No edits to trigger definitions, variables, folders, templates, or consent settings outside selected Floodlight tag parameters.
- Ambiguous trigger names are not auto-resolved.

## Suggested File Name for Export
The script creates a Drive JSON file named like:

`gtm-bulk-edited-container-YYYYMMDD-HHMMSS.json`
