# Data Flow and Processing Logic for Make-Ready Report Generation

This document outlines the overall data flow and the sequence of major processing steps involved in generating the Make-Ready Excel report from SPIDAcalc and Katapult JSON files.

## I. High-Level Data Flow

The following represents the general sequence of operations:

```
+-------------------------+      +-----------------------+
|   SPIDAcalc JSON File   |      |  Katapult JSON File   |
| (spidacalc_data)        |      | (katapult_data)       |
+-------------------------+      +-----------------------+
           |                                |
           +--------------+-----------------+
                          |
                          V
+-------------------------------------------------+
| 1. Load and Parse JSON Files                    |
|    - Read content into Python dictionaries.     |
|    - Basic validation (e.g., is it valid JSON?).|
+-------------------------------------------------+
                          |
                          V
+-------------------------------------------------+
| 2. Pole Matching                                |
|    - Process SPIDAcalc pole labels to           |
|      Canonical Pole ID format (e.g., "PLXXXXXX").|
|    - Create a lookup map from Canonical Pole ID  |
|      to Katapult `node_id`.                     |
|    - Identify poles present in SPIDAcalc to     |
|      drive report generation.                   |
+-------------------------------------------------+
                          |
                          V
+-------------------------------------------------------------+
| 3. Iterate Through Matched Poles (Primary Processing Loop)  |
|    For each SPIDAcalc pole that has a match in Katapult:    |
|    +-----------------------------------------------------+  |
|    | 3a. Extract Pole-Level Data                         |  |
|    |     - Pole #, Owner, Structure, PLA, Grade, etc.    |  |
|    |     - Apply prioritization rules for dual sources.  |  |
|    +-----------------------------------------------------+  |
|    | 3b. Consolidate Attachments                         |  |
|    |     - Gather from SPIDA Measured & Recommended.     |  |
|    |     - Gather from Katapult attachments.             |  |
|    |     - Deduplicate and reconcile into a unique list. |  |
|    +-----------------------------------------------------+  |
|    | 3c. Extract Attachment-Level Data                   |  |
|    |     For each unique attachment on the current pole: |  |
|    |     - Description, Existing Height, Proposed Height.|  |
|    |     - Apply prioritization for existing height.     |  |
|    +-----------------------------------------------------+  |
|    | 3d. Determine Mid-Span Data                         |  |
|    |     - Existing Lowest Com & CPS Electrical.         |  |
|    |     - Proposed Mid-Span for attachments.            |  |
|    |     - Apply fallback logic (NA, UG).                |  |
|    +-----------------------------------------------------+  |
|    | 3e. Determine "From Pole" / "To Pole" Data          |  |
|    |     - Use Katapult connections data.                |  |
|    |     - Identify primary span if multiple.            |  |
|    +-----------------------------------------------------+  |
|    | 3f. Determine "Attachment Action" Summary           |  |
|    |     - Compare SPIDA Measured vs. Recommended.       |  |
|    |     - Consider Katapult work_type for overrides.    |  |
|    +-----------------------------------------------------+  |
+-------------------------------------------------------------+
                          |
                          V
+-------------------------------------------------+
| 4. Populate Internal Data Structure             |
|    - Accumulate rows of data, with each row     |
|      representing an attachment (or a single    |
|      line for poles with no attachments).       |
|    - This structure will be used to create the  |
|      Pandas DataFrame.                          |
+-------------------------------------------------+
                          |
                          V
+-------------------------------------------------+
| 5. Create and Format Excel Report (using openpyxl)|
|    - Create Pandas DataFrame from internal data.|
|    - Write DataFrame to Excel.                  |
|    - Apply static headers (Rows 1-3).           |
|    - Merge cells for pole-level data (Cols A-K).|
|    - Apply text wrapping, bolding, alignment.   |
|    - Adjust column widths.                      |
+-------------------------------------------------+
                          |
                          V
+-------------------------------------------------+
|             Make-Ready Excel Report             |
|                  (.xlsx file)                   |
+-------------------------------------------------+

```

## II. Key Algorithm Descriptions

### A. Pole Matching (Rule 2 & 3 in `project_plan.txt`)

1.  **Normalize SPIDAcalc Pole Identifiers:**
    *   Iterate through SPIDAcalc pole locations (e.g., `spidacalc_data['leads'][0]['locations']`).
    *   For each location, extract the pole label (e.g., `location['label']`).
    *   Apply a normalization function to convert this label to the "Canonical Pole ID Format" (e.g., "1-PL123456" -> "PL123456", "P.O.12345" -> "PO12345"). This function should handle known prefixes, suffixes, or patterns.
2.  **Create Katapult Pole ID Lookup Map:**
    *   Iterate through `katapult_data['nodes']`.
    *   For each Katapult node, extract its Pole Number (e.g., `node['attributes']['PoleNumber']['assessment']`). This should already be in or easily convertible to the Canonical Pole ID Format.
    *   Create a dictionary mapping the Canonical Pole ID to the Katapult `node_id` (e.g., `{"PL123456": "-OJ_NodeIdAbc"}`).
3.  **Match and Iterate:**
    *   The primary loop for report generation will iterate through the normalized SPIDAcalc pole identifiers.
    *   For each SPIDAcalc pole, use its Canonical Pole ID to look up the corresponding Katapult `node_id` in the map created in step 2.
    *   If a match is found, both the SPIDAcalc pole data and the Katapult node data are available for processing.
    *   If no match is found, this pole might be skipped or handled according to error/edge case rules (e.g., logged, reported with missing Katapult data).

### B. Attachment Consolidation (Rule 10 in `project_plan.txt`)

1.  **Gather Attachments from Sources:**
    *   For the current matched pole:
        *   Extract attachments from SPIDAcalc "Measured Design" (`structure['wires']` and `structure['equipments']`).
        *   Extract attachments from SPIDAcalc "Recommended Design" (`structure['wires']` and `structure['equipments']`).
        *   Extract attachments from the corresponding Katapult node (`node['attachments']`).
2.  **Define "Uniqueness":**
    *   An attachment is considered unique based on a combination of its Primary Owner (e.g., 'AT&T', 'CPS Energy') AND a keyword from its Attachment Type/Description (e.g., 'Fiber Optic Com', 'Neutral', 'Riser').
    *   Store key attributes for each gathered attachment: owner, type/description, original source (SPIDA-Measured, SPIDA-Recommended, Katapult), existing height (if applicable), proposed height (if applicable), and any other relevant details like `usageGroup`.
3.  **Reconcile and Deduplicate:**
    *   Create an empty list for the consolidated unique attachments.
    *   Iterate through the gathered attachments. For each attachment:
        *   Determine its unique key (Owner + Type).
        *   If an attachment with the same key already exists in the consolidated list:
            *   Merge information. For example, if Katapult provides a field-verified existing height, prioritize it over SPIDA's Measured Design height for that matched attachment.
            *   Update flags (e.g., if it exists in SPIDA-Recommended, it's part of the proposed state).
        *   If the attachment key is new, add it to the consolidated list.
4.  **Identify State (New, Existing, Modified, Removed):**
    *   Based on its presence in SPIDA Measured, SPIDA Recommended, and Katapult data, determine the state of each unique attachment. This informs the "Attachment Action" summary and how heights are populated.
        *   Present in SPIDA Recommended but not Measured: New.
        *   Present in SPIDA Measured but not Recommended: Removed.
        *   Present in both: Existing (potentially modified if heights differ).
        *   Present in Katapult: Provides field-verified existing state.

### C. "Attachment Action" Summary (Rule 12 in `project_plan.txt`)

1.  **Primary Driver: SPIDA Design Comparison:**
    *   After consolidating attachments for a pole, compare the set of attachments in the "Measured Design" against the "Recommended Design."
    *   If any attachments are present in "Recommended" but were NOT in "Measured" -> Default action is "( I )nstalling".
    *   If any attachments were present in "Measured" but are NOT in "Recommended" -> Consider if "( R )emoving" is appropriate (though this report focuses on make-ready for additions/modifications).
    *   If attachments exist in both, but their properties (like height) change significantly -> Default action is "( E )xisting" (implying modification).
    *   If attachments exist and there are no significant changes, new additions, or removals -> Default action is "( E )xisting".
2.  **Katapult Influence (Overrides/Validation):**
    *   Check `katapult_pole_data.attributes.kat_work_type.button_added` or `work_type.button_added`. If this indicates "Denied", this might override the summary action or be flagged. For the primary Column B value, SPIDA comparison is key unless a "Denied" status dictates a specific non-action.
    *   `katapult_pole_data.attributes.mr_state` can serve as a validation point or for ancillary notes if the report format supported them.

This document provides a high-level overview. Refer to `ai-rules/project_plan.txt` and other specific schema documents for detailed rules and JSON paths.
