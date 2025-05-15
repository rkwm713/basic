# Error Handling and Edge Cases for Make-Ready Report Generation

This document outlines potential error conditions, edge cases, and recommended handling strategies for the Make-Ready Report generation application. Robust error handling is crucial for producing reliable reports and for aiding in debugging.

## I. File Input/Output Errors

1.  **JSON File Not Found:**
    *   **Condition:** Specified SPIDAcalc or Katapult JSON file does not exist at the given path.
    *   **Handling:**
        *   Log a critical error message specifying which file is missing.
        *   Terminate the application gracefully.
        *   Inform the user clearly about the missing file.
2.  **JSON File Unreadable:**
    *   **Condition:** File exists but cannot be read due to permissions issues.
    *   **Handling:** Similar to "File Not Found."
3.  **Malformed JSON Data:**
    *   **Condition:** File content is not valid JSON.
    *   **Handling:**
        *   Attempt to parse the JSON. If `json.JSONDecodeError` (Python) occurs:
            *   Log a critical error indicating the file and, if possible, the approximate location of the syntax error (some libraries provide this).
            *   Terminate gracefully.
            *   Inform the user about the malformed JSON.
4.  **Excel Output Issues:**
    *   **Condition:** Cannot write the output `.xlsx` file (e.g., permissions, disk full, invalid path).
    *   **Handling:**
        *   Log a critical error.
        *   Inform the user that the report could not be saved.

## II. Data Validation and Missing Data

1.  **Missing Critical Keys/Fields in JSON:**
    *   **Condition:** Essential keys or nested fields expected by the extraction logic are absent from a record.
        *   Example: A Katapult node is missing `['attributes']['PoleNumber']`.
        *   Example: A SPIDAcalc attachment is missing `['attachmentHeight']['value']`.
    *   **Handling (per field):**
        *   **Log a Warning:** Indicate the specific pole/attachment and the missing field.
        *   **Use Default/NA:** Populate the corresponding report field with "NA", an empty string, or a predefined default as specified in `excel_gener_details.txt` or `project_plan.txt`.
        *   **Continue Processing:** Generally, the application should attempt to process the rest of the pole/report rather than terminating, unless the missing data makes further processing of that specific item impossible.
2.  **Unexpected Data Types:**
    *   **Condition:** A field contains data of a type different from what's expected (e.g., an expected number is a string that cannot be converted, an expected list is a dictionary).
    *   **Handling:**
        *   **Log a Warning:** Specify the field, the unexpected type, and the item being processed.
        *   **Attempt Conversion:** If applicable (e.g., string "123" to int 123).
        *   **Fallback to Default/NA:** If conversion fails or is not applicable, use "NA" or a default.
3.  **Pole Matching Failures:**
    *   **Condition:** A pole identifier from SPIDAcalc (after normalization) does not have a corresponding match in the Katapult data.
    *   **Handling Options (to be decided based on requirements):**
        *   **Option A (Skip Pole):** Log a warning and skip this SPIDAcalc pole entirely from the report.
        *   **Option B (Report with Missing Katapult Data):** Include the pole in the report, populating SPIDAcalc-derived fields and using "NA" (or similar) for all fields that would normally come from Katapult. Log a warning.
            *   **Details for Option B Implementation:** If Option B is implemented, the report should populate all SPIDAcalc-derivable fields for that pole. This includes:
                *   Pole Owner (if SPIDAcalc is the fallback or primary source as per data prioritization rules)
                *   Pole # (from SPIDAcalc)
                *   Pole Structure (derived from SPIDAcalc data: height, class, species)
                *   Proposed Riser (Yes/No) & (from SPIDAcalc Recommended Design)
                *   Proposed Guy (Yes/No) & (from SPIDAcalc Recommended Design)
                *   PLA (%) with proposed attachment (from SPIDAcalc analysis of Recommended Design)
                *   Construction Grade of Analysis (from SPIDAcalc analysis)
                *   Make Ready Data - Attacher Description (for attachments defined in SPIDAcalc designs)
                *   Make Ready Data - Attachment Height - Existing (from SPIDAcalc Measured Design)
                *   Make Ready Data - Attachment Height - Proposed (from SPIDAcalc Recommended Design)
                *   Fields that are exclusively sourced from Katapult (e.g., field-verified existing mid-span heights if Katapult is primary, From/To Pole from Katapult connections, specific Katapult MR notes/violations if distinct from SPIDA) should be populated with 'NA' or an appropriate placeholder for these unmatched SPIDA poles.
        *   **Recommendation:** Option B is often preferred as it highlights data gaps rather than silently omitting data. This should be configurable or clearly defined.
4.  **Missing SPIDAcalc Design Sections:**
    *   **Condition:** A SPIDAcalc pole location is missing the "Measured Design" or "Recommended Design" array/object within its `designs` list.
    *   **Handling:**
        *   Log a warning for the specific pole and missing design.
        *   If "Measured Design" is missing, existing attachment data from SPIDA will be unavailable.
        *   If "Recommended Design" is missing, proposed attachment data, PLA, proposed guys/risers from SPIDA will be unavailable.
        *   Populate relevant fields with "NA". The "Attachment Action" logic might default to "( E )xisting" or require special handling.
5.  **Empty Arrays/Lists where Data is Expected:**
    *   **Condition:** An array like `structure['guys']` is present but empty.
    *   **Handling:** This is often valid (e.g., "NO" guys). The logic (e.g., `count > 0`) should handle this naturally. No error needed unless the expectation is that it *must* contain items.

## III. Edge Cases in Logic

1.  **Ambiguous Pole Owner/Data from Multiple Sources:**
    *   **Condition:** Both SPIDA and Katapult provide a "Pole Owner," and they differ. Or, Katapult's `pole_owner.multi_added` has multiple entries.
    *   **Handling:** The prioritization rules in `project_plan.txt` should cover this (e.g., "Prioritize Katapult `multi_added[0]`"). If multiple Katapult owners, pick the first or join if appropriate (though typically it's one). Log a notice if discrepancies are significant and not covered by a clear rule.
2.  **Multiple Katapult Connections (Spans) for "From/To Pole":**
    *   **Condition:** A Katapult node has multiple entries in `connections` linking it to different poles.
    *   **Handling:** The `project_plan.txt` (Rule 13) notes: "Logic for selecting the 'primary' span if multiple exist needs to be defined (e.g., the one connecting to SPIDAcalc's `NEXT_POLE` concept, or a specific span type if applicable)." If no such rule can be reliably implemented, the first valid aerial connection could be chosen, with a warning logged if multiple viable spans exist.
3.  **Units Conversion Failures:**
    *   **Condition:** A height value from SPIDA (meters) or Katapult (feet/inches) is present but malformed (e.g., non-numeric) and cannot be converted.
    *   **Handling:** Log a warning. Populate the relevant height field with "NA".
4.  **No Attachments on a Pole:**
    *   **Condition:** A pole has no attachments in SPIDA (Measured or Recommended) or Katapult.
    *   **Handling:** `excel_gener_details.txt` specifies: "If `num_attachments_for_this_pole` is 0, treat as 1 for formatting purposes (to show at least one line for the pole)." Columns L-O for this single line would typically be "NA" or blank.

## IV. Logging Recommendations

*   **Logging Levels:** Use different logging levels (e.g., DEBUG, INFO, WARNING, ERROR, CRITICAL).
*   **Startup Information:** Log application start, input file paths.
*   **Processing Milestones:** Log when major stages begin/end (e.g., "Starting Pole Matching," "Finished processing X poles").
*   **Warnings:** For recoverable issues, missing non-critical data, or potential data quality concerns.
*   **Errors:** For issues that prevent a specific item (e.g., one pole) from being fully processed but allow the application to continue.
*   **Critical Errors:** For issues that force the application to terminate (e.g., file not found, malformed JSON).
*   **Contextual Information:** Logs should include relevant context, such as the current Pole ID or attachment details being processed when an issue occurs.
*   **Summary:** At the end of processing, log a summary (e.g., "Processed X poles. Generated Y rows in report. Encountered Z warnings.").

By considering these error conditions and edge cases proactively, the application can be made more robust, user-friendly, and easier to maintain.
