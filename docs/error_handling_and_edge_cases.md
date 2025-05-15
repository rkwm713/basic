# Error Handling and Edge Cases for Make-Ready Report Generation

This document outlines potential error conditions, edge cases, and recommended handling strategies for the Make-Ready Report generation application. Robust error handling is crucial for producing reliable reports and for aiding in debugging.

## I. File Input/Output Errors

1.  **JSON File Not Found:**
    *   **Condition:** Specified Katapult JSON file does not exist at the given path.
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
5.  **Empty Arrays/Lists where Data is Expected:**
    *   **Condition:** An array like `structure['guys']` is present but empty.
    *   **Handling:** This is often valid (e.g., "NO" guys). The logic (e.g., `count > 0`) should handle this naturally. No error needed unless the expectation is that it *must* contain items.

## III. Edge Cases in Logic

1.  **Ambiguous Pole Owner/Data from Multiple Sources:**
    *   **Condition:** Katapult's `pole_owner.multi_added` has multiple entries, or the primary pole owner field might be ambiguous if sourced from different internal Katapult attributes.
    *   **Handling:** The prioritization rules in `project_plan.txt` (or equivalent Katapult-focused documentation) should cover this (e.g., "Prioritize Katapult `multi_added[0]` if present, otherwise a primary designated owner field"). If multiple Katapult owners are listed in `multi_added`, the first entry is typically chosen. Log a notice if discrepancies are significant and not covered by a clear rule.
2.  **Multiple Katapult Connections (Spans) for "From/To Pole":**
    *   **Condition:** A Katapult node has multiple entries in `connections` linking it to different poles.
    *   **Handling:** The `project_plan.txt` (or equivalent Katapult-focused documentation, e.g., Rule 13 if still applicable) should define logic for selecting the 'primary' span if multiple exist (e.g., based on a specific span type like 'aerial_cable' vs 'reference', or connection attributes). If no such rule can be reliably implemented, the first valid aerial connection could be chosen, with a warning logged if multiple viable spans exist.
3.  **Units Conversion Failures:**
    *   **Condition:** A height value from Katapult (expected in feet/inches) is present but malformed (e.g., non-numeric) and cannot be converted.
    *   **Handling:** Log a warning. Populate the relevant height field with "NA".
4.  **No Attachments on a Pole:**
    *   **Condition:** A pole has no attachments in Katapult.
    *   **Handling:** `excel_gener_details.txt` (or equivalent Katapult-focused documentation) specifies: "If `num_attachments_for_this_pole` is 0, treat as 1 for formatting purposes (to show at least one line for the pole)." Columns for attachment details for this single line would typically be "NA" or blank.

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
