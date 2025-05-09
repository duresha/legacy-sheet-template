# Legacy Sheet Template - Task List

## Outstanding Issues & Required Enhancements

### 1. Fix Spurious New Lines within Sub-paragraphs
- **Problem:** Unwanted extra vertical spaces (like `<br>` effects or empty lines) appear within the content of sub-paragraphs (e.g., marriage details, list of children).
- **Goal:** Ensure text content inside each sub-paragraph flows continuously without unexpected line breaks.
- **Action:**
    - [ ] Investigate how lines constituting a single sub-paragraph are being processed and concatenated in `src/main.js`.
    - [ ] Modify the JavaScript logic to correctly join lines within a sub-paragraph, likely using spaces, similar to how main biographical paragraphs are handled.

### 2. Implement Person Name Hyperlinking
- **Problem:** A previous attempt to auto-link names failed due to a JavaScript regex error.
- **Goal:** Implement robust logic to identify names of people mentioned within the biographical text and link them.
- **Action:**
    - [ ] Revisit/fix the `linkPersonNames` function or create a new one.
    - [ ] The system should compare text against the list of all known persons (`parsedData.persons`).
    - [ ] Identified names (belonging to other individuals, not self-linking the main subject of the entry) should be wrapped in `<a>` tags.
    - [ ] Links should point to `href="#"`.
    - [ ] Links should be styled distinctly: standard blue color with an underline (CSS).
    - [ ] The solution must gracefully handle potential conflicts if the source text already contains Markdown-style links (`[Text](url)`), possibly by preserving those existing links.

### 3. Prevent Content Overflowing Footer & Add Truncation Indicator
- **Problem:** Large amounts of text cause dynamic content to visually overlap the fixed footer area.
- **Goal:** Prevent visual overlap and provide a clear indication of truncation within the single HTML template preview.
- **Action:**
    - [ ] Content should stop rendering cleanly before the reserved footer space (defined by `.legacy-sheet`'s `padding-bottom`).
    - [ ] The last visible line of text before the cutoff should end with an ellipsis (`...`).
    - [ ] Immediately below the truncated text (above the footer line), display an indicator like "[continued on next page]" (italicized).
    - [ ] Implement JavaScript logic to:
        - [ ] Calculate the available vertical space for content within `.legacy-sheet`.
        - [ ] Monitor the height of content as `.person-entry` elements are added.
        - [ ] Detect when adding the next element (or part of it) would exceed the available height.
        - [ ] Truncate the text of the last element appropriately (potentially adding ellipsis via JS or CSS on the last line fragment).
        - [ ] Dynamically insert the "[continued...]" indicator element.

## Next Steps
- Prioritize addressing these three outstanding issues.
- Suggested order:
    1. Fix Spurious New Lines within Sub-paragraphs OR Implement Person Name Hyperlinking.
    2. Tackle the remaining issue from the first step.
    3. Implement Content Overflowing Footer & Add Truncation Indicator (as it's likely more complex).
- Review `applyDataBtn` function in `src/main.js` and relevant CSS in `src/style.css` for all tasks. 
