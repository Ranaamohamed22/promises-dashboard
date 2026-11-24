# EJADA: SNB performance visual preview (safe prototype)

This branch includes a **separate** functional Streamlit entrypoint,
performance_preview.py, reproducing the requested navy sidebar, green file
upload panels, four KPI cards, interactive analysis tabs, charts, report table
and original Excel downloads.

## Safety guarantees for the PROTOTYPE
- The production entrypoint app.py was NOT changed.
- The business calculation module snb_performance.py was NOT changed.
- The comprehensive report snb_comprehensive.py was NOT changed.
- No encrypted data in data/ and no Excel source files were changed.
- No external AI or third-party requests are made with uploaded files.
- The original report generator build_snb_performance() still determines
  treatment status and the exact exported Excel contents.
- The additional design-only percentage is a **row-based** ratio:
  processed_count / number_of_portfolio_rows * 100.
  It does not claim to be unique-account processing rate.
- Product and collector charts appear only if the corresponding columns
  exist in the input. Never fabricate missing groups or values.
- All results derive from the current two uploads only.

## Running the preview
On the preview branch, install dependencies and launch:

    python -m pip install -r requirements.txt
    streamlit run performance_preview.py

The original production command remains:

    streamlit run app.py

The preview has its own sidebar with working Home, Performance and Comprehensive
links. Other report tabs are deliberately labeled as existing only in the
production application rather than showing nonfunctional buttons.

## Test
    python -m unittest discover -s preview_tests -p "test_*.py" -v

## Before production adoption
1. Review on a separate Streamlit preview app or locally.
2. Compare the existing report's KPI values and both Excel exports using
   approved sanitized samples.
3. Validate UI on desktop and smaller screens.
4. Integrate the design into the existing app.py navigation only after approval.
5. Do not merge this branch until the UI is accepted and the checks pass.

**Note:** This preview is not deployed on the live production Streamlit app.
