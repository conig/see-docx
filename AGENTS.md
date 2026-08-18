# See DOCX agent instructions

## Mandatory closeout gate

Before closing out **any** work on See DOCX:

1. Run the complete test battery with `make test-battery` from the repository root. This includes the automated suite and the real GTK pointer/selection/scrolling smoke test on workspace 15.
2. Confirm that the command exits successfully and that every test passes. Do not close out work with failures, skipped verification, or tests still running.
3. Only after the test battery is green, run `make install` from the repository root so the installed application contains the verified code.
4. If any source or test file changes after `make test-battery`, repeat `make test-battery` before running `make install` again.

Report both the passing test result and the successful installation in the final handoff.
