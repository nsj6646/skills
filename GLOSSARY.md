# Safe Work Tool Builder

Domain vocabulary for an Antigravity skill that enables a non-developer to safely design, verify, and package internal work tools across a home build environment and a restricted company environment.

## Language

### Actors

**Maker**:
The non-developer office worker who designs, verifies, and uses Work Tools with Antigravity to automate company tasks.
_Avoid_: User, End-user, Developer, Wife

**Maintainer**:
The developer who authors and updates the skill and configures the Build PC's initial Antigravity environment and permissions.
_Avoid_: Admin, Author, Owner

### Environments & Artifacts

**Build PC**:
The local home computer running Antigravity where the Maker designs, verifies, and packages Work Tools without access to Company Data.
_Avoid_: Home PC, Local PC, Dev machine

**Target PC**:
The restricted, low-spec company computer where a packaged Work Tool executes against Company Data.
_Avoid_: Work PC, Company PC, Prod

**Work Tool**:
A discrete automation or replacement utility created on the Build PC for a specific company workflow, including its source, configuration, and verification criteria.
_Avoid_: App, Script, Project

**Delivery Format**:
The packaging type selected to satisfy Target PC runtime and security constraints so a Work Tool can execute there reliably.
_Avoid_: Build target, Output type

**Delivery Package**:
The complete, transferable bundle produced in a chosen Delivery Format—including the runnable artifact and any user-editable configuration or template files—ready to move to the Target PC.
_Avoid_: Release, Build, Executable

### Data Hygiene

**Company Data**:
Actual company files, emails, database records, or proprietary values from the work environment, which are strictly forbidden from entering the Build PC or AI prompts.
_Avoid_: Real data, Prod data, Confidential data

**Safe Sample Data**:
Newly synthesized data created on the Build PC that mimics the schema, formatting, and edge cases of Company Data without using copied or masked real files.
_Avoid_: Mock data, Dummy data, Test data

### Safeguards & Verification

**Checkpoint**:
A named restore point of a Work Tool on the Build PC captured before major edits so the Maker can revert cleanly if a change breaks behavior.
_Avoid_: Snapshot, Backup, Undo

**Staging Gate**:
An execution boundary on the Target PC where a Work Tool pauses with a preview or draft—such as an email draft or review file—for Maker confirmation before performing an irreversible action.
_Avoid_: Dry run, Soft mode, Draft mode

**Verification Case**:
A Maker-supplied input paired with an independently confirmed exact expected output used to validate calculation and transformation accuracy.
_Avoid_: Test case, Unit test, Acceptance test

**Quarantined Record**:
An anomalous input record—such as a spreadsheet error cell, missing required value, or format mismatch—isolated by a Work Tool for Maker review rather than silently guessed or coerced.
_Avoid_: Bad row, Skipped item, Error row
