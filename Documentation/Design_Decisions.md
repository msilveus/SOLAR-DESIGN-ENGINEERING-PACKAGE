# Solar Design Engineering Package (SDEP)

# Design Decisions

This document records significant architectural and engineering
decisions so future revisions preserve the original design intent.

------------------------------------------------------------------------

## DD-001 -- External Equipment Databases

**Status:** Accepted

**Decision:** Store Panels, Batteries, Inverters, Wire, etc. as CSV
files in `/Data`. `SDEP.xlsm` imports them into hidden `DB_*` worksheets
at startup.

**Reasoning** - Single source of truth - Git-friendly - Future GUI
compatible - Workbook formulas reference only `DB_*`

------------------------------------------------------------------------

## DD-002 -- Separation of Concerns

**Status:** Accepted

**Decision:** Separate the project into: 1. User Interface 2. Databases
3. Calculation Engine 4. Reports

------------------------------------------------------------------------

## DD-003 -- Workbook as the Application

**Status:** Accepted

**Decision:** Keep `SDEP.xlsm` in the repository root.

**Reasoning:** Simplifies startup, relative paths, and user workflow.

------------------------------------------------------------------------

## DD-004 -- Startup Initialization

**Status:** Accepted

Initialization sequence:

1.  Load configuration
2.  Refresh CSV databases
3.  Validate databases
4.  Refresh named ranges
5.  Recalculate workbook
6.  Display dashboard

------------------------------------------------------------------------

## DD-005 -- Single Source of Truth

**Status:** Accepted

Every engineering parameter shall exist in exactly one authoritative
location.

------------------------------------------------------------------------

## DD-006 -- Layout-Driven Design

**Status:** Accepted

The physical array layout drives all downstream calculations.

------------------------------------------------------------------------

## DD-007 -- User-Defined String Topology

**Status:** Accepted

The software analyzes the user's string topology rather than imposing
one.

------------------------------------------------------------------------

## DD-008 -- Column-Oriented Strings

**Status:** Proposed

Evaluate column-oriented strings for sites with progressive east/west
tree shading.

------------------------------------------------------------------------

## DD-009 -- Numerical Engine First

**Status:** Accepted

Complete and validate the numerical model before building an interactive
GUI.

------------------------------------------------------------------------

## Open Questions

-   MPPT pairing optimization
-   Shade engine
-   GUI migration
-   Database versioning
