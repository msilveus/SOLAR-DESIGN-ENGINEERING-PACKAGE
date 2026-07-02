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

------------------------------------------------------------------------

## Project Philosophy

SDEP is being developed as an engineering application whose first
implementation happens to use Microsoft Excel as its host environment.

The project architecture intentionally separates:

-   Data
-   Engineering Logic
-   User Interface
-   Documentation

This separation enables long-term maintainability and provides a
migration path to future desktop or web applications while preserving
the engineering model.

------------------------------------------------------------------------

## DD-010 -- Database Abstraction Layer

**Status:** Accepted

**Decision:** All engineering data shall be accessed exclusively through
the `modDatabase` module. No worksheet, calculation module, or VBA
routine outside `modDatabase` shall directly reference `DB_*`
worksheets.

**Reasoning**

-   Decouples engineering calculations from storage.
-   Eliminates direct worksheet dependencies.
-   Supports future migration to other platforms.

------------------------------------------------------------------------

## DD-011 -- Header-Based Field Resolution

**Status:** Accepted

**Decision:** Database fields shall be identified by header names rather
than fixed column letters or column numbers.

**Reasoning**

-   Column order becomes irrelevant.
-   New fields may be inserted without code changes.
-   Eliminates magic numbers.

------------------------------------------------------------------------

## DD-012 -- Database Cache Architecture

**Status:** Accepted

**Decision:** External CSV files remain the authoritative data source.
Internal `DB_*` worksheets function only as runtime cache tables
populated during workbook initialization.

------------------------------------------------------------------------

## DD-013 -- Layered Software Architecture

**Status:** Accepted

**Decision:**

    CSV Files
        │
    Import Engine
        │
    Database Cache
        │
    Database API (modDatabase)
        │
    Engineering Engine
        │
    User Interface

Each layer communicates only with the layer immediately below it.

------------------------------------------------------------------------

## DD-014 -- Workbook as a Host Application

**Status:** Accepted

**Decision:** Excel is the host environment rather than the engineering
engine. Engineering logic resides primarily in reusable VBA modules
while worksheets provide user interface, visualization, and reporting.

------------------------------------------------------------------------

## DD-015 -- Dynamic User Interface

**Status:** Accepted

**Decision:** User interface elements shall be generated dynamically
from imported databases during workbook initialization.

------------------------------------------------------------------------

## DD-016 -- Schema Agnostic Databases

**Status:** Accepted

**Decision:** The software shall never depend upon a fixed column order
within any CSV database. Field discovery shall always occur through
header lookup.

**Reasoning**

-   Database schemas are expected to evolve.
-   Future expansion should not require code changes.
