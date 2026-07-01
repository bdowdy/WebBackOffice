JetFix2 — Jet->SQL Server SQL rewrites applied to COPIES of the M$ source.

WHAT THIS IS
  Automated, faithful rewrites of the safe Jet-only-SQL patterns, applied to
  copies (never the live Program Files tree). Review, then copy the files in
  /fixed into your VB6 project and COMPILE before using.

AUTO-APPLIED (915 edits across 25 files)
  - cast([x] as date/datetime)      -> TRY_CAST([x] AS date/datetime)   (403)
  - cast([x] as int)                -> TRY_CONVERT(int, [x])            (4)
  - Val([Terminated] & '')          -> ISNULL(TRY_CONVERT(int,[Terminated]),0) (3)
  - format([Hour],'HH:MM ampm')     -> TRY_CAST([Hour] AS time)         (2)
  - cdate([Col])                    -> TRY_CAST([Col] AS datetime)      (268)
  - #" & expr & "#  (simple)        -> '" & Format(CDate(expr),"yyyy-mm-dd hh:nn:ss") & "'  (235)

FILES
  fixed/            corrected copies (drop into project after review)
  *.diff            unified diff per file (what changed)
  NEEDS_REVIEW.txt  lines NOT auto-fixed - handle by hand (see guide)

STILL MANUAL (see NEEDS_REVIEW.txt + ..\JetToSqlServer_Remediation.md)
  - BETWEEN #date + time# ranges (the # literal holds concatenated time).
  - datepart('ww'/'h', cdate(...)) -> DATEPART(ISO_WEEK/hour, TRY_CAST(... AS date/time)).
  - #" & Format(x,"MM/DD/YYYY") & "# cases (these were left half-done: the
    cdate side is converted but the # literal remains - finish them).

NOTE: ignore the older sibling folder "JetFix" (first pass); this JetFix2 is the
good one. The mount wouldn't let me delete it.
