# Employee Data Consolidation - File Format Guide

## Overview
This application processes 7 Excel files with specific tab structures and column mappings.

---

## 1. Core Staff

| Property | Value |
|----------|-------|
| **Upload Key** | `core` |
| **Expected Tab** | `Employee details` |
| **Data Start Row** | Row 2 (skip header) |

### Columns
| Column | Letter | Field |
|--------|--------|-------|
| Emp ID | C | `emp_id` |
| Emp Name | D | `emp_name` |
| Date of Joining | G | `doj` |
| Status | E | `status` |
| LOP | S | `lop` |
| Last Working Day | F | `lwd` |

### Status Detection
- **Active**: "Active", "ACT", "Working"
- **Hold**: "Hold", "Resigned", any other status

---

## 2. Project

| Property | Value |
|----------|-------|
| **Upload Key** | `project` |
| **Expected Tab** | `Project SM` |
| **Structure** | Multi-section (3 sections) |

### Sections (detected by header row text)
| Section | Trigger Text | Data Start |
|---------|--------------|------------|
| Active | "Active Employee Details" | Row after header |
| LOP | "LOP Details" | Row after header |
| Resigned | "Resigned Employee Details" | Row after header |

### Columns (same for all sections)
| Column | Letter | Field |
|--------|--------|-------|
| Emp ID | C | `emp_id` |
| Emp Name | D | `emp_name` |
| DOJ | F | `doj` |
| Status | E | `status` |
| LWD | G | `lwd` |
| LOP | N | `lop` |

### Special Handling
- **LOP rows**: Status set to "Hold"
- **Resigned rows**: Status set to "Hold"

---

## 3. CF (Contractor)

| Property | Value |
|----------|-------|
| **Upload Key** | `cf` |
| **Expected Tab** | `CF Employee Details` |
| **Data Start Row** | Row 2 |

### Columns
| Column | Letter | Field |
|--------|--------|-------|
| Emp ID | B | `emp_id` |
| Emp Name | C | `emp_name` |
| DOJ | E | `doj` |
| Status | D | `status` |
| LOP | K | `lop` |
| LWD | F | `lwd` |

---

## 4. HK (Housekeeping)

| Property | Value |
|----------|-------|
| **Upload Key** | `hk` |
| **Expected Tab** | `HK Employee details` |
| **Structure** | Multi-section (Active + Left) |

### Sections
| Section | Trigger Text | Data Start | Skip Tabs |
|---------|--------------|------------|-----------|
| Active | "Active Employee Details" | Row after header | Yes |
| Left | "Left Employee Details" | Row after header | Yes |

### Skip These HK Tabs
- `LOP Details`
- `Left F&F` 
- `New HK Staff Details`

### Columns
| Column | Letter | Field |
|--------|--------|-------|
| Emp ID | C | `emp_id` |
| Emp Name | D | `emp_name` |
| DOJ | F | `doj` |
| Status | E | `status` |
| LOP | Q | `lop` |
| LWD | G | `lwd` |

### Special Handling
- **LOP rows** (separate tab): Only update LOP on existing employees, never create new records

---

## 5. Retainer

| Property | Value |
|----------|-------|
| **Upload Key** | `retainer` |
| **Expected Tab** | `Retainer LOP` |
| **Data Start Row** | Row 2 |

### Columns
| Column | Letter | Field |
|--------|--------|-------|
| Emp ID | A | `emp_id` |
| Emp Name | B | `emp_name` |
| DOJ | F | `doj` |
| Status | D | `status` |
| LOP | C | `lop` |
| LWD | G | `lwd` |

### Special Handling
- Retainer employees are treated as **Off-roll**
- LOP days are primary purpose of this file

---

## 6. School

| Property | Value |
|----------|-------|
| **Upload Key** | `school` |
| **Expected Tab** | `SM School Staff` |
| **Structure** | Multi-section (Active + LOP + Resigned) |
| **Data Start Row** | Row 3 (skip 2 header rows) |

### Sections
| Section | Trigger Text | Data Start |
|---------|--------------|------------|
| Active | "Active Employee Details" | Row after header |
| LOP | "LOP Details" | Row after header |
| Resigned | "Resigned Employee Details" | Row after header |

### Columns
| Column | Letter | Field |
|--------|--------|-------|
| Emp ID | C | `emp_id` |
| Emp Name | D | `emp_name` |
| DOJ | F | `doj` |
| Status | E | `status` |
| LWD | G | `lwd` |
| LOP | N | `lop` |

### Special Handling
- Skip rows where Emp ID is "Emp ID" (header row detection)
- **LOP rows**: Status set to "Hold"

---

## 7. College

| Property | Value |
|----------|-------|
| **Upload Key** | `college` |
| **Expected Tab** | `SM college Staff` |
| **Data Start Row** | Row 2 |

### Columns
| Column | Letter | Field |
|--------|--------|-------|
| Emp ID | B | `emp_id` |
| Emp Name | C | `emp_name` |
| DOJ | E | `doj` |
| Status | D | `status` |
| LOP | L | `lop` |
| LWD | F | `lwd` |

---

## Employee Status Priority

When an employee appears in multiple sheets with different statuses:

1. **Resigned** → Highest priority (becomes "Hold")
2. **Hold** → Second priority
3. **Active** → Lowest priority

Final status in consolidated output:
- "Active" = Green badge
- "Hold" = Amber badge (includes Resigned)

---

## Duplicate Detection

When same `emp_id` appears in multiple source files:

| Field | Behavior |
|-------|----------|
| `source` | First file encountered |
| `also_in` | Lists additional sources (comma-separated) |
| `status` | Worst status wins (Hold > Active) |
| `lop` | Filled if available from any source |
| `lwd` | Filled if available from any source |

---

## LOP CSV Export (On-roll vs Off-roll)

| Category | Sources |
|----------|---------|
| **On-roll** | Core, Project, CF, HK |
| **Off-roll** | Retainer, School, College |

CSV Columns:
1. Email/Employee ID
2. LOP Period
3. Payout Period
4. LOP (Days)
5. LOP Reversal (Days)
6. Add/Delete

Only employees with **LOP > 0** are included.

---

## New Joinee Export

Processes all files to find employees where:
- `doj` (Date of Joining) falls within the processing month
- Calculates "Amount" based on working days in month

Output: Formatted Excel with:
- Dark blue header row
- Alternating row colors
- Frozen header
- Auto column widths
- Total count at bottom
