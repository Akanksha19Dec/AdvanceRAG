"""
Data Ingestion — Load & Chunk Test Cases from Excel
=====================================================
Each test case = 1 natural chunk (structured chunking).
Converts each row into a rich text representation for embedding.
"""

import pandas as pd


def load_test_cases_from_file(path: str) -> list[dict]:
    """Load test cases from an Excel file path and return as a list of dicts with text chunks."""
    print(f"   -> Reading {path}...")
    df = pd.read_excel(path, sheet_name="Test Cases", engine="openpyxl")

    cases = []
    for _, row in df.iterrows():
        # Build a rich text representation — each test case becomes one chunk
        text = (
            f"Test Case: {row.get('Test Case ID', '')}\n"
            f"Module: {row.get('Module', '')}\n"
            f"Category: {row.get('Category', '')}\n"
            f"Description: {row.get('Description', '')}\n"
            f"Pre-conditions: {row.get('Pre-conditions', '')}\n"
            f"Steps: {row.get('Steps', '')}\n"
            f"Expected Result: {row.get('Expected Result', '')}\n"
            f"Priority: {row.get('Priority', '')}"
        )

        cases.append({
            "id": len(cases),
            "tc_id": str(row.get("Test Case ID", "")),
            "module": str(row.get("Module", "")),
            "category": str(row.get("Category", "")),
            "description": str(row.get("Description", "")),
            "preconditions": str(row.get("Pre-conditions", "")),
            "steps": str(row.get("Steps", "")),
            "expected": str(row.get("Expected Result", "")),
            "priority": str(row.get("Priority", "")),
            "text": text,
            "length": len(text),
        })

    print(f"   -> {len(cases)} test cases loaded")
    return cases


def compute_statistics(cases: list[dict]) -> dict:
    """Compute module, category, and priority distributions."""
    module_stats: dict[str, int] = {}
    cat_stats: dict[str, int] = {}
    pri_stats: dict[str, int] = {}

    for tc in cases:
        module_stats[tc["module"]] = module_stats.get(tc["module"], 0) + 1
        cat_stats[tc["category"]] = cat_stats.get(tc["category"], 0) + 1
        pri_stats[tc["priority"]] = pri_stats.get(tc["priority"], 0) + 1

    return {
        "modules": module_stats,
        "categories": cat_stats,
        "priorities": pri_stats,
    }
