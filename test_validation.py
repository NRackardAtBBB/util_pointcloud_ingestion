#!/usr/bin/env python3
"""
Test script for file validation functions.
"""

from pathlib import Path
from ingest import is_valid_filename, is_supported_extension

# Test cases for file validation
test_cases = [
    # Valid filenames
    ("2586_251231_Floor 1_Wing P.las", True, "Valid with scope"),
    ("2635_240502_Exterior.laz", True, "Valid exterior scan"),
    ("3020_250815_Floor 3_Great Hall.rcp", True, "Valid with multi-word scope"),
    ("2586_251231_Basement.xyz", True, "Valid basement scan"),
    ("1234_260108_Floor 10.pcd", True, "Valid double-digit floor"),
    
    # Invalid filenames
    ("survey_data.las", False, "Missing project and date"),
    ("2586_20251231_Floor 1.las", False, "Wrong date format (8 digits)"),
    ("2586_251231_floor 1.las", False, "Lowercase floor"),
    ("12345_251231_Floor 1.las", False, "5-digit project number"),
    ("2586-251231-Floor 1.las", False, "Using hyphens instead of underscores"),
    ("RENAME_survey_data.las", False, "Already flagged"),
    ("UNSUPPORTED_file.txt", False, "Already flagged as unsupported"),
    
    # Edge cases
    ("2586_251231_Floor 1_Wing P_SubArea.las", True, "Multiple underscores in scope"),
    ("2586_251231_Floor 1_.las", False, "Trailing underscore with empty scope (invalid)"),
]

# Test supported extensions
extension_tests = [
    ("file.las", True),
    ("file.laz", True),
    ("file.pcd", True),
    ("file.ply", True),
    ("file.xyz", True),
    ("file.rcp", True),
    ("file.rcs", True),
    ("file.LAS", True),  # Case insensitive
    ("file.txt", False),
    ("file.pdf", False),
    ("file.docx", False),
]

# Pattern for validation
pattern = r'^(?P<project>\d{4})_(?P<date>\d{6})_(?P<floor>(?:Floor\s*\d+|Exterior|Basement))(?:_(?P<scope>.+))?$'
supported_exts = ['las', 'laz', 'pcd', 'ply', 'xyz', 'rcp', 'rcs']

print("="*70)
print("Testing File Naming Validation")
print("="*70)

passed = 0
failed = 0

for filename, expected, description in test_cases:
    result = is_valid_filename(filename, pattern)
    status = "✓ PASS" if result == expected else "✗ FAIL"
    
    if result == expected:
        passed += 1
    else:
        failed += 1
    
    print(f"{status}: {filename}")
    print(f"       Expected: {expected}, Got: {result} - {description}")
    if result != expected:
        print(f"       ❌ MISMATCH")
    print()

print("="*70)
print("Testing Extension Support")
print("="*70)

ext_passed = 0
ext_failed = 0

for filename, expected in extension_tests:
    result = is_supported_extension(filename, supported_exts)
    status = "✓ PASS" if result == expected else "✗ FAIL"
    
    if result == expected:
        ext_passed += 1
    else:
        ext_failed += 1
    
    print(f"{status}: {filename} - Expected: {expected}, Got: {result}")
    if result != expected:
        print(f"       ❌ MISMATCH")

print("\n" + "="*70)
print("Test Summary")
print("="*70)
print(f"Filename Validation: {passed} passed, {failed} failed")
print(f"Extension Validation: {ext_passed} passed, {ext_failed} failed")
print(f"Total: {passed + ext_passed} passed, {failed + ext_failed} failed")
print("="*70)

if failed + ext_failed == 0:
    print("✓ All tests passed!")
else:
    print("✗ Some tests failed. Please review.")
