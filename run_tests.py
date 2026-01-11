#!/usr/bin/env python
"""
Test Runner for PowerPoint Generator

Runs all tests and provides a summary.
"""

import sys
from test_comic_analysis import run_tests


def main():
    """Run all tests and display results."""
    print("=" * 70)
    print("PowerPoint Generator - Test Suite")
    print("=" * 70)
    print()
    
    result = run_tests()
    
    print()
    print("=" * 70)
    print("Test Summary")
    print("=" * 70)
    print(f"Tests run: {result.testsRun}")
    print(f"Successes: {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"Failures: {len(result.failures)}")
    print(f"Errors: {len(result.errors)}")
    
    if result.wasSuccessful():
        print()
        print("✅ All tests passed!")
        return 0
    else:
        print()
        print("❌ Some tests failed!")
        return 1


if __name__ == '__main__':
    sys.exit(main())
