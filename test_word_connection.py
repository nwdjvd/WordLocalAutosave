"""
Test script to verify Microsoft Word installation and COM access.
This script checks if Word is properly installed and can be accessed via COM.
"""
import sys
import traceback

print("Testing Word connection...")
print("Python version:", sys.version)

try:
    import win32com.client
    from win32com.client import gencache
    print("✓ win32com.client module imported successfully")
    
    # Try to create a Word application
    print("Attempting to connect to Word...")
    word_app = gencache.EnsureDispatch("Word.Application")
    print(f"✓ Connected to Word version: {word_app.Version}")
    
    # Try to access Word properties
    word_app.Visible = False
    print("✓ Changed Word visibility property")
    
    # Try to create a new document
    print("Creating a test document...")
    doc = word_app.Documents.Add()
    print("✓ Created test document")
    
    # Try to manipulate the document
    selection = word_app.Selection
    selection.TypeText("Word COM connection test")
    print("✓ Typed text in the document")
    
    # Close without saving
    doc.Close(SaveChanges=False)
    print("✓ Closed document")
    
    # Close Word
    word_app.Quit()
    print("✓ Closed Word application")
    
    print("\nAll tests PASSED! Word COM connection is working correctly.")
    print("You can run main.py to start the autosave service.")
    
except ImportError as e:
    print("✗ ERROR: Could not import win32com.client module")
    print(f"  {e}")
    print("\nTry installing the required package:")
    print("  pip install pywin32")
    
except Exception as e:
    print("✗ ERROR: Word connection failed")
    print(f"  {e}")
    print("\nPossible causes:")
    print("  1. Microsoft Word is not installed")
    print("  2. COM access is blocked or requires admin rights")
    print("  3. The version of Word is not compatible with this script")
    
    print("\nDetailed error information:")
    traceback.print_exc()

input("\nPress Enter to exit...") 