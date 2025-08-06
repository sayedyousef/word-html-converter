
def main():
    import mammoth_converter
    # test_imports.py
    print("1. Starting...")

    print("2. Testing config...")
    from config import Config
    print("   ✓ Config OK")

    print("3. Testing logger...")
    from logger import setup_logging
    print("   ✓ Logger OK")

    print("4. Testing css_manager...")
    from css_manager import CSSManager
    print("   ✓ CSSManager OK")

    print("5. Testing mammoth_converter...")
    from mammoth_converter import MammothConverter
    print("   ✓ MammothConverter OK")

    print("\nAll imports work! main3.py should run.")

    print("Starting Enhanced Word to HTML Converter...")
if __name__ == "__main__":
    print("Starting Enhanced Word to HTML Converter...")
    main()
