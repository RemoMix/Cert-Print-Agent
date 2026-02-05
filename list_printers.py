#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
سكريبت بسيط لعرض جميع الطابعات المتاحة
Simple script to list all available printers
"""

print("=" * 70)
print("قائمة الطابعات المتاحة / Available Printers")
print("=" * 70)

try:
    import win32print
    
    # 1. الطابعة الافتراضية
    try:
        default_printer = win32print.GetDefaultPrinter()
        print(f"\n✓ الطابعة الافتراضية (Default Printer):")
        print(f"  → {default_printer}")
    except Exception as e:
        print(f"\n✗ لا توجد طابعة افتراضية: {e}")
    
    # 2. كل الطابعات المتاحة
    print(f"\n📋 جميع الطابعات المتاحة (All Available Printers):")
    print("-" * 70)
    
    printers = win32print.EnumPrinters(2)  # 2 = local and network printers
    
    if not printers:
        print("  ✗ لا توجد طابعات متاحة!")
    else:
        for i, printer in enumerate(printers, 1):
            printer_name = printer[2]
            print(f"  {i}. {printer_name}")
    
    # 3. تعليمات الاستخدام
    print("\n" + "=" * 70)
    print("💡 كيفية الاستخدام:")
    print("-" * 70)
    print("انسخ اسم الطابعة بالضبط (مع المسافات وكل حاجة)")
    print("وحطه في ملف config.yaml:")
    print()
    print("printing:")
    if printers:
        example_printer = printers[0][2]
        print(f'  printer_name: "{example_printer}"')
    else:
        print('  printer_name: "اسم الطابعة هنا"')
    print("  retry_attempts: 3")
    print("  retry_delay_seconds: 10")
    print("=" * 70)

except ImportError:
    print("\n✗ مكتبة win32print غير مثبتة!")
    print("\nلتثبيتها:")
    print("  pip install pywin32")
    print("\nبعد التثبيت، شغّل السكريبت مرة تانية.")
    print("=" * 70)

except Exception as e:
    print(f"\n✗ خطأ غير متوقع: {e}")
    import traceback
    print(traceback.format_exc())
    print("=" * 70)

input("\nاضغط Enter للخروج...")
