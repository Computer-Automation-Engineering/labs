#!/usr/bin/env python3
NameF = input("Tell me your first name ").title()
NameL = input("Tell me your last Name: ").title()
if NameF.strip():
    print(f"Hello {NameF} {NameL}!")
else:
    print("Hello there!")

