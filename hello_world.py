#!/usr/bin/env python3
NameF = input("Tell me your first name ").title()
NameL = input("Tell me your last Name: ").title()
if NameF.strip() and NameL.strip():
    print(f"Hello {NameF} {NameL}!")
elif NameF.strip():
    print(f"Hello {NameF}!")
elif NameL.strip():
    print(f"Hello {NameL}!")
else:
    print("Hello there!")

