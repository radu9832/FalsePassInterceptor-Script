# FalsePassInterceptor.vbs

This is a VBScript-based validation tool I wrote to intercept a critical bug in a production test environment. The issue caused the system to report a **false “PASS”** — even though no actual tests were performed.

Instead of escalating the issue upstream, I created this script as a quiet fix, integrating it directly into the test logic without disrupting the main flow.

---

## 💡 What It Solves

At one point, a broken configuration file bypassed the testing process entirely, yet still reported successful results and printed "OK" labels. The system was unaware of the failure.

This script checks whether the items declared by the configuration are actually tested — and blocks the process if anything's missing or malformed.

---

## ⚙️ Key Features

- ✅ **Comparison logic**: Compares two lists — declared items vs. actually tested items  
- ✅ **Input validation**: Flags unexpected or invalid characters to avoid corrupting results  
- ✅ **Duplicate filtering**: Automatically removes repeated entries for cleaner logic  
- ✅ **Safe exit**: Gracefully handles empty input or malformed data  
- ✅ **Mock COM-style variable storage**: Replaces proprietary variable slots using `Scripting.Dictionary`

---

## 🧪 Sample Use

```vbscript
SetVariable "receivedItems", "ITEM001,ITEM002,ITEM003"
SetVariable "testedItems",   "ITEM001,ITEM002"
```

This results in:
```
ScriptStatus: 0
Unmatched Items: ITEM003
```

If everything matches:
```
ScriptStatus: 1
```

---

## 📁 Structure Notes

- `main` is the entry point, simulating how the test system would trigger this
- `filterAndCorrect` ensures characters are safe (`[A-Z0-9_;,-/]`)
- `duplicationRemover` uses `Scripting.Dictionary` to remove repeated values
- `sorting` compares expected vs tested items and flags mismatches

---

## 👨‍🔧 Why It Matters

This wasn't a “nice to have” — it was critical. The script caught real false positives before hardware could ship out untested. It saved time, avoided escalation, and worked without any need to modify upstream applications.
