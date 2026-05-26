# 🔐 hash-xlsx

> Classic cryptographic hash functions implemented entirely in **Microsoft Excel formulas** — no VBA, no macros, no plugins.

```
"hello"  →  SHA-512  →  9b71d224bd62f378...bcdec043
"hello"  →  SHA-256  →  2cf24dba5fb0a30e...938b9824
"hello"  →  SHA-1    →  aaf4c61ddcc5e8a2...9aea9434d
"hello"  →  MD5      →  5d41402abc4b2a76...11017c592
```

---

## What is this?

A collection of cryptographic hash functions built using only standard Excel worksheet formulas (`INT`, `MOD`, `MID`, `CODE`, `INDEX`). Type a word, get a hash — everything recalculates automatically. Each file includes step-by-step sheets explaining exactly how the algorithm works.

No VBA. No macros. Just formulas.

---

## Files

English versions use default filenames. Russian versions use the `_RU` suffix.

| File | Algorithm | Year | Hash size | Formulas | Notes |
|------|-----------|------|-----------|----------|-------|
| `MD5.xlsx` | MD5 | 1992 | 128 bit | 947 | Little-endian, 4 rounds × 16 ops |
| `SHA1.xlsx` | SHA-1 | 1995 | 160 bit | 1115 | Big-endian, 4 rounds × 20 ops |
| `SHA256.xlsx` | SHA-256 | 2001 | 256 bit | 2665 | Big-endian, 64 rounds |
| `SHA512.xlsx` | SHA-512 | 2001 | 512 bit | 7577 | 64-bit words via hi/lo pairs, 80 rounds |

Russian translations are available as `MD5_RU.xlsx`, `SHA1_RU.xlsx`, etc.

---

## Requirements

- **Microsoft Excel 2013 or newer**
- **Maximum input: 7 ASCII characters** (single 512/1024-bit block)
- LibreOffice Calc is not fully supported

---

## How to use

1. Open any `.xlsx` file in Excel 2013+
2. Go to the first sheet (e.g. `🔐 SHA-256`)
3. Type your word (up to 7 ASCII characters) in the yellow input cell
4. The hash updates automatically across all sheets

---

## Sheet structure

| Sheet | Contents |
|-------|----------|
| `🔐 Algorithm` | Input field + hash result |
| `📖 How it Works` | Step-by-step algorithm explanation |
| `📦 Constants` | K constants and initial H values |
| `⚙️ Computation` | All compression rounds with intermediate values |

SHA-256 also includes a `🔤 ASCII Encoding` sheet showing character → byte → bits.

---

## How bitwise operations work in Excel

Excel has no 32-bit integer type. All bitwise operations are implemented via pure arithmetic:

| Operation | Formula |
|-----------|---------|
| `AND32(a, b)` | 4-byte decomposition via `INT`/`MOD`, 8 bits each |
| `OR32(a, b)` | `= a + b − AND32(a, b)` |
| `XOR32(a, b)` | `= a + b − 2 × AND32(a, b)` |
| `NOT32(a)` | `= 4294967295 − a` |
| `ROL32(x, n)` | `= MOD(MOD(x, 2^(32−n)) × 2^n + INT(x / 2^(32−n)), 2³²)` |
| `ADD32(a, b)` | `= MOD(a + b, 4294967296)` |

### SHA-512: 64-bit arithmetic via hi/lo pairs

Excel `double` only holds integers exactly up to 2⁵³. Each 64-bit SHA-512 word is stored as two 32-bit cells (`hi` and `lo`):

```
x = hi × 2³² + lo

ADD64:  lo = MOD(a_lo + b_lo, 2³²)
        carry = INT((a_lo + b_lo) / 2³²)
        hi = MOD(a_hi + b_hi + carry, 2³²)

ROTR64(x, n < 32):
        new_hi = MOD(INT(hi / 2ⁿ) + MOD(lo, 2ⁿ) × 2^(32−n), 2³²)
        new_lo = MOD(INT(lo / 2ⁿ) + MOD(hi, 2ⁿ) × 2^(32−n), 2³²)
```

---

## Verified test vectors

| Input | MD5 | SHA-1 | SHA-256 | SHA-512 |
|-------|-----|-------|---------|---------|
| `hello` | `5d41402a...` | `aaf4c61d...` | `2cf24dba...` | `9b71d224...` |
| `abc` | `90015098...` | `a9993e36...` | `ba7816bf...` | `ddaf35a1...` |
| `test` | `098f6bcd...` | `a94a8fe5...` | `9f86d081...` | `ee26b0dd...` |

All results verified against Python `hashlib`.

---

## Limitations

- **7 characters max** — single-block only; longer input would require multiple blocks
- **ASCII only** — `CODE()` returns single-byte values; Unicode not supported
- **Excel 2013+** only; LibreOffice not fully compatible
- **Slow recalculation** — SHA-512 takes 10–30 sec on slower machines (7577 formulas)
- **Educational only** — not for production use

---

## Why?

Because someone had to find out if it's possible. Turns out it is.

The real SHA-256 in C is ~100 lines. In Excel formulas — 2665 cells. Both produce the same hash.

---
---

# 🔐 hash-xlsx (на русском)

> Классические криптографические хэш-функции реализованы полностью на **формулах Microsoft Excel** — без VBA, без макросов, без плагинов.

---

## Что это такое?

Коллекция криптографических хэш-функций, построенных только на стандартных формулах Excel (`INT`, `MOD`, `MID`, `CODE`, `INDEX`). Вводишь слово — получаешь хэш, всё пересчитывается автоматически. Каждый файл содержит листы с пошаговым объяснением алгоритма.

Никакого VBA. Никаких макросов. Только формулы.

---

## Файлы

Английские версии идут без суффикса. Русские версии используют суффикс `_RU`.

| Файл | Алгоритм | Год | Хэш | Формул | Особенности |
|------|----------|-----|-----|--------|-------------|
| `MD5.xlsx` | MD5 | 1992 | 128 бит | 947 | Little-endian, 4 раунда × 16 операций |
| `SHA1.xlsx` | SHA-1 | 1995 | 160 бит | 1115 | Big-endian, 4 раунда × 20 операций |
| `SHA256.xlsx` | SHA-256 | 2001 | 256 бит | 2665 | Big-endian, 64 раунда |
| `SHA512.xlsx` | SHA-512 | 2001 | 512 бит | 7577 | 64-бит через hi/lo пары, 80 раундов |

Русские версии: `MD5_RU.xlsx`, `SHA1_RU.xlsx` и т.д.

---

## Требования

- **Microsoft Excel 2013 или новее**
- **Максимум 7 символов ASCII** на входе
- LibreOffice Calc поддерживается не полностью

---

## Как пользоваться

1. Открыть любой `.xlsx` файл в Excel 2013+
2. Перейти на первый лист (например `🔐 SHA-256`)
3. Ввести слово (до 7 символов ASCII) в жёлтую ячейку
4. Хэш обновится автоматически на всех листах

---

## Структура листов

| Лист | Содержание |
|------|------------|
| `🔐 Алгоритм` | Поле ввода + результат |
| `📖 Как работает` | Пошаговое объяснение алгоритма |
| `📦 Константы` | Константы K и начальные значения H |
| `⚙️ Вычисление` | Все раунды с промежуточными значениями |

SHA-256 также включает лист `🔤 ASCII-кодирование`.

---

## Побитовые операции

| Операция | Формула |
|----------|---------|
| `AND32(a, b)` | Побайтовое разложение через `INT`/`MOD` |
| `OR32(a, b)` | `= a + b − AND32(a, b)` |
| `XOR32(a, b)` | `= a + b − 2 × AND32(a, b)` |
| `NOT32(a)` | `= 4294967295 − a` |
| `ROL32(x, n)` | `= MOD(MOD(x, 2^(32−n)) × 2^n + INT(x / 2^(32−n)), 2³²)` |

SHA-512 хранит 64-битные числа как пару ячеек hi/lo, поскольку `double` в Excel точен только до 2⁵³.

---

## Ограничения

- **Максимум 7 символов** — один блок 512/1024 бит
- **Только ASCII**
- **Excel 2013+**, LibreOffice не поддерживается полностью
- **SHA-512** пересчитывается 10–30 сек из-за 7577 формул
- **Только для обучения**

---

## Зачем?

Потому что кто-то должен был выяснить, возможно ли это. Оказалось — да.

Реальный SHA-256 на C — около 100 строк. На формулах Excel — 2665 ячеек. Оба дают одинаковый хэш.
