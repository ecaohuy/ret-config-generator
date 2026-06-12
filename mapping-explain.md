# `mapping.json` Explained

Detailed walkthrough of the JSON config that drives the CDD → RETConfigWDTInternal
conversion. (Referred to casually as "config.json"; the real file is `mapping.json`.)

## Top-level structure

```
mapping.json
├── description        ← human note on what this file does
├── source             ← WHERE input comes from (CDD.xlsx)
├── target             ← WHERE output goes (RETConfigWDTInternal.xlsx)
├── extraction         ← HOW to derive X (band) and Y (sector) from CellName
├── band_map           ← X → band attributes (LOOKUP TABLE)
├── sector_map         ← Y → sector attributes (LOOKUP TABLE)
├── constants          ← fixed values (rru_cn=0, rru_sn=0)
├── field_rules        ← HOW each output column is built
├── column_copy_rules  ← extra rule: output H = output B
├── row_expansion      ← 4 device rows per sector
└── validation         ← formats + what to do on bad data
```

---

## 1. `source` — the input map (CDD.xlsx)

Defines the input sheet and labels each column A–J:

```
CDD.xlsx  →  sheet "IMG_8289 (4G CDD)"
 ┌────┬──────────┬──────────┬─────────┬────────────┬───────┬───────────────┬────────┬────────┬────────┐
 │ A  │    B     │    C     │   D     │     E      │   F   │   G  (KEY)    │   H    │   I    │   J    │
 ├────┼──────────┼──────────┼─────────┼────────────┼───────┼───────────────┼────────┼────────┼────────┤
 │site│ site_new │main_site │ ne_name │ enodeb_name│ ne_id │  CellName     │ E_TILT │m_old   │ m_new  │
 │_old│(location)│          │         │            │       │ [the key]     │(degree)│        │        │
 └────┴──────────┴──────────┴─────────┴────────────┴───────┴───────────────┴────────┴────────┴────────┘
        ▲                                    ▲          ▲                       ▲
        used in names                    not used    used in names         becomes RCU Tilt
```

Key fields the conversion actually uses: **B (site_new)**, **F (ne_id)**, **G (CellName)**, **H (E_TILT)**.

---

## 2. `extraction` — parsing the CellName key

Column G is the key. The last two characters encode everything:

```
   CellName  =  A G G P Q C 0 1 B M 4 C B
                └──────── prefix ───────┘ │ │
                                          X Y
                                          │ └─ Y = RIGHT(CellName,1)         → sector
                                          └─── X = LEFT(RIGHT(CellName,2),1) → band
   Example "AGGPQC01BM4CB":  X = C,  Y = B
```

- **X** = 2nd-to-last char → *which frequency band*
- **Y** = last char → *which sector*

---

## 3. `band_map` — X lookup table

X selects one row of band attributes:

```
 X  → band_token   device_group  slot_suffix  color
 ─────────────────────────────────────────────────────
 C  → L1800            1            _1         Lb1
 D  → L1800 F2         2            _2         CLb2
 E  → NSN_U2100        3            _1         CRb3
 _NOT_USED_ → NOT_USED 4            _2         Rb4   ← synthetic 4th device
```

## 4. `sector_map` — Y lookup table

Y selects the sector id and the RRU SRN number:

```
 Y  → sector_id   rru_srn
 ─────────────────────────
 A  →   S1          60
 B  →   S2          61
 C  →   S3          62
 D  →   S4          63
```

## 5. `constants`

```
 rru_cn = 0     (RRU CN, column C — always 0)
 rru_sn = 0     (RRU SN, column E — always 0)
```

---

## 6. `field_rules` — how each output column is built

This is the heart of the config. Each output column (target A–H) comes from a rule:

```
 OUTPUT col │ field         │ rule
 ───────────┼───────────────┼──────────────────────────────────────────────
   A        │ site_name     │ "{site_new}_{ne_id}"            e.g. AGGPQC01_686001
   B        │ rru_name      │ "{site_new}_{ne_id}_{band_token}_{sector_id}{slot}"
   C        │ rru_cn        │ 0   (constant)
   D        │ rru_srn       │ sector_map[Y].rru_srn           (60–63)
   E        │ rru_sn        │ 0   (constant)
   F        │ rcu_coloring  │ band_map[X].color               (Lb1/CLb2/CRb3/Rb4)
   G        │ rcu_tilt      │ round(E_TILT × 10)              degrees → 0.1° units
   H        │ device_name   │ "{rru_name}"  = same as col B
```

> Note: the JSON's `rru_name` example string shows `_{device_group}` at the end, but the
> running code uses `{slot_suffix}` (`_1`/`_2`). The script is the source of truth — it
> builds `..._S1_1`, `..._S1_2`, etc.

---

## 7. `column_copy_rules` — the added rule

```
   ┌────────────────────────────────────────┐
   │  output column H  =  output column B    │
   │  (Device Name)        (RRU Name)        │
   └────────────────────────────────────────┘
```

This formalizes that **Device Name mirrors RRU Name** — verified true for all 32 rows.

---

## 8. `row_expansion` — 1 sector → 4 rows

Even though CDD only lists bands C/D/E, each sector always emits **4 device rows**
(the 4th is the synthetic `NOT_USED`):

```
   one sector (Y)
        │
        ├── device 1 → band C (L1800)
        ├── device 2 → band D (L1800 F2)
        ├── device 3 → band E (NSN_U2100)
        └── device 4 → _NOT_USED_ (tilt = 0)
```

## 9. `validation` — guardrails

```
 site_format       : ^[A-Z]{2,}[0-9]{2,}$
 cell_name_format  : <site><suffix><X><Y>, X∈{C,D,E}, Y∈{A,B,C,D}
 unknown X         : skip row + log
 unknown Y         : skip row + log
```

---

## End-to-end picture

```
 CDD row:  site_new=AGGPQC01 | ne_id=686001 | CellName=AGGPQC01BM4CB | E_TILT=2.5
                │                    │                 │  X=C  Y=B          │
                ▼                    ▼                 ▼                    ▼
          (col A,B,H name parts)  band_map[C]      sector_map[B]      round(2.5×10)=25
                │                  L1800/Lb1/_1      S2 / SRN 61          │
                └──────────────────┬─────────────────┬──────────────────┘
                                   ▼
   OUTPUT row →  Site Name        | RRU Name                  | CN | SRN | SN | Color | Tilt | Device Name
                 AGGPQC01_686001  | AGGPQC01_686001_L1800_S2_1| 0  | 61  | 0  | Lb1   | 25   | AGGPQC01_686001_L1800_S2_1
                                                                                                └── = RRU Name (H=B rule)
```
