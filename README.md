# Termék kereső / Product Finder

**[Magyar](#magyar) · [English](#english)**

![Termék kereső](docs/kepernyo_5005.png)

---

## Magyar

### A probléma

Egy gyártóüzemben egy termék átállításához négy helyről kell összeszedni az információt:

| Mit | Hol van jellemzően |
|---|---|
| Beállítások (gép, szerszám, magasság, programkód) | egy Excelben, vagy csak a beállító fejében |
| Termékadatlap | hálózati meghajtón, ügyfelenkénti mappákban, régi verziókkal keverve |
| 3D modell | szétszórva, egységes elnevezés nélkül |
| Gépi program | gépenkénti mappákban, kódolt névvel |

Ezek egyike sem hivatkozik a többire, ezért minden átállás keresgéléssel kezdődik.

### A megoldás

Beírsz egy termékszámot, és egy ablakban megkapod:
- a termék összes beállítását az Excelből,
- az ügyfelet, amit a program az adatlap mappájából állapít meg,
- az adatlapot, a 3D modellt és a gépi programot, mindegyiket dupla kattintással megnyithatod.

**A meglévő fájlokon semmit nem kell átrendezni.** A program azt olvassa, ami már ott van.

### A nehéz rész: valódi, rendetlen fájlnevek

A valódi fájlrendszer nem következetes. A párosítási szabályok ezt kezelik, és ahol nem egyértelmű a helyzet, **jeleznek, nem találgatnak**:

| Helyzet | Példa | Viselkedés |
|---|---|---|
| A fájlnév csak *kezdődik* a termékszámmal | `1042` → `1042_A_HU-EN.pdf` | megtalálja |
| Hasonló szám, más termék | `1042` → `10421_B.pdf` | **kizárja** |
| Toldalék a cellában | `2210 X` → `2210_C_HU.pdf` | a termékszám alapján megtalálja |
| Szóköz a cellában, kötőjel a fájlban | `450 120 00` → `450 120-00_A.pdf` | megtalálja |
| Archív régi verzió | `ARCHÍV/1042_regi.pdf` | kihagyja |
| Programkód előtag nélkül | `10423` → `P10423.prg` | pótolja az előtagot |
| Több program egy cellában, egy hiányzik | `50051; 50052; 50055` | a meglévőket listázza, a hiányzót **jelzi** |
| Bizonytalan kód | `51001 ?` | megtalálja, de **jelölve** mutatja |
| A program rossz gépnél van | `6020` (gép 1) → `Gep-2/P60201.prg` | megmutatja, **[másik gép]** jelöléssel |
| Változat más néven, azonos gyártással | `3301` → a `3300` sora | az alaptermék adatait hozza |

### Kipróbálás

```
pip install openpyxl
python mintaadat_general.py      # létrehozza a kitalált mintaadatokat
python kereso.pyw                # Windows alatt: dupla kattintás
```

Ezekre érdemes rákeresni: `1042`, `5005`, `5100`, `6020`, `3301`, `450 120`.

### Saját környezethez igazítás

Minden beállítás a `config.json` fájlban van: az Excel helye és a munkalap neve, a fájlforrások (mappa és kiterjesztés), a gépek és a programmappáik, valamint a kihagyandó mappák. A kódhoz nem kell hozzányúlni.

Az összes adat (termékek, ügyfelek, fájlok) **kitalált**.

---

## English

### The problem

On a shop floor, changing over to a product means collecting information from four places:

| What | Where it usually lives |
|---|---|
| Settings (machine, tool, height, program code) | a spreadsheet, or only in the operator's head |
| Product data sheet | network drive, per-customer folders, mixed with old versions |
| 3D model | scattered, no naming convention |
| Machine program | per-machine folders, coded file names |

None of these point to each other, so every changeover starts with a search.

### The solution

Type a product number and one window shows:
- every setting of that product from the spreadsheet,
- the customer, taken from the data sheet's folder,
- the data sheet, 3D model and machine program, each opened with a double-click.

**Nothing has to be moved or renamed.** The tool reads the file system as it is.

### The hard part: real, messy file names

The matching rules handle inconsistent naming, and where a match is ambiguous they **flag it instead of guessing**:

| Case | Example | Behaviour |
|---|---|---|
| File name only *starts* with the number | `1042` → `1042_A_HU-EN.pdf` | found |
| Similar number, different product | `1042` → `10421_B.pdf` | **excluded** |
| Suffix in the cell | `2210 X` → `2210_C_HU.pdf` | found via the core number |
| Space in cell, hyphen in file | `450 120 00` → `450 120-00_A.pdf` | found |
| Archived old version | `ARCHÍV/1042_regi.pdf` | skipped |
| Program code without prefix | `10423` → `P10423.prg` | prefix added |
| Several codes, one missing | `50051; 50052; 50055` | lists the found ones, **flags** the missing |
| Uncertain code | `51001 ?` | found, **marked** as uncertain |
| Program on the wrong machine | `6020` (machine 1) → `Gep-2/P60201.prg` | shown, marked **[other machine]** |
| Variant with another name, same process | `3301` → row of `3300` | returns the base product |

### Try it

```
pip install openpyxl
python mintaadat_general.py      # creates the fictional sample data
python kereso.pyw                # on Windows: double-click
```

Try searching for: `1042`, `5005`, `5100`, `6020`, `3301`, `450 120`.

### Adapting it

Everything is set in `config.json`: spreadsheet path and sheet, file sources (folder, extensions), machines and their program folders, folders to skip. No code changes needed. The user interface is in Hungarian.

All data (products, customers, files) is **fictional**.

---

![Egy termék két gépen / one product on two machines](docs/kepernyo_1042.png)
