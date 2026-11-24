# 📝 Excel VBA: Extract Comments

**[Leer en español](README.es.md)**

Excel has no formula to read a cell's comments. This module adds three ways to get them into the grid, so they can be searched, filtered, counted and exported:

| Tool | What you get | Excel version |
|---|---|---|
| `=ExtraerComentarios(cell)` | All the comments of **one cell**, as text in that cell | Any |
| `=TablaComentarios([range])` | A **table**, one row per comment or reply, that spills by itself | Excel 365 / 2021 (dynamic arrays) |
| `ExportarComentarios` (macro) | The same table for **every sheet of the workbook**, written to a new sheet as an Excel Table | Any |

English aliases do the same thing with English headers: `=ExtractComments()`, `=CommentsTable()` and the `ExportComments` macro.

It handles both kinds of Excel comments: **threaded comments** (the modern ones, with replies) and **notes** (the classic yellow box). When a cell has no comments, the function leaves the cell empty, with no error.

---

## 1. One cell: `=ExtraerComentarios()`

<p align="center">
  <img src="Images/01-single-cell-function.png" alt="ExtraerComentarios next to the data, with and without a custom format" width="95%">
</p>

```
=ExtraerComentarios(D2)
=ExtraerComentarios(D2, "{n}. {author} ({date}): {text}", " | ")
=ExtraerComentarios(A7, "{author}: {text}")
```

| Argument | Required | What it does |
|---|---|---|
| `cell` | yes | The cell whose comments you want. If you pass a range, only the first cell is used. |
| `format` | no | A template for each comment, with `{author}`, `{date}` (`yyyy-mm-dd hh:mm`), `{text}` and `{n}` (1 is the thread, 2 and on are the replies). Without it, the output is the original one: `Author el date: text`. |
| `separator` | no | What goes between one comment and the next. By default, a line break (turn on *Wrap Text* to see it). |

## 2. A table that spills: `=TablaComentarios()`

<p align="center">
  <img src="Images/02-spilling-table.png" alt="TablaComentarios spilling one row per comment" width="95%">
</p>

```
=TablaComentarios()                  every comment on the sheet where the formula is
=TablaComentarios(Ventas!A1:G20)     only the comments inside that range
```

Write it in one cell and Excel fills the rows below by itself, with headers. Leave empty space below and to the right, or you will get `#SPILL!`. Format the `Fecha` column as a date to read it (Excel returns dates from a function as numbers).

## 3. The whole workbook: the `ExportarComentarios` macro

<p align="center">
  <img src="Images/03-macro-comments-sheet.png" alt="The Comentarios sheet created by the macro, as an Excel Table" width="95%">
</p>

Run it from **Developer > Macros** (or `Alt + F8`): choose `ExportarComentarios` and press **Run**. It scans every sheet and writes all the comments in a sheet called **Comentarios** (if it already exists, it is replaced, never duplicated), as an Excel Table called `tblComentarios`, with the dates formatted and the text wrapped. From there, a PivotTable answers questions like *who comments the most*, *open threads by sheet* or *average comments per cell*.

### The table's columns

| Column | Example | Why |
|---|---|---|
| `Hoja` (Sheet) | `Ventas` | Filter by sheet |
| `Columna` (Column) | `D` | Index by column letter |
| `Fila` (Row) | `2` | Index by row number |
| `Celda` (Cell) | `D2` | The key to join back to the data |
| `N comentario` (Comment #) | `1`, `2`, `3` | `1` is the thread, the rest are replies in order |
| `Tipo` (Type) | `Hilo` / `Nota` | Threaded comment or classic note |
| `Autor` (Author) | `Ana` | Who wrote it |
| `Fecha` (Timestamp) | `2026-11-24 20:00` | When. Notes have no date in Excel, so it stays empty |
| `Resuelto` (Resolved) | `TRUE` / `FALSE` | Only on the first comment of a thread. Empty on replies and notes |
| `Texto` (Text) | `Check this price` | The comment itself. For notes, without the `Author:` line Excel adds at the top |

Rows come sorted by sheet, row, column and comment number.

---

## Which one should I use?

| Your situation | Use |
|---|---|
| You want to **read** the comments next to the data, row by row, or copy them into a report | **The function** `=ExtraerComentarios()` |
| You want the comments of **one sheet** as a table inside your model, next to other formulas (`FILTER`, `COUNTIF`, `XLOOKUP`) | **The spilling function** `=TablaComentarios()` (Excel 365 / 2021) |
| You want **every sheet** at once, a **PivotTable**, to **send** the comments to someone, or your Excel has no dynamic arrays | **The macro** `ExportarComentarios` |
| The workbook is **large** (thousands of comments) | **The macro**: it reads everything once and writes it in one go. Many functions dragged down recalculate one by one. |
| You need a **snapshot** that does not change (for example, a month-end audit) | **The macro**: its result is plain values in a table, not formulas. |
| The workbook will be **shared with people who disable macros** | Run the macro, then save a copy as `.xlsx`: the table stays and the code goes away. |

In short: **the function to read, the macro to analyze.**

---

## Installation

**Option A, import the file (recommended):**
1. Download [`VBA-comment-extractor.bas`](VBA-comment-extractor.bas).
2. In Excel, press `Alt + F11` to open the VBA editor.
3. **File > Import File...** and choose the `.bas`. A module called `ExtractorComentarios` appears.
4. Close the editor and use the functions in any cell.

**Option B, copy and paste:** open the `.bas`, copy everything **from the second line on** (the first line, `Attribute VB_Name`, is only for importing), and paste it into a new module (**Insert > Module**).

Save the workbook as **Excel Macro-Enabled Workbook (`.xlsm`)**, or the code is lost.

**Try it first without installing anything:** download [`example/comments-example.xlsm`](example/comments-example.xlsm). It has invented data, threaded comments with replies, notes, the three tools already in use and the test module. Excel will ask you to enable macros.

---

## Refreshing

Excel **does not recalculate a formula when a comment changes**: adding or editing a comment is not a change in any cell. After you change comments, press **`Ctrl + Alt + F9`** (full recalculation) to refresh `ExtraerComentarios` and `TablaComentarios`. The macro is simply run again.

The functions are deliberately **not volatile** (`Application.Volatile`). See *Problems we found* below.

## Limits

- **Notes have no date** in Excel, and their author is only the display name of whoever created them.
- **Threaded comments are created in Microsoft 365.** Older versions show them as notes, and the tools read them as notes.
- **`TablaComentarios` needs dynamic arrays** (Excel 365 or 2021). In older versions, use the macro.
- **The function reads one cell.** If you pass `A1:C10`, it only looks at `A1`.
- **Resolved** exists only for threaded comments, and only on the first comment of the thread.

## Problems we found while building it

- **A volatile function that reads comments can crash Excel when the workbook opens.** With `Application.Volatile`, Excel recalculates the table function while it is still loading the comments from the file, and in our tests Excel closed every time the sample opened. Without `Volatile` it opened every time. That is why the functions are not volatile and this README asks for `Ctrl + Alt + F9`.
- **An Excel Table cannot have the same name as a function.** The first version named the macro's table `TablaComentarios`, like the function, and the workbook stopped opening. The table is now `tblComentarios` (`tblComments` in English).
- **Excel creates a hidden note for every threaded comment.** Reading both collections counts each thread twice. Notes that belong to a thread are skipped.
- **Notes start with `Author:`.** The table removes that line, because the author already has its own column. The single-cell function keeps the original behaviour unless you give it a format.

---

## Tests

[`tests/PruebasExtractor.bas`](tests/PruebasExtractor.bas) has 22 checks that run inside Excel on the sample workbook: the single-cell text, empty cells, the custom format and separator, the order and numbering of the table, *Resolved*, notes without a date, threads not counted twice, a range with no comments, the English headers, and the macro (number of rows, Excel Table, real dates, and that running it twice does not duplicate). Import it next to the main module and run `ProbarExtractor` from the Immediate window (`Ctrl + G`):

```
? ProbarExtractor()
```

Last run: **22 of 22 OK** on Excel for Microsoft 365 (Windows).

---

## Changelog

**v2.0 (November 2026)**
- New: `=TablaComentarios()`, a table that spills, one row per comment or reply.
- New: the `ExportarComentarios` macro, every sheet into an Excel Table.
- New: English aliases `ExtractComments`, `CommentsTable` and `ExportComments`, with English headers.
- New: optional `format` and `separator` for the single-cell function.
- New: a sample workbook and a test module.
- Change: the default separator is a line break (`vbLf`) instead of `vbCrLf`, which Excel shows better inside a cell.

**v1.0**: `=ExtraerComentarios(cell)`.

## Roadmap

- [x] One comment per row, as a macro and as a spilling function.
- [x] English aliases.
- [x] A sample `.xlsm` with test comments.
- [x] A custom format for the single-cell function.
- [ ] An `ImportarComentarios` macro, the way back: edit the table and write the comments into the cells.
- [ ] Filter the macro by sheet or by date.
- [ ] A ready-made PivotTable sheet in the sample.

Suggestions are welcome as [issues](https://github.com/mauroemartinez/VBA-extract-comments-function/issues).

## License

[MIT](LICENSE).
