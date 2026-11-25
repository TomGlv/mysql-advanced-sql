# mysql-advanced-sql

> Advanced SQL labs on MySQL 8: schema, joins, GROUP BY/HAVING, CTEs, views, plus Python scripts that run the lab and build the Word report.

## Content
- `tp-sql-avance/` — company database (employees, departments, projects…): `LEFT JOIN`, `GROUP BY` / `HAVING`, sub-queries, **CTEs**, **views**
- `tp2/` — second MySQL lab (`src/script tp2.txt`)
- `run_tp.py` runs every query against a local MySQL server and logs the outputs to `log.json`
- `make_docx.py` turns that log into a formatted Word report (`python-docx`): SQL blocks, terminal output, tables

## Stack
MySQL 8 · SQL · Python · python-docx

---
Made by [Tom Galvan](https://tomglv.dev) · Bachelor in AI, ECE Paris
