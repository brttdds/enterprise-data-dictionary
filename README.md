# Enterprise Data Dictionary

A pipeline for extracting database metadata and generating structured field definitions at scale using SQL and LLM prompt engineering.

Built during a data science internship where no enterprise data dictionary existed. The goal was to create a central reference that data engineers, data scientists, and business stakeholders could all consult.

## What This Does

1. Extracts column-level metadata from a relational database using a standard `INFORMATION_SCHEMA` query
2. Organizes fields into a structured template with placeholder columns for definitions, notes, and join relationships
3. Sends field batches to an LLM using an engineered prompt with a full style guide
4. Returns consistent, reviewable definitions validated through subject matter expert review

## Scale of the Original Work

| Environment | Tables | Fields Defined |
|---|---|---|
| Data Warehouse | 121 | 1,246 |
| Operational Data Store (ODS) | 197 | 3,652 |

## Repository Contents

| File | Description |
|---|---|
| `query.sql` | Metadata extraction query using `INFORMATION_SCHEMA.COLUMNS` |
| `prompt_template.md` | Generalized LLM prompt with style guide and output format rules |
| `demo.py` | End-to-end Python demo running against a local SQLite database |
| `sample_output.xlsx` | Sample output showing the definition format across two fictional tables |

## How to Run the Demo

```
pip install openai pandas openpyxl
```

Run without an API key to use mock definitions:

```bash
python demo.py
```

Run with a live LLM:

```bash
export OPENAI_API_KEY=your_key_here
python demo.py
```

The demo creates a local SQLite database with two fictional healthcare tables, extracts metadata, generates definitions, and writes formatted output to `sample_output.xlsx`.

## Prompt Engineering

The prompt was built around a structured style guide covering surrogate key identification, boolean and enum handling, confidence scoring for human review flagging, and language rules that kept definitions consistent across thousands of fields. See `prompt_template.md` for the full structure.

## Background

This work was done during a data science internship at a healthcare technology company. No data dictionary existed, which created friction for engineers querying unfamiliar tables and scientists building models on undocumented fields. The pipeline here produced the first version of that dictionary. A sanitized version of the methodology is presented in this repo.

## Skills Demonstrated

- SQL metadata extraction from relational databases
- LLM prompt engineering with structured style guides
- Python scripting for data pipeline automation
- Data documentation and governance
