# Bilance: compound batch dosing and production records for calendering lines

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.23186695.svg)](https://doi.org/10.5281/zenodo.23186695)

*Dosaggio delle mescole e registrazione della produzione per le linee di calandratura*

2002–2003 · version 7.4.7  
Author: **Massimo Sbarbaro** ([ORCID 0009-0006-8965-9013](https://orcid.org/0009-0006-8965-9013))

## Overview

Bilance (“scales”) records the preparation of the compound batches (*mescole*) that feed the calenders of a plastic-film plant. For each batch the operator enters the production order, the compound, the raw materials with their lots and weighed kilograms, the multiplier and the number of cycles; the program checks the quantities against the order and closes the production. Production data are then sent back to the MFG/PRO ERP as CIM batch-load files that create work orders and record material issues and receipts.

I designed and programmed this application in 2002–2003. It is published here in 2026 as part of the documented record of my software work, with a permanent DOI on Zenodo.

## Main functions

- Batch registration per batch ID with raw materials, lots and weights (`frmBilance`, `frmID`).
- List of batches to produce, shift and operator management (`frmSE`, `frmTurno`).
- Weekly planning of orders by calender and moving of orders between lines (`frmImposta`, `frmSposta`).
- Download of new orders for a date range from the ERP (`frmNewOrd`, `frmImportMFG`).
- Generation of MFG/PRO CIM files (`wowomt`, `wowoisrc`) to create work orders and post issues and receipts.
- Password-protected editing of master data (`frmPass`).

## Data

Microsoft Access database through DAO 3.6. Main tables: `Formule`, `SE`, `MP`, `DescrizMesc`, `Turno`, `Operaio`, `Controtipo`, `Parametri`, plus staging tables for ERP imports.

## Technology

DAO 3.6, Microsoft Access object library, Crystal Reports 8.5 (RDC, viewer, export), Winsock, Common Controls, Tab control, Masked Edit, DBGrid.

## Repository contents

| Path | Content |
|---|---|
| `src/` | Project file (`.vbp`), forms (`.frm` with their binary resources `.frx`) and modules (`.bas`), as listed in the project file. |
| `config-example/` | Templates of the `.ini` configuration files read at start-up, with placeholder values. |

## What is not included

Crystal Reports layouts (`.rpt`), compiled executables, installers, scripts for the host connection and the production databases are **not** included, because they contain operational data of the organisations for which the program was built. Configuration files are published as templates in `config-example/` with placeholder values: server addresses, accounts and passwords have been removed. Names of client organisations and products have been removed from captions and comments; local paths have been normalised.

## Related repositories

- [calender-production-orders](https://github.com/massimosbarbaro/calender-production-orders)
- [compound-cost-analysis](https://github.com/massimosbarbaro/compound-cost-analysis)

## How to cite

Use the citation metadata in [`CITATION.cff`](CITATION.cff) (GitHub: *Cite this repository*). The release is archived on Zenodo with the DOI [10.5281/zenodo.23186695](https://doi.org/10.5281/zenodo.23186695).

> Sbarbaro, Massimo. 2003. *Bilance: compound batch dosing and production records for calendering lines*. Software (2002–2003), version 7.4.7. Zenodo. https://doi.org/10.5281/zenodo.23186695.

## License

Released under the [MIT License](LICENSE). © 2002 Massimo Sbarbaro.
