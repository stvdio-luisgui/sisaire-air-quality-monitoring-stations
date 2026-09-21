# SISAIRE Air Quality Monitoring Stations

Python workflow for organising and analysing Colombian SISAIRE air-quality monitoring data.

## Overview

This project processes large SISAIRE datasets and organises air-quality monitoring records by environmental authority, monitoring station, and measured variable.

The workflow:

- Reads compressed SISAIRE CSV databases.
- Validates and converts `MED_CONCENTRACION_ESTANDAR` to numeric values.
- Identifies missing and non-convertible measurements.
- Groups records by environmental authority (`NOMBRE_FGDA`).
- Identifies monitoring stations (`NOMBRE_EST`) belonging to each authority.
- Summarises records, zero values, and missing values by station.
- Produces summaries by measured variable (`MSFL_CODE`).
- Creates Excel workbooks for each environmental authority.
- Creates separate worksheets for each measured variable.
- Generates consolidated summary databases for subsequent analysis in Python.

## Main SISAIRE fields

- `NOMBRE_FGDA`: Environmental authority.
- `NOMBRE_EST`: Monitoring station.
- `MSFL_CODE`: Measured variable.
- `MED_CONCENTRACION_ESTANDAR`: Standardised measurement value.

## Requirements

Install the required Python packages with:

```bash
pip install -r requirements.txt