# TOPSIS Implementation in Python

## 🔗 Live Application
👉 **Live Demo:** https://topsis-102303372.vercel.app/
👉 **pypi link:** https://pypi.org/project/Topsis-Atharva-102303372/

*(Command-line based Python package published on PyPI)*

---

**Course Name:** UCS654 – Predictive Analytics using Statistics  
**Assignment:** Assignment-1 (TOPSIS)

**Name:** Atharva Pandey  
**Roll Number:** 102303372  

---

## Overview

This package implements the **Technique for Order of Preference by Similarity to Ideal Solution (TOPSIS)**, a widely used Multi-Criteria Decision Making (MCDM) method.
It ranks alternatives based on their distance from an ideal best and an ideal worst solution.

The package is designed to work directly from the command line and processes data provided in CSV format.

---

## Features

* Accepts CSV input files with multiple decision criteria
* Supports user-defined weights for each criterion
* Supports benefit (`+`) and cost (`-`) impacts
* Generates a ranked output file with TOPSIS scores
* Handles invalid inputs with clear error messages

---

## Installation

Install the package using pip:

```bash
pip install Topsis-Atharva-102303372
```

---

## Usage

Run the package from the command line using:

```bash
python -m topsis <InputDataFile> <Weights> <Impacts> <ResultFileName>
```

### Example

```bash
python -m topsis input.csv "1,1,1,2,3" "+,+,-,+,+" output.csv
```

---

## Parameters

* **`<InputDataFile>`**
  Path to the CSV file containing the decision matrix.
  The first column must contain alternative names, and the remaining columns must contain numeric criteria values.

* **`<Weights>`**
  Comma-separated weights corresponding to each criterion.

* **`<Impacts>`**
  Comma-separated impacts for each criterion (`+` for benefit, `-` for cost).

* **`<ResultFileName>`**
  Name of the output CSV file where results will be saved.

---

## Sample Input File

The input CSV file should contain alternatives in the first column, followed by numerical values for each criterion.

| Model | Storage (GB) | Camera (MP) | Price     | Looks |
| ----- | ------------ | ----------- | --------- | ----- |
| A1    | 64           | 48          | 32000     | 4     |
| A2    | 128          | 64          | 45000     | 5     |
| A3    | 64           | 32          | 28000     | 3     |
| A4    | 256          | 108         | 60000     | 5     |
| A5    | 128          | 48          | 40000     | 4     |

---

## Output

The output CSV file contains the original data along with the following additional columns:

* **Topsis Score**
* **Rank**

The alternative with **Rank 1** is considered the **best choice**.

### Example Output

| Model | Storage | Camera | Price | Looks | Topsis Score | Rank |
| ----- | ------- | ------ | ----- | ----- | ------------ | ---- |
| A2    | 128     | 64     | 45000 | 5     | 0.812        | 1    |
| A5    | 128     | 48     | 40000 | 4     | 0.684        | 2    |
| A1    | 64      | 48     | 32000 | 4     | 0.552        | 3    |
| A4    | 256     | 108    | 60000 | 5     | 0.421        | 4    |
| A3    | 64      | 32     | 28000 | 3     | 0.305        | 5    |

---

## Error Handling

The package checks for:

* Missing or invalid input files
* Non-numeric values in criteria columns
* Mismatch between number of criteria, weights, and impacts
* Invalid impact symbols (must be `+` or `-`)

Appropriate error messages are displayed in each case.

---

## Author

**Atharva Pandey**

---

## License

This project is intended for academic and educational use.
